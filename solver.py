"""シフト最適化ソルバー(OR-Tools CP-SAT)

データ構造
- staff: list[dict]  name, max_days(週上限), is_leader(bool)
- days: list[date]
- shifts: list[str]  例 ["早番", "遅番"]  ※最後のシフトを「遅番」、最初を「早番」とみなす
- required: dict[(day_index, shift_index)] -> 必要人数
- prefs: dict[(staff_index, day_index)] -> "◯" | "△" | "×"
"""
from __future__ import annotations

from dataclasses import dataclass, field
from ortools.sat.python import cp_model


@dataclass
class Params:
    max_consecutive: int = 5          # 連勤上限
    forbid_late_to_early: bool = True # 遅番の翌日は早番にしない
    w_pref: int = 10                  # △に入れたペナルティ
    w_fair: int = 5                   # 勤務日数の偏りペナルティ
    w_leader: int = 20                # リーダー不在シフトのペナルティ
    time_limit: float = 10.0
    weekly_cap: bool = True           # 週の上限日数を守る


@dataclass
class Result:
    status: str                       # OPTIMAL / FEASIBLE / INFEASIBLE / UNKNOWN
    assign: dict = field(default_factory=dict)  # (s, d) -> shift index
    penalties: dict = field(default_factory=dict)
    diagnosis: list = field(default_factory=list)


def _build(staff, days, shifts, required, prefs, p: Params, soft_relax: set | None = None):
    """モデルを構築。soft_relax に制約名を入れると、その制約を外す(原因診断用)。"""
    soft_relax = soft_relax or set()
    S, D, T = len(staff), len(days), len(shifts)
    m = cp_model.CpModel()
    x = {(s, d, t): m.NewBoolVar(f"x_{s}_{d}_{t}")
         for s in range(S) for d in range(D) for t in range(T)}

    # 1日1シフトまで
    for s in range(S):
        for d in range(D):
            m.Add(sum(x[s, d, t] for t in range(T)) <= 1)

    # 必要人数
    if "required" not in soft_relax:
        for d in range(D):
            for t in range(T):
                m.Add(sum(x[s, d, t] for s in range(S)) == required.get((d, t), 0))

    # ×の日は入れない
    if "off" not in soft_relax:
        for s in range(S):
            for d in range(D):
                if prefs.get((s, d)) == "×":
                    for t in range(T):
                        m.Add(x[s, d, t] == 0)

    # 連勤上限
    if "consecutive" not in soft_relax:
        k = p.max_consecutive
        for s in range(S):
            for d0 in range(D - k):
                m.Add(sum(x[s, d, t] for d in range(d0, d0 + k + 1) for t in range(T)) <= k)

    # 遅番→早番禁止
    if p.forbid_late_to_early and T >= 2 and "interval" not in soft_relax:
        for s in range(S):
            for d in range(D - 1):
                m.Add(x[s, d, T - 1] + x[s, d + 1, 0] <= 1)

    # 週の上限日数(月曜始まりの暦週。日付から判定)
    if p.weekly_cap and "weekly" not in soft_relax:
        weeks: dict = {}
        for d, day in enumerate(days):
            key = day.isocalendar()[:2]
            weeks.setdefault(key, []).append(d)
        for s in range(S):
            cap = staff[s]["max_days"]
            for idxs in weeks.values():
                m.Add(sum(x[s, d, t] for d in idxs for t in range(T)) <= cap)

    # ---- ソフト制約 ----
    pref_terms = []
    for s in range(S):
        for d in range(D):
            if prefs.get((s, d)) == "△":
                pref_terms.append(sum(x[s, d, t] for t in range(T)))
    pref_pen = sum(pref_terms) if pref_terms else 0

    work = [sum(x[s, d, t] for d in range(D) for t in range(T)) for s in range(S)]
    mx = m.NewIntVar(0, D, "mx")
    mn = m.NewIntVar(0, D, "mn")
    for s in range(S):
        m.Add(mx >= work[s])
        m.Add(mn <= work[s])
    fair_pen = mx - mn

    leader_pen_terms = []
    leaders = [s for s in range(S) if staff[s]["is_leader"]]
    for d in range(D):
        for t in range(T):
            if required.get((d, t), 0) > 0:
                has = m.NewBoolVar(f"lead_{d}_{t}")
                m.Add(sum(x[s, d, t] for s in leaders) >= 1).OnlyEnforceIf(has)
                m.Add(sum(x[s, d, t] for s in leaders) == 0).OnlyEnforceIf(has.Not())
                leader_pen_terms.append(has.Not())
    leader_pen = sum(leader_pen_terms) if leader_pen_terms else 0

    m.Minimize(p.w_pref * pref_pen + p.w_fair * fair_pen + p.w_leader * leader_pen)
    return m, x, (pref_pen, fair_pen, leader_pen)


def solve(staff, days, shifts, required, prefs, p: Params | None = None) -> Result:
    p = p or Params()
    S, D, T = len(staff), len(days), len(shifts)
    m, x, (pref_pen, fair_pen, leader_pen) = _build(staff, days, shifts, required, prefs, p)
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = p.time_limit
    solver.parameters.num_workers = 8
    st = solver.Solve(m)

    if st in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        assign = {}
        for s in range(S):
            for d in range(D):
                for t in range(T):
                    if solver.Value(x[s, d, t]):
                        assign[(s, d)] = t
        pen = {
            "△に入れた回数": int(solver.Value(pref_pen)) if not isinstance(pref_pen, int) else 0,
            "勤務日数の差(最大-最小)": int(solver.Value(fair_pen)),
            "リーダー不在シフト数": int(solver.Value(leader_pen)) if not isinstance(leader_pen, int) else 0,
        }
        return Result("OPTIMAL" if st == cp_model.OPTIMAL else "FEASIBLE", assign, pen)

    if st == cp_model.INFEASIBLE:
        return Result("INFEASIBLE", diagnosis=diagnose(staff, days, shifts, required, prefs, p))
    return Result("UNKNOWN", diagnosis=["制限時間内に解が見つかりませんでした。時間を延ばすか条件を緩めてください。"])


_LABELS = {
    "off": "「×(休み希望)」をすべて守る条件",
    "consecutive": "連勤上限",
    "interval": "遅番の翌日は早番にしない条件",
    "weekly": "スタッフの週あたり上限日数",
}


def diagnose(staff, days, shifts, required, prefs, p: Params) -> list[str]:
    """解が無いとき、日本語で原因の候補を返す。"""
    msgs: list[str] = []
    S, D, T = len(staff), len(days), len(shifts)

    # 1) 単純な人数不足(その日の必要人数 > ×でない人数)
    for d in range(D):
        avail = sum(1 for s in range(S) if prefs.get((s, d)) != "×")
        need = sum(required.get((d, t), 0) for t in range(T))
        if need > avail:
            msgs.append(f"{days[d]:%m/%d}:必要人数 {need}人に対し、出勤可能(×でない)な人が {avail}人しかいません。")

    # 2) 制約を1つずつ外して解けるか確認
    base = {"required"}  # 必要人数は外さない(目的そのもの)
    for key in ["off", "consecutive", "interval", "weekly"]:
        if key == "interval" and not p.forbid_late_to_early:
            continue
        if key == "weekly" and not p.weekly_cap:
            continue
        m, _, _ = _build(staff, days, shifts, required, prefs, p, soft_relax={key})
        sv = cp_model.CpSolver()
        sv.parameters.max_time_in_seconds = 3
        sv.parameters.num_workers = 8
        if sv.Solve(m) in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            msgs.append(f"{_LABELS[key]}を緩めると解が見つかります。これが主な原因の可能性が高いです。")

    # 3) 総量チェック
    total_need = sum(required.get(k, 0) for k in [(d, t) for d in range(D) for t in range(T)])
    n_weeks = len({day.isocalendar()[:2] for day in days})
    cap_total = sum(min(st["max_days"] * n_weeks, D) for st in staff)
    if total_need > cap_total:
        msgs.append(f"必要延べ人数 {total_need} が、全員の上限日数の合計 {cap_total} を超えています。人員または上限日数を増やしてください。")

    if not msgs:
        msgs.append("複数の条件が重なって解がありません。必要人数・休み希望・連勤上限のいずれかを見直してください。")
    return msgs
