"""シフト自動作成デモ(Streamlit)  起動: streamlit run app.py"""
from __future__ import annotations

import time
from io import BytesIO

import pandas as pd
import streamlit as st

from io_excel import (
    SHIFT_COLORS, WEEKDAY, TemplateError, build_shift_table, make_template,
    read_template, result_summary, sample_data, write_result,
)
from solver import Params, solve

st.set_page_config(page_title="シフト自動作成デモ", page_icon="🗓️", layout="wide")

st.title("🗓️ シフト自動作成")
st.caption("希望シフトのExcelをアップロードして、ボタン1つで条件を満たすシフト表を作成します。")

# ---------------- サイドバー ----------------
with st.sidebar:
    st.header("条件")
    max_consecutive = st.slider("連勤の上限(日)", 3, 7, 5)
    forbid_l2e = st.checkbox("遅番の翌日は早番にしない", value=True)
    st.subheader("重視度")
    w_pref = st.slider("△の日を避ける", 0, 30, 10)
    w_fair = st.slider("勤務日数を公平に", 0, 20, 5)
    w_leader = st.slider("全シフトにリーダーを配置", 0, 50, 20)
    time_limit = st.slider("計算時間の上限(秒)", 5, 60, 10)
    st.divider()
    manual_min = st.number_input("手作業でかかっている時間(分)※比較表示用の入力値", 10, 1000, 180, step=10)

# ---------------- 入力 ----------------
c1, c2, c3 = st.columns([1, 2, 1])
with c1:
    buf = BytesIO()
    make_template(buf)
    st.download_button("① テンプレートをダウンロード", buf.getvalue(), "shift_template.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
with c2:
    uploaded = st.file_uploader("② 記入したExcelをアップロード", type=["xlsx"], label_visibility="visible")
with c3:
    st.write("")
    st.write("")
    if st.button("サンプルデータで試す"):
        st.session_state["use_sample"] = True
        st.session_state.pop("res", None)

if uploaded is not None:
    st.session_state["use_sample"] = False

data = None
source = ""
try:
    if uploaded is not None:
        data = read_template(uploaded)
        source = uploaded.name
    elif st.session_state.get("use_sample"):
        data = sample_data()
        source = "サンプル(飲食店 10人×30日・早番/遅番)"
except TemplateError as e:
    st.error(str(e))
    st.stop()

if data is None:
    st.info("テンプレートに記入してアップロードするか、「サンプルデータで試す」を押してください。")
    st.stop()

staff, days, shifts, required, prefs = data
need_total = sum(required.values())
st.success(f"読み込み: {source}  /  スタッフ {len(staff)}人・{len(days)}日・必要延べ人数 {need_total}")

# ---------------- 求解 ----------------
if st.button("③ シフトを作成", type="primary"):
    params = Params(max_consecutive=max_consecutive, forbid_late_to_early=forbid_l2e,
                    w_pref=w_pref, w_fair=w_fair, w_leader=w_leader, time_limit=float(time_limit))
    t0 = time.perf_counter()
    with st.spinner("最適なシフトを計算中..."):
        res = solve(staff, days, shifts, required, prefs, params)
    st.session_state["res"] = {"res": res, "sec": time.perf_counter() - t0,
                               "key": (source, len(staff), len(days))}

saved = st.session_state.get("res")
if not saved or saved["key"] != (source, len(staff), len(days)):
    st.stop()

res, sec = saved["res"], saved["sec"]

if res.status in ("INFEASIBLE", "UNKNOWN"):
    st.error("条件を満たすシフトが作れませんでした。" if res.status == "INFEASIBLE"
             else "制限時間内に解が見つかりませんでした。")
    for m in res.diagnosis:
        st.write("・" + m)
    st.stop()

# ---------------- 結果 ----------------
info = result_summary(staff, days, shifts, prefs, res.assign)
m1, m2, m3, m4 = st.columns(4)
m1.metric("希望反映率", f"{info['pref_rate']:.0%}",
          help=f"休み希望(×){info['n_off']}件はすべて反映。△{info['n_tri']}件のうち{info['tri_assigned']}件は出勤になりました。")
m2.metric("計算時間", f"{sec:.1f}秒")
m3.metric("手作業(入力値)", f"{manual_min}分")
m4.metric("リーダー不在シフト", f"{res.penalties.get('リーダー不在シフト数', 0)}")
if res.status == "FEASIBLE":
    st.caption("※制限時間内で見つかった最良の解です(最適性は未確認)。計算時間を延ばすと改善する場合があります。")

table = build_shift_table(staff, days, shifts, res.assign)
cols = [f"{d.month}/{d.day}({WEEKDAY[d.weekday()]})" for d in days]
df = pd.DataFrame(table, index=[s["name"] + ("★" if s["is_leader"] else "") for s in staff], columns=cols)
color_of = {sh: SHIFT_COLORS[i % len(SHIFT_COLORS)] for i, sh in enumerate(shifts)}


def paint(v: str) -> str:
    return f"background-color: #{color_of[v]}; color: #222; text-align: center" if v in color_of else ""


st.subheader("シフト表")
st.dataframe(df.style.map(paint), width="stretch")
st.caption("★=リーダー可")

left, right = st.columns(2)
with left:
    st.subheader("スタッフ別 勤務日数")
    st.bar_chart(pd.Series(info["work_days"], name="勤務日数"))
with right:
    st.subheader("評価内訳")
    st.table(pd.DataFrame({"項目": list(res.penalties.keys()), "値": list(res.penalties.values())}).set_index("項目"))

st.download_button("④ Excelでダウンロード", write_result(staff, days, shifts, prefs, res.assign),
                   "shift_result.xlsx",
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
