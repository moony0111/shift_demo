"""サンプルExcelを生成する: python make_sample.py
sample.xlsx            ... 解けるデータ
sample_infeasible.xlsx ... わざと解けないデータ(原因表示のデモ用)
"""
from io_excel import make_template, sample_data

if __name__ == "__main__":
    make_template("sample.xlsx", *sample_data())
    make_template("sample_infeasible.xlsx", *sample_data(infeasible=True))
    print("sample.xlsx / sample_infeasible.xlsx を作成しました")
