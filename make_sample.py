"""sample.xlsx を生成する: python make_sample.py"""
from io_excel import make_template, sample_data

if __name__ == "__main__":
    staff, days, shifts, required, prefs = sample_data()
    make_template("sample.xlsx", staff, days, shifts, required, prefs)
    print("sample.xlsx を作成しました")
