from reportman import *
from dbfread2 import DBF
import pandas as pd


def read_dbf(
    filename: str,
    column_names: list[str],
    group_by: str,
    agg_dict: dict,
    alias: list[str],
):
    table = DBF(filepath=filename)
    df = pd.DataFrame(iter(table))
    dbf = (
        df.query("RE == 1")
        .filter(items=column_names)
        .groupby(group_by)
        .agg(agg_dict)
        .reset_index()
    )
    dbf.columns = alias

    return dbf


def read_dwbm(filename: str):
    dwbm = pd.read_excel(filename)
    return dwbm


def build_data(dwbm, dbf, on_field, with_how) -> list[list]:
    result = pd.merge(dwbm, dbf, on=on_field, how=with_how)

    return result


def main():
    dbf = read_dbf(
        filename=r"E:\企业补贴\数据\企业补贴202505\bt_ltx.dbf",
        column_names=["DWBM", "实发补贴"],
        group_by="DWBM",
        agg_dict={"实发补贴": ["count", "sum"]},
        alias=["单位编码", "实发人数", "企业补贴发放额"],
    )
    dwbm = read_dwbm(r"E:\dwbm.xlsx")
    data = build_data(dwbm, dbf, "单位编码", "left").values.tolist()
    wb = load_workbook(filename=r"E:\老人补贴202505.xlsx")
    ws = wb["单位汇总202505"]
    ws_1 = wb["银行表202505"]
    dbf_1 = read_dbf(
        filename=r"E:\企业补贴\数据\企业补贴202505\bt_ltx.dbf",
        column_names=["发放银行", "X_银行帐号", "补发_其它", "其它扣款", "实发补贴"],
        group_by="发放银行",
        agg_dict={
            "X_银行帐号": "count",
            "补发_其它": "sum",
            "其它扣款": "sum",
            "实发补贴": "sum",
        },
        alias=["发放银行", "X_银行帐号", "补发_其它", "其它扣款", "企业补贴发放额"],
    )
    bank_data = dbf_1.values.tolist()
    fill(ws_1, "G4", bank_data)
    fill(ws, "B4", data)
    wb.save(r"E:\老人补贴202505_new.xlsx")

    dbf2 = read_dbf(
        filename=r"E:\企业补贴\数据\集体工企业补贴202505\bt_ltx.dbf",
        column_names=["DWBM", "X_银行帐号", "补发_其它", "实发补贴"],
        group_by="DWBM",
        agg_dict={
            "X_银行帐号": "count",
            "补发_其它": "sum",
            "实发补贴": "sum",
        },
        alias=["单位编码", "实发人数", "提高待遇", "实发金额"],
    )
    wb2 = load_workbook(
        filename=r"E:\企业补贴\银行报盘\2025年\5月\退休集体工企业补贴发放总信息202505.xlsx"
    )
    ws2 = wb2["单位汇总202505工会（企业补贴）"]
    ws_2 = wb2["单位汇总202505工会（企业补贴） 含遗孀"]
    dwbm2 = read_dwbm(filename=r"E:\dwbm2.xlsx")
    data2 = build_data(dwbm=dwbm2, dbf=dbf2, on_field="单位编码", with_how="left")
    data2["企业补贴"] = data2["实发金额"] - data2["提高待遇"]
    data2f = data2.filter(
        items=["单位名称", "实发人数", "企业补贴", "提高待遇", "实发金额", "性质"]
    )
    fqm = data2f[data2f["性质"] != "工程"]
    data2f_ = data2.filter(
        items=["单位名称", "实发人数", "提高待遇", "实发金额", "性质"]
    )
    gc = data2f_[data2f_["性质"] == "工程"]
    fqm_data = fqm.values.tolist()
    gc_data = gc.values.tolist()
    ws3 = wb2["单位征集(企业补贴)202505"]
    fill(ws2, "B4", fqm_data)
    fill(ws_2, "B4", fqm_data)
    fill(ws3, "J4", gc_data)
    ws_2 = wb2["银行汇总202505(企业补贴) "]
    dbf_2 = read_dbf(
        filename=r"E:\企业补贴\数据\集体工企业补贴202505\bt_ltx.dbf",
        column_names=["发放银行", "X_银行帐号", "补发_其它", "实发补贴"],
        group_by="发放银行",
        agg_dict={
            "X_银行帐号": "count",
            "补发_其它": "sum",
            "实发补贴": "sum",
        },
        alias=["发放银行", "实发人数", "提高待遇", "实发金额"],
    )
    bank_data2 = dbf_2.values.tolist()
    fill(ws_2, "H4", bank_data2)
    wb2.save(r"E:\退休集体工企业补贴202505.xlsx")


if __name__ == "__main__":
    main()
