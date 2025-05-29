# tmp.py

import xlwings as xw
import pandas as pd
from dbfread2 import DBF


# Paths 文件路径"
path_a = r"E:\企业补贴\数据\企业补贴202506"
path_b = r"E:\企业补贴\数据\集体工企业补贴202506"
path_c = r"E:\企业补贴\数据\提高待遇202507"
dbf_path_a = path_a + r"\bt_ltx.dbf"
dbf_path_b = path_b + r"\bt_ltx.dbf"
dbf_path_c = path_c + r"\bt_ltx.dbf"
temp_path_icbc = r"E:\企业补贴\银行报盘\工商银行报盘模板.xlsx"
temp_path_cbc = r"E:\企业补贴\银行报盘\建设银行报盘模板.xlsx"
temp_path_boc = r"E:\企业补贴\银行报盘\中国银行报盘模板.xlsx"
temp_path_bocny = r"E:\企业补贴\银行报盘\中国银行南阳报盘模板.xlsx"


def export_datagrid(
    dbf_path: str,
    template: str,
    sheet: str,
    arch: str,
    add_columns: dict,
    condition: str,
    columns: list[str],
    output: str,
    index=False,
):
    dbf = DBF(dbf_path)
    df = pd.DataFrame(iter(dbf))

    app = xw.App(visible=False, add_book=False)
    wb = app.books.open(template)
    sht = wb.sheets[sheet]

    for k, v in add_columns.items():
        df[k] = v

    result = df.query(condition).filter(items=columns)
    if index == True:
        result = result.reset_index(drop=True)
        result.index = result.index + 1
        result = result.reset_index()

    data = result.values.tolist()
    sht.range(arch).value = data

    wb.save(output)
    wb.close()
    app.quit()


# 老人企业补贴 -- 工行跨行
export_datagrid(
    dbf_path=dbf_path_a,
    template=temp_path_icbc,
    sheet="工行跨行",
    arch="A2",
    add_columns={"行别": 1, "业务种类": "00602", "协议书号": ""},
    condition="RE == 1 and 发放银行 == '工行异地'",
    columns=[
        "姓名",
        "X_银行帐号",
        "行别",
        "银行帐号",
        "业务种类",
        "协议书号",
        "发放地点",
        "实发补贴",
    ],
    output=r"E:\企业补贴\银行报盘\2025年\6月\工行报盘\老人企业补贴(工行报盘-跨行)_202506.xlsx",
)

# 老人企业补贴 -- 工行本行
export_datagrid(
    dbf_path=dbf_path_a,
    template=temp_path_icbc,
    sheet="工行跨行",
    arch="A2",
    add_columns={
        "行别": "",
        "跨行行号": "",
        "业务种类": "",
        "协议书号": "",
        "账号地址": "",
    },
    condition="RE == 1 and 发放银行 == '工商银行'",
    columns=[
        "姓名",
        "X_银行帐号",
        "行别",
        "跨行行号",
        "业务种类",
        "协议书号",
        "账号地址",
        "实发补贴",
    ],
    output=r"E:\企业补贴\银行报盘\2025年\6月\工行报盘\老人企业补贴(工行报盘-本行)_202506.xlsx",
)

# 老人企业补贴 -- 建设银行
export_datagrid(
    dbf_path=dbf_path_a,
    template=temp_path_cbc,
    sheet="sheet1",
    arch="A2",
    add_columns={
        "行别": "",
        "跨行行号": "",
        "业务种类": "",
        "协议书号": "",
        "账号地址": "",
    },
    condition="RE == 1 and 发放银行 == '建设银行'",
    columns=["X_银行帐号", "姓名", "实发补贴"],
    index=True,
    output=r"E:\企业补贴\银行报盘\2025年\6月\建行报盘\老人企业补贴(建行报盘)_202506.xls",
)

# 老人企业补贴 -- 中行南阳
export_datagrid(
    dbf_path=dbf_path_a,
    template=temp_path_bocny,
    sheet="sheet1",
    arch="A2",
    add_columns={
        "行别": "",
        "跨行行号": "",
        "业务种类": "",
        "协议书号": "",
        "账号地址": "",
    },
    condition="RE == 1 and 发放银行 == '中国银行_南阳'",
    columns=["姓名", "身份证", "发放银行", "X_银行帐号", "实发补贴"],
    index=True,
    output=r"E:\企业补贴\银行报盘\2025年\6月\中行报盘\老人企业补贴(中行南阳报盘)_202506.xlsx",
)

# 集体工企业补贴 -- 工行跨行
export_datagrid(
    dbf_path=dbf_path_b,
    template=temp_path_icbc,
    sheet="工行跨行",
    arch="A2",
    add_columns={"行别": 1, "业务种类": "00602", "协议书号": ""},
    condition="RE == 1 and 发放银行 in ('工商银行异地','商业银行（工行代发）')",
    columns=[
        "姓名",
        "X_银行帐号",
        "行别",
        "银行帐号",
        "业务种类",
        "协议书号",
        "发放地点",
        "实发补贴",
    ],
    output=r"E:\企业补贴\银行报盘\2025年\6月\工行报盘\集体工企业补贴(工行报盘-跨行)_202506.xlsx",
)

# 集体工企业补贴 -- 工行本行
export_datagrid(
    dbf_path=dbf_path_b,
    template=temp_path_icbc,
    sheet="工行跨行",
    arch="A2",
    add_columns={
        "行别": "",
        "跨行行号": "",
        "业务种类": "",
        "协议书号": "",
        "账号地址": "",
    },
    condition="RE == 1 and 发放银行 == '工商银行'",
    columns=[
        "姓名",
        "X_银行帐号",
        "行别",
        "跨行行号",
        "业务种类",
        "协议书号",
        "账号地址",
        "实发补贴",
    ],
    output=r"E:\企业补贴\银行报盘\2025年\6月\工行报盘\集体工企业补贴(工行报盘-本行)_202506.xlsx",
)

# 集体工企业补贴 -- 建设银行
export_datagrid(
    dbf_path=dbf_path_b,
    template=temp_path_cbc,
    sheet="sheet1",
    arch="A2",
    add_columns={
        "行别": "",
        "跨行行号": "",
        "业务种类": "",
        "协议书号": "",
        "账号地址": "",
    },
    condition="RE == 1 and 发放银行 == '建设银行'",
    columns=["X_银行帐号", "姓名", "实发补贴"],
    index=True,
    output=r"E:\企业补贴\银行报盘\2025年\6月\建行报盘\集体工企业补贴(建行报盘)_202506.xls",
)

# 集体工企业补贴 -- 中行南阳
export_datagrid(
    dbf_path=dbf_path_b,
    template=temp_path_bocny,
    sheet="sheet1",
    arch="A2",
    add_columns={
        "行别": "",
        "跨行行号": "",
        "业务种类": "",
        "协议书号": "",
        "账号地址": "",
    },
    condition="RE == 1 and 发放银行 == '中国银行_南阳'",
    columns=["姓名", "身份证", "发放银行", "X_银行帐号", "实发补贴"],
    index=True,
    output=r"E:\企业补贴\银行报盘\2025年\6月\中行报盘\集体工企业补贴(中行南阳报盘)_202506.xlsx",
)

# 集体工企业补贴 -- 中行油区
export_datagrid(
    dbf_path=dbf_path_b,
    template=temp_path_boc,
    sheet="sheet1",
    arch="A2",
    add_columns={
        "开户行": "中国银行",
        "行号": "41",
    },
    condition="RE == 1 and 发放银行 == '中国银行_油区'",
    columns=["实发补贴", "姓名", "X_银行帐号", "开户行", "行号"],
    index=False,
    output=r"E:\企业补贴\银行报盘\2025年\6月\中行报盘\集体工企业补贴(中行油区报盘)_202506.xlsx",
)

# 中人提高待遇 -- 工行跨行
export_datagrid(
    dbf_path=dbf_path_c,
    template=temp_path_icbc,
    sheet="工行跨行",
    arch="A2",
    add_columns={"行别": 1, "业务种类": "00602", "协议书号": ""},
    condition="RE == 1 and 发放银行 in ('工商银行（异地）','交通银行')",
    columns=[
        "姓名",
        "X_银行帐号",
        "行别",
        "收款行行号",
        "业务种类",
        "协议书号",
        "发放地点",
        "实发补贴",
    ],
    output=r"E:\企业补贴\银行报盘\2025年\6月\工行报盘\中人提高待遇(工行报盘-跨行)_202507.xlsx",
)

# 中人提高待遇 -- 工行本行
export_datagrid(
    dbf_path=dbf_path_c,
    template=temp_path_icbc,
    sheet="工行跨行",
    arch="A2",
    add_columns={
        "行别": "",
        "跨行行号": "",
        "业务种类": "",
        "协议书号": "",
        "账号地址": "",
    },
    condition="RE == 1 and 发放银行 == '工商银行'",
    columns=[
        "姓名",
        "X_银行帐号",
        "行别",
        "跨行行号",
        "业务种类",
        "协议书号",
        "账号地址",
        "实发补贴",
    ],
    output=r"E:\企业补贴\银行报盘\2025年\6月\工行报盘\中人提高待遇(工行报盘-本行)_202507.xlsx",
)

# 中人提高待遇 -- 建设银行
export_datagrid(
    dbf_path=dbf_path_c,
    template=temp_path_cbc,
    sheet="sheet1",
    arch="A2",
    add_columns={
        "行别": "",
        "跨行行号": "",
        "业务种类": "",
        "协议书号": "",
        "账号地址": "",
    },
    condition="RE == 1 and 发放银行 == '建设银行'",
    columns=["X_银行帐号", "姓名", "实发补贴"],
    index=True,
    output=r"E:\企业补贴\银行报盘\2025年\6月\建行报盘\中人提高待遇(建行报盘)_202507.xls",
)

# 中人提高待遇 -- 中行油区
export_datagrid(
    dbf_path=dbf_path_c,
    template=temp_path_boc,
    sheet="sheet1",
    arch="A2",
    add_columns={
        "开户行": "中国银行",
        "行号": "41",
    },
    condition="RE == 1 and 发放银行 == '中国银行_油区'",
    columns=["实发补贴", "姓名", "X_银行帐号", "开户行", "行号"],
    index=False,
    output=r"E:\企业补贴\银行报盘\2025年\6月\中行报盘\中人提高待遇(中行油区报盘)_202507.xlsx",
)
