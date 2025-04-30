# export-tmp.py

import xlwings as xw
import pandas as pd
from dbfread2 import DBF

# 老人补贴
tab_a = DBF(r"E:\企业补贴\数据\企业补贴202505\bt_ltx.dbf")


app = xw.App(visible=False, add_book=False)
wb = app.books.open(r"E:\企业补贴\银行报盘\工商银行报盘模板.xlsx")
sht = wb.sheets["工行跨行"]

df = pd.DataFrame(iter(tab_a))

bankA_95588 = df.query("RE == 1 and 发放银行 == '工商银行'").filter(
    items=["姓名", "X_银行帐号", "实发补贴"]
)
bankA_95588_ = df.query("RE == 1 and 发放银行 == '工行异地'").filter(
    items=["姓名", "X_银行帐号", "银行帐号", "发放地点", "实发补贴"]
)
bankA_95588_["行别"] = "1"
bankA_95588_["跨行行号"] = bankA_95588_["银行帐号"]
bankA_95588_["业务种类"] = "00602"
bankA_95588_["协议书号"] = ""
bankA_95588_["账号地址"] = bankA_95588_["发放地点"]
bankA_95588_ = bankA_95588_.filter(
    items=[
        "姓名",
        "X_银行帐号",
        "行别",
        "跨行行号",
        "业务种类",
        "协议书号",
        "账号地址",
        "实发补贴",
    ]
)
bankA_95588_ = bankA_95588_.values.tolist()


bankA_95588["行别"] = ""
bankA_95588["跨行行号"] = ""
bankA_95588["业务种类"] = ""
bankA_95588["协议书号"] = ""
bankA_95588["账号地址"] = ""
bankA_95588 = bankA_95588.filter(
    items=[
        "姓名",
        "X_银行帐号",
        "行别",
        "跨行行号",
        "业务种类",
        "协议书号",
        "账号地址",
        "实发补贴",
    ]
)
bankA_95588 = bankA_95588.values.tolist()
len_list = len(bankA_95588_) + 2

sht.range("A2").value = bankA_95588_
sht.range(f"A{len_list}").value = bankA_95588

wb.save(
    path=r"E:\企业补贴\银行报盘\2025年\5月\工行报盘\老人企业补贴工行报盘_202505.xlsx"
)
wb.close()
app.quit()


# 集体工补贴
tab_b = DBF(filepath=r"E:\企业补贴\数据\集体工企业补贴202505\bt_ltx.dbf")
# for r in tab_b:
#    print(r)
app = xw.App(visible=False, add_book=False)
wb = app.books.open(r"E:\企业补贴\银行报盘\工商银行报盘模板.xlsx")
sht = wb.sheets["工行跨行"]

df = pd.DataFrame(iter(tab_b))

bankB_95588 = df.query("RE == 1 and 发放银行 == '工商银行'").filter(
    items=["姓名", "X_银行帐号", "实发补贴"]
)

bankB_95588_ = df.query(
    "RE == 1 and 发放银行 in ('工商银行异地','商业银行（工行代发）' )"
).filter(items=["姓名", "X_银行帐号", "银行帐号", "发放地点", "实发补贴"])

bankB_95588_["行别"] = "1"
bankB_95588_["跨行行号"] = bankB_95588_["银行帐号"]
bankB_95588_["业务种类"] = "00602"
bankB_95588_["协议书号"] = ""
bankB_95588_["账号地址"] = bankB_95588_["发放地点"]
bankB_95588_ = bankB_95588_.filter(
    items=[
        "姓名",
        "X_银行帐号",
        "行别",
        "跨行行号",
        "业务种类",
        "协议书号",
        "账号地址",
        "实发补贴",
    ]
)
bankB_95588_ = bankB_95588_.values.tolist()
# print(bankB_95588_)

bankB_95588["行别"] = ""
bankB_95588["跨行行号"] = ""
bankB_95588["业务种类"] = ""
bankB_95588["协议书号"] = ""
bankB_95588["账号地址"] = ""
bankB_95588 = bankB_95588.filter(
    items=[
        "姓名",
        "X_银行帐号",
        "行别",
        "跨行行号",
        "业务种类",
        "协议书号",
        "账号地址",
        "实发补贴",
    ]
)
bankB_95588 = bankB_95588.values.tolist()
list_len = len(bankB_95588_) + 2

sht.range("A2").value = bankB_95588_
sht.range(f"A{list_len}").value = bankB_95588

wb.save(
    path=r"E:\企业补贴\银行报盘\2025年\5月\工行报盘\集体工企业补贴工行报盘_202505.xlsx"
)
wb.close()
app.quit()

# 建行报盘
app = xw.App(visible=False, add_book=False)
wb = app.books.open(r"E:\企业补贴\银行报盘\建设银行报盘模板.xlsx")
sht = wb.sheets["sheet1"]

df = pd.DataFrame(iter(tab_a))
bankA_cbc = df.query("RE == 1 and 发放银行 == '建设银行'").filter(
    items=["X_银行帐号", "姓名", "实发补贴"]
)


bankA_cbc = bankA_cbc.reset_index(drop=True)

bankA_cbc.index = bankA_cbc.index + 1
bankA_cbc = bankA_cbc.reset_index()

arr = bankA_cbc.values.tolist()
# print(arr)
sht.range("A2").value = arr
wb.save(
    path=r"E:\企业补贴\银行报盘\2025年\5月\建行报盘\老人企业补贴建行报盘_202505.xlsx"
)
wb.close()
app.quit()

# 建行集体工
app = xw.App(visible=False, add_book=False)
wb = app.books.open(r"E:\企业补贴\银行报盘\建设银行报盘模板.xlsx")
sht = wb.sheets["sheet1"]

df = pd.DataFrame(iter(tab_b))
bankA_cbc = df.query("RE == 1 and 发放银行 == '建设银行'").filter(
    items=["X_银行帐号", "姓名", "实发补贴"]
)


bankA_cbc = bankA_cbc.reset_index(drop=True)

bankA_cbc.index = bankA_cbc.index + 1
bankA_cbc = bankA_cbc.reset_index()
arr = bankA_cbc.values.tolist()

sht.range("A2").value = arr
wb.save(
    path=r"E:\企业补贴\银行报盘\2025年\5月\建行报盘\集体工企业补贴建行报盘_202505.xlsx"
)
wb.close()
app.quit()


# 中行油区
app = xw.App(visible=False, add_book=False)
wb = app.books.open(r"E:\企业补贴\银行报盘\中国银行报盘模板.xlsx")
sht = wb.sheets["sheet1"]

df = pd.DataFrame(iter(tab_b))
boc = df.query("RE == 1 and 发放银行 == '中国银行_油区'").filter(
    items=["实发补贴", "姓名", "X_银行帐号"]
)
boc["开户行"] = "中国银行"
boc["行号"] = "41"
boc.filter(items=["实发补贴", "姓名", "X_银行帐号", "开户行", "行号"])


arr = boc.values.tolist()
# print(arr)
sht.range("A2").value = arr
wb.save(
    path=r"E:\企业补贴\银行报盘\2025年\5月\中行报盘\集体工企业补贴中行报盘_202505.xlsx"
)
wb.close()
app.quit()

# 中行南阳
app = xw.App(visible=False, add_book=False)
wb = app.books.open(r"E:\企业补贴\银行报盘\中国银行南阳报盘模板.xlsx")
sht = wb.sheets["sheet1"]

df = pd.DataFrame(iter(tab_a))
boc_a = df.query("RE == 1 and 发放银行 == '中国银行_南阳'").filter(
    items=["姓名", "身份证", "发放银行", "X_银行帐号", "实发补贴"]
)
df = pd.DataFrame(iter(tab_b))
boc_b = df.query("RE == 1 and 发放银行 == '中国银行_南阳'").filter(
    items=["姓名", "身份证", "发放银行", "X_银行帐号", "实发补贴"]
)


boc_comb = pd.concat([boc_a, boc_b], axis=0)
boc_uni = boc_comb.drop_duplicates()
boc_uni = boc_uni.reset_index(drop=True)
boc_uni.index = boc_uni.index + 1
boc_uni = boc_uni.reset_index()

arr = boc_uni.values.tolist()
# print(arr)
sht.range("A2").value = arr


wb.save(
    path=r"E:\企业补贴\银行报盘\2025年\5月\中行报盘\企业补贴中行南阳报盘_202505.xlsx"
)
wb.close()
app.quit()
