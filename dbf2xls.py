from dbfread2 import DBF
import pandas as pd
import xlwings as xw

# Stream records (memory-efficient)
# for record in DBF(r"E:\bt_ltx.dbf"):
#    print(record)
# {'NAME': 'Alice', 'BIRTHDATE': datetime.date(1987, 3, 1)}
# {'NAME': 'Bob', 'BIRTHDATE': datetime.date(1980, 11, 12)}
table = DBF(r"E:\bt_ltx.dbf")
df = pd.DataFrame(iter(table))
# print(df)

nyzh = df.query(" RE == 0 and 死亡登记 == '2025-06'").filter(
    items=["收款行行号", "单位名称", "姓名", "身份证", "SWSJ", "补贴更正"]
)
arr = nyzh.values.tolist()
print(arr)
arr_len = len(arr)
app = xw.App(visible=False)
wb = app.books.add()
sht = wb.sheets[0]

title = sht.range("A1:G2")
title.merge()
title.value = "2025年6月离退休职工（老人）死亡登记汇总表"
title.api.font.name = "黑体"
# title.api.font.bold = True
title.api.font.size = 18

signature = sht["A3"]
signature.value = "单位：社保中心养老保险室"
signature.font.name = "楷体"
signature.font.size = 12
signature.api.HorizontalAlignment = -4131

header_rng = f"A4:G4"
header = sht[header_rng]
header.value = [
    "序号",
    "员工编号",
    "单位",
    "姓名",
    "身份证",
    "死亡时间",
    "企业补贴金额",
]


sht["A5"].options(transpose=True).value = list(range(1, arr_len + 1))
sht[f"A5:A{arr_len+4}"].api.HorizontalAlignment = -4108

sht[f"B5:E{arr_len+4}"].api.NumberFormat = "@"
sht["B5"].value = arr


sht[f"A{arr_len+5}:D{arr_len+5}"].merge()
sht[f"A{arr_len+5}:D{arr_len+5}"].value = "合计"
sht[f"A{arr_len+5}"].api.HorizontalAlignment = -4108

total = sht[f"G{arr_len+5}"]
total.formula = f"=SUM(G5:G{arr_len+4})"

comment = sht[f"B{arr_len+7}"]
comment.value = "*注：此表仅供参考，以实际变动为准*"
comment.api.font.name = "仿宋"
comment.api.font.size = 14

# 水平居中
sht["A1"].api.HorizontalAlignment = -4108
sht[f"A4:G{arr_len+5}"].api.HorizontalAlignment = -4108

# 设置列宽
sht[f"A3:A{arr_len+4}"].column_width = 7
sht[f"B3:B{arr_len+4}"].column_width = 9
sht[f"C3:C{arr_len+4}"].column_width = 18
sht[f"D3:D{arr_len+4}"].column_width = 12
sht[f"E3:E{arr_len+4}"].column_width = 22
sht[f"F3:F{arr_len+4}"].column_width = 10
sht[f"G3:G{arr_len+4}"].column_width = 15

# 设置行高
sht["A1"].row_height = 44.5
sht["A3"].row_height = 27.75
sht[f"A4:F{arr_len+5}"].row_height = 21.25

# 设置边框
content = sht[f"A4:G{arr_len+5}"]
content.api.BorderAround(LineStyle=1, Weight=2)

xborder = content.api.Borders(11)
xborder.LineStyle = 1
xborder.Weight = 2

yborder = content.api.Borders(12)
yborder.LineStyle = 1
yborder.Weight = 2


# 设置打印区域
print_area = f"$A$1:$G${arr_len+7}"
sht.page_setup.print_area = print_area

# 横向打印
# sheet.api.PageSetup.Orientation = 2  # 1为纵向，2为横向
# 纵向打印
sht.api.PageSetup.Orientation = 1

sht.api.PageSetup.Zoom = False
sht.api.PageSetup.FitToPagesWide = 1  # 宽度调整为1页
# sheet.api.PageSetup.FitToPagesTall = 1  # 高度调整为1页

# 设置页边距（单位：厘米）
# sheet.api.PageSetup.LeftMargin = 1.5
# sheet.api.PageSetup.RightMargin = 1.5
# sheet.api.PageSetup.TopMargin = 2
# sheet.api.PageSetup.BottomMargin = 2

# 水平居中
sht.api.PageSetup.CenterHorizontally = True


wb.save("离退休职工（老人）死亡登记汇总表-202506.xlsx")
wb.close()
