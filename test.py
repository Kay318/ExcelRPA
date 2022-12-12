
import xlwings as xw

path = r"C:\Users\82109\Desktop\다국어자동화.xlsx"
 
wb = xw.Book(path)
print(wb.name)

# 모든 시트명 표시
li_sh = wb.sheets
li = [s.name for s in li_sh]
print(li)

sh1 = wb.sheets("SUMMARY")

# 자동 줄바꿈
sh1.range("A1:K1").api.WrapText = True

ex = sh1.range("A1:A3").expand("right")
print(ex)