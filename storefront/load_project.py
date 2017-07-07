import xlrd
dic = {"num":1}
rb = xlrd.open_workbook('storefront_project1.xlsx')
sheet = rb.sheet_by_index(0)
for rownum in range(2, sheet.nrows):
    row = sheet.row_values(rownum)
    if (row[0] != 0):
        dic[row[0]] = row[1]

for key in sorted(dic.keys()):
    print("%s: %s" % (key, dic[key]))
#print (dic)