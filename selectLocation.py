import xlrd
import xlsxwriter
outWorkbook=xlsxwriter.Workbook("a.xlsx")
outSheet=outWorkbook.add_worksheet()
location=("newDATA.xlsx")
wb=xlrd.open_workbook(location)
#print(wb)
liste=[]
sheet=wb.sheet_by_index(0)
sheet.cell_value(0,0)
j=0
t=0
y=[]
sum1=0
sum2=0
sum3=0
for i in range(1,2842):
    liste=sheet.row_values(i)
    outSheet.write(j, 0, liste[0])
    outSheet.write(j, 1, liste[1])
    outSheet.write(j, 4, liste[4])


    sum1 += liste[2]
    sum2 += liste[3]
    if liste[3]>140:
        outSheet.write(j, 3, "True")
    else:
        outSheet.write(j, 3, "False")
    if liste[2]<45:
        outSheet.write(j, 2, "True")
    else:
        outSheet.write(j, 2, "False")



    j+=1
print("aver",sum1/j)
print("max",sum2/j)
outWorkbook.close()




