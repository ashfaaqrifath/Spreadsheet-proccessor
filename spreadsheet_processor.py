import openpyxl as xl
from openpyxl.chart import BarChart, Reference

confirm = input("Did you upload the desired file? (y/n): ")
if confirm.lower() == "y" :
        filename = input("Enter file name: ")
        rename = input("Rename your file: ")

        wb = xl.load_workbook(filename)
        sheet = wb['Sheet1']

        for row in range(2, sheet.max_row + 1):
            cell = sheet.cell(row, 3)
            corrected_price = cell.value * 0.9
            corrected_price_cell = sheet.cell(row, 4)
            corrected_price_cell.value = corrected_price

        values = Reference(sheet,
                  min_row=2,
                  max_row=sheet.max_row,
                  min_col=4,
                  max_col=4)

        chart = BarChart()
        chart.add_data(values)
        sheet.add_chart(chart, 'a6')

        wb.save(rename)
        print("--------------------------------")
        print("SPREADSHEET SUCCESSFULLY UPDATED")
        print("OPEN IN FILE EXPLORER TO SEE THE CHANGES")
else:
    print("-----------------------")
    print("PLEASE UPLOAD YOUR FILE")