from graphics import *
import xlrd
path = "C:\\Users\\Faruk\\Desktop\\yemek\\ücret2.xlsx"
exel_workbook = xlrd.open_workbook(path)
exel_worksheet = exel_workbook.sheet_by_index(0)
total = 0

for row in range(1, exel_worksheet.nrows):
    total += float(exel_worksheet.cell_value(row, 1))

win = GraphWin("Total", 300, 200)
win.setBackground(color_rgb(95, 183, 190))
win.setCoords(0, 0, 100, 100)
rect = Rectangle(Point(25, 40), Point(75, 70))
rect.draw(win)
header = Text(Point(50, 62.5), "The total amount is")
print_total = Text(Point(50, 48), str(total))
print_total.draw(win)
header.draw(win)
win.getMouse()
win.close()
