
import xlwings

if __name__ == '__main__':
    wb = xlwings.books.active
    sh = wb.sheets.active
    cell = sh.cells(1, 1)
    print(cell.address)
