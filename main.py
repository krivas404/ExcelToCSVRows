import os.path
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import csv
import settings as s

file_name = s.file_name
new_file = s.new_file
sheet_name = s.sheet_name

wb = load_workbook(filename=file_name, read_only=True)
sheet = wb[sheet_name]


def main():
    print('started')
    val1 = sheet['B1'].value
    print('val1 =', val1)
    i = 0
    with open(new_file, 'w', newline='') as csv_file:
        csvwriter = csv.writer(csv_file, dialect='excel', delimiter=';')
        for row in sheet.rows:
            if i == 0:
                i += 1
                continue
            csvwriter.writerow([row[1].value, row[6].value])
            i += 1


if __name__ == "__main__":
    main()