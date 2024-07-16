from codecs import open as op
from openpyxl import Workbook
from csv import reader, Sniffer
from datetime import datetime
from socket import gethostname
from os import getlogin

hostname = gethostname()  # get PC-name using socket
username = getlogin()  # get username using os

now = datetime.now()  # Current time
today = now.strftime("%d.%m.%Y_%H-%M")  # Filtered current time
final_path = f"C:\\Users\\{username}\\Desktop"  # path where the file will be saved to
filename = f'Отчет_по_моточасам_{today}'  # name of resulted .xlsx file

def read_csv_with_encoding(file_path, encoding):
    with op(file_path, 'r', encoding=encoding) as f:
        sample = f.read(1024)
        f.seek(0)
        dialect = Sniffer().sniff(sample)
        if not dialect.delimiter in [',', ';']:
            dialect.delimiter = ';'
        csv_data = [row for row in reader(f, dialect)]
        return csv_data

# Try to read the CSV file with different encodings
csv_data = []
encodings = ['utf-8-sig', 'utf-16', 'utf-8']
for encoding in encodings:
    try:
        csv_data = read_csv_with_encoding('Archive.csv', encoding)
        if csv_data:
            break
    except Exception as e:
        print(f"Failed to read with encoding {encoding}: {e}")

row_count = len(csv_data)
column_count = max(len(row) for row in csv_data)

# Write to Excel file
workbook = Workbook()
worksheet = workbook.active

for row in csv_data:
    worksheet.append(row)

workbook.save(f"{final_path}\\{filename}.xlsx")
print(f"Total rows: {row_count}")
print(f"Total columns: {column_count}")