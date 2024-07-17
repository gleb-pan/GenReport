from codecs import open as op
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from csv import reader, Sniffer
from datetime import datetime as dt


# Get current time
now = dt.now()
today = now.strftime("%d.%m.%Y")

# Path and filename
filename = f'Архив_событий_{now.strftime("%d.%m.%Y")}'


def read_csv_with_encoding(file_path, encoding):
    with op(file_path, 'r', encoding=encoding) as f:
        sample = f.read(1024)
        f.seek(0)
        dialect = Sniffer().sniff(sample)
        if dialect.delimiter not in [',', ';']:
            dialect.delimiter = ';'
        csv_data = [row for row in reader(f, dialect)]
        return csv_data

# Try to read the CSV file with different encodings
csv_data = []
encodings = ['utf-8-sig', 'utf-16', 'utf-8']
for encoding in encodings:
    try:
        csv_data = read_csv_with_encoding('Archive — копия.csv', encoding)
        if csv_data:
            break
    except Exception as e:
        print(f"Failed to read with encoding {encoding}: {e}")

# Count rows and columns
row_count = len(csv_data)
column_count = max(len(row) for row in csv_data)

# Write to Excel file
workbook = Workbook()
worksheet = workbook.active

for row in csv_data:
    worksheet.append(row)

# 1. Automatically adjust the width of the columns
for col in range(1, column_count + 1):
    max_length = 0
    column = get_column_letter(col)
    for cell in worksheet[column]:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    worksheet.column_dimensions[column].width = adjusted_width

# 2. Fill the very top row (only five columns) with blue color
fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
for col in range(1, column_count + 1):
    worksheet.cell(row=1, column=col).fill = fill

# 3. Apply "All borders" to the entire table
thin = Side(border_style="thin", color="000000")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

for row in range(1, row_count + 1):
    for col in range(1, column_count + 1):
        worksheet.cell(row=row, column=col).border = border

# Save the workbook
workbook.save(f"{filename}.xlsx")

# Print total rows and columns
print(f"Total rows: {row_count}")
print(f"Total columns: {column_count}")