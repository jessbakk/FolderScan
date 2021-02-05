import os
import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell
import math


# Converts the file size from bytes into a readable format (B, KB, MB, GB)
def convert_size(size_bytes):
    if size_bytes == 0:
        return "0B"
    size_name = ("B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB")
    i = int(math.floor(math.log(size_bytes, 1024)))
    p = math.pow(1024, i)
    s = round(size_bytes / p, 2)
    return "%s %s" % (s, size_name[i])


path = input("Copy and paste path to files: ")

if __name__ == '__main__':
    # Get all files
    files = os.listdir(path)

    pairs = []
# Loop and add files to list
for file in files:
    # Use join to get full file path
    location = os.path.join(path, file)

    # Get size and add to list of tuples
    size = os.path.getsize(location)
    pairs.append((file, convert_size(size)))

# Sort list of tuples (file, size)
pairs.sort(key=lambda s: s[0])

#Display pairs
for pair in pairs:
   print(pair)

# Create Pandas data frame from the list of tuples
df = pd.DataFrame(pairs, columns=['File Name', 'File Size'])
print(df)

# Takes the data frame and creates Excel sheet with path where it will be saved
writer = pd.ExcelWriter(path + "/" + "FolderScan.xlsx", engine='xlsxwriter')
df.to_excel(writer, index=False, sheet_name='Sheet1')

# Formats columns in Excel sheet using XlsxWriter library
workbook = writer.book
worksheet = writer.sheets['Sheet1']

column_format = workbook.add_format()
column_format.set_align('right')

worksheet.set_column('A:A', 45)
worksheet.set_column('B:B', 10, column_format)

writer.save()
