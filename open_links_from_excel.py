"""
put this script to the same file with websites.xlsx which contains the links
author: Caner Erden
"""

import xlrd, webbrowser

workbook = xlrd.open_workbook('websites.xlsx')
# Sheet name
sheet = workbook.sheet_by_name('Sayfa1')

# Suppose your URLs are in column 5 which is F column, rows 2 to 6 which are F2, F3, F4, F5 cells
url_column = 5

# This will open 28 webpages on your browser
for row in range(1, 6):
    url = sheet.cell_value(row, url_column)
    webbrowser.open_new_tab(url)