"""
put this script to the same file with websites.xlsx which contains the links
author: Caner Erden
"""

import xlrd, webbrowser

workbook = xlrd.open_workbook('websites.xlsx')
# Sheet name
sheet = workbook.sheet_by_name('Sayfa1')

# Suppose your URLs are in column 5, rows 2 to 30
url_column = 5

# This will open 28 webpages on your browser
for row in range(2, 30):
    url = sheet.cell_value(row, url_column)
    webbrowser.open_new_tab(url)