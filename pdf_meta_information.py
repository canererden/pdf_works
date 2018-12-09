"""
This script gets all PDF files meta information in a file and give them within a xlsx file.
author: Caner Erden
"""
# -*- coding: utf-8 -*-

import glob 
import csv
import PyPDF2
from openpyxl import Workbook

# Get all pdfs in a folder
files = glob.glob('pdf_samples\\*.pdf')  # Working Directory

# define excel workbook
kitap = Workbook()

# Choose excel sheet
sayfa = kitap.active

# Enter First 2 cell name
sayfa['A1'] = "File Name"
sayfa['B1'] = "Page Number"

# Give the numbers
for file in files:
    pdf=PyPDF2.PdfFileReader(open(file,'rb'))
    sayfa.append([file, pdf.getNumPages()])

# Save the workbook
kitap.save("extracted_page_numbers_from_PDF.xlsx")

# Close the workbook
kitap.close()