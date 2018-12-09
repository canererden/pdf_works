"""
This script gets all PDF files meta information in a file and give them within a CSV file.
author: Caner Erden
"""
# -*- coding: utf-8 -*-

import os
import csv
import PyPDF2
from openpyxl import Workbook

# Get all pdfs in a folder
files = [f for f in os.listdir('.') if os.path.isfile(
    f) and f.endswith('.pdf')]  # Working Directory

# define excel workbook
kitap = Workbook()

# Choose excel sheet
sayfa = kitap.active

# Enter First 2 cell name
sayfa['A1'] = "Dosya İsmi"
sayfa['B1'] = "Sayfa Numarası"

# Give the numbers
for file in files:
    pdf=PyPDF2.PdfFileReader(open(file,'rb'))
    sayfa.append([file, pdf.getNumPages()])

# Save the workbook
kitap.save("sayfa_numaralari.xlsx")

# Close the workbook
kitap.close()