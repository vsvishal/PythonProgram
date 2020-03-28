"""
Author: Vishal Sharma (VS)
Date: 28-Mar-2020

This Program extract the text from
PDF file Tabular Data using "Camelot"
module and generate csv

Use Anaconda/Miniconda to install
Camelot Library not from PIP
"""

#Import library
import camelot

#ENTER THE PDF FILE PATH
#load the pdf file
tables = camelot.read_pdf(r"C:\Users\VS\Desktop\ENG_CD_1103426_D6.pdf")

#ENTER THE CSV FILE NAME
#Export the table data to csv
#If "compress=True" is set it will create zip file of CSV
tables.export('ENG_CD_1103426.csv', f='csv', compress=False)
