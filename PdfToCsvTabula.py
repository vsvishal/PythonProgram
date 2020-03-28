"""
Author: Vishal Sharma (VS)
Date: 28-Mar-2020

This Program extract the text from
PDF file Tabular Data using "Tabula"
module and generate csv

Use Anaconda/Miniconda to install
Tabula Library
"""

#import library
import tabula

#ENTER THE PDF FILE PATH
#Convert PDF into CSV file
tabula.convert_into(r"C:\Users\VS\Desktop\191640006.pdf", "Output.csv", output_format="csv", pages='all')


