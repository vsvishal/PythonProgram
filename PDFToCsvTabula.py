"""
Author: Vishal Sharma (VS)
Date: 28-Mar-2020

This Program extract the text from
PDF file Tabular Data using "Tabula"
module and generate csv

Use Anaconda/Miniconda to install
Tabula Library
"""

#import Library
import tkinter as tk

HEIGHT = 500
WIDTH = 600

#All code will will inside root & before mainloop()
root = tk.Tk()

root.title("Extract PDF tables from PDF using Camelot")

canvas = tk.Canvas(root, height=HEIGHT, width=WIDTH)
canvas.pack()

frame = tk.Frame(root, bg='#b3ffb3', bd=5)
frame.place(relx=0.5, rely=0.02, relwidth=0.95, relheight=0.95, anchor='n')

lbl1 = tk.Label(frame, text='Convert PDF tables to CSV', font=('comicsansms',15,'bold') )
lbl1.place(relx=0.2, rely=0.0, relwidth=0.50, relheight=0.1)

lbl2 = tk.Label(frame, text='Enter PDF file path: ', font=('comicsansms',9,'bold'))
lbl2.place(relx=0.02, rely=0.2, relwidth=0.20, relheight=0.1)

lbl3 = tk.Label(frame, text='CSV location path: ', font=('comicsansms',9,'bold'))
lbl3.place(relx=0.02, rely=0.35, relwidth=0.20, relheight=0.1)

filePath = tk.StringVar(frame)
filePathEntry = tk.Entry(frame, textvariable=filePath, font=40)
filePathEntry.place(relx=0.3, rely=0.2, relwidth=0.60, relheight=0.1)

savePath = tk.StringVar(frame)
savePathEntry = tk.Entry(frame, textvariable=savePath, font=40)
savePathEntry.place(relx=0.3, rely=0.35, relwidth=0.60, relheight=0.1)

btn = tk.Button(root, text="Covert to CSV", bg='#ffa64d', font=('comicsansms',14,'bold'))
btn.place(relx=0.3, rely=0.6, relheight=0.1, relwidth=0.3)

#ENTER THE PDF FILE PATH
#load the pdf tables

pdfLoc = filePath.get()

#print(tables)

#ENTER THE CSV FILE NAME
#Export the table data to csv
#If "compress=True" is set it will create zip file of CSV
saveLoc = savePath.get()
tables.export(saveLoc, f='csv', compress=False)
#tables.export('191640006.csv', f='csv', compress=False)


root.mainloop()



# #import library
# import tabula
#
# #Convert PDF into CSV file
# tabula.convert_into(r"C:\Users\VS\Desktop\191640006.pdf", "Output.csv", output_format="csv", pages='all')


