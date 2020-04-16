"""
Author: Vishal Sharma (VS)
Date: 28-Mar-2020
This Program extract the text from
PDF file Tabular Data using "Tabula"
module and generate csv
Use Anaconda/Miniconda to install
Tabula Library
"""

# import library
import tkinter as tk
import tabula


class PdfToCsvTabula():

    def camelotMain(self):

        HEIGHT = 500
        WIDTH = 600

        def btnClick():
            ConvertToCSv()

        #All code will will inside root & before mainloop()
        root = tk.Tk()

        root.title("Extract PDF tables from PDF using Camelot")

        canvas = tk.Canvas(root, height=HEIGHT, width=WIDTH)
        canvas.pack()

        frame = tk.Frame(root, bg='#b3ffb3', bd=5)
        frame.place(relx=0.5, rely=0.02, relwidth=0.95, relheight=0.95, anchor='n')

        frame1 = tk.Frame(root, bg='white', bd=5)
        frame1.place(relx=0.5, rely=0.75, relwidth=0.65, relheight=0.1, anchor='n')

        lbl1 = tk.Label(frame, text='Convert PDF tables to csv/txt', font=('comicsansms',15,'bold') )
        lbl1.place(relx=0.15, rely=0.05, relwidth=0.70, relheight=0.10)

        lbl2 = tk.Label(frame, text='Enter PDF file path: ', font=('comicsansms',9,'bold'))
        lbl2.place(relx=0.02, rely=0.2, relwidth=0.20, relheight=0.1)

        lbl3 = tk.Label(frame, text='CSV location path: ', font=('comicsansms',9,'bold'))
        lbl3.place(relx=0.02, rely=0.35, relwidth=0.20, relheight=0.1)

        lower_frame = tk.Frame(root, bg='white', bd=5)
        lower_frame.place(relx=0.5, rely=0.75, relwidth=0.65, relheight=0.15, anchor='n')

        labelLowerFrame = tk.Label(lower_frame, bg='white', font=('comicsansms',11,'bold'))
        labelLowerFrame.place(relx=0.5, rely=0.1, relwidth=1, relheight=0.85, anchor='n')

        filePath = tk.StringVar(frame)
        filePathEntry = tk.Entry(frame, textvariable=filePath, font=('comicsansms',8))
        filePathEntry.place(relx=0.25, rely=0.2, relwidth=0.72, relheight=0.1)

        savePath = tk.StringVar(frame)
        savePathEntry = tk.Entry(frame, textvariable=savePath, font=('comicsansms',8))
        savePathEntry.place(relx=0.25, rely=0.35, relwidth=0.72, relheight=0.1)

        btn = tk.Button(root, text="Covert to CSV", bg='#ffa64d',command=btnClick, font=('comicsansms',14,'bold'))
        btn.place(relx=0.3, rely=0.55, relheight=0.1, relwidth=0.3)

        def getPdfPath():
            pdfPath = filePath.get()
            return pdfPath

        def getCsvPath():
            csvPath = savePath.get()
            return csvPath


        def ConvertToCSv():

            try:
                #Get the Pdf & CSV Path
                pdf_path = getPdfPath()
                csv_path = getCsvPath()

                tabula.convert_into(pdf_path, csv_path, output_format="csv",pages='all')

                labelLowerFrame['text'] = 'File converted'

            except Exception as e:
                labelLowerFrame['text'] = e
                print(e)

        root.mainloop()


if __name__ == '__main__':
    tabulaObj = PdfToCsvTabula()
    tabulaObj.camelotMain()


#import library
#import tabula

#ENTER THE PDF FILE PATH
#Convert PDF into CSV file
#tabula.convert_into(r"C:\Users\VS\Desktop\191640006.pdf", "Output.csv", output_format="csv", pages='all')
