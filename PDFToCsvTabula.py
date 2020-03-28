import tabula

#Read PDF into list of DataFrame
df = tabula.read_pdf(r"C:\Users\VS\Desktop\191640003.pdf", pages="all")

#Convert PDF into CSV file
tabula.convert_into(r"C:\Users\VS\Desktop\191640003.pdf", "Output.csv", output_format="csv", pages='all')


