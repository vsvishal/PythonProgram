import cv2
import numpy as np
import pandas as pd
import os

try:
    from PIL import Image
except ImportError:
    import Image
import pytesseract


import tkinter as tk
import tabula


class ImageTableToExcel:

    def OpenCVMain(self):

        HEIGHT = 500
        WIDTH = 600

        def btnClick():
            ImageTableDataExtract()

        #All code will will inside root & before mainloop()
        root = tk.Tk()

        root.title("Extract Table Data from Image into Excel")

        canvas = tk.Canvas(root, height=HEIGHT, width=WIDTH)
        canvas.pack()

        frame = tk.Frame(root, bg='#b3ffb3', bd=5)
        frame.place(relx=0.5, rely=0.02, relwidth=0.95, relheight=0.95, anchor='n')

        frame1 = tk.Frame(root, bg='white', bd=5)
        frame1.place(relx=0.5, rely=0.75, relwidth=0.65, relheight=0.1, anchor='n')

        lbl1 = tk.Label(frame, text='Extract Image Table data to Excel', font=('comicsansms',15,'bold') )
        lbl1.place(relx=0.15, rely=0.05, relwidth=0.70, relheight=0.10)

        lbl2 = tk.Label(frame, text='Enter Image file \n path: ', font=('comicsansms',9,'bold'))
        lbl2.place(relx=0.02, rely=0.2, relwidth=0.20, relheight=0.1)

        lbl3 = tk.Label(frame, text='Excel path with\n .xlsx extension  ', font=('comicsansms',9,'bold'))
        lbl3.place(relx=0.02, rely=0.35, relwidth=0.20, relheight=0.1)

        lower_frame = tk.Frame(root, bg='white', bd=5)
        lower_frame.place(relx=0.5, rely=0.75, relwidth=0.65, relheight=0.15, anchor='n')

        labelLowerFrame = tk.Label(lower_frame, bg='white', font=('comicsansms',11,'bold'))
        labelLowerFrame.place(relx=0.5, rely=0.1, relwidth=1, relheight=0.85, anchor='n')

        filePath = tk.StringVar(frame)
        filePathEntry = tk.Entry(frame, textvariable=filePath, font=('comicsansms', 8))
        filePathEntry.place(relx=0.25, rely=0.2, relwidth=0.72, relheight=0.1)

        savePath = tk.StringVar(frame)
        savePathEntry = tk.Entry(frame, textvariable=savePath, font=('comicsansms', 8))
        savePathEntry.place(relx=0.25, rely=0.35, relwidth=0.72, relheight=0.1)

        btn = tk.Button(root, text="Generate Excel", bg='#ffa64d',command=btnClick, font=('comicsansms',14,'bold'))
        btn.place(relx=0.3, rely=0.55, relheight=0.1, relwidth=0.3)

        def getImagePath():
            pdfPath = filePath.get()
            return pdfPath

        def getExcelPath():
            csvPath = savePath.get()
            return csvPath


        def ImageTableDataExtract():

            try:
                #Get the Pdf & CSV Path
                # read your file
                #drawing_file = r'C:\Users\u88ltuc\PycharmProjects\untitled1\Cropped\194340001_sd.pdf'
                drawing_file = getImagePath()
                extension = os.path.splitext(drawing_file)
                new_extension = extension[0] + ".png"
                img = Image.open(drawing_file)
                img.save(new_extension)

                img = cv2.imread(new_extension, 0)
                # img.shape

                # thresholding the image to a binary image
                thresh, img_bin = cv2.threshold(img, 128, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)

                # inverting the image
                img_bin = 255 - img_bin
                cv2.imwrite(r'C:\Users\u88ltuc\PycharmProjects\untitled1\Cropped\cv_inverted.png', img_bin)
                # Plotting the image to see the output
                #plotting = plt.imshow(img_bin, cmap='gray')
                # plt.show()

                # countcol(width) of kernel as 100th of total width
                kernel_len = np.array(img).shape[1] // 100
                # Defining a vertical kernel to detect all vertical lines of image
                ver_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, kernel_len))
                # Defining a horizontal kernel to detect all horizontal lines of image
                hor_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (kernel_len, 1))
                # A kernel of 2x2
                kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 1))

                # Use vertical kernel to detect and save the vertical lines in a jpg
                image_1 = cv2.erode(img_bin, ver_kernel, iterations=2)
                vertical_lines = cv2.dilate(image_1, ver_kernel, iterations=4)
                cv2.imwrite(r'C:\Users\u88ltuc\PycharmProjects\untitled1\Cropped\vertical.jpg', vertical_lines)
                # Plot the generated image
                # plotting = plt.imshow(image_1, cmap='gray')
                # plt.show()

                # Use horizontal kernel to detect and save the horizontal lines in a jpg
                image_2 = cv2.erode(img_bin, hor_kernel, iterations=4)
                horizontal_lines = cv2.dilate(image_2, hor_kernel, iterations=4)
                cv2.imwrite(r'C:\Users\u88ltuc\PycharmProjects\untitled1\Cropped\horizontal.jpg', horizontal_lines)
                # Plot the generated image
                # plotting = plt.imshow(image_2, cmap='gray')
                # plt.show()

                # Combine horizontal and vertical lines in a new third image, with both having same weight.
                img_vh = cv2.addWeighted(vertical_lines, 0.5, horizontal_lines, 0.5, 0.0)
                # Eroding and thesholding the image
                img_vh = cv2.erode(~img_vh, kernel, iterations=2)
                thresh, img_vh = cv2.threshold(img_vh, 128, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
                cv2.imwrite(r'C:\Users\u88ltuc\PycharmProjects\untitled1\Cropped\img_vh.jpg', img_vh)
                bitxor = cv2.bitwise_xor(img, img_vh)
                bitnot = cv2.bitwise_not(bitxor)
                # Plotting the generated image
                # plotting = plt.imshow(bitnot, cmap='gray')
                # plt.show()

                # Detect contours for following box detection
                contours, hierarchy = cv2.findContours(img_vh, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)

                def sort_contours(cnts, method="left-to-right"):
                    # initialize the reverse flag and sort index
                    reverse = False
                    i = 0
                    # handle if we need to sort in reverse
                    if method == "right-to-left" or method == "bottom-to-top":
                        reverse = True
                    # handle if we are sorting against the y-coordinate rather than
                    # the x-coordinate of the bounding box
                    if method == "top-to-bottom" or method == "bottom-to-top":
                        i = 1
                    # construct the list of bounding boxes and sort them from top to
                    # bottom
                    boundingBoxes = [cv2.boundingRect(c) for c in cnts]
                    (cnts, boundingBoxes) = zip(*sorted(zip(cnts, boundingBoxes),
                                                        key=lambda b: b[1][i], reverse=reverse))
                    # return the list of sorted contours and bounding boxes
                    return (cnts, boundingBoxes)

                # Sort all the contours by top to bottom.
                contours, boundingBoxes = sort_contours(contours, method="top-to-bottom")

                # Creating a list of heights for all detected boxes
                heights = [boundingBoxes[i][3] for i in range(len(boundingBoxes))]

                # Get mean of heights
                mean = np.mean(heights)

                # Create list box to store all boxes in
                box = []
                # Get position (x,y), width and height for every contour and show the contour on image
                for c in contours:
                    x, y, w, h = cv2.boundingRect(c)
                    if (w < 1000 and h < 500):
                        image = cv2.rectangle(img, (x, y), (x + w, y + h), (0, 255, 0), 2)
                        box.append([x, y, w, h])

                #plotting = plt.imshow(image, cmap='gray')
                # plt.show()

                # Creating two lists to define row and column in which cell is located
                row = []
                column = []
                j = 0

                # Sorting the boxes to their respective row and column
                for i in range(len(box)):

                    if (i == 0):
                        column.append(box[i])
                        previous = box[i]

                    else:
                        if (box[i][1] <= previous[1] + mean / 2):
                            column.append(box[i])
                            previous = box[i]

                            if (i == len(box) - 1):
                                row.append(column)

                        else:
                            row.append(column)
                            column = []
                            previous = box[i]
                            column.append(box[i])

                # print(column)
                # print(row)

                # calculating maximum number of cells
                countcol = 0
                for i in range(len(row)):
                    countcol = len(row[i])
                    if countcol > countcol:
                        countcol = countcol

                # Retrieving the center of each column
                center = [int(row[i][j][0] + row[i][j][2] / 2) for j in range(len(row[i])) if row[0]]

                center = np.array(center)
                center.sort()
                #print(center)
                # Regarding the distance to the columns center, the boxes are arranged in respective order

                finalboxes = []
                for i in range(len(row)):
                    lis = []
                    for k in range(countcol):
                        lis.append([])
                    for j in range(len(row[i])):
                        diff = abs(center - (row[i][j][0] + row[i][j][2] / 4))
                        minimum = min(diff)
                        indexing = list(diff).index(minimum)
                        lis[indexing].append(row[i][j])
                    finalboxes.append(lis)

                # from every single image-based cell/box the strings are extracted via pytesseract and stored in a list
                outer = []
                for i in range(len(finalboxes)):
                    for j in range(len(finalboxes[i])):
                        inner = ''
                        if (len(finalboxes[i][j]) == 0):
                            outer.append(' ')
                        else:
                            for k in range(len(finalboxes[i][j])):
                                y, x, w, h = finalboxes[i][j][k][0], finalboxes[i][j][k][1], finalboxes[i][j][k][2], \
                                             finalboxes[i][j][k][3]
                                finalimg = bitnot[x:x + h, y:y + w]
                                kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 1))
                                border = cv2.copyMakeBorder(finalimg, 2, 2, 2, 2, cv2.BORDER_CONSTANT, value=[255, 255])
                                resizing = cv2.resize(border, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
                                dilation = cv2.dilate(resizing, kernel, iterations=1)
                                erosion = cv2.erode(dilation, kernel, iterations=3)

                                out = pytesseract.image_to_string(erosion)
                                if (len(out) == 0):
                                    out = pytesseract.image_to_string(erosion, config='--psm 11')
                                inner = inner + " " + out
                            outer.append(inner)

                # Creating a dataframe of the generated OCR list5
                arr = np.array(outer)
                dataframe = pd.DataFrame(arr.reshape(len(row), countcol))

                # dataframe.dropna(axis=1, how='all', thresh=2, subset=None, inplace=True)

                #print(dataframe)
                # data = dataframe.style.set_properties(align="left")
                # Converting it in a excel-file
                excel_path =getExcelPath()
                dataframe.to_excel(excel_path)

                #Delete new png file
                os.remove(new_extension)

                labelLowerFrame['text'] = 'Data exported to Excel'

            except Exception as e:
                labelLowerFrame['text'] = e
                #print(e)

        root.mainloop()


if __name__ == '__main__':
    tabulaObj = ImageTableToExcel()
    tabulaObj.OpenCVMain()


#####################################################################
