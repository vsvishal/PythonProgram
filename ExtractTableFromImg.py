import cv2
import numpy as np
import matplotlib.pyplot as plt
import pytesseract
import pandas as pd


pytesseract.pytesseract.tesseract_cmd = r"C:\Users\VS\AppData\Local\Tesseract-OCR\tesseract.exe"

file = r"C:\Users\VS\Desktop\08900319.tif"

im1 = cv2.imread(file, 0)
im = cv2.imread(file)

ret, thresh_value = cv2.threshold(im1, 180, 255, cv2.THRESH_BINARY_INV)
#thresh_value = cv2.adaptiveThreshold(im1, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY,11, 2)

kernel = np.ones((5, 5), np.uint8)
dilated_value = cv2.dilate(thresh_value, kernel, iterations=1)

contours, hierarchy = cv2.findContours(dilated_value, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
cordinates = []
for cnt in contours:
    x, y, w, h = cv2.boundingRect(cnt)
    cordinates.append((x, y, w, h))
    # bounding the images
    if y < 50:
        cv2.rectangle(im, (x, y), (x + w, y + h), (0, 0, 255), 1)

#plt.imshow(im)
cv2.namedWindow('detecttable', cv2.WINDOW_NORMAL)
cv2.imwrite('tableData.jpg', im)

img = cv2.imread('tableData.jpg')
text = pytesseract.image_to_string(img)

file = open("ImageText.csv", "a")
file.write(text)
file.write("\n")
file.close()


#text = pytesseract.image_to_data(img, output_type=csv)
#print(text)