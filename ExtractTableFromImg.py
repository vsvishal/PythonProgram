# This Program Extracts the tabular data from image
# and save to the csv
#import libraries
import cv2
import numpy as np
import matplotlib.pyplot as plt
import pytesseract
import pandas as pd

#ENTER THE TESSERACT FILE PATH
#give the tesseract installion file location
pytesseract.pytesseract.tesseract_cmd = r"C:\Users\VS\AppData\Local\Tesseract-OCR\tesseract.exe"

#ENTE THE FILE PATH OF IMAGE
#image file path
file = r"C:\Users\VS\Desktop\08900319.tif"

#load image
im1 = cv2.imread(file, 0)
im = cv2.imread(file)

#Set the threshold for separating object from background
ret, thresh_value = cv2.threshold(im1, 180, 255, cv2.THRESH_BINARY_INV)

kernel = np.ones((5, 5), np.uint8)
#dilation is used for increasing object area and joining broken parts of object
dilated_value = cv2.dilate(thresh_value, kernel, iterations=1)

#contours are used for extractig the contours from image
#Contours are defined as the line joining all the points
#along the boundary of an image that are having the same intensity
contours, hierarchy = cv2.findContours(dilated_value, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
cordinates = []
for cnt in contours:
    x, y, w, h = cv2.boundingRect(cnt)
    cordinates.append((x, y, w, h))
    # bounding the images
    if y < 50:
        cv2.rectangle(im, (x, y), (x + w, y + h), (0, 0, 255), 1)

cv2.namedWindow('detecttable', cv2.WINDOW_NORMAL)
#save the processed image with new name
cv2.imwrite('tableData.jpg', im)

#read the new image file
img = cv2.imread('tableData.jpg')

#convert image to text
text = pytesseract.image_to_string(img)

#open the csv file and write the text in it
#YOU CAN GIVE THE LOCATION WHERE YOU HAVE TO THE SAVE THE CSV
#IF LOCATION NOT GIVEN IT WILL SAVE THE FILE TO LOCATION WHERE YOUR PROGRAM IS RUNNING
file = open("ImageText.csv", "a")
file.write(text)
file.write("\n")
file.close()


#text = pytesseract.image_to_data(img, output_type=csv)
#print(text)