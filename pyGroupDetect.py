#Written by Christopher Masselli
#christophermasselli@gmail.com


import pytesseract
import argparse
import cv2
import xlsxwriter
import re

#Creates an array of images corresponding to groups
def crop_Image(img): 
    h, w = img.shape[:2]
    hoffset = 38
    woffset =  2317
    x = 60
    leftsides = []
    rightsides =[]
    for y in range (48, 199, hoffset):
            y1 = y + hoffset
            x1 = x + woffset
            tiles = img[y:y+hoffset,x:x+woffset]
            leftsides.append(tiles)
            #under used for debug
            #cv2.rectangle(img, (x,y), (x1,y1), (0,255,0))
           # cv2.imwrite("save/" + str(x) +' _' + str(y) + ".png", tiles)
            #cv2.imwrite("test.png", img)

    for y in range (612, 615+123, hoffset):
            y1 = y + hoffset
            x1 = x + woffset
            tiles = img[y:y+hoffset,x:x+woffset]
            rightsides.append(tiles)
            #under used for debug
            #cv2.rectangle(img, (x,y), (x1,y1), (0,255,0))
            #cv2.imwrite("save/" + str(x) +' _' + str(y) + ".png", tiles)
            #cv2.imwrite("test.png", img)
    return leftsides, rightsides


#Makes string into an array of floats
def processString(s):
   splits = s.splitlines()
   nums = []
   for line in splits[1:]: #starts at line 1 to avoid group
        nums = nums + [float(re.sub('[^0-9.]', '', x)) for x in line.split()]
   return nums

#Prints data to excel
def printToText(dataLS, dataRS, outFile):
    workbook = xlsxwriter.Workbook(outFile)
    worksheet = workbook.add_worksheet()
    row = 0
    for col, arrays in enumerate(dataLS):
        worksheet.write_column(row, col, arrays)
    for col, arrays in enumerate(dataRS):
        worksheet.write_column(row, col + 4, arrays)
    workbook.close()

#MAIN
arg = argparse.ArgumentParser()
arg.add_argument("-i", "--image", required=True, help= "path to input image")
arg.add_argument("-o", "--out", type=str, default= "out.xlsx")
args = vars(arg.parse_args())

inPic = args["image"]
outFile = args["out"]
image = cv2.imread(inPic) 
ls , rs =crop_Image(image)
dataLS = []
dataRS = []
for gPic in ls:
    tempstring = pytesseract.image_to_string(gPic)
    dataLS = dataLS + [processString(tempstring.encode("utf-8"))]
for gPic in rs:
    tempstring = pytesseract.image_to_string(gPic)
    dataRS = dataRS + [processString(tempstring.encode("utf-8"))]

printToText(dataLS, dataRS, outFile) 
