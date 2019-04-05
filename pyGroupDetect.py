#Written by Christopher Masselli
#christophermasselli@gmail.com

from PIL import Image
import pytesseract
import argparse
import cv2
import os
import xlsxwriter
import re
import numpy
class Group:
    def __init__(self, groupname, data):
        self.groupname = groupname
        self.data = data
    
def crop_Image(img): 
    h, w = img.shape[:2]
    hoffset = 70
    woffset =  1140
    x = 0
    leftsides = []
    for y in range (0, h, hoffset):
            y1 = y + hoffset
            x1 = x + woffset
            tiles = img[y:y+hoffset,x:x+woffset]
            leftsides.append(tiles)
            cv2.rectangle(img, (x,y), (x1,y1), (0,255,0))
            cv2.imwrite("save/" + str(x) +' _' + str(y) + ".png", tiles)
            cv2.imwrite("test.png", img)
    x = 1140 + 80
    rightsides = []
    for y in range (0, h, hoffset):
            y1 = y + hoffset
            x1 = x + woffset
            tiles = img[y:y+hoffset,x:x+woffset]
            rightsides.append(tiles)
            cv2.rectangle(img, (x,y), (x1,y1), (0,255,0))
            cv2.imwrite("save/" + str(x) +' _' + str(y) + ".png", tiles)
            cv2.imwrite("test.png", img)
    return leftsides, rightsides

#    outFile = open('./out.txt', 'w+')

def processString(s):
   splits = s.splitlines()
#   group = splits[0][0:6]
   nums = []
   for line in splits[1:]:
        nums = nums + [float(re.sub('[^0-9.]', '', x)) for x in line.split()]
   return nums
#   return [Group(group, nums)]

def printToText(groupls, grouprs, outFile):
    workbook = xlsxwriter.Workbook('out.xlsx')
    worksheet = workbook.add_worksheet()
    row = 0
    for col, arrays in enumerate(groupls):
        worksheet.write_column(row, col, arrays)
    for col2, arrays2 in enumerate(grouprs):
        worksheet.write_column(row, col2, arrays2)
    workbook.close()
#    az, bs, cs, ds = groupls
#    z= map(None, az, bs, cs, ds)


#    for x in groupls:
#        print >> outFile, x.groupname + '\n'
#        for datas in x.data:
#            print >> outFile, datas


pathname = "./test.bmp"
image = cv2.imread(pathname)
image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)




#image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
image = cv2.threshold(image, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
#image = cv2.bilateralFilter(image,9,75,75)
#temp = cv2.GaussianBlur(image,(0,0),3)
#temp = cv2.addWeighted(1.5, image, -0.5,0,image)

image = cv2.equalizeHist(image)
ls , rs =crop_Image(image)
groupls = []
grouprs = []
for primes in ls:
    tempstring = pytesseract.image_to_string(primes)
    groupls = groupls +[processString(tempstring.encode("utf-8"))]
for primes2 in rs:
    tempstring = pytesseract.image_to_string(primes2)
    groupsrs = grouprs + [processString(tempstring.encode("utf-8"))]
printToText(groupls, grouprs, "txt.out")
#outstring = pytesseract.image_to_string(image)

#pytesseract('./test.jpg', 'output', lang=None, boxes=False, config="hocr")
#print(dir(pytesseract))





# print(pytesseract.image_to_boxes(Image.open(pathname)))



    
