#Importing all of the libs
from PIL import Image
import pytesseract
import sys
from pdf2image import convert_from_path
import os
import xlwt
from xlwt import Workbook
import argparse

msg = "Script to perform OCR on PDF and add values to excel"
parser= argparse.ArgumentParser(description=msg)
parser.add_argument("--pdf_name",help="the name of pdf on which to perform OCR without extension",required=True,type=str)
parser.add_argument("--excel_name",help="the name of excel where the result will be stored without extension",required=True,type=str)
args = parser.parse_args()

global count
count = 0  
  
#Method to parse the list values 4 at a time
def pairwise(iterable):
    "s -> (s0, s1), (s2, s3), (s4, s5), ..."
    a = iter(iterable)
    return zip(a, a, a, a)

#Cropping and then using tessaract on it
def cropInfer(coordinateList,image): 
    fileName =[]
    imageName =[]
    im = Image.open(image)
    for l, t, r, b in pairwise(coordinateList):
        global count  
        # Setting the points for cropped image
        # left = 650
        left = l
        top = t
        right = r
        bottom = b
        imgName = "img"+str(count)+".png"
        txtName = "img"+str(count)+".txt"
        # # (It will not change orginal image)
        imgCropped = im.crop((left, top, right, bottom))
        # Shows the image in image viewer
        imgCropped.save(imgName,"PNG")
        os.system(f"tesseract {imgName} {txtName[:-4]} --psm 6 >/dev/null 2>&1")
        count+=1
        fileName.append(txtName)
        imageName.append(imgName)
    return fileName,imageName

#This function will be used to parse the txt files and write to the xml files
def parseTxt(txtList,sheetName,start):
    i = 0
    j= start
    count = 0
    for fileName in txtList:
        with open(fileName) as input_file:
            for line in input_file: 
                if count == 0:
                    line = line.replace(" ", ": ")
                else:    
                    line = line.replace(". ", ": ")
                if len(line) > 1:
                    items= [item.strip() for item in line.split(':')]
                    if i!=4 and j==0: # for the headings
                        sheetName.write(j,i,items[0])
                        sheetName.write(j+1,i,items[1])
                    elif i!=4 and j>0: #for normal values
                        sheetName.write(j+1,i,items[1])
                    elif i==4 and j==0:
                        sheetName.write(j,i,"Type")
                    else:
                        sheetName.write(j+1,i,items[0])
                    i+=1
                count+=1   
        wb.save(args.excel_name+'.xls')
    return j

#Parsing of the Matrix
def parseBlock(txtList,sheetName):
#This will always give us 4 file
    totalLineOne =""
    totalLineTwo =""
    upperQuad = []
    lowerQuad = []
    with open(txtList[0]) as input_file,open(txtList[1]) as input_file_two, open(txtList[2]) as input_file_three,open(txtList[3]) as input_file_four:
            for line1, line2, line3, line4 in zip(input_file, input_file_two,input_file_three,input_file_four): 
                line1 = line1.replace("= ", "")
                line1 = line1.replace("\n", " ")
                line2 = line2.replace("= ","")
                line2 = line2.replace("\n", " ")
                line3 = line3.replace("= ", "")
                line3 = line3.replace("\n", " ")
                line4 = line4.replace("= ","")
                line4 = line4.replace("\n", " ")
                totalLineOne += line1 + line2    
                totalLineTwo += line3 + line4
            upperQuad.append(totalLineOne.split(" "))
            lowerQuad.append(totalLineTwo.split(" "))
            upperQuad = [x for x in upperQuad[0] if x]
            lowerQuad = [x for x in lowerQuad[0] if x]
            startUpper = 0 #the place to start
            readUpper = 4 # the amount of numbers to read in 1st pass
            endUpper =4 # end upper is where we will read till
            incUpper = 2 # the amount to add in every pass



            startLower = 0 
            readLower = 9
            endLower = 9
            decLower = 8
            lenLower = len(lowerQuad)
            lenUpper = len(upperQuad)
            for upper in upperQuad:
                    if startUpper <=27 and len(upper) > 0 and lenLower >0:
                        print(upperQuad[startUpper:endUpper])
                        readUpper+= incUpper
                        startUpper = endUpper
                        endUpper = startUpper + readUpper       
            
            for lower in lowerQuad:
                if startLower <=27 and lenUpper > 0 and len(lower) > 0:
                    print(lowerQuad[startLower:endLower])
                    readLower = decLower
                    startLower = endLower
                    endLower = startLower + readLower
                    decLower -= 2


# Path of the pdf
PDF_file = args.pdf_name+".pdf"

# Store all the pages of the PDF in a variable
pages = convert_from_path(PDF_file, 500)
  
# Counter to store images of each page of PDF to image
image_counter = 1

origList = [0,0,1280,720,2960,800,3904,912,248,976,1608,1536,1710,960,2952,1480,3120,960,4000,1336]
blockList = [880,1650,2168,2736]
cropList = [0,0,619,507,665,0,1250,500,25,550,615,1075,650,550,1270,1075]
wb = Workbook()
sheet1 = wb.add_sheet('Patient Details')
sheet1_Block = wb.add_sheet('Patient Block Data')
# Iterate through all the pages stored above
startRow = 0
for page in pages:
    # PDF page n -> page_n.jpg
    filename = "page_"+str(image_counter)+".png" 
    print("Processing",filename)
    # Save the image of the page in system
    page.save(filename, 'PNG')
    # Increment the counter to update filename
    image_counter = image_counter + 1
    files,images =cropInfer(origList,filename)
    parseTxt(files,sheet1,startRow)
    files,images = cropInfer(blockList,filename)
    #the last image will be the block image where all of the int values are stored
    #This will create 4 distinct quadrants of the big blob
    files,images =cropInfer(cropList,images[-1])
    parseBlock(files,sheet1_Block)
    startRow+=1
print("OCR Completed for",image_counter)
