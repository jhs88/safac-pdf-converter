import os  # if you have a macbook
import xlrd  # to access excel docs
from PyPDF2 import PdfFileReader, PdfFileWriter  # to read pdfs
import glob  # to do paths
import shutil  # to be able to do many operations
import time  # to slow it down to complete the code

#Stephen:
#For this code to work, one must change the path accordingly depending on the nature of the OS ie MAC/Windows/Linux etc. Maybe someone can get a more
#global version in the future :)
#I also have changed the structure of the folders because of changing the shutil.copy function. Essentially, within your folder there should be
#two folders. One containing the master pdf file and the other empty. This folder will be where all the pdfs will appear.
#the code and the excel file can remain in the root folder.
#When changing the paths, I initially had some unicode errors. This is resolved by putting 'r' before the path as you can see. This converts the path
#to a raw string.
#Finally, there is an issue in the excel file. Basically special characters such as ':' or even ';' will cause an error. Make sure to change that
#to a hypen or something.

def split(path, name_of_split):  # to split up the pdf
    pdf = PdfFileReader(path)  # to read the pdf
    for page in range(pdf.getNumPages()):  # to pull each page
        pdf_writer = PdfFileWriter()  # to give the new pages a new name
        pdf_writer.addPage(pdf.getPage(page))  # to add new pages

        output = f'{name_of_split}{page}.pdf'  # to split it up
        with open(output, 'wb') as output_pdf:  # to make it new
            pdf_writer.write(output_pdf)  # write the new pdf
    for f in glob.glob(r"/Users/joe/Downloads/Thingy/*.pdf"):  # for each file
        shutil.move(f, r"/Users/joe/Downloads/Thingy/filesnew")  # put it into this folder


def rename_pdfs():  # to rename the pdfs
    path = r"/Users/joe/Downloads/Thingy/filesnew"  # path of the files
    excelFile = xlrd.open_workbook(
        r"/Users/joe/Downloads/Thingy/FY 23 Master Sheet.xls")  # the excel file with all the orgs listed
    workSheet = excelFile.sheet_by_index(0)  # to begin at the initial position
    fileNum = 0  # to begin at the initial number

    for rownum in range(workSheet.nrows):  # for each row in the excel
        inv = workSheet.cell(rownum, 0).value  # pull the word in the excel
        newFilename = inv  # new variable
        os.rename(os.path.join(path, str(fileNum) + ".pdf"),  # rename the file
                  os.path.join(r"/Users/joe/Downloads/Thingy/filesnew",
                                newFilename + " " + "-" + " " + "SAFAC FY '23 Info Letter" + ".pdf"))  # move the new file to a new folder
        fileNum += 1  # to move on to the next row and file


if __name__ == "__main__":  # to begin the code
    path = r"/Users/joe/Downloads/Thingy/filesnew/FY23 Prelim Award Email Template.pdf"  # whatever the giant file is called
    split(path, '')  # to put a space in between and split the files
    time.sleep(2)  # without this, the names are not changed
    rename_pdfs()  # rename the files

# By Vanessa Papadimitriou
# Iterative Version 2.0 by Stephen A.(The Professor)
# Since I wrote this, I would like to keep my name here for future people to read, thanks!