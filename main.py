import sys # for argv
import os  # if you have a macbook
import xlrd  # to access excel docs
from PyPDF2 import PdfFileReader, PdfFileWriter  # to read pdfs
import glob  # to do paths
import shutil  # to be able to do many operations
import time # to slow it down to complete the code

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

# Joe Scherrek Fixes:
# global version is now suppported which requires the following arguments 
#           python3 main.py <master pdf path> <new pdf folder path> <master sheet path>
# paths are no longer dependent on the OS just pass the relative paths to each. This is always so difficult!
# added error checking to make debugging easier in the future :)
# The sheet needs to be laid out the same way as the Excel file in this example and needs to be a .xls.
# If the generation fails delete all generated pdfs files in the folder and try again.

# For another update I suggest changing workbook.sheet_by_index() to another function in the xlrd library
# Looks at the spreedsheet pages names and it will look at the right one.
# The next update should also use the get_sheet_data() function to add get the infomation from the spreadsheet
# and then generate a new pdf based on the data.

relative_path = os.path.dirname(__file__) # path to current script for relative path
master_pdf_path = os.path.realpath(os.path.join(relative_path,sys.argv[1])) # first arg should be relative path for master pdf this creates absoulute path to pdf 
new_files_path = os.path.realpath(os.path.join(relative_path,sys.argv[2])) # folder to store new pdfs
master_sheet_path = os.path.realpath(os.path.join(relative_path,sys.argv[3])) # path to master sheet 

# arrays for master sheet data collection
orgs = []
prelim_amount = []
applied_grades = []
actual_grades = []

try: # master pdf
    pdf = PdfFileReader(master_pdf_path)  # to read the pdf
except Exception as e:
    print("ERROR: Could not open: " + master_pdf_path +".\nReturned:\n")
    print(e)
    print("\n\nCLOSING...")
    sys.exit(1)

try: # master sheet
    workbook = xlrd.open_workbook(master_sheet_path) # the excel file with all the orgs listed
except Exception as e:
    print("ERROR: Opening Master Sheet spreadsheet.\nReturned:\n")
    print(e)
    print("\n\nCLOSING...")
    sys.exit(1)

def split(name_of_split):  # to split up the pdf
    try:
        for page in range(pdf.getNumPages()):  # to pull each page
            pdf_writer = PdfFileWriter()  # to give the new pages a new name
            pdf_writer.addPage(pdf.getPage(page))  # to add new pages

            output = f'{name_of_split}{page}.pdf'  # to split it up
            with open(output, 'wb') as output_pdf:  # to make it new
                pdf_writer.write(output_pdf)  # write the new pdf
    except Exception as e:
        print("ERROR: Could not split pdfs.\nReturned:\n")
        print(e)
        print("\n\nCLOSING...")
        sys.exit(1)
    
    try:
        for f in glob.glob(os.path.join(relative_path,r"*.pdf")):  # for each file
            shutil.move(f,new_files_path)  # put it into this folder
    except Exception as e:
        print("ERROR: Moving pdfs to " + new_files_path + "failed.\nReturned:\n")
        print(e)
        print("\n\nCLOSING...")
        sys.exit(1)

def rename_pdfs():  # to rename the pdfs    
    try: 
        fileNum = 0  # to begin at the initial number
        for org in orgs:  # for each row in the excel rename the pdf with orgname
            print(org)
            os.rename(os.path.join(new_files_path, str(fileNum) + ".pdf"),  # rename the file
                    os.path.join(new_files_path,
                                    org + " - SAFAC FY '23 Info Letter.pdf"))  # move the new file to a new folder
            fileNum += 1  # to move on to the next row and file
    except Exception as e:
        print("ERROR: Renaming PDF: " + str(fileNum) + ".pdf.\nReturned:\n")
        print(e)
        print("\n\nCLOSING...")
        sys.exit(1)

def get_sheet_data():   # pull data from master sheet and add to arrays
    try:                # this can create problems in the future if the formatting of the worksheet changes
        sheet = workbook.sheet_by_index(0) # open desired sheet in workbook
        for col in range(sheet.nrows - 1):  # its -1 because of the titles in the first row
            org = str(sheet.cell_value(col + 1, 2)) # grab org name
            prelim = str(sheet.cell_value(col + 1, 28))
            app = str(sheet.cell_value(col + 1, 3)) # grab applied grade for org
            act = str(sheet.cell_value(col + 1, 4)) # grab actual grade from sheet
            orgs.append(org) # append to arrays accordingly
            prelim_amount.append(prelim)
            applied_grades.append(app)
            actual_grades.append(act)
        # print(str(orgs) + "\n" + str(prelim_amount) + "\n" + str(applied_grades) + "\n" + str(actual_grades))
    except Exception as e:
        print("ERROR: Getting data from master sheet.\nReturned:\n")
        print(e)
        print("\n\nCLOSING...")
        sys.exit(1)

if __name__ == "__main__":  # to begin the code
    if sys.argv[1] == "" or sys.argv[2] == "" or sys.argv[3] == "": 
        print("Error: Paths to files not found! \nPlease specify locations of the master sheet, master pdf, and the new pdf folder.")
        sys.exit(1)
    print("Welcome!\nMaster PDF: " + master_pdf_path + "\nPDF's location: " + new_files_path + "Master Sheet: " + master_sheet_path + "\n\nGetting sheet data...")
    get_sheet_data()
    print("DONE.\n\nSplitting files...")
    split('')  # to put a space in between and split the files
    time.sleep(2)  # without this, the names are not changed
    print("DONE.\n\nRenaming Files...")
    rename_pdfs()  # rename the files
    print("DONE.\nSUCCESS: PDF's generated successfully.")

# By Vanessa Papadimitriou
# Iterative Version 2.0 by Stephen A.(The Professor)
# Global Version 3.0 by Joseph Scherreik
# Since I wrote this, I would like to keep my name here for future people to read, thanks!
