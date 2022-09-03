# SAFAC PDF Converter

## There are three varibles that are requried to run the script:

### 1. Master PDF path

- Location of main PDF to break up

### 2. New PDF folder

- Location for new generated PDFs

### 3. Master Sheet path

- Location of master spreadsheet to get org names
- File must be an .xls file .xlsx will not work

# To Run

Make sure python is installed on your machine and added to the PATH. Go into `cmd` on Windows or `Terminal` in MacOS. Make sure you `cd` into this folder and run the command below:
> `python3 main.py <master pdf path> <new pdf folder path> <master sheet path>`

### Windows example:

> `> cd C:\Users\<your user>\Documents\safac-pdf-converter`
>
> `> python3 main.py '.\FY 23 Master Sheet.xls' filesnew '.\filesnew\FY 23 Master Sheet.xls'`

### MacOS example:

> `$ cd ~/Documents/safac-pdf-converter`
>
> `$ python3 main.py "'./FY 23 Master Sheet.xls'" filesnew "'./FY 23 Master Sheet.xls'"`

All the PDF should now be located in your New PDF folder. If the converstion fails make sure you delete the contents of your New PDF folder and start over again.
