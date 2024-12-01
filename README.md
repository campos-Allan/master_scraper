# pdf-and-excel-scraping

## Disclaimer
It’s not possible to share the PDF and Excel files that this script was fetching, but I can tell you that the files were not clean and lacked standardization, which made the task more complicated than I initially expected.

## Structure
* `app.py` -> Basic GUI to make running this script more accessible.
* `script_final.py` -> Does the dirty work
  * sap(cod): a bot created with PyAutoGUI to generate spreadsheets within a software used to fetch information. The COD variable changes the region to be searched.
  * excel_write(dic_descarga: dict, action: str): part of the extracted data needed to be inserted into Excel for the first time, while the other part had to be added to already existing information to update it. dic_descarga is a variable with old values and their corresponding updated data. action determines whether the function will only insert new data or check the spreadsheet to update the existing information.
  * pdf_reader(operador: str): reads and extracts data from PDF. Initially, I tried using read_pdf from the Tabula library and converting it to a DataFrame, but some PDFs didn't extract all the information this way, so I had to use PdfReader and search for values in a huge string. In a scenario with more standardized PDFs, this function could be streamlined a lot.
  * excel_reader(operador: str): reads and extracts data from Excel files. I used openpyxl without major issues since the spreadsheets were much more standardized and with 'clean' data for extraction.

## Approach
* sap: clicks on specific areas on the screen and types values to navigate a software and obtain Excel spreadsheets that will be read later. These spreadsheets should be saved in the same folder as the file.
* excel_write: using the action of typing new information, the function searches for the last row of the spreadsheet and pastes the new data extracted from Excel and PDFs, following a certain format. In the action of updating information, a check is first performed to match the data that will be updated afterward.
* pdf_reader: reads more standardized PDFs using the Tabula library, correcting some potential errors and cleaning things up to fit the necessary format. For less standardized PDFs, the PdfReader library is used to search for values in a string. Movements are calculated in the script to display accumulated data based on several conditions (type of modal and region where movements occur).
* excel_reader: reads Excel files more simply than the previous function.

## Results
![Final](https://i.imgur.com/KUrXufj.png)

