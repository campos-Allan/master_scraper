# master_scraper

## Disclaimer
It’s not possible to share the PDF and Excel files that this script was using, so I made some examples files to give a demonstration, they can be found in the file's folder. I couldn't make good examples to be used for 2 of the files, so I just commented part of the code and created a variable with what the code would scrape from the PDF. Also made names anonymous, so general understanding of the context may be a little harder.

## Previously
![Final](https://i.imgur.com/KUrXufj.png)
I first made this script early this year, and it was was GUI-heavy, with lots of clicks and new windows, and could only run with files from one day, so I just made a new version with a lot less GUI implementation to copy the data quicker into the shared spreadsheet, and made it possible to run a lot of files from different days.

## Task
Update a shared spreadsheet daily, with data that comes from 5 PDF files and 5 Excel spreadsheets. This update was based on a scheme of product transportation (products X1, Y1 and Y2), monitoring storage and product transportation trips between a few storage spots. Between two specific spots, transport could be done by two modals, A and B, the rest of the spots could only transport in modal A.
![Scheme](https://i.imgur.com/iiXjSai.png)

The PDF files were very hard to read due to bad formatting and lack of standardization, suddenly changing little things in its format (and sometimes going back on those changes a few days later, I had no control over this). So the code would have to be easy to change, that's also why I made this updated version, relying more on pandas power to manage data, and less in scraping big text strings, with a lot of exceptions and conditions, like the first version.

## Approach
There's still a GUI.py, but now there's only a button to read instructions and other to run main.py, so people that don't work with code can run the script without worries. main.py will get variables and file location from file.py and var.py, these variables are structured as DataFrames containg the columns that are also in a model spreadsheet called 'modelo.xlsx', the final place where all the info is going, so the user can easily copy from that spreadsheet into the shared one. This Excel file has sheets for every storage location and one for the transportation trips.

Then main.py starts T1_reader for reading storage files and trips from T1 (these trips had to be made into a variable with the content as I could not replicate the style of the original file), every time the scrip gets a file to read, it annotates its day of reference, so it can organize the day each information is coming from.  T1_reader also reads storage info from T2, as the file is very similar to T1's. Going back to main.py, the info passes through some formatting and transformation. The info from the trips have to be divided into two types of modal (A and B), to enter in different columns of the spreadsheet, as T1's files don't make this distinction.

Finally, main.py calls for T1_insert_reader to read how ongoing trips to T2 ended. This part of the code also has to calculate which trip ended per which modal, getting the total per modal. This part of the code would read the PDF's, but I had the same trouble as trips coming from T1, so just made a variable with the info I need to show how it would extract. Then it opens the model Excel file, to write destination info in the trips that were ongoing but now ended to reach T2.

For the Excel files, first the code reads trips from T2 to S1 and R1 to R2, as they come in the same format. Trips to S1 from T2 had to be deducted separately from the storage info, as they don't detail that there, and the shared spreadsheet needs this info (which portion of trips from T2 went to S1). Then, it ends reading the spreadsheet from R2 with storage and trips from R1, and S1 files for storage and trips from T2. Every time the code looks at ending ongoing trips, it makes the same process described in the end of the last paragraph. Finally, the files as put in a trash folder, so the next time the program runs, there's no repetition of info going into the model spreadsheet.

## Result
Every white cell is what already was in the sheet, orange cells are what departed from T1 to T2, green cells what departed from R1 to R2, and blue cells what departed from T2 to S1. Grey cells are info from trips that ended on day 13/12 (T1>T2, R1>R2 and T2>S1), purple cells the same but for the 14th day of December.

![Results](https://i.imgur.com/9tgvQ5d.png)

Each storage spot has a sheet with all the storage info extracted from the files and divided per day, now it's only a matter of copying the info into the shared master spreadsheet.

![Results2](https://i.imgur.com/JSGzAlR.png)




