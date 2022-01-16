# Excel file merge
The program merges the first sheets of the selected Excel-files.

Working with: .xlsx

Important:
the program creates a new Excel-file, containing the text from the first sheets of the selected files. The sheet names are taken from the names of the source files.

### Features 
* Clicking the "Merge Excel files" button, opens a dialog box in which you can select the files you want to merge.
* The final file name in the format: "Alarm MBH_H-W {YY} {WW}", where {YY} is the year for last week, {WW} is the number of the previous week.
* Clicking the "Exit Program" button, close the program window.
* Program execution can be monitored through the progress bar.

Dependencies: threading.Thread, pandas, tkinter, os.path, datetime, time, xlrd(ver.1.2.0)

Convert to exe: use pyinstaller or other package for create execution file