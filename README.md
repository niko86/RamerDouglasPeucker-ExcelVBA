# RamerDouglasPeucker-ExcelVBA
I was frustrated by the lack of a fast excel implementation of the Ramer–Douglas–Peucker algorithm, so after quite a few days of getting lost in recursion I made my own. My work is based on the pseudocode presented on the wikipedia article https://en.wikipedia.org/wiki/Ramer%E2%80%93Douglas%E2%80%93Peucker_algorithm.
My code requires three things to be present in an excel file; though you can obviously edit the code to fit your worksheet.
1. A named range called "epsilon",
2. An excel table on "Sheet1" named "Table1"
3. An excel table on "Sheet2" named "Table2"
You can plot the values of Tables 1 and 2 on the same chart and use the count function to determine the table lengths to get a measure on the amount of data reduction.
