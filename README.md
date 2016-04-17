# Big-Data-Charts
Optimises Excel Charts for Big Data (VBA Excel Add-In)

![Alt text](/screenshot01.jpg?raw=true "Big Data Charts")

Charts for Big Data is a chart optimiser for big databases and is:
  - Fast: charts are light, fast, and optimised for Big Data
  - Memory-efficient: the file size of the spreadsheet is reduced by up to 80%!
  - Powerful: Optimised charts can handle millions of rows of data without any slow down
  - Compatible: with Excel 2007, 2010, 2013

--------INSTALLATION---------
In Excel, go to File > Options > Add-Ins.
Select "Manage Excel Add-ins" and click "Go".
Click "Browse", select the file "Charts for Big Data - Excel Add-in.xlam", and click "OK".
You should see a new "Charts for Big Data" menu in Excel now!



--------HELP-----------------
1) Create some charts with data and leave the first two columns of the charts' sheet free.
2) Click "Setup and Optimise Charts", select the charts to optimise and click OK. 
3) Check the newly created blocks of values [Chart #, DateMin#, DateMax#] at the top of the first two columns of the active sheet. Change the values of DateMin and DateMax and click "Refresh Range" to fine-tune the ranges of the charts. 
4) "Print" the selected charts to the default printer.


Setup and Optimise Charts:
Optimises the selected charts to dynamic charts for fast performance thanks to named ranges and blocks of range values. When possible, the data series of the charts are recreated and converted to named ranges (CBD_*) adjusted to the minimum ranges needed by the charts. As a result, the charts are light, memory-efficient and fast to work with. 
The Blocks of Range Values [Chart #, DateMin# and DateMax#] are automatically created at the top of the first two columns of the active sheet. Each chart is linked to its block of range values and can be easily adjusted.

DateMin#:
Minimum date of Chart #'s range. 

DateMax#:
Maximum date of Chart #'s range.

Refresh Range:
Applies DateMin and DateMax values to the selected charts - recreates the blocks of range values if needed -, and readjusts the dynamic named ranges of the charts. Please always click Refresh Range after changing DateMin or DateMax values.

Print:
Prints the selected charts to the default printer.


TROUBLESHOOTING:
If any data problems, reselect the data for your chart ("Chart tools > Select Data") and click "Setup and Optimise Charts" again.
If you can't see the blocks of values, they are located at the top of the first two columns of the charts' sheet. Please note that they can be hidden by a chart if the chart is over the first two columns.
