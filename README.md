ExcelToCsv
==========

Sample Fault Tolerant Java Utility Class to convert one or more Excel Files (XLS/XLSXs) to csv.

Seems like I go thought this sort of exercise every few years - need to convert an excel spread sheet to comma separated format for subsequent processing.
Two basic approaches - read the spreadsheet itself without using Excel (this approach) or rely upon an Excel installation and essentially write a small 
script that uses OLE automation and opens the xls/clsx, and does a "File-> Save As."  This project is an example of the first approach.


See also:
1)  http://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/ss/examples/ToCSV.java (example from POI Site)
2)  https://gist.github.com/991207
