chevron-parse
=============

Many (if not all) Chevron gas stations use "Blue Cube" software for transaction management inside of the store. Unfortunately, the software does not currently support any sort of analytics on sales (such as gas volumes, etc) beyond rudimentary reports for the entire day.

This script is a one-off solution to help my parents get a better look at their sales so that they can make educated decisions about their prices.

Here's how it works:

1. The first argument to the script is a directory of month directories. The Chevron register software outputs transaction data by month, with a text file for each day containing the sales data. 
2. For each month directory, the script grabs information about gas and carwash sales.
3. The second argument to the script is the destination of an excel spreadsheet file. For each month, a new sheet is created in the spreadsheet file, and each days information is placed in a column.

To run it:
```bash
$ python parse.py <PATH TO DIRECTORY OF DAY REPORT TEXT FILES> <PATH TO EXCEL WORKBOOK FILE>
```
