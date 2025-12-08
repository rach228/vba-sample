# Clean Vulnerability Reports using VBA Macros
To automate the creation of pivot tables for CVE analysis, you can use the sample VBA macros with the following steps:

First, download the vulnerability report for the target cluster from Red Hat Advanced Cluster Security or from the scheduled email reports. Convert the vulnerability report from CSV format to XLSM format to enable macros. Copy the VBA macro code from `create_pivot_table.txt`.

Second, open Excel and navigate to the developer tab. If the developer tab is not visible, right-click on the ribbon, select “customize the ribbon,” and tick the checkbox for “developer.”

Once the developer tab is visible, click on “Visual Basic” to open the VBA editor. Create a new module within your workbook and paste the copied VBA code.

Locate the section of the code that defines the `wsData` variable and update it to match the name of your worksheet. For example:
`wsData = wb.Sheets("RHACS_Vulnerability_Report_Work")`

The default name for your Excel sheet in the vulnerability report is usually something like `RHACS_Vulnerability_Report_<first 4 letters of the report name>`. Make sure this matches the actual name in your Excel file. Save the file.

Once the macro is set up, return to your Excel sheet, click on “macros,” select the macro you just created, and run it. The macro will generate the pivot table automatically, allowing you to filter and visualize the CVE data efficiently.

Refer to this [video](https://youtu.be/Ian18xhB0Xc?si=OTPPdOKy4g4xeN5T) for reference.
