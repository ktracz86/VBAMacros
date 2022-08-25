# VBAMacros
Various macros for Office apps

Excel Toolkit: This file contains various Macros that can help with different tasks in Excel

Trackers: This is the code that is used in order to merge previous submissions statuses with new ones. 
Certain reports need to be created before hand and formatted in a particular way in order for this to work.
  
OutlookScripts: Scripts required to enable the "auto-sort sent mail" function.
  After users hit Send on an email, it will prompt the user to select which folder the sent mail should be filed away. 
  This is different than the built-in feature, as the built-in feature always saves to the same folder and doesn't give you an option to save in a different
  location with each email. 
  
XMLVerification: This script is used on a particular Excel Sheet. User runs the script, selects a particular eIndex.xml file to check, and the script will validate the eIndex file. 
  
    LIST OF MACROS IN TOOLKIT:
    
    HyperlinkPMRA
      When an excel spreadsheet contains a list of PMRA numbers, this macro allows the officer to
      select the column where the PMRA numbers are, and it will automatically generate a hyperlinks 
      to the file. 
    
    CommaSeparate
      This essentially concatenates all values in a single column, and separates them with a comma. 
      This is useful if running a report that will generate search criteria for another search query.
      Eg.
        Run a Active Source report to determine which products are used a source of active ingredient.  
        Use the results of that report to run a Product report to determine the Class of each product.  
    
    HighlightDifferentRows
      This is useful if the user wants to identify logical breaks in unique data, when there are data 
      represented for a smiilar set. 
      For example, the user wants to visually group all SUBNs based on their REGN values. The user would 
      first sort by REGN number to keep all of the same REGNs together, and then run the script 
      selecting the REGN column as the "unique value" The script will run and highlight all SUBN records 
      of a particular REGN the same colour. The next group of REGNs will have colour 2 and then 
      the following set will revert back to Colour 1. This goes back and forth until all groups are done. 
      Note: Script will not over-ride pre-existing Table Formatting or Conditional Formatting
    
    FindNeonic
      Asks user to identify which column has the list of active ingredient codes.
      If the cell contains either THE, COD, or IMI, then highlights the cell RED.
    
    FormatTableStandard
      Just a quick table format. 
        Creates a grid
        Bolds the headers
        Heavier line width for the bottom of the headers
        Asks if you want to freeze the top row, and does so accordingly.
    
    BulkReplace
      If user wants to replace many different values with other values. Instead of using the Replace feature 
      in excel, they can prepare a table with 2 columns.
      The first column is the lookup value, and the second column is the replaceWith value. 
      Create this reference table on a new sheet in the Workbook and run the script.
      The script will go through every worksheet in the workbook and replace all values that were identified 
      on the source table. 
    
    GetFileNamesDate
      User selects a Folder. The script then goes into that folder and writes all the Filenames, Dates, and KB Sizes 
      into the spreadsheet.
    
    CreateHyperlinkLabelSearch
      Assuming the REGN is on the First Column, it will create a hyperlink to the OLD Label Search, using the value 
      of the first columnn as a Free Text Search.
      

