<h1>This directory contains a small file that contains a number of hash colision for Excel versions before 2010 (Worksheet and Workbook protection)</h1>
<br/>First column: The password
<br/>Second columns: Hash in 64" format
<br/>Third columns: Hash in hexadecimal format (the same as in the .xml file)

<h1>Where to find the hash of an Excel file ?</h1>
<br/>Create a copy
<br/>Change the extension to .zip
<br/>Unzip the file and go to: xl/worksheets/
<br/>Open the / a ".xml" file
<br/>Then look for the tag workbookProtection for the binder or sheetProtection for the sheet and find the hash value, for example :<br/>
sheetProtection password="D1E6" sheet="1" objects="1" scenarios="1"
<br/>Here the hash value is "D1E6".
<br/>Find the correspondence in the ExcelHash file (first column).
<br/>Enter this password to unlock the sheet.
<br/><br/>There are many other ways to "crack" an Excel sheet, this file just shows the weakness of this type of protection.
