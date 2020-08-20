# ExcelReader

This PowerShell script uses OleDb to query a Excel file using SQL statements.
You can use the Cmd-Let by calling:
```powershell
Get-ExcelFileContent
```
The function returns an object of type System.Data.DataTable.

You may have to install the OleDb Provider "Provider=Microsoft.ACE.OLEDB.12.0" first.
