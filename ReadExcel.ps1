$strFileName = "C:\example\file.xlsx"
$strSheetName = "Tabelle1$"
#You may have to install the provider first
#Installer: https://download.microsoft.com/download/0/6/A/06AD225D-42E3-4A58-A35C-71CCF694C9C6/AccessDatabaseEngine_X64.exe
$strProvider = "Provider=Microsoft.ACE.OLEDB.12.0"
$strDataSource = "Data Source = $strFileName"
$strExtend = "Extended Properties=Excel 8.0"
#SQL Statement
$strQuery = "Select * from [$strSheetName]"
 
#Instantiate System.Data.OleDb.OleDbConnection to establish a db connection later
$objConn = New-Object System.Data.OleDb.OleDbConnection("$strProvider;$strDataSource;$strExtend")
#Instantiate System.Data.OleDb.OleDbCommand with you query
$sqlCommand = New-Object System.Data.OleDb.OleDbCommand($strQuery)
$sqlCommand.Connection = $objConn

$objConn.open()

Write-Host("Querying file...")

#Instantiate System.Data.DataTable
$table = New-Object System.Data.DataTable 'Excelfile'
#Instantiate Data-Adapter to read data
$sqlDataAdapter = New-Object System.Data.OleDb.OleDbDataAdapter
#Set Command
$sqlDataAdapter.SelectCommand = $sqlCommand
#Fill DataTable
$sqlDataAdapter.Fill($table)
#dont forget to close the connection
$objConn.close()
