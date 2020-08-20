<#
.SYNOPSIS
    This function reads an excel file.

.DESCRIPTION
    A databaseconnection gets established and through a SQL statement you query the Excel File.
    The function returns an object of type System.Data.DataTable

.PARAMETER strProvider
    The OleDb Provider Microsoft.ACE.OLEDB.12.0

.PARAMETER strDataSource
    Basically the path to the file

.PARAMETER strExtend
    The information that it is an Excel File

.PARAMETER strQuery
    The SQL statement

.EXAMPLE
    The example below shows how to call the function
    PS C:\> Get-ExcelFileContent $strProvider $strDataSource $strExtend $strQuery

.Example
    This example shows example parameters
    PS C:\> Get-ExcelFileContent "Provider=Microsoft.ACE.OLEDB.12.0" "Data Source = C:\Users\Me\Desktop\example.xlsx "Extended Properties=Excel 8.0" "Select * from [Tabelle1$]"

#>
function Get-ExcelFileContent
{
    [OutputType([System.Data.DataTable])]
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [string]
        $strProvider,

        [Parameter(Mandatory=$true)]
        [string]
        $strDataSource,

        [Parameter(Mandatory=$true)]
        [string]
        $strExtend,

        [Parameter(Mandatory=$true)]
        [string]
        $strQuery
    )

    BEGIN
    {
        #instatntiate a new object of type 'System.Data.OleDb.OleDbConnection' using the provided variables
        #you have to pass OleDb Provider, the Datasource, which is basically the path and specify that it is an Excel File
        #this is the connection to the database
        $objConn = New-Object System.Data.OleDb.OleDbConnection("$strProvider;$strDataSource;$strExtend")

        #instatntiate a new object of type 'System.Data.OleDb.OleDbCommand'
        #you have to pass the previous initialized SQL statement
        #this is the command
        $sqlCommand = New-Object System.Data.OleDb.OleDbCommand($strQuery)

        #set the connection for the command
        $sqlCommand.Connection = $objConn
        
        try
        {
            #establish new databaseconnection
            $objConn.open()
        }
        catch
        {
            throw("Es ist ein Fehler aufgetreten. Die Excel Datei konnte nicht geöffnet werden. Bitte überprüfen Sie den Dateipfad und stellen Sie sicher, dass die Datei nciht geöffnet ist.")
        }

        #give some output for the user
        Write-Host("Lese Exceldatei...")
    }
    PROCESS
    {
        #instatntiate a new object of type 'System.Data.DataTable' with the variable tablename 'Exceldatei'
        $table = New-Object System.Data.DataTable 'Exceldatei'

        #instatntiate a new object of type 'System.Data.OleDb.OleDbDataAdapter'
        #the DataAdapter reads the data from the file
        $sqlDataAdapter = New-Object System.Data.OleDb.OleDbDataAdapter

        #set the command for the DataAdapter
        $sqlDataAdapter.SelectCommand = $sqlCommand
    
        #the DataAdapter fills the DataTable object
        $sqlDataAdapter.Fill($table)
    }
    END
    {
        #close the databaseconnection
        $objConn.close()

        #return the DataTable object
        return $table
    }
}
