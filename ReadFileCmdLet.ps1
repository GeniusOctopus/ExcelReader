#read the excel file
function Get-ExcelFileContent
{
    [OutputType([System.Data.DataTable])]
    [CmdLetBinding()]
    param()

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
    }
    PROCESS
    {
        #establish new databaseconnection
        $objConn.open()

        #give some output for the user
        Write-Host("Lese Exceldatei...")

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
