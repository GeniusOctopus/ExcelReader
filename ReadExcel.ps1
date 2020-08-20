#Pfad zur Exceldatei fest
$strFileName = "C:\example\file.xlsx"
#Name der Tabelle
$strSheetName = "Tabelle1$"
#Der Provider für den Datenbankzugriff
#Muss eventuell erst installiert werden
#Installer: https://download.microsoft.com/download/0/6/A/06AD225D-42E3-4A58-A35C-71CCF694C9C6/AccessDatabaseEngine_X64.exe
$strProvider = "Provider=Microsoft.ACE.OLEDB.12.0"
#Datenquelle für die Datenbankverbindung
$strDataSource = "Data Source = $strFileName"
#Es handelt sich um eine Exceldatei
$strExtend = "Extended Properties=Excel 8.0"
#SQL Statement
$strQuery = "Select * from [$strSheetName]"
 
#Instanziieren eines neuen Objektes des Typs 'System.Data.OleDb.OleDbConnection' mit den vorher initialisierten Variablen
#Hier wird die Verbindung zur Datenbank erzeugt
#Benötigt werden der OLEDB Provider, die Datenquelle und der Verweis, dass es sich um Excel handelt
$objConn = New-Object System.Data.OleDb.OleDbConnection("$strProvider;$strDataSource;$strExtend")
#Instanziieren eines neuen Objektes des Typs 'System.Data.OleDb.OleDbCommand'
#Hier wird ein neues SQL Statement als Objekt erzeugt, welches die Abfrage als String übergeben bekommt
$sqlCommand = New-Object System.Data.OleDb.OleDbCommand($strQuery)
#Setzt die Verbindung zur Datenbank für das SQL Statement
$sqlCommand.Connection = $objConn

#Öffnet die Datenbankverbindung
$objConn.open()

#Ausgabe auf der Konsole
Write-Host("Lese Exceldatei...")

#Instanziieren eines neuen Objektes des Typs 'System.Data.DataTable' mit dem frei wählbaren Tabellennamen 'Exceldatei'
$table = New-Object System.Data.DataTable 'Exceldatei'
#Instanziieren eines neuen Objektes des Typs 'System.Data.OleDb.OleDbDataAdapter'
#Der DataAdapter liest dann die Daten
$sqlDataAdapter = New-Object System.Data.OleDb.OleDbDataAdapter
#Setzt das vorher instanziierte OleDbCommand Objekt als Kommando für den DataAdapter
$sqlDataAdapter.SelectCommand = $sqlCommand
#DataAdapter füllt das DataTable Objekt
$sqlDataAdapter.Fill($table)
#Schließt die Datenbankverbindung
$objConn.close()
