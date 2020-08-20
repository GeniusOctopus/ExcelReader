# ExcelReader

This PowerShell script uses OleDb to query a Excel file using SQL statements.
You can use the Cmd-Let by calling:
```powershell
Get-ExcelFileContent $strProvider $strDataSource $strExtend $strQuery
```
The function returns an object of type System.Data.DataTable.

You may have to install the OleDb Provider "Provider=Microsoft.ACE.OLEDB.12.0" first.

Check for your providers by using this piece of code:
```powershell
$list = New-Object ([System.Collections.Generic.List[psobject]])
foreach($provider in [System.Data.OleDb.OleDbEnumerator]::GetRootEnumerator())
{
    $p = New-Object psobject
    for ($i = 0; $i -lt $provider.FieldCount; $i++)
    {
        Add-Member -in $p NoteProperty $provider.GetName($i) $provider.GetValue($i)
    }
    $list.Add($p)
}
$list
```

If it isn't installed get it [here](https://download.microsoft.com/download/0/6/A/06AD225D-42E3-4A58-A35C-71CCF694C9C6/AccessDatabaseEngine_X64.exe).
