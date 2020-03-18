if (-not (Get-Module SqlServer)) {
      Import-Module -Name SqlServer
}



$SQL_SET = @{'server'='HQDBT01';'Database'='SQLShackDemo'}
$Paths = Get-ChildItem C:\Users\*  -Directory -recurse

Function SQL_ADD
{
param($Path)
# Data preparation for loading data into SQL table 
$InsertResults = @"
INSERT INTO [SQLShackDemo].[dbo].[tbl_PosHdisk](FullPath)
VALUES ('$Path')
"@      
#call the invoke-sqlcmdlet to execute the query
         Invoke-sqlcmd $SQL_SET -Query $InsertResults
}

foreach ($Path in $Paths) {

    $Path.FullName | Out-File C:\test\Dirs.txt -Append
    SQL_ADD -Path $Path.FullName 
}

