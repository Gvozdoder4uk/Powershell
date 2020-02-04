$Config_File = "C:\Test\cfg.ini"

Get-Content $Config_File| foreach-object -begin {$START=@{}} -process { $k = [regex]::split($_,'='); if(($k[0].CompareTo("") -ne 0) -and ($k[0].StartsWith("[") -ne $True)) { $START.Add($k[0], $k[1]) } }
$START

$Configuration_Start = $START.Programm_Mode
$Configuration_Start