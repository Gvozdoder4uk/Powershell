# Выгрузка в CSV списка активных ПК
Get-ADComputer -Filter {Name -Like "W00-*"} -Properties Description |
Where-Object {$a=$_.name; $_.DistinguishedName -ne "CN=$a,OU=Computers,OU=Disabled,DC=rusagrotrans,DC=ru"} |
Sort-Object NAME | Select-Object NAME,DESCRIPTION | Export-csv -NoTypeInformation C:\TEST\LOCAL.csv -Encoding UTF8

# Определение группы
$GroupName = "Администраторы"

# Подруб CSV
$ImportCsv = import-csv c:\Test\LOCAL.csv

"===================================" | out-file -Append "C:\Test\Local_Admins.txt" 

# Тело программы, получение списка и выгрузка в TXT
$ImportCsv | ForEach-Object {
$a= $_.Name
    if ((Test-connection $a -count 2 -quiet) -eq "True")
    {
        $CompName[$i]
        $Computer = $a
        $a | out-file -Append "C:\test\Local_Admins.txt"
        $qwe = ([ADSI]"WinNT://$Computer/$GroupName").psbase.invoke("members") |%  {$_.GetType().InvokeMember("Name", ‘GetProperty’, $null, $_, $null)}
        for ($j=0; $j -lt $qwe.count;$j++){
                $qwe[$j] | out-file -Append "C:\Test\Local_Admins.txt"
        
        
    }
    "===================================" | out-file -Append "C:\Test\Local_Admins.txt"
}
    }