chcp  65001
$Archive = Get-ChildItem -Path C:\Users\fokin_ok\Downloads -Filter "Diadoc.Documents*" #| Where-Object -Property {$_.LastWriteTime -like "*07.02.2020*"} 
$Date =  Get-date -Format "d MMMM yyyy"
$DiadocR_Date =  Get-Date -Format "dd.MM.yy"
$Month = Get-Date -Format "MMMM.yyyy"


if(Test-Path -Path C:\Reestr\$Month\Формы\)
{
}
else
{
New-Item -ItemType Directory C:\Reestr\$Month\Формы\
}

$Diadoc_Reesrt = Get-ChildItem -Path C:\Users\fokin_ok\Downloads -Filter "Diadoc $DiadocR_Date*"
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $true
$Workbook = $Excel.Workbooks.Open($Diadoc_Reesrt.FullName)
$Reestr = $WorkBook.Worksheets.Item(1)
$Reestr.Columns.Item(1).Delete()
$Reestr.Columns.Item(1).Delete()
$Reestr.Columns.Item(3).Delete()
$Reestr.Columns.Item(3).Delete()
$Reestr.Columns.Item(5).Delete()
$Reestr.Columns.Item(5).Delete()
$Reestr.Columns.Item(5).Delete()
for($i=1;$i -le 8;$i++)
{
   $Reestr.Columns.Item(6).Delete() 
}
$Range = $Reestr.UsedRange()
$Row = $Range.Rows.Count
$Range.EntireColumn.AutoFit() | Out-Null
$DataRangeInventory = $Reestr.Range(("A{0}" -f 1), ("E{0}" -f $Row))
7..12 | ForEach-Object `
{
    $DataRangeInventory.Borders.Item($_).LineStyle = 1
    $DataRangeInventory.Borders.Item($_).Weight = 2
}
$Reestr.Range("A1","E1").Interior.ColorIndex = 15
$Reestr.Rows.Item(1).Font.Bold = $True
$Excel.Quit()  
 

$Workbook.SaveAs("C:\Reestr\Февраль.2020\Формы\Реестр.csv")
$Workbook.Close("C:\Reestr\Февраль.2020\Формы\Реестр.csv")


#Expand-Archive -Path $Archive.FullName -DestinationPath C:\Reestr\$Month 
$shell = new-object -com shell.application
$zip = $shell.NameSpace($Archive.FullName)
foreach($item in $zip.items())
{
$shell.Namespace("C:\Reestr\$Month").copyhere($item)
}
$i = 1
#New-Item -ItemType Directory C:\Reestr\$Month\Формы\
Get-ChildItem -Path C:\Reestr\$Month  -Recurse -Include "*.pdf"  | ? { $_.FullName -notmatch 'Формы' } |  ForEach-Object {
 #$_.Name  
if($_.Name -like "*Печатная*")
{
    $NewNameD = $_.DirectoryName+"\$i. " + $_.Name 
    Rename-Item -Path $_.FullName -NewName  $NewNameD
    Copy-Item $NewNameD -Destination C:\Reestr\$Month\Формы -Force
    $Names = $_.Name -split "форма"
    $Grep = $Names[1]
}
elseif($_.Name -like "Протокол*" -and $_.Name -notlike "*sgn*")
{
    $NewNameP = $_.Name -replace ".pdf",$Grep
    $NewDirNameP = $_.DirectoryName+"\$i. " + $NewNameP
    Rename-Item -Path $_.FullName -NewName $NewDirNameP
    Copy-Item $NewDirNameP -Destination C:\Reestr\$Month\Формы -Force
    #$NameP
    $i++
}
else
{
}  

} 


Get-ChildItem -Path C:\Reestr\$Month\ -Exclude "Формы" | Remove-Item -Recurse -Force
#Get-ChildItem C:\Reestr\$Month\Формы | ForEach-Object {start-process $_.FullName –Verb Print}
