chcp  65001
$Archive = Get-ChildItem -Path C:\Users\fokin_ok\Downloads -Filter "Diadoc.Documents*" #| Where-Object -Property {$_.LastWriteTime -like "*07.02.2020*"} 
$Date =  Get-date -Format "d MMMM yyyy"
$Month = Get-Date -Format "MMMM.yyyy"

#Expand-Archive -Path $Archive.FullName -DestinationPath C:\Reestr\$Month 
$shell = new-object -com shell.application
$zip = $shell.NameSpace($Archive.FullName)
foreach($item in $zip.items())
{
$shell.Namespace("C:\Reestr\$Month").copyhere($item)
}

New-Item -ItemType Directory C:\Reestr\$Month\Формы\
Get-ChildItem -Path C:\Reestr\$Month -Recurse -Include "*.pdf" -Exclude "*Формы" | ForEach-Object {

 #$_.Name  
if($_.Name -like "Печатная*")
{
    Copy-Item $_.FullName -Destination C:\Reestr\$Month\Формы -Force
    $Names = $_.Name -split "форма"
    $Grep = $Names[1]
}
elseif($_.Name -like "Протокол*" -and $_.Name -notlike "*sgn*")
{
     
    $NameP = $_.FullName -replace ".pdf",$Grep
    Rename-Item -Path $_.FullName -NewName $NameP
    Copy-Item $NameP -Destination C:\Reestr\$Month\Формы -Force
    #$NameP
}
else
{
}  

} 
