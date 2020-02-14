$Excel = New-Object -ComObject Excel.Application

# Делаем его видимым
$Excel.Visible = $true

# Добавляем рабочую книгу
$WorkBook = $Excel.Workbooks.Add()

$Sheet = $WorkBook.Worksheets.Item(1)
$Sheet.Name = 'Инвентаризация ЦО'
$Sheet.Cells.Item(1,1) = 'Имя Контейнера'
$Sheet.Cells.Item(1,2) = 'Владелец'
$Sheet.Cells.Item(1,3) = 'Аккаунт'
$Sheet.Cells.Item(1,4) = 'Полный Путь'

$Row = 2
$Column = 1

Get-ADOrganizationalUnit -SearchBase 'OU=Русагротранс,OU=Users,OU=MSK,DC=rusagrotrans,DC=ru' -Filter 'Name -like "*"' | Select-Object Name, ManagedBy,DistinguishedName | ForEach-Object `
{ 
    $Sheet.Cells.Item($Row,$Column) = $_.Name
    $Column++
    if($_.ManagedBy -eq $Null)
    {
    $Column++
    $Sheet.Cells.Item($Row,$Column) = $_.ManagedBy
    $Column++
    }
    else
    {
    
    $var = $_.ManagedBy 
    $L = $var.Length
    $var = $var.Substring(3,$L-4)
    $var = $var.Split(",OU")[0]
    $Login = Get-AdUser -Filter "Name -like '$var*'" -SearchBase 'OU=Русагротранс,OU=Users,OU=MSK,DC=rusagrotrans,DC=ru'

    $Sheet.Cells.Item($Row,$Column) = $var
    $Column++
    $Sheet.Cells.Item($Row,$Column) = $Login.SamAccountName
    $Column++
    }
    $Sheet.Cells.Item($Row,$Column) = $_.DistinguishedName
    $Row++
    $Column=1
    
}
$Range = $Sheet.Range("A1","D1")
$Range.EntireColumn.AutoFit()