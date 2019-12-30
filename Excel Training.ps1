# Созадём объект Excel
$Excel = New-Object -ComObject Excel.Application

# Делаем его видимым
$Excel.Visible = $true

# Добавляем рабочую книгу
$WorkBook = $Excel.Workbooks.Add()

$LogiclDisk = $WorkBook.Worksheets.Item(1)
$LogiclDisk.Name = "Логические диски"

$LogiclDisk.Cells.Item(1,1) = 'Буква диска'
$LogiclDisk.Cells.Item(1,2) = 'Метка'
$LogiclDisk.Cells.Item(1,3) = 'Размер (ГБ)'
$LogiclDisk.Cells.Item(1,4) = 'Свободно (ГБ)'

# Выделяем жирным шапку таблицы
$LogiclDisk.Rows.Item(1).Font.Bold = $true


$Row = 2
$Column = 1

# ... и заполняем данными в цикле по логическим разделам
Get-WmiObject Win32_LogicalDisk | ForEach-Object `
{
    # DeviceID
    $LogiclDisk.Cells.Item($Row, $Column) = $_.DeviceID
    $Column++
    
    # VolumeName
    $LogiclDisk.Cells.Item($Row, $Column) = $_.VolumeName
    $Column++
    
    # Size
    $LogiclDisk.Cells.Item($Row, $Column) = ([Math]::Round($_.Size/1GB, 2))
    $Column++
    
    # Free Space
    $LogiclDisk.Cells.Item($Row, $Column) = ([Math]::Round($_.FreeSpace/1GB, 2))
    
    # Переходим на следующую строку и возвращаемся в первую колонку
    $Row++
    $Column = 1
}


# Выравниваем для того, чтобы их содержимое корректно отображалось в ячейке
$UsedRange = $LogiclDisk.UsedRange
$UsedRange.EntireColumn.AutoFit() | Out-Null
$WorkBook.Worksheets.Add()
$PhysicalDrive = $WorkBook.Worksheets.Item(1)

# Переименовываем лист
$PhysicalDrive.Name = 'Физические диски'

# Заполняем ячейки - шапку таблицы
$PhysicalDrive.Cells.Item(1,1) = 'Модель'
$PhysicalDrive.Cells.Item(1,2) = 'Размер (ГБ)'
$PhysicalDrive.Cells.Item(1,3) = 'Кол-во разделов'
$PhysicalDrive.Cells.Item(1,4) = 'Тип'

# Переходим на следующую строку...
$Row = 2
$Column = 1

# ... и заполняем данными в цикле по физическим дискам
Get-WmiObject Win32_DiskDrive | ForEach-Object `
{
    # Model
    $PhysicalDrive.Cells.Item($Row, $Column) = $_.Model
    $Column++
    
    # Size
    $PhysicalDrive.Cells.Item($Row, $Column) = ([Math]::Round($_.Size /1GB, 1))
    $Column++
    
    # Partitions
    $PhysicalDrive.Cells.Item($Row, $Column) = $_.Partitions
    $Column++
    
    # InterfaceType
    $PhysicalDrive.Cells.Item($Row, $Column) = $_.InterfaceType
    
    # Переходим на следующую строку и возвращаемся в первую колонку
    $Row++
    $Column = 1
}

# Выделяем жирным шапку
$PhysicalDrive.Rows.Item(1).Font.Bold = $true

# Выравниваем для того, чтобы их содержимое корректно отображалось в ячейке
$UsedRange = $PhysicalDrive.UsedRange
$UsedRange.EntireColumn.AutoFit() | Out-Null


$WorkBook.Worksheets.Add()
$MY = $WorkBook.Worksheets.Item(1)
$My.Name = "ТЕСТОВАЯ СТРАНИЦА"



$My.Cells.Item(1,1) = 'Имя процесса'
$My.Cells.Item(1,2) = 'PID'
$My.Cells.Item(1,3) = 'CLASS %'

$PhysicalDrive.Rows.Item(1).Font.Bold = $true

$Row = 2
$Column = 1



Get-WmiObject -Class Win32_Process | ForEach-Object `
{
    $my.Cells.Item($Row,$Column) = $_.Name
    $Column++
    $my.Cells.Item($Row,$Column) = $_.Handle
    $Column++
    $my.Cells.Item($Row,$Column) = $_.__SUPERCLASS
    $Column= 1
    $Row++
}

$UsedRange = $My.UsedRange
$UsedRange.EntireColumn.AutoFit() | Out-Null