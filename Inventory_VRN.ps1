﻿##[Ps1 To Exe]
##
##Kd3HDZOFADWE8uO1
##Nc3NCtDXTlaDjpvW7Do31UrtSXsXZ8aU7Piux47c
##Kd3HFJGZHWLWoLaVvnQnhQ==
##LM/RF4eFHHGZ7/K1
##K8rLFtDXTiW5
##OsHQCZGeTiiZ4NI=
##OcrLFtDXTiW5
##LM/BD5WYTiiZ4tI=
##McvWDJ+OTiiZ4tI=
##OMvOC56PFnzN8u+Vs1Q=
##M9jHFoeYB2Hc8u+Vs1Q=
##PdrWFpmIG2HcofKIo2QX
##OMfRFJyLFzWE8uK1
##KsfMAp/KUzWJ0g==
##OsfOAYaPHGbQvbyVvnQX
##LNzNAIWJGmPcoKHc7Do3uAuO
##LNzNAIWJGnvYv7eVvnQX
##M9zLA5mED3nfu77Q7TV64AuzAgg=
##NcDWAYKED3nfu77Q7TV64AuzAgg=
##OMvRB4KDHmHQvbyVvnQX
##P8HPFJGEFzWE8tI=
##KNzDAJWHD2fS8u+Vgw==
##P8HSHYKDCX3N8u+Vgw==
##LNzLEpGeC3fMu77Ro2k3hQ==
##L97HB5mLAnfMu77Ro2k3hQ==
##P8HPCZWEGmaZ7/K1
##L8/UAdDXTlaDjpvb9TF58UT8W1SbnHVFCFYXGjFhTz9cbRnve7QrYFphkyfoC1mkF/cKUJU=
##Kc/BRM3KXhU=
##
##
##fd6a9f26a06ea3bc99616d4851b372ba
#############################################
# Soft For Inventarization
#
# Created by Fokin L@B 
#############################################

############################################
#Import Module ACTIVE DIRECTORY
#
import-module ActiveDirectory


# INITIALIZE FOLDER AND FILES #
$INI_FOLDER = "C:\Inventory\Воронеж\5.Инвентаризация Воронеж.xlsx"
$AD_GREP_FILE = "C:\Inventory\Воронеж\VRN_PC.csv"


############################################
# Получаем список ПК из AD

Get-ADComputer -Filter {Name -Like "VRN_*"}  -Properties Description |
Where-Object {$a=$_.name; $_.DistinguishedName -ne "CN=$a,OU=Computers,OU=Disabled,DC=rusagrotrans,DC=ru"} |
Sort-Object NAME | Select-Object NAME,DESCRIPTION | Export-csv -NoTypeInformation "$AD_GREP_FILE" -Encoding UTF8

# Инициализация Конфигурационного Файла:
$Config_File = "C:\Inventory\cfg_filials.ini"
Get-Content $Config_File| foreach-object -begin {$START=@{}} -process { $k = [regex]::split($_,'='); if(($k[0].CompareTo("") -ne 0) -and ($k[0].StartsWith("[") -ne $True)) { $START.Add($k[0], $k[1]) } }
$Configuration_Start = $START.Programm_Mode


Write-Host "Инициализирован режим программы "$Configuration_Start

if($Configuration_Start -eq 0)
{

# Созадём объект Excel
$Excel = New-Object -ComObject Excel.Application

# Делаем его видимым
$Excel.Visible = $true

# Добавляем рабочую книгу
$WorkBook = $Excel.Workbooks.Add()


# Определение стратовой площадки записи в файл.
$Row = 2
$Column = 1
$BadColumn = 1
$BadRow = 2
$InitialRow = 2
$Initial_BadRow = 2






#Основная инвентаризационная страница ЦО
$Archive = $Excel.Worksheets.Add()
$Archive.Name = 'Архив'
$Archive = $WorkBook.Worksheets.Item('Архив')
$Archive.columns.item('i').NumberFormat = "@"
$Archive.Rows.Item(1).HorizontalAlignment = -4108
$Archive.Columns.Item('u').HorizontalAlignment = -4108
$Archive.Columns.Item('w').HorizontalAlignment = -4108
$Archive.Columns.Item('y').HorizontalAlignment = -4108
$Archive.Cells.Item(1,1) = 'Имя Пользователя'
$Archive.Cells.Item(1,2) = 'Сетевое имя'
$Archive.Cells.Item(1,3) = 'Дата Проверки'
$Archive.Cells.Item(1,4) = 'OS'
$Archive.Cells.Item(1,5) = 'Процессор'
$Archive.Cells.Item(1,6) = 'Модель'
$Archive.Cells.Item(1,7) = 'Материнская плата'
$Archive.Cells.Item(1,8) = 'Модель'
$Archive.Cells.Item(1,9) = 'Серийный номер'
#Column HDD Start 8
$Archive.Cells.Item(1,10) = 'HDD 1'
$Archive.Cells.Item(1,11) = 'HDD 2'
$Archive.Cells.Item(1,12) = 'HDD 3'
$Archive.Cells.Item(1,13) = 'HDD 4'
#$Archive.Cells.Item(1,9) = 'Объем (Гб)'
#Column OZY START 14
$Archive.Cells.Item(1,14) = 'Суммарно ОЗУ (Гб)'
$Archive.Cells.Item(1,15) = 'Тип Памяти'
$Archive.Cells.Item(1,16) = 'ОЗУ 1 (ГБ)'
$Archive.Cells.Item(1,17) = 'ОЗУ 2 (ГБ)'
$Archive.Cells.Item(1,18) = 'ОЗУ 3 (ГБ)'
$Archive.Cells.Item(1,19) = 'ОЗУ 4 (ГБ)'
# Column Video Start 20
$Archive.Cells.Item(1,20) = 'Видеокарта 1'
$Archive.Cells.Item(1,21) = 'Объем памяти (MB)'
$Archive.Cells.Item(1,22) = 'Видеокарта 2'
$Archive.Cells.Item(1,23) = 'Объем памяти (MB)'
$Archive.Cells.Item(1,24) = 'Видеокарта 3'
$Archive.Cells.Item(1,25) = 'Объем памяти (MB)'
$Archive.Cells.Item(1,25) = 'Видеокарта 3'
$Archive.Cells.Item(1,26) = 'Объем памяти (MB)'
#Column Network Start
$Archive.Cells.Item(1,27) = 'Cетевая Карта 1'
$Archive.Cells.Item(1,28) = 'MAC'
$Archive.Cells.Item(1,29) = 'Cетевая Карта 2'
$Archive.Cells.Item(1,30) = 'MAC'
#Column Availabilyty  31
$Archive.Cells.Item(1,31) = 'Монитор №1'
$Archive.Cells.Item(1,32) = 'Монитор №2'
$Archive.Cells.Item(1,33) = 'Монитор №3'
$Archive.Cells.Item(1,34) = 'Монитор №4'

$Range = $Archive.Range("A1","AI1")
$Range.WrapText = $True
$Range.AutoFilter() | Out-Null
$Range.Interior.ColorIndex = 15
# Выделяем жирным шапку таблицы
$Archive.Rows.Item(1).Font.Bold = $true

#Страница необработанных ПК
$Bad_PC = $Excel.Worksheets.Add()
$Bad_PC.Name = "Недоступные ПК"
$Bad_PC  = $WorkBook.Worksheets.Item("Недоступные ПК")


$Bad_PC.Cells.Item(1,1) = 'Имя Пользователя'
$Bad_PC.Cells.Item(1,2) = 'Сетевое имя'
$Bad_PC.Cells.Item(1,3) = 'Статус'
$Bad_PC.Cells.Item(1,4) = 'Дата Потери Соединения'
$Bad_PC.Cells.Item(1,5) = 'Дата Восстановления'
$Bad_PC.Cells.Item(1,6) = 'Дата Сканирования'
$Bad_PC.Cells.Item(1,7) = 'Недоступен (Дней)'
$Bad_PC.Columns.Item('D').HorizontalAlignment = -4108
$Bad_PC.Columns.Item('E').HorizontalAlignment = -4108
$Bad_PC.Columns.Item('G').HorizontalAlignment = -4108
$Range = $Bad_PC.Range("A1","G1")
$Range.AutoFilter() | Out-Null
$Range.Interior.ColorIndex = 15
$Bad_PC.Rows.Item(1).Font.Bold = $true
$Bad_PC.Rows.Item(1).WrapText = $true
$Bad_PC.Rows.Item(1).HorizontalAlignment = -4108

#Страница Изменений
$Change_History = $Excel.Worksheets.Add()
$Change_History.Name = "История Изменений"
$Change_History  = $WorkBook.Worksheets.Item("История Изменений")


$Change_History.columns.item('i').NumberFormat = "@"
$Change_History.Rows.Item(1).HorizontalAlignment = -4108
$Change_History.Cells.Item(1,1) = 'Имя Пользователя'
$Change_History.Cells.Item(1,2) = 'Сетевое имя'
$Change_History.Cells.Item(1,3) = 'Дата Проверки'
$Change_History.Cells.Item(1,4) = 'OS'
$Change_History.Cells.Item(1,5) = 'Процессор'
$Change_History.Cells.Item(1,6) = 'Модель'
$Change_History.Cells.Item(1,7) = 'Материнская плата'
$Change_History.Cells.Item(1,8) = 'Модель'
$Change_History.Cells.Item(1,9) = 'Серийный номер'
#Column HDD Start 8
$Change_History.Cells.Item(1,10) = 'HDD 1'
$Change_History.Cells.Item(1,11) = 'HDD 2'
$Change_History.Cells.Item(1,12) = 'HDD 3'
$Change_History.Cells.Item(1,13) = 'HDD 4'
#$Change_History.Cells.Item(1,9) = 'Объем (Гб)'
#Column OZY START 14
$Change_History.Cells.Item(1,14) = 'Суммарно ОЗУ (Гб)'
$Change_History.Cells.Item(1,15) = 'Тип Памяти'
$Change_History.Cells.Item(1,16) = 'ОЗУ 1 (ГБ)'
$Change_History.Cells.Item(1,17) = 'ОЗУ 2 (ГБ)'
$Change_History.Cells.Item(1,18) = 'ОЗУ 3 (ГБ)'
$Change_History.Cells.Item(1,19) = 'ОЗУ 4 (ГБ)'
# Column Video Start 20
$Change_History.Cells.Item(1,20) = 'Видеокарта 1'
$Change_History.Cells.Item(1,21) = 'Объем памяти (MB)'
$Change_History.Cells.Item(1,22) = 'Видеокарта 2'
$Change_History.Cells.Item(1,23) = 'Объем памяти (MB)'
$Change_History.Cells.Item(1,24) = 'Видеокарта 3'
$Change_History.Cells.Item(1,25) = 'Объем памяти (MB)'
$Change_History.Cells.Item(1,25) = 'Видеокарта 3'
$Change_History.Cells.Item(1,26) = 'Объем памяти (MB)'
#Column Network Start
$Change_History.Cells.Item(1,27) = 'Cетевая Карта 1'
$Change_History.Cells.Item(1,28) = 'MAC'
$Change_History.Cells.Item(1,29) = 'Cетевая Карта 2'
$Change_History.Cells.Item(1,30) = 'MAC'
$Change_History.Cells.Item(1,31) = 'Cетевая Карта 3'
$Change_History.Cells.Item(1,32) = 'MAC'
$Change_History.Cells.Item(1,33) = 'Cетевая Карта 4'
$Change_History.Cells.Item(1,34) = 'MAC'
#Column Availabilyty  33



$Range = $Change_History.Range("A1","AI1")
$Range.AutoFilter() | Out-Null
$Range.Interior.ColorIndex = 15

$Row_Change = 2
$Column_Change = 1
$Initial_Change_Row = 2



#Основная инвентаризационная страница ЦО
$InventoryFile = $Excel.Worksheets.Add()
$InventoryFile.Name = 'Инвентаризация Воронеж'
$InventoryFile = $WorkBook.Worksheets.Item('Инвентаризация Воронеж')
$InventoryFile.columns.item('i').NumberFormat = "@"
$InventoryFile.Rows.Item(1).HorizontalAlignment = -4108
$InventoryFile.Columns.Item('u').HorizontalAlignment = -4108
$InventoryFile.Columns.Item('w').HorizontalAlignment = -4108
$InventoryFile.Columns.Item('y').HorizontalAlignment = -4108
$InventoryFile.Cells.Item(1,1) = 'Имя Пользователя'
$InventoryFile.Cells.Item(1,2) = 'Сетевое имя'
$InventoryFile.Cells.Item(1,3) = 'Дата Проверки'
$InventoryFile.Cells.Item(1,4) = 'OS'
$InventoryFile.Cells.Item(1,5) = 'Процессор'
$InventoryFile.Cells.Item(1,6) = 'Модель'
$InventoryFile.Cells.Item(1,7) = 'Материнская плата'
$InventoryFile.Cells.Item(1,8) = 'Модель'
$InventoryFile.Cells.Item(1,9) = 'Серийный номер'
#Column HDD Start 8
$InventoryFile.Cells.Item(1,10) = 'HDD 1'
$InventoryFile.Cells.Item(1,11) = 'HDD 2'
$InventoryFile.Cells.Item(1,12) = 'HDD 3'
$InventoryFile.Cells.Item(1,13) = 'HDD 4'
#$InventoryFile.Cells.Item(1,9) = 'Объем (Гб)'
#Column OZY START 14
$InventoryFile.Cells.Item(1,14) = 'Суммарно ОЗУ (Гб)'
$InventoryFile.Cells.Item(1,15) = 'Тип Памяти'
$InventoryFile.Cells.Item(1,16) = 'ОЗУ 1 (ГБ)'
$InventoryFile.Cells.Item(1,17) = 'ОЗУ 2 (ГБ)'
$InventoryFile.Cells.Item(1,18) = 'ОЗУ 3 (ГБ)'
$InventoryFile.Cells.Item(1,19) = 'ОЗУ 4 (ГБ)'
# Column Video Start 20
$InventoryFile.Cells.Item(1,20) = 'Видеокарта 1'
$InventoryFile.Cells.Item(1,21) = 'Объем памяти (MB)'
$InventoryFile.Cells.Item(1,22) = 'Видеокарта 2'
$InventoryFile.Cells.Item(1,23) = 'Объем памяти (MB)'
$InventoryFile.Cells.Item(1,24) = 'Видеокарта 3'
$InventoryFile.Cells.Item(1,25) = 'Объем памяти (MB)'
$InventoryFile.Cells.Item(1,25) = 'Видеокарта 3'
$InventoryFile.Cells.Item(1,26) = 'Объем памяти (MB)'
#Column Network Start
$InventoryFile.Cells.Item(1,27) = 'Cетевая Карта 1'
$InventoryFile.Cells.Item(1,28) = 'MAC'
$InventoryFile.Cells.Item(1,29) = 'Cетевая Карта 2'
$InventoryFile.Cells.Item(1,30) = 'MAC'
#Column Availabilyty  31
$InventoryFile.Cells.Item(1,31) = 'Монитор №1'
$InventoryFile.Cells.Item(1,32) = 'Монитор №2'
$InventoryFile.Cells.Item(1,33) = 'Монитор №3'
$InventoryFile.Cells.Item(1,34) = 'Монитор №4'

$Range = $InventoryFile.Range("A1","AI1")
$Range.WrapText = $True
$Range.AutoFilter() | Out-Null
$Range.Interior.ColorIndex = 15
# Выделяем жирным шапку таблицы
$InventoryFile.Rows.Item(1).Font.Bold = $true




}
# ДРУГОЙ РЕЖИМ!
elseif($Configuration_Start -eq 1)
{
    $FilePath = $INI_FOLDER

    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $true
    $Workbooks = $Excel.Workbooks.Open($FilePath)
    #$InventoryFile.Names("data12").Delete()
    

    #Sheets("data").Names("_FilterDatabase").Delete


# Main Window CO Selection
$InventoryFile = $WorkBooks.Worksheets.Item(1)

#$Range = $InventoryFile.Range("A1","AJ1")
#$Range.AutoFilter() | Out-Null

# Bad_PC Selection
$Bad_PC  = $WorkBooks.Worksheets.Item(3)

# Change_History Selection
$Change_History  = $WorkBooks.Worksheets.Item(2)

# Archive Selection
$Archive = $WorkBooks.Worksheets.Item(4)



$UsedRangeMain = $InventoryFile.UsedRange
$Row_New = $UsedRangeMain.Rows.Count

$UsedRangeBad = $Bad_PC.UsedRange
$RowBad_New = $UsedRangeBad.Rows.Count

$UsedRangeChange = $Change_History.UsedRange
$RowChange_New = $UsedRangeChange.Rows.Count


$UsedRangeArchive = $Archive.UsedRange
$Row_Archive_New = $UsedRangeArchive.Rows.Count

# INVENTORY
$Column = 1
$Row = $Row_new+1
$InitialRow = $Row_new+1

# BAD
$BadColumn = 1
$BadRow = $RowBad_New+1
$Initial_Bad_Row = $RowBad_New+1


# CHANGE
$Column_Change = 1
$Row_Change = $RowChange_New+1
$Initial_Change_Row = $RowChange_New+1

# ARCHIVE 
$Column_Archive = 1
$Row_Archive = $Row_Archive_New+1
$Initial_Archive_Row = $Row_Archive_New+1




}




$ImportCsv = import-csv "$AD_GREP_FILE"

$Current_Date = Get-Date -format "dd.MM.yyyy"
 
 $ImportCsv | ForEach-Object {
$a=$_.name
$b=$_.Description
if(($a -like "*srt_wsus*") -or ($a -like "*W00-0602*") -or ($a -like "*W00-0642*") -or ($a -like "W00-0656") -or ($a -like"W00-0366"))
{
}
else
{
if ((Test-Connection $a -count 1 -quiet) -eq "True")
{ 

        # Bad_PC Initialize
        <#if($Configuration_Start -eq 0)
        {
        # Заполнение Доступных ПК
        $Bad_PC.Cells.Item($BadRow, $BadColumn) = $b
        $BadColumn++
        $Bad_PC.Cells.Item($BadRow, $BadColumn) = $a
        $BadColumn++
        $Check = $Bad_PC.UsedRange.find("$a")
        $BadColumn = $Check.Column
        $BadColumn++
        if($Bad_PC.Cells.Item($Check.Row,$BadColumn).Text -eq "НЕДОСТУПЕН" -or $Bad_PC.Cells.Item($Check.Row,$BadColumn).Value2 -eq $Null)
        {
            $RRW = $Check.Row
            #Cтатус
            $Bad_PC.Cells.Item($Check.Row,$BadColumn) = "ДОСТУПЕН"
            $Bad_PC.Cells.Item($Check.Row,$BadColumn).font.ColorIndex = 10
            $BadColumn++
            # Дата Падения
            $Bad_PC.Cells.Item($Check.Row,$BadColumn) = ""
            $BadColumn++
            # Дата восстановления
            $Bad_PC.Cells.Item($Check.Row,$BadColumn) = ""
            $BadColumn++
            #Дата сканирования
            $Bad_PC.Cells.Item($Check.Row,$BadColumn) = $Current_Date
            $BadColumn++
            # Расчет кол-ва дней
            $Bad_PC.Cells.Item($Check.Row,$BadColumn).Formula = "=IF(C$RRW=`"`Недоступен`"`,DATEDIF(D$RRW,F$RRW,`"`d`"`),`"`")"
        }
        


        }
        #>
        if($Configuration_Start -eq 1)
        {
         $Check = $Bad_PC.UsedRange.find($a)   
        if($Check -ne $null)
        {
            $BadColumn = $Check.Column
            $BadColumn++
            if($Bad_PC.Cells.Item($Check.Row,$BadColumn).Text -eq "НЕДОСТУПЕН")
            {
                $Bad_PC.Cells.Item($Check.Row,6) = $Current_Date
            }
            if($Bad_PC.Cells.Item($Check.Row,$BadColumn).Text -eq "НЕДОСТУПЕН" -or $Bad_PC.Cells.Item($Check.Row,$BadColumn).Value2 -eq $Null)
            {
                $RRW = $Check.Row
                #Cтатус
                $Bad_PC.Cells.Item($Check.Row,$BadColumn) = "ДОСТУПЕН"
                $Bad_PC.Cells.Item($Check.Row,$BadColumn).font.ColorIndex = 10
                $BadColumn++
                # Дата Падения
                $Bad_PC.Cells.Item($Check.Row,$BadColumn) = ""
                $BadColumn++
                # Дата восстановления
                $Bad_PC.Cells.Item($Check.Row,$BadColumn) = $Current_Date
                $BadColumn++
                #Дата сканирования
                $Bad_PC.Cells.Item($Check.Row,$BadColumn) = $Current_Date
                $BadColumn++
            }
            elseif($Bad_PC.Cells.Item($Check.Row,$BadColumn).Text -eq "ДОСТУПЕН")
            {
                $Bad_PC.Rows($Check.Row).Delete()
            }
            }
        else
        {

        }
        $Check = ""
        }

        $RowStart = $Row
        Write-Host "$A PC - Доступен!" -ForegroundColor Cyan
        Write-Host "Проверка компьютера " -ForeGroundColor Green $a 
        #Запись имени пользователя и имени ПК
        $InventoryFile.Cells.Item($Row, $Column) = $b
        $Column++
        $InventoryFile.Cells.Item($Row, $Column) = $a
        $Column++
        $InventoryFile.Cells.Item($Row, $Column) = Get-Date
        $Column++
        # Получение сведений об ОС
        $Parameter  = Get-WmiObject -computername $a Win32_OperatingSystem | Select-Object csname, caption, Serialnumber, csdVersion  -ErrorAction Stop
        $InventoryFile.Cells.Item($Row, $Column) = $Parameter.caption
        $Column++
        
###########################################################################################

        #Модель процессора и прочая ересь
        $Parameter = Get-WmiObject -computername $a Win32_Processor | Select-Object name, SocketDesignation, Description -ErrorAction Stop
        $InventoryFile.Cells.Item($Row, $Column) = $Parameter.name
        $Column++
        $InventoryFile.Cells.Item($Row, $Column) = $Parameter.SocketDesignation
        $Column++
        
###########################################################################################

        #Модель материнской платы
        #"Материнская плата" 
        $Parameter = Get-WmiObject -computername $a Win32_BaseBoard | Select-Object Manufacturer, Product, SerialNumber -ErrorAction Stop
        $InventoryFile.Cells.Item($Row, $Column) = $Parameter.Manufacturer
        $Column++
        $InventoryFile.Cells.Item($Row, $Column) = $Parameter.Product
        $Column++
        $InventoryFile.Cells.Item($Row, $Column) = $Parameter.Serialnumber
        $Column++
###########################################################################################

        # HDD + SSD
        #"Жесткие диски" 

        $ColemnTemp = $Column
        $RowTemp = $Row
        $ColOfElements = 0
        if($a -eq "W00-0626")
        {
           
        }
        else
        {
        Get-WmiObject -computername $a Win32_DiskDrive | Where-Object {$_.Model -notlike "*usb*" -or $_.Model -notlike "*USB*"}| ForEach-Object `
        {
            $InventoryFile.Cells.Item($Row, $Column) = $_.Model
            $Column++
            #$InventoryFile.Cells.Item($Row, $Column) = ($_.Size/1GB).ToString("F00")
            #$Row++
            $ColOfElements++
            #Write-Host "Жесткие диски "$Column
            
            #$RowHDD = $Row
            #$Column = $ColemnTemp

        } -ErrorAction Stop
        }
        $Row = $RowTemp
        $Column=14
        
###########################################################################################
        # ОЗУ
        # "Оперативная память"
        $ColemnTemp = $Column
        $RowTemp =$Row
        $ColOfElements = 0

        

        $T = Get-WmiObject -computername $a Win32_Physicalmemory | Measure-Object -Property capacity -Sum
        $T= $T.Sum/1GB
        $InventoryFile.Cells.Item($Row, $Column) = $T
        $Column++
        $Speed = (Get-WmiObject -computername $a Win32_Physicalmemory).Speed | select -First 1
        if($Speed -le 1666 -and $Speed  -ge 1333 )
            {
                $Type = "DDR3"
            }
            elseif($Speed  -lt 1333)
            {
                $Type = "DDR2"
            }
            else
            {
                $Type = "DDR4"
            }
        $InventoryFile.Cells.Item($Row, $Column) = $Type
        $Column++
        Get-WmiObject -computername $a Win32_Physicalmemory | ForEach-Object `
        {
            $InventoryFile.Cells.Item($Row, $Column) = ([Math]::Round($_.Capacity/1GB, 2))
            $Column++
            $ColOfElements++

            $RowOZY = $Row

        } -ErrorAction Stop
        $Row = $RowTemp
        $Column=20
###########################################################################################
        
        # Видеокарта
        $ColemnTemp = $Column
        $RowTemp =$Row
        $ColOfElements = 0
        #Write-Host "Перед Видяхой "$Column
        Get-WmiObject -computername $a Win32_videoController | ForEach-Object `
        {
                
                if($_.Name -like "Radmin*" -or $_.Name -like "*Remote*")
                {
                }
                else
                {
                $InventoryFile.Cells.Item($Row, $Column) = $_.name
                $Column++
                $InventoryFile.Cells.Item($Row, $Column) = ($_.AdapterRAM/1MB).tostring("F00")
                $Column++
                $ColOfElements++

                $RowVideo = $Row
                }
                
        } -ErrorAction Stop
        $Column = 27
        $Row = $RowTemp
        #$Column+=2
###########################################################################################
        
       # Сетевая Карта
        $ColemnTemp = $Column
        $RowTemp =$Row
        $ColOfElements = 0   

        
        $OS=Get-WmiObject -computername $a Win32_OperatingSystem | foreach {$_.caption}
        if ($OS -eq "Microsoft Windows 2000 Professional")
        { 
        $Parameter = Get-WmiObject -computername $a Win32_NetworkAdapterConfiguration -Filter "DHCPEnabled=True" | ForEach-Object `
            {

            $InventoryFile.Cells.Item($Row, $Column) = $_.caption
            $Column++
            $InventoryFile.Cells.Item($Row, $Column) = $_.MACAddress
            $Column++

            }
        }
        else
        {
        $Parameter = Get-WmiObject -computername $a Win32_NetworkAdapter | Where-Object {$_.Name -like "*Realtek*" -or $_.Name -like "*Ethernet*" -and $_.Name -notlike "*Wireless*" -and  $_.Name -notlike "*Bluetooth*" -and $_.Name -notlike "*Apple*" -and $_.Name -notlike "*Hyper-V*" -and $_.MACAddress -notlike ""} | ForEach-Object `
            {
            $InventoryFile.Cells.Item($Row, $Column) = $_.Name
            $Column++
            $InventoryFile.Cells.Item($Row, $Column) = $_.MACAddress
            $Column++
            }
        }

        $ColemnTemp = $Column
        $RowTemp =$Row
        $ColOfElements = 0
        $Column = 31

        $Monitors = Get-WmiObject WmiMonitorID -ComputerName $a -Namespace root\wmi | ForEach-Object {($_.UserFriendlyName | foreach {[char]$_}) -join "";}
        $Monitors | ForEach-Object {

            $InventoryFile.Cells.Item($Row, $Column) = $_
            $Column++


        } -ErrorAction Stop




$Range_Current = $InventoryFile.Range("B"+$Row,"Y"+$Row)
#$Range_Current.font.ColorIndex = 10
#$Range_Current.copy()

#Range_Previous = $InventoryFile.Range("B"+($Row-1).ToString(),"Y"+($Row-1).ToString())
#$InventoryFile.Paste($Range_Previous)


$Row++
$BadColumn = 1
$RowFinish = $Row
$Column = 1
Write-Host "Проверка компьютера Завершена! " -ForeGroundColor DarkYellow $a
$Set = 1

# Formula Excel
$Formula = "=IF(C$RRW=`"`Недоступен`"`,DATEDIF(D$RRW,F$RRW,`"`d`"`),`"`")"


}
elseif ((Test-connection $a -count 1 -quiet) -ne "True")
{

        if($Configuration_Start -eq 0)
        {
        Write-Host "$A PC - НЕДОСТУПЕН" -ForeGroundColor DarkRed
        #Запись имени ПК и Имени пользователя
        $Bad_PC.Cells.Item($BadRow, $BadColumn) = $b
        $BadColumn++
        $Bad_PC.Cells.Item($BadRow, $BadColumn) = $a
        $BadColumn++

# Заполнение Недоступных ПК
        $Check = $Bad_PC.UsedRange.find("$a")
        $RRW = $Check.Row
                    
        $BadColumn = $Check.Column
        $BadColumn++
        if($Bad_PC.Cells.Item($Check.Row,$BadColumn).Text -eq "ДОСТУПЕН" -or $Bad_PC.Cells.Item($Check.Row,$BadColumn).Value2 -eq $Null)
        {

            #Cтатус
            $Bad_PC.Cells.Item($Check.Row,$BadColumn) = "НЕДОСТУПЕН"
            $Bad_PC.Cells.Item($Check.Row,$BadColumn).font.Color = 255
            $BadColumn++
            # Дата Падения
            If($Bad_PC.Cells.Item($Check.Row,$BadColumn).Value2 -eq $Null)
            {
                $Bad_PC.Cells.Item($Check.Row,$BadColumn) = $Current_Date
            }
            else
            {
                #Установлена старая дата! Был недоступен Ранее
            }
            $BadColumn++
            # Дата восстановления
            $Bad_PC.Cells.Item($Check.Row,$BadColumn) = ""
            $BadColumn++
            #Дата сканирования
            $Bad_PC.Cells.Item($Check.Row,$BadColumn) = $Current_Date
            $BadColumn++
            # Расчет кол-ва дней 
            #$Bad_PC.Cells.Item($Check.Row,$BadColumn).Formula = $Formula
        }

        $BadRow++
        $BadColumn = 1
        }
        elseif($Configuration_Start -eq 1)
        {
            Write-Host "$A PC - НЕДОСТУПЕН" -ForeGroundColor DarkRed
            #Панель дислокаций	W00-0289
            $Check = $null
            $Check = $Bad_PC.UsedRange.find($a)
            if($Check.Text -eq $null)
            {
                Write-Host "$A PC обнаружена потеря соединения с ПК" -ForeGroundColor DarkRed
                #Запись имени ПК и Имени пользователя
                $Bad_PC.Cells.Item($BadRow, $BadColumn) = $b
                $BadColumn++
                $Bad_PC.Cells.Item($BadRow, $BadColumn) = $a
                $BadColumn++
                $Bad_PC.Cells.Item($BadRow, $BadColumn) = "НЕДОСТУПЕН"
                $Bad_PC.Cells.Item($BadRow, $BadColumn).font.Color = 255
                $BadColumn++
                $Bad_PC.Cells.Item($BadRow, $BadColumn) = $Current_Date
                $Bad_PC.Cells.Item($BadRow,6) = $Current_Date
                #$Bad_PC.Cells.Item($BadRow,7).Formula = $Formula
                $BadRow++
                $BadColumn = 1
            }
            elseif($Check.Text -ne $null)
            {
                $Check_Col = $False
                $Target = $Check
                $First = $Target
                Do
                {
                    #Write-Host $Target.Row
                    # Взяли строку
                    #
                    # Cравниваем чекируем
                       if(($Bad_PC.Cells.Item($Target.Row,1).Text -eq $b) -and ($Bad_PC.Cells.Item($Target.Row,2).Text -eq $a))
                        {
                            Write-Host "ПК был ранее недоступен" -ForegroundColor DarkRed
                            $Check_Col = $true
                            $Bad_PC.Cells.Item($Target.Row,6) = $Current_Date
                        }
                        elseif(($Bad_PC.Cells.Item($Target.Row,1).Text -ne $b) -and ($Bad_PC.Cells.Item($Target.Row,2).Text -eq $a))
                        {
                            "Проставляем дату на несовпадающем"
                            if($Check_Col -eq $False)
                            {
                                "Проставляем дату на несовпадающем"
                                $Bad_PC.Cells.Item($Target.Row,6) = $Current_Date
                            }

                        }
                    $Target = $Bad_PC.UsedRange.FindNext($Target)
                }
                While ($Target -ne $NULL -and $Target.AddressLocal() -ne $First.AddressLocal())

                if($Check_Col -eq $False)
                { 
                   $BadColumn = 1  
                   $Bad_PC.Cells.Item($BadRow, $BadColumn) = $b
                   $BadColumn++
                   $Bad_PC.Cells.Item($BadRow, $BadColumn) = $a
                   $BadColumn++
                   $Bad_PC.Cells.Item($BadRow, $BadColumn) = "НЕДОСТУПЕН"
                   $Bad_PC.Cells.Item($BadRow, $BadColumn).font.Color = 255
                   $BadColumn++
                   $Bad_PC.Cells.Item($BadRow, $BadColumn) = $Current_Date
                   $Bad_PC.Cells.Item($BadRow,6) = $Current_Date
                   #$Bad_PC.Cells.Item($BadRow,7).Formula = $Formula
                   $BadRow++
                   $BadColumn = 1  
                        
                }
                
            }
            $Check = $null
        }

}
}
}



# Проверка всего списка недоступных ПК - Актуализация даты недоступности
$Bad_Pc_Range = $Bad_PC.UsedRange
foreach($PC in $Bad_Pc_Range.Rows)
{
    $ColOfCompare = 0
    $Row_PC = $PC.Row -as [int]
    if($Bad_PC.Cells.Item($Row_PC,3).Formula -eq "НЕДОСТУПЕН")
    {
        #Write-Host $Bad_PC.Cells.Item($Row_PC,2).Formula
        $Bad_PC.Cells.Item($Row_PC,6).Formula = $Current_Date
    }
    
}


$Row--
$DataRangeInventory = $InventoryFile.Range(("A{0}" -f 1), ("AI{0}" -f $Row))
7..12 | ForEach-Object `
{
    $DataRangeInventory.Borders.Item($_).LineStyle = 1
    $DataRangeInventory.Borders.Item($_).Weight = 2
}

$BadRow--
$DataRangeInventory = $Bad_PC.Range(("A{0}" -f 1), ("G{0}" -f $BadRow))
7..12 | ForEach-Object `
{
    $DataRangeInventory.Borders.Item($_).LineStyle = 1
    $DataRangeInventory.Borders.Item($_).Weight = 2
}

$Row_Change++
$DataRangeInventory = $Change_History.Range(("A{0}" -f 1), ("AI{0}" -f $Row_Change))
7..12 | ForEach-Object `
{
    $DataRangeInventory.Borders.Item($_).LineStyle = 1
    $DataRangeInventory.Borders.Item($_).Weight = 2
}
$Row_Change--

# Последняя строка файла и сортировка по имени!
$Filler = [System.Type]::Missing
$UsedRange = $InventoryFile.UsedRange
$UsedRange.EntireColumn.AutoFit() | Out-Null
$T = "A" + $UsedRange.Rows.Count
$Sorting_Space = $InventoryFile.range("A2:$T" )
#$Sorting_Space.Select()
$UsedRange.Sort($Sorting_Space,1,$Filler,$Filler,$Filler,$Filler,$Filler,1)



$Filler = [System.Type]::Missing
$UsedRange = $Bad_PC.UsedRange
$UsedRange.EntireColumn.AutoFit() | Out-Null
$T = "C" + $UsedRange.Rows.Count
$Sorting_Space = $Bad_PC.range("C2:$T" )
#$Sorting_Space.Select()
$UsedRange.Sort($Sorting_Space,2,$Filler,$Filler,$Filler,$Filler,$Filler,1)


# HDD Width
$InventoryFile.columns.item('k').ColumnWidth = 7
$InventoryFile.columns.item('l').ColumnWidth = 7
$InventoryFile.columns.item('m').ColumnWidth = 7

# ОЗУ Width

$InventoryFile.columns.item('p').ColumnWidth = 5
$InventoryFile.columns.item('q').ColumnWidth = 5
$InventoryFile.columns.item('r').ColumnWidth = 5
$InventoryFile.columns.item('s').ColumnWidth = 5

# Video Width
$InventoryFile.columns.item('T').ColumnWidth = 27
$InventoryFile.columns.item('u').ColumnWidth = 7
$InventoryFile.columns.item('v').ColumnWidth = 7
$InventoryFile.columns.item('w').ColumnWidth = 7
$InventoryFile.columns.item('X').ColumnWidth = 7
$InventoryFile.columns.item('Y').ColumnWidth = 7

#Network Width
$InventoryFile.columns.item('AA').ColumnWidth = 25
$InventoryFile.columns.item('AB').ColumnWidth = 17
$InventoryFile.columns.item('AC').ColumnWidth = 7
$InventoryFile.columns.item('AD').ColumnWidth = 5
$InventoryFile.columns.item('AE').ColumnWidth = 7
$InventoryFile.columns.item('AF').ColumnWidth = 5

$UsedBadRange = $Bad_PC.UsedRange
$UsedBadRange.EntireColumn.AutoFit() | Out-Null



# Добавление новых строк их сравнение и выоплнение определения изменений

if($Configuration_Start -eq 1)
{
#Блок проверки поступивших данных и удаление совпадающих.

$Work_Range = $InventoryFile.UsedRange
#$Work_Range.Rows

foreach($Name in $Work_Range.Rows)
{
    $ColOfCompare = 0
    $Test = $Name.Row -as [int]
    #$Test
    $Username = $InventoryFile.Cells.Item($Test,1).Formula

    if($InventoryFile.Cells.Item($Test,3).Formula -eq "") 
    {
        $InventoryFile.Rows($Test).Delete()
    }
    elseif($InventoryFile.Cells.Item($Test+1,3).Formula -eq "" )
    {
        $InventoryFile.Rows($Test).Delete()   
    }


    if($InventoryFile.Cells.Item($Test,1).Formula -eq $InventoryFile.Cells.Item($Test+1,1).Formula -and ($InventoryFile.Cells.Item($Test,1).Formula -ne "" -or $InventoryFile.Cells.Item($Test+1,1).Formula -ne "") -and ($InventoryFile.Cells.Item($Test,3).Formula -ne "" -or $InventoryFile.Cells.Item($Test+1,3).Formula -ne "" ))
    {
        Write-Host "Пользователи Совпадают!" $InventoryFile.Cells.Item($Test,1).Formula  " "  $InventoryFile.Cells.Item($Test+1,1).Formula
        $TESTO = $Test+1
        for($i=4;$i -lt 30;$i++)
        {
           if($InventoryFile.Cells.Item($Test,$i).Formula -eq $InventoryFile.Cells.Item($Test+1,$i).Formula)
           {
            $ColOfCompare++
            }
           elseif($InventoryFile.Cells.Item($Test,$i).Formula -ne $InventoryFile.Cells.Item($Test+1,$i).Formula -and $InventoryFile.Cells.Item($Test,$i).Formula -notcontains "*USB*" -or $InventoryFile.Cells.Item($Test+1,$i).Formula -notcontains "*USB*")
           {
            # Если есть изменения в совпадающих строках
            "Ячейка не равны" + $InventoryFile.Cells.Item($Test,$i).Formula + " " + $InventoryFile.Cells.Item($Test+1,$i).Formula
            $InventoryFile.Cells.Item($Test,$i).Interior.ColorIndex = 6
           }
            
        }

        if($ColOfCompare -eq 26)
        {
            $InventoryFile.Rows($Test+1).Delete() | Out-Null
        }
        else
        {
            $Range = $InventoryFile.Rows($Test)
            $Range.Cut()
            $Insert_Into = $Change_History.Rows($Row_Change)
            $Change_History.Paste($Insert_Into)
            $InventoryFile.Rows($Test).Delete()
            $Row_Change++
        }
    }
}
    #$WorkBooks.SaveAs("C:\Test\Инвентаризация.xlsx")
}
# Cортировка по Имени ПК
$Filler = [System.Type]::Missing
$UsedRange = $InventoryFile.UsedRange
$UsedRange.EntireColumn.AutoFit() | Out-Null
$T = "B" + $UsedRange.Rows.Count
$Sorting_Space = $InventoryFile.range("B2:$T" )
#$Sorting_Space.Select()
$UsedRange.Sort($Sorting_Space,1,$Filler,$Filler,$Filler,$Filler,$Filler,1)

# Тело Сортировки
$Work_Range = $InventoryFile.UsedRange
foreach($NamePC in $Work_Range.Rows)
{
    
    $RRP = $NamePC.Row -as [int]
    if(($Work_Range.Cells.Item($RRP,1).Formula -eq $Work_Range.Cells.Item($RRP+1,1).Formula) -and ($Work_Range.Cells.Item($RRP,2).Formula -eq $Work_Range.Cells.Item($RRP+1,2).Formula))
    {
     "Все параметры равны"
    }
    elseif(($Work_Range.Cells.Item($RRP,1).Formula -ne $Work_Range.Cells.Item($RRP+1,1).Formula) -and ($Work_Range.Cells.Item($RRP,2).Formula -eq $Work_Range.Cells.Item($RRP+1,2).Formula))
    {
    $Work_Range.Cells.Item($RRP,2).Formula
    $Work_Range.Cells.Item($RRP+1,2).Value2
        if($Work_Range.Cells.Item($RRP,3).Formula -lt $Work_Range.Cells.Item($RRP+1,3).Formula -and ($Work_Range.Cells.Item($RRP,3).Interior.ColorIndex -eq -4142 -or 0 ))
        {
                $INDEX = Get-Random -Minimum 2 -Maximum 24
                $Work_Range.Range("A$RRP","AI$RRP").Interior.ColorIndex = $INDEX
                $Work_Range.Cells.Item($RRP,35) = $Work_Range.Cells.Item($RRP+1,1)
                $Work_Range.Cells.Item($RRP+1,28).Interior.ColorIndex = $INDEX
                $Work_Range.Cells.Item($RRP+1,35) = $Work_Range.Cells.Item($RRP,1)
                $InventoryFile.Hyperlinks.Add( `
                $InventoryFile.Cells.Item($RRP+1,35) , `
                "" , "Архив!A$Row_Archive", "", $Work_Range.Cells.Item($RRP,1).Value2)
                $Range = $Work_Range.Range("A$RRP","AI$RRP")
                $Range.Cut()
                $Insert_Into = $Archive.Rows($Row_Archive)
                $Archive.Paste($Insert_Into)
                $InventoryFile.Rows($RRP).Delete() 
                $Row_Archive++
        }
        elseif($Work_Range.Cells.Item($RRP,3).Formula -gt $Work_Range.Cells.Item($RRP+1,3).Formula -and ($Work_Range.Cells.Item($RRP+1,3).Interior.ColorIndex -eq -4142 -or 0 ))
        {
                $Set = $RRP+1
                $INDEX = Get-Random -Minimum 2 -Maximum 24
                $Work_Range.Range("A$SET","AI$SET").Interior.ColorIndex = $INDEX
                $Work_Range.Cells.Item($RRP+1,35) = $Work_Range.Cells.Item($RRP,1) 
                $Work_Range.Cells.Item($RRP,28).Interior.ColorIndex = $INDEX 
                $Work_Range.Cells.Item($RRP,35) = $Work_Range.Cells.Item($RRP+1,1)
                $InventoryFile.Hyperlinks.Add( `
                $InventoryFile.Cells.Item($RRP,35) , `
                "" , "Архив!A$Row_Archive", "", $Work_Range.Cells.Item($RRP+1,1).Value2)
                $Range = $Work_Range.Range("A$SET","AI$SET")
                $Range.Cut()
                $Insert_Into = $Archive.Rows($Row_Archive)
                $Archive.Paste($Insert_Into)
                $InventoryFile.Rows($RRP+1).Delete()
                $Row_Archive++
        }
        else
        {
         
        }
    }
    
}


# Cортировка по MAC Адресу
$Filler = [System.Type]::Missing
$UsedRange = $InventoryFile.UsedRange
$UsedRange.EntireColumn.AutoFit() | Out-Null
$T = "AB" + $UsedRange.Rows.Count
$Sorting_Space = $InventoryFile.range("AB2:$T" )
$UsedRange.Sort($Sorting_Space,1,$Filler,$Filler,$Filler,$Filler,$Filler,1)
# Проверка на совпадения МАК Адресов
$Work_Range = $InventoryFile.UsedRange
foreach($NamePC in $Work_Range.Rows)
{
    $RRP = $NamePC.Row -as [int]
    if(($Work_Range.Cells.Item($RRP,28).Formula -eq $Work_Range.Cells.Item($RRP+1,28).Formula) -and ($Work_Range.Cells.Item($RRP,28).Formula -ne "" -or $Work_Range.Cells.Item($RRP+1,28).Formula -ne ""))
    {
     "МАКИ РАВНЫ!"
     $Work_Range.Cells.Item($RRP,2).Formula
    }
    elseif(($Work_Range.Cells.Item($RRP,28).Formula -ne $Work_Range.Cells.Item($RRP+1,28).Formula))
    {
    #$Work_Range.Cells.Item($RRP,2).Formula
    #$Work_Range.Cells.Item($RRP+1,2).Value2 
    }
}







# Восстановление сортировки по имени пользователя.
$Filler = [System.Type]::Missing
$UsedRange = $InventoryFile.UsedRange
$UsedRange.EntireColumn.AutoFit() | Out-Null
$T = "A" + $UsedRange.Rows.Count
$Sorting_Space = $InventoryFile.range("A2:$T" )
#$Sorting_Space.Select()
$UsedRange.Sort($Sorting_Space,1,$Filler,$Filler,$Filler,$Filler,$Filler,1)

$InventoryFile.Range("AJ1:AN200").Delete()

if($Configuration_Start -eq 0){
$WorkBook.SaveAs($INI_FOLDER)
}
else
{

$WorkBooks.SaveAs($INI_FOLDER)
}
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)



