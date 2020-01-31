#############################################
# Soft For Inventarization
#
# Created by Fokin L@B 
#############################################

############################################
#Import Module ACTIVE DIRECTORY
#
import-module ActiveDirectory

############################################
# Получаем список ПК из AD

#Get-ADComputer -Filter {Name -Like "W00-*"}  -Properties Description |
#Where-Object {$a=$_.name; $_.DistinguishedName -ne "CN=$a,OU=Computers,OU=Disabled,DC=rusagrotrans,DC=ru"} |
#Sort-Object NAME | Select-Object NAME,DESCRIPTION | Export-csv -NoTypeInformation C:\TEST\AllComputers.csv  -Encoding UTF8

# Инициализация Конфигурационного Файла:
$Config_File = "C:\Test\cfg.ini"
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

#Страница необработанных ПК
$Bad_PC = $Excel.Worksheets.Add()
$Bad_PC  = $WorkBook.Worksheets.Item(2)
$Bad_PC.Name = "Недоступные ПК"

<#$Bad_PC.Cells.Item(1,1) = 'Необработанные ПК'
$Bad_PC.Cells.Item(1,1).Font.Size = 12
$Bad_PC.Cells.Item(1,1).Font.Bold = $true
$Bad_PC.Cells.Item(1,1).Font.ThemeFont = 1
$Bad_PC.Cells.Item(1,1).Font.ThemeColor = 4
$Bad_PC.Cells.Item(1,1).Font.ColorIndex = 15
$Bad_PC.Cells.Item(1,1).Font.Color = 8210719
$Bad_PC.Rows.Item(1).HorizontalAlignment = -4108
$Bad_PC.Rows.Item(2).HorizontalAlignment = -4108
$Range = $Bad_PC.Range('A1','C1')
$Range.Merge()
#>

$Bad_PC.Cells.Item(2,1) = 'Имя Пользователя'
$Bad_PC.Cells.Item(2,2) = 'Сетевое имя'
$Range = $Bad_pc.Range("A2","B2")
$Range.Interior.ColorIndex = 15
$Bad_PC.Rows.Item(2).Font.Bold = $true
$Bad_PC.Rows.Item(1).HorizontalAlignment = -4108

#Основная инвентаризационная страница ЦО
<#

$InventoryFile.Name = "Инвентаризация Инфраструктуры"
$InventoryFile.Cells.Item(1,1) = 'Инвентаризационные данные ПК в ЦО'
$InventoryFile.Cells.Item(1,1).Font.Size = 18
$InventoryFile.Cells.Item(1,1).Font.Bold = $true
$InventoryFile.Cells.Item(1,1).Font.ThemeFont = 1
$InventoryFile.Cells.Item(1,1).Font.ThemeColor = 4
$InventoryFile.Cells.Item(1,1).Font.ColorIndex = 55
$InventoryFile.Cells.Item(1,1).Font.Color = 8210719
$InventoryFile.Rows.Item(1).HorizontalAlignment = -4108
$InventoryFile.Rows.Item(2).HorizontalAlignment = -4108
$Range = $InventoryFile.Range('A1','Y1')
$Range.Merge()
#>

$InventoryFile = $WorkBook.Worksheets.Item(1)
$InventoryFile.columns.item('i').NumberFormat = "@"
$InventoryFile.Rows.Item(1).HorizontalAlignment = -4108
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
#Column Next Start 23
#$InventoryFile.Cells.Item(2,23) = 'Сетевая карта'
#$InventoryFile.Cells.Item(2,24) = 'MAC Адрес'
$Range = $InventoryFile.Range("A1","Y1")
$Range.AutoFilter() | Out-Null
$Range.Interior.ColorIndex = 15

$InventoryFile.Name = 'Инвентаризация ЦО'
# Выделяем жирным шапку таблицы
$InventoryFile.Rows.Item(1).Font.Bold = $true

# Определение стратовой площадки записи в файл.
$Row = 2
$Column = 1
$BadColumn = 1
$BadRow = 2
$InitialRow = 2


}
elseif($Configuration_Start -eq 1)
{
    $FilePath = "C:\Test\MyExcel.xlsx"

    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $true
    $Workbooks = $Excel.Workbooks.Open($FilePath)
    
#Main Window CO
$InventoryFile = $WorkBooks.Worksheets.Item(1)
$Range = $InventoryFile.Range("A1","Y1")
$Range.AutoFilter() | Out-Null

#Bad_PC
$Bad_PC  = $WorkBook.Worksheets.Item(2)

$UsedRangeMain = $InventoryFile.UsedRange
$Row_New = $UsedRangeMain.Rows.Count

$UsedRangeBad = $Bad_PC.UsedRange
$RowBad_New = $UsedRangeBad.Rows.Count

$Row = $Row_new+1
$Column = 1
$BadColumn = $RowBad_New+1
$BadRow = 2
$InitialRow = $Row_new+1


}




$ImportCsv = import-csv c:\Test\AllComputers.csv

$Current_Date = Get-Date -format "dd.MM.yyyy"
 
 $ImportCsv | ForEach-Object {
$a=$_.name
$b=$_.Description
if ((Test-Connection $a -count 1 -quiet) -eq "True")
{ 
        $a
        $b
        $RowStart = $Row
        Write-Host "$A PC - Доступен!" -ForegroundColor Cyan
        Write-Host "Проверка компьютера " -ForeGroundColor Green $a "Компьютер" | Out-File C:\Test\Comp\$a.txt
        #Запись имени пользователя и имени ПК
        if($Configuration_Start -eq 0)
        {
            $InventoryFile.Cells.Item($Row, $Column) = $b
            #$Range = $InventoryFile.Range("B"+$Row,"Y"+$Row)
            #$Range.Merge()
        $Row++
        }
        $InventoryFile.Cells.Item($Row, $Column) = $b
        $Column++
        $InventoryFile.Cells.Item($Row, $Column) = $a
        $Column++
        $InventoryFile.Cells.Item($Row, $Column) = $Current_Date
        $Column++
        # Получение сведений об ОС
        $Parameter  = Get-WmiObject -computername $a Win32_OperatingSystem | Select-Object csname, caption, Serialnumber, csdVersion  -ErrorAction SilentlyContinue
        $InventoryFile.Cells.Item($Row, $Column) = $Parameter.caption
        $Column++
        
###########################################################################################

        #Модель процессора и прочая ересь
        "Процессор" | Out-File C:\Test\Comp\$a.txt -Append
        $Parameter = Get-WmiObject -computername $a Win32_Processor | Select-Object name, SocketDesignation, Description -ErrorAction SilentlyContinue
        $InventoryFile.Cells.Item($Row, $Column) = $Parameter.name
        $Column++
        $InventoryFile.Cells.Item($Row, $Column) = $Parameter.SocketDesignation
        $Column++
        
###########################################################################################

        #Модель материнской платы
        "Материнская плата" | Out-File C:\Test\Comp\$a.txt -append
        $Parameter = Get-WmiObject -computername $a Win32_BaseBoard | Select-Object Manufacturer, Product, SerialNumber -ErrorAction SilentlyContinue
        $InventoryFile.Cells.Item($Row, $Column) = $Parameter.Manufacturer
        $Column++
        $InventoryFile.Cells.Item($Row, $Column) = $Parameter.Product
        $Column++
        $InventoryFile.Cells.Item($Row, $Column) = $Parameter.Serialnumber
        $Column++
###########################################################################################

        # HDD + SSD
        "Жесткие диски" | out-file C:\Test\Comp\$a.txt -Append

        $ColemnTemp = $Column
        $RowTemp = $Row
        $ColOfElements = 0

        Get-WmiObject -computername $a Win32_DiskDrive | ForEach-Object `
        {
            $InventoryFile.Cells.Item($Row, $Column) = $_.Model
            $Column++
            #$InventoryFile.Cells.Item($Row, $Column) = ($_.Size/1GB).ToString("F00")
            #$Row++
            $ColOfElements++
            #Write-Host "Жесткие диски "$Column
            
            #$RowHDD = $Row
            #$Column = $ColemnTemp

        } -ErrorAction SilentlyContinue
        $Row = $RowTemp
        $Column=14

###########################################################################################
       # ОЗУ
        "Оперативная память" | out-file C:\Test\Comp\$a.txt -Append
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
            #$Column = $ColemnTemp
        } -ErrorAction SilentlyContinue
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
                
                if($_.Name -like "Radmin*")
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
                
        } -ErrorAction SilentlyContinue

        $Row = $RowTemp
        #$Column+=2
###########################################################################################
        
       <# Сетевая Карта
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
            $Row++

            $RowNet = $Row
            $Column = $ColemnTemp 
            }
        }
        else
        {
        $Parameter = Get-WmiObject -computername $a Win32_NetworkAdapter -Filter "NetConnectionStatus>0" | ForEach-Object `
            {
            $InventoryFile.Cells.Item($Row, $Column) = $_.Name
            $Column++
            $InventoryFile.Cells.Item($Row, $Column) = $_.MACAddress
            $Row++

            $RowNet = $Row
            $Column = $ColemnTemp 
            }
        }
        #>
<#if($Row -lt $RowHDD)
{
    $Row = $RowHDD
}
elseif($Row -lt $RowOZY)
{
    $Row = $RowOZY
}
elseif($Row -lt $RowVideo)
{
    $Row = $RowVideo
}
else
{

}
#>

$Range_Current = $InventoryFile.Range("B"+$Row,"Y"+$Row)
$Range_Current.font.ColorIndex = 10
#$Range_Current.copy()

#Range_Previous = $InventoryFile.Range("B"+($Row-1).ToString(),"Y"+($Row-1).ToString())
#$InventoryFile.Paste($Range_Previous)


$Row++
$RowFinish = $Row
$Column = 1

$Set = 1
For($i = $RowStart;$i -lt $RowFinish;$i++)
{
    if($InventoryFile.Cells.Item($i+1, $Set) -eq "")
    {
      $Range = $InventoryFile.Range('A'+$i.ToString(),'A'+$I+1)
      $Range.Merge()  
    }
}
$TEST

}
elseif ((Test-connection $a -count 1 -quiet) -ne "True")
{
        Write-Host "$A PC - НЕДОСТУПЕН"
        #Запись имени ПК и Имени пользователя
        $Bad_PC.Cells.Item($BadRow, $BadColumn) = $b
        $BadColumn++
        $Bad_PC.Cells.Item($BadRow, $BadColumn) = $a
        $BadColumn++
        
        $BadRow++
        $BadColumn = 1

}
}




$Row--
$DataRangeInventory = $InventoryFile.Range(("A{0}" -f 1), ("Y{0}" -f $Row))
7..12 | ForEach-Object `
{
    $DataRangeInventory.Borders.Item($_).LineStyle = 1
    $DataRangeInventory.Borders.Item($_).Weight = 2
}

#Последняя строка файла 
$Filler = [System.Type]::Missing
$UsedRange = $InventoryFile.UsedRange
$UsedRange.EntireColumn.AutoFit() | Out-Null
$T = "A" + $UsedRange.Rows.Count
$Sorting_Space = $InventoryFile.range("A2:$T" )
#$Sorting_Space.Select()
$UsedRange.Sort($Sorting_Space,1,$Filler,$Filler,$Filler,$Filler,$Filler,1)


$UsedBadRange = $Bad_PC.UsedRange
$UsedBadRange.EntireColumn.AutoFit() | Out-Null

$WorkBook.SaveAs("C:\Test\MyExcel.xlsx")
$InventoryFile

#Блок проверки поступивших данных и удаление совпадающих.

$Work_Range = $InventoryFile.UsedRange
#$Work_Range.Rows

foreach($Name in $Work_Range.Rows)
{
    $Test = $Name.Row -as [int]
    #$Test
    $Username = $InventoryFile.Cells.Item($Test,1).Formula

    if($InventoryFile.Cells.Item($Test,1).Formula -eq $InventoryFile.Cells.Item($Test+1,1).Formula -and ($InventoryFile.Cells.Item($Test,1).Formula -ne "" -or $InventoryFile.Cells.Item($Test+1,1).Formula -ne ""))
    {
          #$Grep = $InventoryFile.Rows($Test+1)
          #$Grep_Current = $InventoryFile.Rows($Test)
          Write-Host $Username
          Write-host $Name.Row
          $Next = $Test+1
          $Grep_Current = $InventoryFile.Range("D"+$Test,"Y"+$Test)
          $Grep = $InventoryFile.Range("D"+$Next,"Y"+$Next)
          $TESTGREP = $Grep_Current.Value2 -eq $Grep.Value2
          if($Grep_Current.Value2 -eq $Grep.Value2)
          {
            Write-Host "Удалена Строка "$Grep.Row
            $Grep.Delete()
          }
          elseif($Grep_Current.Value2 -ne $Grep.Value2)
          {
            Write-Host "Строки не равны!!!!"
          }
          else
          {
            Write-Host "ШТА?"
          }  
    }
}

