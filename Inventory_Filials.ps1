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


# Определение стратовой площадки записи в файл.
$Row = 2
$Column = 1
$BadColumn = 1
$BadRow = 2
$InitialRow = 2
$Initial_BadRow = 2

#Страница необработанных ПК
$Bad_PC = $Excel.Worksheets.Add()
$Bad_PC  = $WorkBook.Worksheets.Item(2)
$Bad_PC.Name = "Недоступные ПК"

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
$Change_History  = $WorkBook.Worksheets.Item(2)
$Change_History.Name = "История Изменений ЦО"

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
$Change_History.Cells.Item(1,33) = 'Cетевая Карта 4'
$Change_History.Cells.Item(1,34) = 'MAC'
#Column Availabilyty  33
$Change_History.Cells.Item(1,35) = 'Недоступен (День)'
$Change_History.Cells.Item(1,36) = 'Дата Обнаружения'


$Range = $Change_History.Range("A1","AJ1")
$Range.AutoFilter() | Out-Null
$Range.Interior.ColorIndex = 15

$Row_Change = 2
$Column_Change = 1
$Initial_Change_Row = 2



#Основная инвентаризационная страница ЦО
$InventoryFile = $WorkBook.Worksheets.Item(1)
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
$InventoryFile.Cells.Item(1,31) = 'Cетевая Карта 3'
$InventoryFile.Cells.Item(1,32) = 'MAC'
$InventoryFile.Cells.Item(1,33) = 'Cетевая Карта 4'
$InventoryFile.Cells.Item(1,34) = 'MAC'
#Column Availabilyty  33
$InventoryFile.Cells.Item(1,35) = 'Недоступен (День)'
$InventoryFile.Cells.Item(1,36) = 'Дата Обнаружения'

$InventoryFile.Name = 'Инвентаризация ЦО'
$Range = $InventoryFile.Range("A1","AJ1")
$Range.AutoFit()
$Range.WrapText = $True
$Range.AutoFilter() | Out-Null
$Range.Interior.ColorIndex = 15
# Выделяем жирным шапку таблицы
$InventoryFile.Rows.Item(1).Font.Bold = $true


}
# ДРУГОЙ РЕЖИМ!
elseif($Configuration_Start -eq 1)
{
    $FilePath = "C:\Test\MyExcel.xlsx"

    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $true
    $Workbooks = $Excel.Workbooks.Open($FilePath)
    
# Main Window CO Selection
$InventoryFile = $WorkBooks.Worksheets.Item(1)
#$Range = $InventoryFile.Range("A1","AJ1")
#$Range.AutoFilter() | Out-Null

# Bad_PC Selection
$Bad_PC  = $WorkBooks.Worksheets.Item(3)

# Change_History Selection
$Change_History  = $WorkBooks.Worksheets.Item(2)

$Row_Change = 2
$Column_Change = 1
$Initial_Change_Row = 2


$UsedRangeMain = $InventoryFile.UsedRange
$Row_New = $UsedRangeMain.Rows.Count

$UsedRangeBad = $Bad_PC.UsedRange
$RowBad_New = $UsedRangeBad.Rows.Count

$UsedRangeChange = $Change_History.UsedRange
$RowChange_New = $UsedRangeChange.Rows.Count

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





}




$ImportCsv = import-csv c:\Test\AllComputers.csv

$Current_Date = Get-Date -format "dd.MM.yyyy"
 
 $ImportCsv | ForEach-Object {
$a=$_.name
$b=$_.Description
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
        }




        $a
        $b
        $RowStart = $Row
        Write-Host "$A PC - Доступен!" -ForegroundColor Cyan
        Write-Host "Проверка компьютера " -ForeGroundColor Green $a "Компьютер" | Out-File C:\Test\Comp\$a.txt
        #Запись имени пользователя и имени ПК
        $InventoryFile.Cells.Item($Row, $Column) = $b
        $Column++
        $InventoryFile.Cells.Item($Row, $Column) = $a
        $Column++
        $InventoryFile.Cells.Item($Row, $Column) = $Current_Date
        $Column++
        # Получение сведений об ОС
        $Parameter  = Get-WmiObject -computername $a Win32_OperatingSystem | Select-Object csname, caption, Serialnumber, csdVersion  -ErrorAction Stop
        $InventoryFile.Cells.Item($Row, $Column) = $Parameter.caption
        $Column++
        
###########################################################################################

        #Модель процессора и прочая ересь
        "Процессор" | Out-File C:\Test\Comp\$a.txt -Append
        $Parameter = Get-WmiObject -computername $a Win32_Processor | Select-Object name, SocketDesignation, Description -ErrorAction Stop
        $InventoryFile.Cells.Item($Row, $Column) = $Parameter.name
        $Column++
        $InventoryFile.Cells.Item($Row, $Column) = $Parameter.SocketDesignation
        $Column++
        
###########################################################################################

        #Модель материнской платы
        "Материнская плата" | Out-File C:\Test\Comp\$a.txt -append
        $Parameter = Get-WmiObject -computername $a Win32_BaseBoard | Select-Object Manufacturer, Product, SerialNumber -ErrorAction Stop
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
        if($a -eq "W00-0626")
        {
           
        }
        else
        {
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

        } -ErrorAction Stop
        }
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
        $Parameter = Get-WmiObject -computername $a Win32_NetworkAdapter -Filter "NetConnectionStatus>0" | ForEach-Object `
            {
            $InventoryFile.Cells.Item($Row, $Column) = $_.Name
            $Column++
            $InventoryFile.Cells.Item($Row, $Column) = $_.MACAddress
            $Column++
            }
        }



$Range_Current = $InventoryFile.Range("B"+$Row,"Y"+$Row)
#$Range_Current.font.ColorIndex = 10
#$Range_Current.copy()

#Range_Previous = $InventoryFile.Range("B"+($Row-1).ToString(),"Y"+($Row-1).ToString())
#$InventoryFile.Paste($Range_Previous)


$Row++
$BadRow++
$BadColumn = 1
$RowFinish = $Row
$Column = 1

$Set = 1
#For($i = $RowStart;$i -lt $RowFinish;$i++)
#{
#    if($InventoryFile.Cells.Item($i+1, $Set) -eq "")
#    {
#      $Range = $InventoryFile.Range('A'+$i.ToString(),'A'+$I+1)
#      $Range.Merge()  
#    }
#}


$TEST

}
elseif ((Test-connection $a -count 1 -quiet) -ne "True")
{

        if($Configuration_Start -eq 0)
        {
        Write-Host "$A PC - НЕДОСТУПЕН"
        #Запись имени ПК и Имени пользователя
        $Bad_PC.Cells.Item($BadRow, $BadColumn) = $b
        $BadColumn++
        $Bad_PC.Cells.Item($BadRow, $BadColumn) = $a
        $BadColumn++

# Заполнение Недоступных ПК
        $Check = $Bad_PC.UsedRange.find("$a")
        $BadColumn = $Check.Column
        $BadColumn++
        if($Bad_PC.Cells.Item($Check.Row,$BadColumn).Text -eq "ДОСТУПЕН" -or $Bad_PC.Cells.Item($Check.Row,$BadColumn).Value2 -eq $Null)
        {
            $RRW = $Check.Row
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
            $Bad_PC.Cells.Item($Check.Row,$BadColumn).Formula = "=IF(C$RRW=`"`Недоступен`"`,DATEDIF(D$RRW,F$RRW,`"`d`"`),`"`")"
        }

        $BadRow++
        $BadColumn = 1
        }
        elseif($Configuration_Start -eq 1)
        {
            $Check = $Bad_PC.UsedRange.find($a)
            $BadColumn = $Check.Column
            $BadColumn++
        if($Bad_PC.Cells.Item($Check.Row,$BadColumn).Text -eq "ДОСТУПЕН" -or $Bad_PC.Cells.Item($Check.Row,$BadColumn).Value2 -eq $Null)
        {
            $RRW = $Check.Row
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
        }
        elseif($Bad_PC.Cells.Item($Check.Row,$BadColumn).Text -eq "ДОСТУПЕН")
        {
            $Bad_PC.Cells.Item($Check.Row,6) = $Current_Date
        }
        }

}
}




$Row--
$DataRangeInventory = $InventoryFile.Range(("A{0}" -f 1), ("AJ{0}" -f $Row))
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
$DataRangeInventory = $Change_History.Range(("A{0}" -f 1), ("AJ{0}" -f $Row_Change))
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
        for($i=4;$i -lt 36;$i++)
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

        if($ColOfCompare -eq 32)
        {
            $InventoryFile.Rows($Test+1).Delete()
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
    $WorkBooks.SaveAs("C:\Test\Инвентаризация.xlsx")
}
else
{
    "Первое Заполнение таблицы Выполнено!"
    $WorkBook.SaveAs("C:\Test\Инвентаризация.xlsx")
}




#[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
