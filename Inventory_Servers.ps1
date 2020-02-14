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

#Get-ADComputer -Filter {Name -Like "msk*"} -Properties Description |
#Where-Object {$a=$_.name; $_.DistinguishedName -ne "CN=$a,OU=Computers,OU=Disabled,DC=rusagrotrans,DC=ru"} |
#Sort-Object NAME | Select-Object NAME,DESCRIPTION | Export-csv -NoTypeInformation C:\Servers\AllComputers.csv  -Encoding UTF8

# Инициализация Конфигурационного Файла:
$Config_File = "C:\Servers\cfg.ini"
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
$Change_History.Cells.Item(1,5) = 'Процессор 1'
$Change_History.Cells.Item(1,6) = 'Слот'
$Change_History.Cells.Item(1,7) = 'Процессор 2'
$Change_History.Cells.Item(1,8) = 'Слот'
$Change_History.Cells.Item(1,9) = 'Процессор 3'
$Change_History.Cells.Item(1,10) = 'Слот'
$Change_History.Cells.Item(1,11) = 'Процессор 4'
$Change_History.Cells.Item(1,12) = 'Слот'
#HDD START 13
$Change_History.Cells.Item(1,13) = 'HDD 1'
$Change_History.Cells.Item(1,14) = 'HDD 2'
$Change_History.Cells.Item(1,15) = 'HDD 3'
$Change_History.Cells.Item(1,16) = 'HDD 4'
$Change_History.Cells.Item(1,17) = 'HDD 5'
$Change_History.Cells.Item(1,18) = 'HDD 6'
$Change_History.Cells.Item(1,19) = 'HDD 7'
$Change_History.Cells.Item(1,20) = 'HDD 8'
$Change_History.Cells.Item(1,21) = 'HDD 9'
$Change_History.Cells.Item(1,22) = 'HDD 10'
#Column OZY START 23
$Change_History.Cells.Item(1,23) = 'Суммарно ОЗУ (Гб)'
$Change_History.Cells.Item(1,24) = 'Тип Памяти'
#Column Availabilyty  33



$Range = $Change_History.Range("A1","X1")
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
$InventoryFile.Cells.Item(1,5) = 'Процессор 1'
$InventoryFile.Cells.Item(1,6) = 'Слот'
$InventoryFile.Cells.Item(1,7) = 'Процессор 2'
$InventoryFile.Cells.Item(1,8) = 'Слот'
$InventoryFile.Cells.Item(1,9) = 'Процессор 3'
$InventoryFile.Cells.Item(1,10) = 'Слот'
$InventoryFile.Cells.Item(1,11) = 'Процессор 4'
$InventoryFile.Cells.Item(1,12) = 'Слот'
#HDD START 13
$InventoryFile.Cells.Item(1,13) = 'HDD 1'
$InventoryFile.Cells.Item(1,14) = 'HDD 2'
$InventoryFile.Cells.Item(1,15) = 'HDD 3'
$InventoryFile.Cells.Item(1,16) = 'HDD 4'
$InventoryFile.Cells.Item(1,17) = 'HDD 5'
$InventoryFile.Cells.Item(1,18) = 'HDD 6'
$InventoryFile.Cells.Item(1,19) = 'HDD 7'
$InventoryFile.Cells.Item(1,20) = 'HDD 8'
$InventoryFile.Cells.Item(1,21) = 'HDD 9'
$InventoryFile.Cells.Item(1,22) = 'HDD 10'
#Column OZY START 23
$InventoryFile.Cells.Item(1,23) = 'Суммарно ОЗУ (Гб)'
$InventoryFile.Cells.Item(1,24) = 'Тип Памяти'


$InventoryFile.Name = 'Инвентаризация Серверов'
$Range = $InventoryFile.Range("A1","X1")
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
    $FilePath = "C:\Servers\MyExcel.xlsx"

    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $true
    $Workbooks = $Excel.Workbooks.Open($FilePath)
    

    #Sheets("data").Names("_FilterDatabase").Delete


# Main Window CO Selection
$InventoryFile = $WorkBooks.Worksheets.Item(1)

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




$ImportCsv = import-csv C:\Servers\AllComputers.csv

$Current_Date = Get-Date -format "dd.MM.yyyy"
 
 $ImportCsv | ForEach-Object {
$a=$_.name
$b=$_.Description
if ((Test-Connection $a -count 1 -quiet) -eq "True")
{ 
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

        $a
        $b
        $RowStart = $Row
        Write-Host "$A PC - Доступен!" -ForegroundColor Cyan
        Write-Host "Проверка компьютера " -ForeGroundColor Green $a "Компьютер"
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
        $CheckSlot = Get-WmiObject -computername $a Win32_Processor | Select-Object name, SocketDesignation, Description | Select -First 1
        #Модель процессора и прочая ересь
        if($Parameter -like "*2003*" -and $CheckSlot.SocketDesignation -like "*None*")
        {
            Get-WmiObject -computername $a Win32_Processor | Select-Object name, SocketDesignation, Description | Select -First 1 | ForEach-Object `
            {
                    $InventoryFile.Cells.Item($Row, $Column) = $_.name
                    $Column++
                    $InventoryFile.Cells.Item($Row, $Column) = $_.SocketDesignation
                    $Column++
            } -ErrorAction Stop            
        }
        elseif($Parameter -like "*2003*" -and $CheckSlot.SocketDesignation -notlike "*None*")
        {
            Get-WmiObject -computername $a Win32_Processor | Select-Object name, SocketDesignation, Description | Select -First 2 | ForEach-Object `
            {
                    $InventoryFile.Cells.Item($Row, $Column) = $_.name
                    $Column++
                    $InventoryFile.Cells.Item($Row, $Column) = $_.SocketDesignation
                    $Column++
            } -ErrorAction Stop            
        }
        else
        {
        Get-WmiObject -computername $a Win32_Processor | Select-Object name, SocketDesignation, Description | ForEach-Object `
        {
                    $InventoryFile.Cells.Item($Row, $Column) = $_.name
                    $Column++
                    $InventoryFile.Cells.Item($Row, $Column) = $_.SocketDesignation
                    $Column++
        } -ErrorAction Stop
        }
        $Column=13
###########################################################################################

###########################################################################################

        # HDD + SSD
        "Жесткие диски" 

        $ColemnTemp = $Column
        $RowTemp = $Row
        $ColOfElements = 0
        Get-WmiObject -computername $a Win32_DiskDrive | Where-Object {$_.Model -notlike "*usb*" -or $_.Model -notlike "*USB*"}| ForEach-Object `
        {
            $SizeDisk = ($_.Size/1GB)
            
            if($SizeDisk -gt 999)
            {
                $SizeDisk =($_.Size/1TB)
                $SizeDisk = $SizeDisk.ToString("F00")
                $SizeDisk = "$SizeDisk TB"
                $InventoryFile.Cells.Item($Row, $Column) = $SizeDisk
                $Column++
                $ColOfElements++
                $SizeDisk = $Null
            }
            else
            {

                if($SizeDisk -lt 50)
                {
    
                }
                else
                {
                $SizeDisk = $SizeDisk.ToString("F00")
                $SizeDisk = "$SizeDisk GB"
                $InventoryFile.Cells.Item($Row, $Column) = $SizeDisk
                $Column++
                $ColOfElements++
                $SizeDisk = $Null
                }

        }
            
            


        } -ErrorAction Stop

        $Row = $RowTemp
        $Column=23
        
###########################################################################################
       # ОЗУ
        "Оперативная память" 
        $ColemnTemp = $Column
        $RowTemp =$Row
        $ColOfElements = 0

        

        $T = Get-WmiObject -computername MSKTS5 Win32_Physicalmemory | Measure-Object -Property capacity -Sum
        $T= $T.Sum/1GB.ToString("F00")
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
        $Row = $RowTemp
        #$Column=20
###########################################################################################

###########################################################################################
        



$Range_Current = $InventoryFile.Range("B"+$Row,"Y"+$Row)

$Row++
$BadColumn = 1
$RowFinish = $Row
$Column = 1

$Set = 1

# Formula Excel
$Formula = "=IF(C$RRW=`"`Недоступен`"`,DATEDIF(D$RRW,F$RRW,`"`d`"`),`"`")"


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
            Write-Host "$A PC - НЕДОСТУПЕН"
            #Панель дислокаций	W00-0289
            $Check = $null
            $Check = $Bad_PC.UsedRange.find($a)
            if($Check.Text -eq "")
            {
                Write-Host "$A PC - НЕДОСТУПЕН"
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
            elseif($Check.Text -ne "")
            {
                $Check_Col = $False
                $Target = $Check
                $First = $Target
                Do
                {
                    Write-Host $Target.Row
                    # Взяли строку
                    #
                    # Cравниваем чекируем
                       if(($Bad_PC.Cells.Item($Target.Row,1).Text -eq $b) -and ($Bad_PC.Cells.Item($Target.Row,2).Text -eq $a))
                        {
                            "Проставляем дату yf CОВПАДЕНИИ"
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




$Row--
$DataRangeInventory = $InventoryFile.Range(("A{0}" -f 1), ("X{0}" -f $Row))
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

$Row_Change--
$DataRangeInventory = $Change_History.Range(("A{0}" -f 1), ("X{0}" -f $Row_Change))
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
        for($i=5;$i -lt 30;$i++)
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
    #$WorkBooks.SaveAs("C:\Servers\Инвентаризация.xlsx")
}
else
{
    "Первое Заполнение таблицы Выполнено!"
    
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

    }
    elseif(($Work_Range.Cells.Item($RRP,1).Formula -ne $Work_Range.Cells.Item($RRP+1,1).Formula) -and ($Work_Range.Cells.Item($RRP,2).Formula -eq $Work_Range.Cells.Item($RRP+1,2).Formula))
    {
        if($Work_Range.Cells.Item($RRP,3).Formula -lt $Work_Range.Cells.Item($RRP+1,3).Formula -and ($Work_Range.Cells.Item($RRP,3) -eq -4142 -or 0 ))
        {
                $INDEX = Get-Random -Minimum 2 -Maximum 24
                $Work_Range.Range("A$RRP","AH$RRP").Interior.ColorIndex = $INDEX
                $Work_Range.Cells.Item($RRP+1,28).Interior.ColorIndex = $INDEX 
        }
        elseif($Work_Range.Cells.Item($RRP,3).Formula -gt $Work_Range.Cells.Item($RRP+1,3).Formula -and ($Work_Range.Cells.Item($RRP+1,3) -eq -4142 -or 0 ))
        {
                $Set = $RRP+1
                $INDEX = Get-Random -Minimum 2 -Maximum 24
                $Work_Range.Range("A$SET","AH$SET").Interior.ColorIndex = $INDEX 
                $Work_Range.Cells.Item($RRP,28).Interior.ColorIndex = $INDEX  

        }
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

$InventoryFile.Range("Y1:AZ200").Delete()

if($Configuration_Start -eq 0){
$WorkBook.SaveAs("C:\Servers\Инвентаризация_Cерверов.xlsx")
}
else
{

$WorkBooks.SaveAs("C:\Servers\Инвентаризация_Серверов.xlsx")
}
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)



