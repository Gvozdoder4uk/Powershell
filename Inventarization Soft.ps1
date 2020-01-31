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
##
#Get-ADComputer -Filter * |
#Where-Object {$a=$_.name; $_.DistinguishedName -ne "CN=$a,OU=Computers,OU=Disabled,DC=rusagrotrans,DC=ru"} |
#Sort-Object NAME | Select-Object NAME | Export-csv C:\TEST\AllComputers.csv -NoTypeInformation


# Созадём объект Excel
$Excel = New-Object -ComObject Excel.Application

# Делаем его видимым
$Excel.Visible = $true

# Добавляем рабочую книгу
$WorkBook = $Excel.Workbooks.Add()

$InventoryFile = $WorkBook.Worksheets.Item(1)

$InventoryFile.Name = "Инвентаризация Инфраструктуры"
$InventoryFile.Cells.Item(1,1) = 'Инвентаризационные данные ПК в ЦО'
$InventoryFile.Cells.Item(1,1).Font.Size = 18
$InventoryFile.Cells.Item(1,1).Font.Bold = $true
$InventoryFile.Cells.Item(1,1).Font.ThemeFont = 1
$InventoryFile.Cells.Item(1,1).Font.ThemeColor = 4
$InventoryFile.Cells.Item(1,1).Font.ColorIndex = 55
$InventoryFile.Cells.Item(1,1).Font.Color = 8210719
$InventoryFile.Cells[1,1].Style.HorizontalAlignment = "Center"
$Range = $InventoryFile.Range('A1','N1')
$Range.Merge()



$InventoryFile.Cells.Item(2,1) = 'Сетевое имя'
$InventoryFile.Cells.Item(2,2) = 'OS'
$InventoryFile.Cells.Item(2,3) = 'Процессор'
$InventoryFile.Cells.Item(2,4) = 'Модель'
$InventoryFile.Cells.Item(2,5) = 'Материнская плата'
$InventoryFile.Cells.Item(2,6) = 'Модель'
$InventoryFile.Cells.Item(2,7) = 'Жесткий Диск'
$InventoryFile.Cells.Item(2,8) = 'Объем (Гб)'
$InventoryFile.Cells.Item(2,9) = 'ОЗУ (Гб)'
$InventoryFile.Cells.Item(2,10) = 'Частота (Mhz)'
$InventoryFile.Cells.Item(2,11) = 'Видеокарта'
$InventoryFile.Cells.Item(2,12) = 'Объем памяти (MB)'
$InventoryFile.Cells.Item(2,13) = 'Сетевая карта'
$InventoryFile.Cells.Item(2,13) = 'MAC Адрес'
$Range = $InventoryFile.Range("A2","N2")
$Range.Interior.ColorIndex = 15

$InventoryFile.Name = 'Инвентаризация ЦО'
# Выделяем жирным шапку таблицы
$InventoryFile.Rows.Item(2).Font.Bold = $true


$Row = 3
$Column = 1
$InitialRow = 2




import-csv c:\Test\AllComputers.csv | foreach {
$a=$_.name
if ((Test-Connection $a -count 1 -quiet) -eq "True")
{ 
        Write-Host "$A PC - Доступен!" -ForegroundColor Cyan
        Write-Host "Проверка компьютера " -ForeGroundColor Green $a "Компьютер" | Out-File C:\Test\Comp\$a.txt
        # Получение сведений об ОС
        $Parameter  = Get-WmiObject -computername $a Win32_OperatingSystem | Select-Object csname, caption, Serialnumber, csdVersion 
        #  Out-File C:\Test\Comp\$a.txt -Append
            #@{label="Наименование"; Expression={$_.caption}},
            #@{label="Версия"; Expression={$_.csdVersion}},
            #@{label="Серийный номер"; Expression={$_.SerialNumber}} -auto -Wrap | Out-File C:\Test\Comp\$a.txt -Append
        $InventoryFile.Cells.Item($Row, $Column) = $Parameter.csname
        $Column++
        $InventoryFile.Cells.Item($Row, $Column) = $Parameter.caption
        $Column++
        

        # Хрен пока пойми что
        #Get-WmiObject -computername $a Win32_ComputerSystemProduct | Select-Object UUID |
        #ft UUID -AutoSize | Out-File C:\Test\Comp\$a.txt -Append

        #Модель процессора и прочая ересь
        "Процессор" | Out-File C:\Test\Comp\$a.txt -Append
        $Parameter = Get-WmiObject -computername $a Win32_Processor | Select-Object name, SocketDesignation, Description
        $InventoryFile.Cells.Item($Row, $Column) = $Parameter.name
        $Column++
        $InventoryFile.Cells.Item($Row, $Column) = $Parameter.SocketDesignation
        $Column++

        <#Get-WmiObject -computername $a Win32_Processor | Select-Object name, SocketDesignation, Description |
        ft  @{label="Имя"; Expression={$_.name}} | Out-File C:\Test\Comp\$a.txt -Append
            #@{label="Разъем"; Expression={$_.SocketDesignation}},
            #@{label="Описание"; Expression={$_.Description}} -auto -Wrap | Out-File C:\Test\Comp\$a.txt -Append#>

        #Модель материнской платы
        "Материнская плата" | Out-File C:\Test\Comp\$a.txt -append
        $Parameter = Get-WmiObject -computername $a Win32_BaseBoard | Select-Object Manufacturer, Product, SerialNumber
        $InventoryFile.Cells.Item($Row, $Column) = $Parameter.Manufacturer
        $Column++
        $InventoryFile.Cells.Item($Row, $Column) = $Parameter.Product
        $Column++

        <#Get-WmiObject -computername $a Win32_BaseBoard | Select-Object Manufacturer, Product, SerialNumber |
        ft  @{label="Производитель"; Expression={$_.manufacturer}}  | Out-File C:\Test\Comp\$a.txt -Append
            @{label="Модель"; Expression={$_.Product}},
            #@{label="Серийный номер"; Expression={$_.SerialNumber}} -auto -Wrap | Out-File C:\Test\Comp\$a.txt -Append #>

        # HDD + SSD
        "Жесткие диски" | out-file C:\Test\Comp\$a.txt -Append
        $Parameter = Get-WmiObject -computername $a Win32_DiskDrive | Select-Object Model, Partitions, Size, interfacetype
        $InventoryFile.Cells.Item($Row, $Column) = $Parameter.Model
        $Column++
        $InventoryFile.Cells.Item($Row, $Column) = ($Parameter.Size/1GB).ToString("F00")
        $Column++

        <#Get-WmiObject -computername $a Win32_DiskDrive | Select-Object Model, Partitions, Size, interfacetype |
        ft @{Label="Модель"; Expression={$_.Model}},
           #@{Label="Количество разделов"; Expression={$_.Partitions}},
           @{Label="Размер (гб)"; Expression={($_.Size/1GB).tostring("F00")}},
           @{Label="Интерфейс"; Expression={$_.interfaceType}} -AutoSize -Wrap | Out-File C:\Test\Comp\$a.txt -Append #>

       <# Все логические диски.
        "Логические диски" | out-file C:\Test\Comp\$a.txt -Append
        Get-WmiObject -computername $a Win32_LogicalDisk -Filter "DriveType=3" | Select-Object DeviceID, FileSystem, Size, FreeSpace |
        ft  @{Label="Наименование"; Expression={$_.DeviceID}},
            @{Label="Файловая система"; Expression={$_.FileSystem}},
            @{Label="Размер (гб)"; Expression={($_.Size/1GB).tostring("F00")}},
            @{Label="Свободное место (гб)"; Expression={($_.FreeSpace/1GB).tostring("F00")}} -AutoSize -Wrap | Out-File C:\Test\Comp\$a.txt -Append #>
 
       # ОЗУ
        "Оперативная память" | out-file C:\Test\Comp\$a.txt -Append
        $ColemnTemp = $Column
        $RowTemp =$Row
        $ColOfElements = 0
        Get-WmiObject -computername $a Win32_Physicalmemory | ForEach-Object `
        {
            
            #Write-Host "Плашка "$_.Capacity
            $InventoryFile.Cells.Item($Row, $Column) = ([Math]::Round($_.Capacity/1GB, 2))
            #$Column++
            $Column++
            $InventoryFile.Cells.Item($Row, $Column) = $_.Speed
            $Row++
            $ColOfElements++


            $Column = $ColemnTemp
        }
        $Row = $RowTemp
        $Column+= $ColOfElements


        <#Get-WmiObject -computername $a Win32_Physicalmemory | Select-Object capacity, DeviceLocator |
        ft  @{Label="Размер (мб)"; Expression={($_.capacity/1MB).tostring("F00")}} -AutoSize -Wrap | Out-File C:\Test\Comp\$a.txt -append #>
            #@{Label="Расположение"; Expression={$_.DeviceLocator}} -AutoSize -Wrap | Out-File C:\Test\Comp\$a.txt -append
       
       # Видеокарта
        "Видеокарта" | out-file C:\Test\Comp\$a.txt -append
        $Parameter = Get-WmiObject -computername $a Win32_videoController | Select-Object name, AdapterRAM, VideoProcessor
        $InventoryFile.Cells.Item($Row, $Column) = $Parameter.name
        $Column++
        $InventoryFile.Cells.Item($Row, $Column) = ($Parameter.AdapterRAM/1MB).tostring("F00")
        $Column++

        <#Get-WmiObject -computername $a Win32_videoController |
        Select-Object name, AdapterRAM, VideoProcessor |
        ft @{Label="Наименование"; Expression={$_.name}},
           @{Label="Объем памяти (мб)"; Expression={($_.AdapterRAM/1MB).tostring("F00")}} -AutoSize -Wrap | Out-File C:\Test\Comp\$a.txt -append
           #@{Label="Видеопроцессор"; Expression={$_.VideoProcessor}} -AutoSize -Wrap | Out-File C:\Test\Comp\$a.txt -append #>

       # Сетевая Карта
        "Сетевая карта" | Out-File C:\Test\Comp\$a.txt -append
        $ColemnTemp = $Column
        $RowTemp =$Row
        $ColOfElements = 0
        
        $OS.Caption
        $OS=Get-WmiObject -computername $a Win32_OperatingSystem | foreach {$_.caption}
        if ($OS -eq "Microsoft Windows 2000 Professional")
        {

        $Parameter = Get-WmiObject -computername $a Win32_NetworkAdapterConfiguration -Filter "DHCPEnabled=True" | ForEach-Object `
        {

            $InventoryFile.Cells.Item($Row, $Column) = $_.caption
            $Column++
            $InventoryFile.Cells.Item($Row, $Column) = $_.MACAddress
            $Row++
            $ColOfElements++


            $Column = $ColemnTemp 
        }


        <#Get-WmiObject -computername $a Win32_NetworkAdapterConfiguration -Filter "DHCPEnabled=True" |
        Select-Object caption,MACaddress |
        ft  @{Label="Наименование"; Expression={$_.caption}},
            @{Label="MAC адрес"; Expression={$_.MACAddress}} -AutoSize -Wrap | Out-File C:\Test\Comp\$a.txt -append
        #>
        }
        else
        {
        $Parameter = Get-WmiObject -computername $a Win32_NetworkAdapter -Filter "NetConnectionStatus>0" | ForEach-Object `
        {
           #Write-Host "Плашка "$_.Capacity
            $InventoryFile.Cells.Item($Row, $Column) = $_.Name
            #$Column++
            $Column++
            $InventoryFile.Cells.Item($Row, $Column) = $_.MACAddress
            $Row++
            $ColOfElements++


            $Column = $ColemnTemp 
        }

       <# Get-WmiObject -computername $a Win32_NetworkAdapter -Filter "NetConnectionStatus>0" | Select-Object name, AdapterType, MACAddress |
        ft @{Label="Наименование"; Expression={$_.name}},
           @{Label="MAC адрес"; Expression={$_.MACAddress}},
           @{Label="Тип"; Expression={$_.AdapterType}} -AutoSize -Wrap | Out-File C:\Test\Comp\$a.txt -append
           #>
        }
        
        if($RowTemp > $Row)
        {
         $Row = $RowTemp + 1
        }
        else
        {
        }
        $Column = 1




}
elseif ((Test-connection $a -count 1 -quiet) -ne "True")
{
        Write-Host "$A PC - НЕДОСТУПЕН"
}
}





$Row--
$DataRange = $InventoryFile.Range(("A{0}" -f $InitialRow), ("N{0}" -f $Row))
7..12 | ForEach-Object `
{
    $DataRange.Borders.Item($_).LineStyle = 1
    $DataRange.Borders.Item($_).Weight = 2
}

$UsedRange = $InventoryFile.UsedRange
$UsedRange.EntireColumn.AutoFit() | Out-Null



#if ((Get-WmiObject -computername $a Win32_OperatingSystem) -eq $null)
#{
#Write-Host "$a Не определена ОС"
#}
