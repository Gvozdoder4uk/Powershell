##[Ps1 To Exe]
##
##Kd3HDZOFADWE8uO1
##Nc3NCtDXTlaDjofG5iZk2UD9fW4kZcyVhZKo04+w8OvoqBntQZsYQEB22x3zFlm4Sr89U/Mct9RcehssJvEOr6HTE+imSe8Yh+96efbAr7EmdQ==
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
##L8/UAdDXTlaDjofG5iZk2UD9fW4kZcyVhZKi14qo8PrQmC3cSpErelFlgCD/AVjzXOoXNQ==
##Kc/BRM3KXhU=
##
##
##fd6a9f26a06ea3bc99616d4851b372ba


#PO For Check Services Status on Remote Servers

DO
{
cls

$Choice = ""
Function CheckServices([string]$Server)
{
    $Wildfly = Get-Service -Name Wildfly -ComputerName $Server -ErrorAction SilentlyContinue

    if ($Wildfly.status -eq "running")
    {
     Write-Host "============================================================="
     Write-Host "Служба Wildfly запущена и работает!" -ForegroundColor Green
     Write-Host "============================================================="
    }
    else
    {
     Write-Warning "============================================================="
     Write-Warning "Служба Wildfly НЕ ЗАПУЩЕНА!!! "
     Write-Warning "============================================================="
     Write-Host "-------------------------------------------------------------"
     Write-Host " Не желаете запустить службу сейчас? Y/N" -ForegroundColor Green
     Write-Host "-------------------------------------------------------------"
     $Choice = Read-Host " Сделайте выбор: Y/N"
     Switch ($Choice) 
     {
        Y {Get-Service -Name Wildfly -ComputerName $Server | Start-Service
           Write-Warning "Выполняется запуск службы Wildfly"}
        N {Write-Warning "Служба Wildfly останется остановленной!"}

     }
    }


    $NTSwincash = Get-Service -Name "NTSwincash distributor" -ComputerName $Server -ErrorAction SilentlyContinue
    if ($NTSwincash.status -eq "running")
    {
     Write-Host "============================================================="
     Write-Host "Служба NTSwincash запущена и работает!" -ForegroundColor Green
     Write-Host "============================================================="
     Write-Host " Не желаете ПЕРЕзапустить службу сейчас? Y/N" -ForegroundColor Green
     $SET = Read-Host " Сделайте выбор: Y/N"
     Switch ($SET)
     {
        Y {Get-Service -Name "NTSwincash distributor" -ComputerName $Server | Restart-Service
           Write-Warning "Выполняется Перезапуск службы NTSwincash"}
        N {Write-Warning "Служба NTSwincash работает!"}
     }
    }
    else
    {
     Write-Warning "============================================================="
     Write-Warning "Служба NTSwincash НЕ ЗАПУЩЕНА!!! "
     Write-Warning "============================================================="
     Write-Host "-------------------------------------------------------------"
     Write-Host " Не желаете запустить службу сейчас? Y/N" -ForegroundColor Green
     Write-Host "-------------------------------------------------------------"
     $Choice = Read-Host " Сделайте выбор: Y/N"
     
     Switch ($Choice) 
     {
        Y {Get-Service -Name "NTSwincash distributor" -ComputerName $Server | Start-Service
           Write-Warning "Выполняется запуск службы NTSwincash"}
        N {Write-Warning "Служба NTSwincash останется остановленной!"}
     }
        
    }
 pause
}


Write-Host "=====================
Выберите контур:   ||
|| 1: VRQ          ||
|| 2: VRX          ||
|| 3: VRA          ||
|| 0: EXIT         ||
=====================" -BackgroundColor DarkYellow -ForegroundColor Black
$Contur = Read-Host "Выберите контур"

if($Contur -eq 0){ break } 
#FORCE BREAK AFTER CHOICE

if ($Contur -eq 1)
{
$shopName = Read-Host "Введите магазин или нажмите Enter"
    if ($shopName -eq '')
    {
    $MAG = 0
    $conturName = "vrq"
    $conturNumber = Read-Host "Укажите номер контура"
    }
    else
    {
    $MAG = 1
    $conturName = "vrq"
    }
}
elseif ($Contur -eq 2)
{
$shopName = Read-Host "Введите магазин или нажмите Enter"
    if ($shopName -eq '')
    {
    $MAG = 0
    $conturName = "vrx"
    $conturNumber = Read-Host "Укажите номер контура"
    }
    else
    {
    $MAG = 1
    $conturName = "vrx"
    }
}
elseif ($Contur -eq 3)
{
$shopName = Read-Host "Введите магазин или нажмите Enter"
    if ($shopName -eq '')
    {
    $MAG = 0
    $conturName = "vra"
    $conturNumber = Read-Host "Укажите номер контура"
    }
    else
    {
    $MAG = 1
    $conturName = "vra"
    }
}



if ($MAG -eq 0)
{
    $Server = "fobo-$conturName-ajb$conturNumber"
    Write-Host "Вы выбрали сервер:" $Server -ForegroundColor green 
    CheckServices($Server) 
}
elseif ($MAG -eq 1)
{
    $Server = "fobo-$conturName-a$shopName"
    Write-Host "Вы выбрали сервер:" $Server -ForegroundColor green
    CheckServices($Server)
}
} while ($Contur -ne 0)