##[Ps1 To Exe]
##
##Kd3HDZOFADWE8uO1
##Nc3NCtDXTlaDjofG5iZk2UD9fW4kZcyVhZKo04+w8OvoqBnbRZMaQmt6lyfDFE6ySf4XavcQtcMVahQpIPw0s+CETr/xC6sJnYM=
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
##L8/UAdDXTlaDjofG5iZk2UD9fW4kZcyVhZKi14qo8PrQiQDqfbsmfRREsxvKLWiYarwXTfB1
##Kc/BRM3KXhU=
##
##
##fd6a9f26a06ea3bc99616d4851b372ba
#Create by Fokin Oleg 16.10.2019
##################################################################################################################
#Программа для перезапуска служб на удаленных серверах тестовых контуров VRQ,VRX,VRA
#1.Согласно меню предлагается выбор Релиза (ввод числовой)
#2.Согласно меню предлагается выбор контура (ввод числовой)
#3.В зависимости от того нужен центральный сервер или сервер магазина предлагается меню выбора
# a) При вводе номера магазина будет сформирован магазин для соединения с сервером магазина
# б) При нажатии клавиши Enter поле останется пустым и будет предложено указание номера контура.
#4.Выполняется отработка функций согласно сформированному шаблону подключения (fobo-(контрур)-(ajb№ при выборе контура)(а№№№ при выборе магазина)
#5.Выполнение остановки процесса java.exe и остановка службы Wildfly удаленного ПК.
#6.Выполнение очистки файлов NTSwincash\jboss\wildfly10\standalone\deployments\ по шаблону "*.backup","*.deployed","*.failed"
#7.Выполнение удаление директорий \NTSwincash\jboss\wildfly10\standalone\data\  
#                                 \NTSwincash\jboss\wildfly10\standalone\tmp\
#8.Запуск службы Wildfly
#9.Завершение выполнения программы
##################################################################################################################







chcp 65001
<#Проведение восстановления релиза после падения сервисов#>
#=======================================================================================================================
$Server = ''
$shopName = ''
$MAG = ''
$service =''
#$PATH = "C`$\NTSwincash\jboss\wildfly10\standalone\deployments"
#========================================================================================================================================================================================
#БЛОК ФУНКЦИИ
Write-Host "SOFTINA Для проверки и рестарта службы WILDFLY"
Function CheckRelease([string]$SRV)
{
   $BUILD = Select-String -Path "\\$SRV\C`$\NTSwincash\build.txt" -SimpleMatch "/19","/20","/21","/22"
   $BUILD -split ("/")
}
   

Function KillWildfly([string]$SRV)
{
    Get-Process -Name java -ComputerName $SRV -ErrorAction SilentlyContinue | Format-List
    $Status = Get-WmiObject -Class Win32_Process -ComputerName $SRV -Filter "name='java.exe'"
    $Status.terminate()
    Get-Service -Name Wildfly -ComputerName $SRV -ErrorAction SilentlyContinue | Stop-Service
}

Function Release19_changes ([string]$SRV)
{
        Get-ChildItem -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\*" -Include "*.backup","*.deployed","*.failed" | Remove-Item
        #Remove-Item -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\tmp\" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\data\" -Recurse -Force -ErrorAction SilentlyContinue

}
Function Release20_changes ([string]$SRV)
{
        Get-ChildItem -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\*" -Include "*.backup","*.deployed",".readclaim*.","*.failed","*.facade*","*.transfer*" | Remove-Item
        #Remove-Item -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\tmp\" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\data\" -Recurse -Force -ErrorAction SilentlyContinue
}

Function Release21_changes ([string]$SRV)
{
        Get-ChildItem -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\*" -Include "*.backup","*.deployed",".readclaim*.","*.failed","*.facade*","*.transfer*" | Remove-Item        
        #Remove-Item -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\tmp\" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\data\" -Recurse -Force -ErrorAction SilentlyContinue

}

Function Release22_changes ([string]$SRV)
{
        Get-ChildItem -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\*" -Include "*.backup","*.deployed",".readclaim*.","*.failed","*.facade*","*.transfer*" | Remove-Item
        #Remove-Item -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\tmp" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\data" -Recurse -Force -ErrorAction SilentlyContinue


}
#========================================================================================================================================================================================

#МЕНЮ
Write-Host "=====================
||Выберите релиз:  ||
|| 1: 19.0.0       ||
|| 2: 20.0.0       ||
|| 3: 21.0.0       ||
|| 4: 22.0.0       ||
=====================" -BackgroundColor DarkBlue

$select_menu = Read-Host "Выберите пункт меню"

#Write-Host "Вы выбрали релиз" $select_menu -ForegroundColor DarkYellow
Write-Host "=====================
Выберите контур:   ||
|| 1: VRQ          ||
|| 2: VRX          ||
|| 3: VRA          ||
=====================" -BackgroundColor DarkGreen
$Contur = Read-Host "Выберите контур"

#Заполнение Контура
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


if ($select_menu -eq 1)
{
    Write-Host "Релиз 19.0.0"
    if ($MAG -eq 0)
    {
    $Server = "fobo-$conturName-ajb$conturNumber"
    
    Write-Host "Вы выбрали сервер:" $Server -ForegroundColor green
     KillWildfly($Server)
     Start-Sleep -Seconds 3
     Release19_changes($Server)
     Start-Sleep -Seconds 5
     Get-Service -Name Wildfly -ComputerName $server | Start-Service    
    }

    elseif ($MAG -eq 1)
    {
    $Server = "fobo-$conturName-a$shopName"
    Write-Host "Вы выбрали сервер:" $Server -ForegroundColor green
    KillWildfly($Server)
    Start-Sleep -Seconds 3
    Release19_changes($server)
    Start-Sleep -Seconds 5
    Get-Service -Name Wildfly -ComputerName $server | Start-Service
    }
}
elseif ($select_menu -eq 2)
{
    Write-Host "Релиз 20.0.0"
    if ($MAG -eq 0)
    {

    $Server = "fobo-$conturName-ajb$conturNumber"
    Write-Host "Вы выбрали сервер:" $Server -ForegroundColor green
    KillWildfly($Server)
    Start-Sleep -Seconds 3
    Release20_changes($server)
    Start-Sleep -Seconds 5
    Get-Service -Name Wildfly -ComputerName $server | Start-Service
    }
    elseif ($MAG -eq 1)
    {

    $Server = "fobo-$conturName-a$shopName"
    Write-Host "Вы выбрали сервер:" $Server -ForegroundColor green
    KillWildfly($Server)
    Start-Sleep -Seconds 3
    Release20_changes($server)
    Start-Sleep -Seconds 5
    Get-Service -Name Wildfly -ComputerName $server | Start-Service 
    }
}
elseif ($select_menu -eq 3)
{
    Write-Host "Релиз 21.0.0"
    if ($MAG -eq 0)
    {
    $Server = "fobo-$conturName-ajb$conturNumber"
    Write-Host "Вы выбрали сервер:" $Server -ForegroundColor green
    KillWildfly($Server)
    Start-Sleep -Seconds 3
    Release21_changes($Server)
    Start-Sleep -Seconds 5
    Get-Service -Name Wildfly -ComputerName $server | Start-Service


    }
    elseif ($MAG -eq 1)
    {
    $Server = "fobo-$conturName-a$shopName"
    Write-Host "Вы ввели" $Server
    Write-Host "Вы выбрали сервер:" $Server -ForegroundColor green
    KillWildfly($Server)
    Start-Sleep -Seconds 3
    Release21_changes($Server)
    Start-Sleep -Seconds 5
    Get-Service -Name Wildfly -ComputerName $server | Start-Service
    }
}
if ($select_menu -eq 4)
{
    Write-Host "Релиз 22.0.0"
    if ($MAG -eq 0)
    {
    $Server = "fobo-$conturName-ajb$conturNumber"
    
    Write-Host "Вы выбрали сервер:" $Server -ForegroundColor green
        KillWildfly($Server)
        Start-Sleep -Seconds 3
        Release22_changes($server)
        Start-Sleep -Seconds 5
        Get-Service -Name Wildfly -ComputerName $server | Start-Service
    }

    elseif ($MAG -eq 1)
    {
    $Server = "fobo-$conturName-a$shopName"
    Write-Host "Вы выбрали сервер:" $Server -ForegroundColor green
        KillWildfly($Server)
        Start-Sleep -Seconds 3
        Release22_changes($server)
        Start-Sleep -Seconds 5
        Get-Service -Name Wildfly -ComputerName $server | Start-Service
    }
}
pause