#PO For Check Services Status on Remote Servers



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
        Write-Warning "Служба Wildfly НЕ ЗАПУЩЕНА!!! "
    }


    $NTSwincash = Get-Service -Name "NTSwincash distributor" -ComputerName $Server -ErrorAction SilentlyContinue
    if ($NTSwincash.status -eq "running")
    {
     Write-Host "============================================================="
     Write-Host "Служба NTSwincash запущена и работает!" -ForegroundColor Green
     Write-Host "============================================================="
    }
    else
    {
        Write-Warning "Служба NTSwincash НЕ ЗАПУЩЕНА!!! "
    }
    
}


Write-Host "=====================
Выберите контур:   ||
|| 1: VRQ          ||
|| 2: VRX          ||
|| 3: VRA          ||
=====================" -BackgroundColor DarkYellow
$Contur = Read-Host "Выберите контур"

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