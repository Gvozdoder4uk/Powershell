﻿#chcp 65001
#schtasks /Query /s "fobo-vrx-int2"
$CNUM =''
$CNAME = ''
$Contur = ''
$Server = ''
$CNT_TBL = @{"1"="vrq";"2"="vrx"}
$abs = ''

Function CONTUR([string]$CNT)
{
    
    if ($CNT -eq '1')
    {
        $CNAME = $CNT_TBL[$CNT]
        $СNUM = Read-Host "Введите номер контура VRQ"
        $Server = "fobo-$CNAME-int$CNUM"
        return $Server
        
    }
    elseif ($CNT -eq '2')
    {
        $CNAME = $CNT_TBL[$CNT]
        $CNUM = Read-Host "Введите номер контура VRX"
        $Server = "fobo-$CNAME-int$CNUM"
        return $Server
          
    }
}

Function INTERFACES([string]$SRV)
{
    $INTERFACE = Read-host "Введите номер интерфейса в формате XXX"
    $PORT = Get-ScheduledTask -CimSession $SRV -TaskName "FOBO*$INTERFACE*"
    $PORT.TaskName
    foreach ($p in $PORT)
    {
        Write-Host "Желаете запустить задачу?" $p.taskname
        if ($p.State -eq 'Disabled')
        {
        Write-Host "JOB В ГОВНЕ"
        }
        else
        {
        Write-host "JOB не в Говне"
        }
    }

}


do
{
    

Write-Host "|IIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIII|" -BackgroundColor DarkGreen
Write-Host "|ЧТО ЖЕЛАЕШЬ ПОВЕЛИТЕЛЬ?                              |
|1. Показать весь список задач на удаленной машине?   |
|2. Запустим конкретный джоб?                         |
|3. Выход из программы!                               |" -BackgroundColor DarkGreen 
Write-Host "|IIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIII|" -BackgroundColor DarkGreen

$SelectJhin = Read-Host "|Слушаю и повинуюсь " 

if ($SelectJhin -eq '0'){break}

if ($SelectJhin -eq '1')
{
Write-Host "IIIIIIIIIIIIIIIIIIIII
||Выберите контур: ||
|| 1: VRQ          ||
|| 2: VRX          ||
IIIIIIIIIIIIIIIIIIIII" -BackgroundColor DarkGray -ForegroundColor Black
    $Contur = Read-Host "Выберите контур"
    CONTUR($Contur)
    $Server
    #Get-ScheduledTask -CimSession $Server 
}
elseif ($SelectJhin -eq '2')
{
Write-Host "IIIIIIIIIIIIIIIIIIIII
||Выберите контур: ||
|| 1: VRQ          ||
|| 2: VRX          ||
IIIIIIIIIIIIIIIIIIIII" -BackgroundColor DarkCyan -ForegroundColor Black

    $Contur = Read-Host "Выберите контур"

    $Connect_to_Server = CONTUR($Contur)
    Write-Host "FUNC CONTUR FINISHED -  REULT ="$Server
    $Result_INT = INTERFACES($Connect_to_Server)

}

}
while ($SelectJhin -eq '0')



#foreach ($p in $PORT)
#{
#    Write-Host "Желаете запустить задачу?" $p.taskname
#    if ($p.State -eq 'Disabled')
#    {
#     Write-Host "JOB В ГОВНЕ"
#    }
#    else
#    {
#    Write-host "JOB не в Говне"
#    }
#}