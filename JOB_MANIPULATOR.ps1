#chcp 65001
#schtasks /Query /s "fobo-vrx-int2"
Write-Host "|INT|INT|INT|INT|INT|INT|INT|INT|INT|INT|INT||INT||INT|" -BackgroundColor DarkGreen
Write-Host "|Что будем делать повелитель?:                        |
|1. Показать весь список задач на удаленной машине?   |
|2. Запустим конкретный джоб?                         |
|3. Выход из программы!                               |" -BackgroundColor Green 
Write-Host "|INT|INT|INT|INT|INT|INT|INT|INT|INT|INT|INT||INT||INT|" -BackgroundColor DarkGreen

$SelectJhin = Read-Host "Слушаю и повинуюсь: " 


if ($SelectJhin -eq '1')
{
  Get-ScheduledTask -CimSession $Server 
}
elsif ($SelectJhin -eq '2')
{
Write-Host "=====================
Выберите контур:   ||
|| 1: VRQ          ||
|| 2: VRX          ||
=====================" -BackgroundColor DarkYellow -ForegroundColor Black
$Contur = Read-Host "Выберите контур"
    if ($Contur -eq 1)
    {
        $CNAME = "vrq"
        $СNUM = Read-Host "Введите номер контура"
        $Server = "fobo-$CNAME-int$CNUM"
        
    }
    elseif ($Contur -eq 2)
    {
        $CNAME = "vrx"
        $CNUM = Read-Host "Введите номер контура"
        $Server = "fobo-$CNAME-int$CNUM"
          
    }
}
$INTERFACE = "086"

$PORT = Get-ScheduledTask -CimSession fobo-vrx-int2 -TaskName "FOBO*$INTERFACE*"
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
