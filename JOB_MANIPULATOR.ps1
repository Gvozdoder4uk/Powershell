#chcp 65001
#schtasks /Query /s "fobo-vrx-int2"
$CNUM =''
$CNAME = ''
$Contur = ''
$Server = ''
$CNT_TBL = @{"1"="vrq";"2"="vrx"}

Function CONTUR([string]$CNT)
{
    
    if ($CNT -eq '1')
    {
        $CNAME = $CNT_TBL[$CNT]
        $COUNT = Read-Host "Введите номер контура VRQ"
        $Server = "fobo-$CNAME-int$COUNT"
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
        $tst = $p.taskname
        Write-Host "Желаете запустить задачу?" $p.taskname
        $Choice = Read-Host "Сделайте выбор Y/N "
        Switch ($Choice) 
        {
            Y {
            if ($p.State -eq 'Disabled')
            {
            Write-Host "JOB в статусе:" $p.State
            Enable-ScheduledTask -CimSession $SRV -TaskName $p.TaskName
            Start-ScheduledTask -CimSession $SRV -TaskName $p.TaskName
            Start-Sleep -Seconds 2
            $p.State
            }
            else
            {
            Write-host "JOB в статусе:" $p.State
            Start-ScheduledTask -CimSession $SRV -TaskName $p.TaskName
#           Disable-ScheduledTask -CimSession $SRV -TaskName $p.TaskName
            }}
         
            N {Write-Warning "JOB $tst не будет запущен"}
        }
        
       }

}


Function ENABLE_DISABLE([string]$SRV)
{
   $INTERFACE = Read-host "Введите номер интерфейса в формате XXX"
    $PORT = Get-ScheduledTask -CimSession $SRV -TaskName "FOBO*$INTERFACE*"
    $PORT.TaskName
    foreach ($p in $PORT)
    {
        $tst = $p.taskname
        Write-Host "Текущий статус JOB" $p.taskname " - " $p.State -BackgroundColor DarkGreen
        Write-Host "        IIIIIIIIIIIIIIIIIIIII
        |1.Включить задачу  |
        |2.Остановить задачу|
        |3.Ничего не делать |
        IIIIIIIIIIIIIIIIIIIII" -BackgroundColor DarkBlue
        $Choice = Read-Host "Сделайте выбор "
        Switch ($Choice) 
        {
            '1' {

            Enable-ScheduledTask -CimSession $SRV -TaskName $p.TaskName
            Start-Sleep -Seconds 2
            $p.TaskName
            $p.State
                }
         
            '2' {
            Write-Warning "JOB $tst будет отключен"
            Disable-ScheduledTask -CimSession $SRV -TaskName $p.TaskName
            Start-Sleep -Seconds 2
            $p.TaskName
            $p.State
              }
             
            '3' {
             Write-Warning "Job $tst останется в статусе - " $p.State
                }
              
        }
    } 

}


do
{
    

Write-Host "|IIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIII|" -BackgroundColor DarkGreen
Write-Host "|ЧТО ЖЕЛАЕШЬ ПОВЕЛИТЕЛЬ?                              |
|1. Показать весь список задач на удаленной машине?   |
|2. Запустим конкретный джоб?                         |
|3. Отключить или включить джоб?                      |
|0. Выход из программы!                               |" -BackgroundColor DarkGreen 
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
    $Server = CONTUR($Contur)
    Get-ScheduledTask -CimSession $Server -TaskName "FOBO*" | Format-Table TaskName,State
    pause
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
#    Write-Host "FUNC CONTUR FINISHED -  REULT ="$Server
    $Result_INT = INTERFACES($Connect_to_Server)
    pause
}

elseif ($SelectJhin -eq '3')

{
 Write-Host "IIIIIIIIIIIIIIIIIIIII
||Выберите контур: ||
|| 1: VRQ          ||
|| 2: VRX          ||
IIIIIIIIIIIIIIIIIIIII" -BackgroundColor DarkCyan -ForegroundColor Black

    $Contur = Read-Host "Выберите контур"
    $Connect_to_Server = CONTUR($Contur)
    ENABLE_DISABLE($Connect_to_Server)
     
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
