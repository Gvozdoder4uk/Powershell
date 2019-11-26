##[Ps1 To Exe]
##
##Kd3HDZOFADWE8uO1
##Nc3NCtDXTlaDjofG5iZk2UD9fW4kZcyVhZKo04+w8OvoqBn6RogaRFV5giH5V1m5FMEDQPQHscEdUAVnGuYC7rvEEuK6CLYLgehsJeyNqdI=
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
##L8/UAdDXTlaDjofG5iZk2UD9fW4kZcyVhZKi14qo8PrQC95opyrOCeWY1pgetJkNh0LLFUXKBAGhiq3wnyI30r3HHvW6TacGnqNPeeqEo7E9KeogUnJEDKlMf0Zww72N5olPsBO0pf6bdYVtm91DOV/Ybl+fXczehp9ovEuN23002VkOG7KZ9g==
##Kc/BRM3KXhU=
##
##
##fd6a9f26a06ea3bc99616d4851b372ba

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
        Write-Host "JOB в статусе:" $p.State -ForegroundColor DarkYellow
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
cls    

Write-Host "|IIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIII|" -ForegroundColor DarkGreen
Write-Host "|ЧТО ЖЕЛАЕШЬ ПОВЕЛИТЕЛЬ?                              |
|1. Показать весь список задач на удаленной машине?   |
|2. Запустим конкретный джоб?                         |
|3. Отключить или включить джоб?                      |
|0. Выход из программы!                               |" -ForegroundColor DarkGreen
Write-Host "|IIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIII|" -ForegroundColor DarkGreen
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
while ($SelectJhin -ne '0')



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
