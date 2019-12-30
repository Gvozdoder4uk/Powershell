$Server = "fobo-vrx-a543"
$VRXPACKAGES = @{
##########################
# VRX 1
##########################
    '166'='1 WS-M0';
    '279'='1 WS-M0';
    '105'='1 WS-M1';
    '660'='1 WS-M1';
    '024'='1 WS-M2';
    '050'='1 WS-M2';
    '175'='1 WS-M3';
    '180'='1 WS-M4';
    '061'='1 WS-M4';
##########################
# VRX 2
##########################
    '134'='2 WS-M0';
    '465'='2 WS-M0';
    'A01'='2 WS-M1';
    '217'='2 WS-M1';
    '266'='2 WS-M1';
    '064'='2 WS-M2';
    '299'='2 WS-M3';
    '482'='2 WS-M3';
    '208'='2 WS-M4';
    '469'='2 WS-M5';
    '018'='2 WS-M5';
##########################
# VRX 3
##########################
    '111'='3 WS-M0';
    '123'='3 WS-M0';
    '190'='3 WS-M1';
    '401'='3 WS-M1';
    '191'='3 WS-M1';
    '444'='3 WS-M2';
    '306'='3 WS-M5';
##########################
# VRX 4
##########################
    '025'='4 WS-M0';
    '099'='4 WS-M0';
    '118'='4 WS-M1';
    '119'='4 WS-M1';
    '102'='4 WS-M2';
    '112'='4 WS-M2';
    '230'='4 WS-M3';
##########################
# VRX 5
##########################
    '122'='5 WS-M00';
    '146'='5 WS-M01';
    '494'='5 WS-M02';
    '106'='5 WS-M03';
    '284'='5 WS-M04';
    '127'='5 WS-M05';
    '269'='5 WS-M06';
    '400'='5 WS-M07';
    '224'='5 WS-M08';
    '139'='5 WS-M09';
    '158'='5 WS-M09';
    '110'='5 WS-M10';
    '399'='5 WS-M10';
    '107'='5 WS-M11';
    '278'='5 WS-M11';
    '152'='5 WS-M12';
    '258'='5 WS-M13';
    '434'='5 WS-M14';
    '056'='5 WS-M15';
    '067'='5 WS-M15';
##########################
# VRX 6
##########################
    '014'='6 WS-M0';
    '015'='6 WS-M0';
    '196'='6 WS-M1';
    '461'='6 WS-M1';
    '141'='6 WS-M2';
    '188'='6 WS-M2';
    '130'='6 WS-M3';
    '132'='6 WS-M3';
    '128'='6 WS-M4';
    '131'='6 WS-M4';
    '543'='6 WS-M5';
    '754'='6 WS-M5';

}



$VRQPACKAGES = @{
##########################
# VRQ 1
##########################
    '111'='1 WS';
    '123'='1 WS';
    '105'='1 WS-M0';
    '166'='1 WS-M0';
    '279'='1 WS-M0';
    '190'='1 WS-M0';
    'A02'='1 WS-M1';
    '061'='1 WS-M1';
    '660'='1 WS-M2';
##########################
# VRQ 2
##########################
    '064'='2 WS';
    '299'='2 WS';
    '482'='2 WS';
    '266'='2 WS-M0';
    '306'='2 WS-M1';
    '208'='2 WS-M2';
    'A01'='2 WS-M3';
##########################
# VRQ 3
##########################
    '142'='3 WS';
    '102'='3 WS-M0';
    '112'='3 WS-M1';
##########################
# VRQ 4
##########################
    '118'='4 WS-M0';
    '119'='4 WS-M0';
    '120'='4 WS-M1';
##########################
# VRQ 5
##########################
    '122'='5 WS-M00';
    '146'='5 WS-M01';
    '494'='5 WS-M02';
    '106'='5 WS-M03';
    '284'='5 WS-M04';
    '127'='5 WS-M05';
    '269'='5 WS-M06';
    '400'='5 WS-M07';
    '224'='5 WS-M08';
    '139'='5 WS-M09';
    '158'='5 WS-M09';
    '110'='5 WS-M10';
    '399'='5 WS-M10';
    '107'='5 WS-M11';
    '278'='5 WS-M11';
    '152'='5 WS-M12';
    '258'='5 WS-M13';
    '434'='5 WS-M14';
    '056'='5 WS-M15';
    '067'='5 WS-M15';
##########################
# VRQ 6
##########################
    '014'='6 WS-M0';
    '015'='6 WS-M0';
    '196'='6 WS-M1';
    '461'='6 WS-M1';
    '141'='6 WS-M2';
    '188'='6 WS-M2';
    '130'='6 WS-M3';
    '132'='6 WS-M3';
    '128'='6 WS-M4';
    '131'='6 WS-M4';
    '543'='6 WS-M5';
    '754'='6 WS-M5';

}


$ZONESX = $VRXPACKAGES
$ZONESQ = $VRQPACKAGES

If($Server -like "*vrx*")
{
    
    foreach($T in $ZONESX.Keys)
    {
        #Write-Host $T
        if($Server -like "*"+$T)
        {
            $PacKet = $ZONESX.$T
            $PacKet = $PacKet.remove(0,2)
            Write-Host "ВЫБРАН ПАКЕТ: "$PacKet
            $STR = $ZONESX.$T
            $T = "fobo-vrx-ajb" +  $STR[0]
            Write-Host "ВЫБРАН СЕРВЕР: "$T
            $Source = "\\$T\C$\EtalonR3\$PacKet"
            Invoke-Item $Source
            #Copy-Item -Path $Source\* -Destination C:\NTSwincash -Recurse -PassThru

        }
    }

}
elseif($Server -like "*vrq*")
{
    foreach($T in $ZONESQ.Keys)
    {
        #Write-Host $T
        if($Server -like "*"+$T)
        {
            $PacKet = $ZONESQ.$T
            $PacKet = $PacKet.remove(0,2)
            Write-Host "ВЫБРАН ПАКЕТ: "$PacKet
            $STR = $ZONESQ.$T
            $T = "fobo-vrq-ajb" +  $STR[0]
            Write-Host "ВЫБРАН СЕРВЕР: "$T
            $Source = "\\$T\C$\EtalonR3\$PacKet"
            Invoke-Item $Source
            #Copy-Item -Path $Source\* -Destination C:\NTSwincash -Recurse -PassThru

        }
    }


}


New-Service -Name "FOBO_TEST" -BinaryPathName "C:\NTSwincash\jbin\DistributorService.exe" -DisplayName "MY FOBO" -Description "PROVERKA"