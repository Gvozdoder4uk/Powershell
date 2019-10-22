##[Ps1 To Exe]
##
##Kd3HDZOFADWE8uO1
##Nc3NCtDXTlaDjofG5iZk2UD9fW4kZcyVhZKo04+w8OvoqBn+W5sEQV10hWTKDUm4FNodQbg3ocMfXBMtYsEE4bvRF6qaTrsal612aOru
##Kd3HFJGZHWLWoLaVvnQnhQ==
##LM/RF4eFHHGZ7/K1
##K8rLFtDXTiW5
##OsHQCZGeTiiZ4dI=
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
##L8/UAdDXTlaDjofG5iZk2UD9fW4kZcyVhZKi14qo8PrQiC7AWZ9Ub31P2CzkASs=
##Kc/BRM3KXhU=
##
##
##fd6a9f26a06ea3bc99616d4851b372ba
<#Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings]
"ProxyEnable"=dword:00000000		;Отключает прокси-сервер
"ProxyServer"=""			;Стирает существующий прокси-сервер
"GlobalUserOffline"=dword:00000000	;Отключает автономный режим
"SecureProtocols"=dword:00000028	;Разрешает использование протоколов SSL 2.0 и SSL 3.0
"WarnOnZoneCrossing"=dword:00000001	;Включает опцию "Предупреждать о переключении режима безопасности"
"AutoConfigURL"=""			;Выключает галочку "Использовать сценарий автоматической настройки"

[HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main]
"Start Page"="http://www.chayka-net.ru/" 	;Установка домашней страницы

;------------------------------------------------------------------------------------------
;Снимает галочку "Автоматическое определение параметров"
[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Connections]
"DefaultConnectionSettings"=hex:3c,00,00,00,1e,00,00,00,01,00,00,00,00,00,00,\
  00,00,00,00,00,00,00,00,00,01,00,00,00,00,00,00,00,60,79,ba,70,e8,ce,c4,01,\
  01,00,00,00,c0,a8,00,bd,00,00,00,00,00,00,00,00
"SavedLegacySettings"=hex:3c,00,00,00,42,07,00,00,01,00,00,00,00,00,00,00,00,\
  00,00,00,00,00,00,00,01,00,00,00,00,00,00,00,60,79,ba,70,e8,ce,c4,01,01,00,\
  00,00,c0,a8,00,bd,00,00,00,00,00,00,00,00
;------------------------------------------------------------------------------------------

;------------------------------------------------------------------------------------------
;Восстановление стандартных префиксов URL
[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\URL\DefaultPrefix]
@="http://"

[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\URL\Prefixes]
"ftp"="ftp://"
"gopher"="gopher://"
"home"="http://"
"mosaic"="http://"
"www"="http://"
;-----------------------------------------------------------------------

#>

get-process iexplore | Stop-Process -Force
$path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings'
Set-ItemProperty -Path $path -Name 'Autodetect' -Value 0
Set-ItemProperty -Path $path -Name 'AutoConfigURL' -Value "http://pac.mvideo.ru/proxy.pac"
Invoke-Item "C:\Program Files\Internet Explorer\iexplore.exe"