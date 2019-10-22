##[Ps1 To Exe]
##
##Kd3HDZOFADWE8uO1
##Nc3NCtDXTlaDjofG5iZk2UD9fW4kZcyVhZKo04+w8OvoqBnxTJkcQFdyhCbzSWazXP8dGMYEuJ8YVhJK
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
##L8/UAdDXTlaDjofG5iZk2UD9fW4kZcyVhZKi14qo8PrQlDfYSpYRCX1ZpR3dKGfzXOoXNQ==
##Kc/BRM3KXxU=
##
##
##fd6a9f26a06ea3bc99616d4851b372ba
#Create by Fokin Oleg 16.10.2019
#################################################################################################
#Программа установки клиента sqldeveloper на удаленный пк версия x86 for testers
#Выводится запрос на ввод имени ПК пользователя, которому необходима установка sqldeveloper
#Выполняется подключение к удаленной машине пользователя
#Создается переменная среды ORACLE_HOME
#Редактируется переменная среды PATH
#Выполняется копирование клиента на локальную машину пользователя в корень диска C
#################################################################################################


Write-Host "MEGA PROJECT FOR DEPLOY SQLx86!"
$server = Read-Host "Введите имя машины"
C:\Windows\system32\psexec.exe \\$server cmd /c winrm quickconfig -quiet
#$sessionUP = New-PSSession $server
Invoke-Command -ComputerName $server -ScriptBlock {
    [System.Environment]::SetEnvironmentVariable("ORACLE_HOME","P:\Oracle\Ora11_g","Machine")
    [System.Environment]::SetEnvironmentVariable("PATH", $Env:Path + ";P:\Oracle\Ora11_g", "Machine")
    xcopy "\\dubovenko\D\sqldeveloper" "C:\" /E /D
}
