############################################
#                   Made by                #
#                  Fokin L@B               #
# Сбор и проверка пользовательских паролей #
############################################

# Проверка наличия модуля, если нет то импортируем
if (-not (Get-Module DSInternals)) {
      Import-Module ".\lib\DSInternals.psd1"
}
if (-not (Get-Module ActiveDirectory)) {
      Import-Module ActiveDirectory
}

# Инициализация Переменных
$datA=Get-Date -Format "dd.MM.yyyy(HH:mm)" | foreach {$_ -replace ":", "."}
$Date_With_Hours = Get-Date -Format "dd.MM.yyyy(HH:mm)"
$DC = "mskdc7.rusagrotrans.ru"
$Domain = "DC=rusagrotrans,DC=ru"
$Pth = "C:\Scripts\"
$Pfile = 'C:\Scripts\LOGS\AD_PSSW_'+"$data"+'.csv'
$DictFile = "C:\Scripts\weak_passwords.txt"
$OU='*DC=rusagrotrans,DC=ru'
$MailServer = "webmail.rusagrotrans.ru"
$FromAddress = "weak-passwords@rusagrotrans.ru"
$strBlankPasswordNThash = '31d6cfe0d16ae931b73c59d7e0c089c0'
$htBadPasswords = @{}
$htBadPasswords.Add($strBlankPasswordNThash,"")

#Значение количества юзеров предварительно обнулим
$intBadPasswordsFound = 0


#Тут все колдовство и шаманство по изъятию с домена хэшей
    Function Get-NTHashFromClearText
    {
        Param ([string]$ClearTextPassword)
        Return ConvertTo-NTHash $(ConvertTo-SecureString $ClearTextPassword -AsPlainText -Force)
    }

Foreach ($WordlistPath in $DictFile)
    {
        If (Test-Path $WordlistPath)
        {
            If ($bolWriteToLogFile) {LogWrite -Logfile $LogFileName -LogEntryString "Word list file found: $WordlistPath" -LogEntryType INFO -TimeStamp}
            Write-Verbose "Word list file found: $WordlistPath"
            $BadPasswordList = Get-Content -Path $WordlistPath
            $cnt_LineNumber = 1
            Foreach ($BadPassword in $BadPasswordList)
            {
                #Detect and ignore emty strings
               If ($BadPassword -eq '')
               {
                   #If ($bolWriteToLogFile) {LogWrite -Logfile $LogFileName -LogEntryString "| Empty input line ignored (line#: $cnt_LineNumber)." -LogEntryType INFO -TimeStamp}
                   Write-Verbose "| Empty input line ignored (line#: $cnt_LineNumber)."
                   $cnt_LineNumber++
                   Continue
               }
                $NTHash = $(Get-NTHashFromClearText $BadPassword)
                If ($htBadPasswords.ContainsKey($NTHash)) # NB! Case-insensitive on purpose
                {
                    $intBadPasswordsInListsDuplicates++
                    If ($bolWriteToLogFile -and $bolWriteVerboseInfoToLogfile) {LogWrite -Logfile $LogFileName -LogEntryString "| Duplicate password: '$BadPassword' = $NTHash" -LogEntryType INFO -TimeStamp}
                    Write-Verbose "| Дублирующий пароль: '$BadPassword' = $NTHash (line#: $cnt_LineNumber)"
                }
                Else # New password to put into hash table
                {
                    If ($bolWriteToLogFile -and $bolWriteVerboseInfoToLogfile) {LogWrite -Logfile $LogFileName -LogEntryString "| Adding to hashtable: '$BadPassword' = $NTHash" -LogEntryType INFO -TimeStamp}
                    Write-Verbose "| Добавление в хештаблицу : '$BadPassword' = $NTHash (line#: $cnt_LineNumber)"
                    $htBadPasswords.Add($NTHash,$BadPassword)
                }
                # Counting line numbers
              $cnt_LineNumber++
            } # Foreach BadPassword
        }
    }

#Смещение верхней строки
$intBadPasswordsInLists = $htBadPasswords.Count - 1
Add-content -path $Pfile "Name,Username,NTLMPassword"

$arrUsersAndHashes = Get-ADReplAccount -All -Server $DC -NamingContext $Domain | Where {$_.Enabled -eq $true -and $_.SamAccountType -eq 'User' -and $_.DistinguishedName -like $OU} `
| Select SamAccountName,@{Name="NTHashHex";Expression={ConvertTo-Hex $_.NTHash}}
$intUsersAndHashesFromAD = $arrUsersAndHashes.Count

Foreach ($hashuser in $arrUsersAndHashes)
    {
        $strUserSamAccountName = $hashuser.SamAccountName
        $Nametable = Get-ADUser -LDAPFilter "(sAMAccountName=$strUserSamAccountName)" | Select Name
        $Name = $Nametable.name
        $strUserNTHashHex = $hashuser.NTHashHex
If ($htBadPasswords.ContainsKey($strUserNTHashHex)) # NB! Case-insensitive on purpose
        {
            $intBadPasswordsFound++
            $strUserBadPasswordClearText = $htBadPasswords.Get_Item($strUserNTHashHex)
            Add-content $Pfile "$Name,$strUserSamAccountName,$strUserBadPasswordClearText"  -Encoding UTF8
        }
    }
$UsersAll = Import-Csv $Pfile -Encoding UTF8
# Исключение пользователей
$Users = Import-Csv $Pfile -Encoding UTF8 | where-object { $_.Username -NotMatch 'УверенныйАдмин|testuser|СУПЕРБОСС|_*' }
# Мылинг
function SendNotification {
$Msg = New-Object Net.Mail.MailMessage
$Smtp = New-Object Net.Mail.SmtpClient($MailServer)
$Msg.From = $FromAddress
$Msg.To.Add($ToAddress)
$Msg.Subject = "Внимание! $datA Обнаружен слабый пароль!"
$Msg.Body = $EmailBody
$Msg.IsBodyHTML = $true
$Msg.Priority = [System.Net.Mail.MailPriority]::High
$Smtp.Send($Msg)
}
$to_Owners='fokin_ok@rusagrotrans.ru'
$head = "Внимание! $Date_With_Hours Обнаружены слабые пароли на $intBadPasswordsFound из $intUsersAndHashesFromAD :"
$emailBodyTXT = $UsersAll | Sort NTLMPassword,Username | ConvertTo-Html -Head $head | Format-Table -Autosize | Out-String
#
# Мега Мылинг
#
# Заполнение данных для письма
$EmailFrom = "Аудит Паролей <weak_passwords@rusagrotrans.ru>"
$ToAddress = "fokin_ok@rusagrotrans.ru" 
$Subject = "Внимание! $Date_With_Hours Обнаружены слабые пароли!"
$Body = $emailBodyTXT
$Msg = New-Object Net.Mail.MailMessage
$Msg.From = $EmailFrom
$Msg.To.Add($ToAddress)
$Msg.Subject = $Subject
$Msg.Body = $Body
$Msg.IsBodyHTML = $true
$Msg.Priority = [System.Net.Mail.MailPriority]::High
 
$SMTPServer = "webmail.rusagrotrans.ru" 
$SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 25)
$SMTPClient.EnableSsl = $False
#$SMTPClient.Credentials = New-Object System.Net.NetworkCredential($CredUser, $CredPassword); 
$SMTPClient.Send($Msg)


#Send-MailMessage -smtpServer $MailServer -from $FromAddress -to $to_Owners -subject "Внимание! $datA Обнаружены слабые пароли!" -body $emailBodyTXT -BodyAsHTML -Encoding UTF8 -Port 25
<#
Foreach ($User in $Users){
$FIO = $User.name
$ToAddress = "$FIO <" + $User.Username + '@rusagrotrans.ru>'
#$How = $User.NTLMState
$Pass = $User.NTLMPassword
$Pwd = $Pass.Substring(0,$Pass.Length-4)
$login = $User.Username
$Name = Get-ADUser -LDAPFilter "(sAMAccountName=$login)" | Select Name
$Who = $Name -replace "@{Name=" -replace "}"
$EmailBody = @"
#Красивости
<html>
  <style>
h1{ine-height:70%;}
h5{line-height:1px;}
#nol {font-size:11px;line-height:17px;padding:10px;border:5px solid #456;}
</style>
<head>
</head>
<body>
<div id="nol">

<h2 style="color:green"><img style="width:50px;height:45px" class=Gimage src=data:image/png;base64,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></img>  Добрый день <strong>$Who,</strong></h2>
<b>
<h4>Согласно ВНД "Положение об организации системы управления информационной безопасностью в АО «yourcompany»" ПО – 4568-9870 от 12.04.2019 г.  Приложение 1 \ Глава 2. Парольная защита \ § 1. Правила формирования пароля</h4>
<h4>!В результе проведённого поверхностного аудита паролей сотрудников компании, Ваш пароль не прошел проверку на сложность и не может гарантировать безопасность данных.</h4>
<h3 style="color:grey">Ваш пароль: <strong style="color:red">$Pwd***** </strong></h3>
<p></p>
<h4>Необходимо сменить на более сложный пароль.</h4>
<h4>Подобные проверки будут проводиться регулярно для большей уверенности в сохранности данных.</h4>
<h4>Если необходима консультация касательно сложности пароля, Вы можете обратиться к системному администратору. <br>
<p><a href="https://password.kaspersky.com/ru/">Либо воспользоваться сервисом</a></p></h4>
<h4><strong>Сменить пароль сейчас можно нажав сочитание клавиш ctrl+alt+del / Изменить пароль.</strong></h4>
  <h4><strong>В случае игнорирования, Ваша учетная запись будет заблокирована.</strong></h4>
<h3><strong>!!! После прочтения это письмо следует удалить !!!</strong></h3>
<h3 style="color:red">Данное письмо сформировано автоматически и отвечать на него не нужно.</h3>
</div>

<p>С Уважением,</p>
  <p>Вумный админ yourcompany</p>
</body>
</html>
"@
#Раскоментить строку ниже для отображения процесса отправки
#Write-Host "Sending notification to $Who $Pass ($ToAddress)" -ForegroundColor Yellow
#Send-MailMessage -smtpServer $MailServer -from $FromAddress -to $ToAddress -subject "Внимание! $datA Обнаружен слабый пароль!" -body $emailBody -BodyAsHTML -Encoding UTF8
#Start-Sleep -Seconds 60
}
#>