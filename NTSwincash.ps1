
#Бета-скрипт
 
# Создано - Кирьянов Артем
    # 1. Остановка службы NTSwincash
    # 2. Kill процессов javaw.exe | Logistics.exe | Sales.exe |Sales_old.exe
    # 3. Удаление файлов в папке C:NTSwincash\userdata
    # 4. Запуск NTSwincash
 
#Функция записи лога скрипта
 
function ZipArchive {
    Add-Type -Assembly "System.IO.Compression.FileSystem"
    [System.IO.Compression.ZipFile]::CreateFromDirectory("C:\NTSwincash\log", ("C:\NTSwincash\userdata\Logs_" + ((Get-Date).ToString('dd-MM-yyyy HH-mm')) + ".zip"))
}
 
function CheckServices {
   $service =  Get-Service -Name "NTSwincash distributor"
   $comp = $env:computername
   #Если служба запущена,то
   if( $service.status -eq "running"){
       #Останавливаем ее
       Stop-Service $service -ErrorAction SilentlyContinue
   }
   #Stop-Process -Name 'Sales','Sales_old','Configurator', 'javaw','Logistics', 'SalesReporter' -Force
   Get-Process -name 'Sales','Sales_old','Configurator', 'java', 'javaw','Logistics', 'SalesReporter','office','SyncAgent','wmDisconnected'  | Where-Object {$_.Path -NE 'C:\jenkins\jdk1.8.0_152\bin\java.exe'} | Stop-Process -Force
   ZipArchive
   Remove-Item C:\NTSwincash\log\* -Recurse -Force -Confirm:$false
   Start-Service $service -ErrorAction SilentlyContinue
   Write-Output "Все ок! Папки почищены,служба снова работает!"
}  
 
Start-Transcript -path D:\not_delete\logging.txt -append 
CheckServices -ServiceName "NTSwincash distributor"
Stop-Transcript