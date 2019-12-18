$a = [System.Windows.MessageBox]::Show('Would  you like to play a game?','Game input','YesNoCancel','Error')
if ($a -eq 'YES')
{
  [System.Windows.MessageBox]::Show('Ну что сыграем?','ПИПОООС КАК КРУТО','YesNoCancel','Warning')  
}

 switch  ($msgBoxInput) {

  'Yes' {

  ## Do something 

  }

  'No' {

  ## Do something

  }

  'Cancel' {

  ## Do something

  }

  }


  Add-Type -AssemblyName System.Windows.Forms
  $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
  Filter = 'Картинки (*.png)|*.png|Файлы (*.war)|*.war'
  Title = 'Выберите файл сервиса для деплоя'}
  $Null = $FileBrowser.ShowDialog()
  Write-Host $Null
  
  $DEST = Read-Host "SERVER?"
  $Dest +=$FileBrowser.SafeFileName
  Rename-Item $DEST -NewName "$Dest.backup" 
  Copy-Item -Path $FileBrowser.FileName -Destination $DEST 
  $Hash1 = Get-FileHash $FileBrowser.FileName
  $Hash2 = Get-FileHash $DEST
  Compare-Object -ReferenceObject $Hash1 -DifferenceObject $Hash2

  if ($Hash1.Hash -eq $Hash2.Hash)
  {
    
    [System.Windows.MessageBox]::Show("Файл успешно перенесен $DEST","Перенос файла успешен")
    
  }
  else
  {
    [System.Windows.MessageBox]::Show("Файл перенесен в $DEST с ошибками
    Выполнить восстановление файла?","Перенос файла провалился","YesNoCancel")
    Rename-Item "$Dest.backup" -NewName "$Dest"
  }