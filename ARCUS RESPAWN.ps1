$Server = "fobo-ws-tst-012","fobo-ws-tst-013"

foreach($S in $Server)
{
$S
if(Test-Path "\\$S\c$\Arcus2")
{
    Write-Host "EFT СУЩЕСТВУЕТ"
}
else
{
    Write-Host "EFT НЕ СУЩЕСТВУЕТ"
}
}