


$SERVERS = "fobo-vrx-ajb1","fobo-vrx-ajb2","fobo-vrx-ajb3","fobo-vrx-ajb4","fobo-vrx-ajb5","fobo-vrx-ajb6"
foreach($S in $SERVERS){

$T = Get-ChildItem -Path "\\$S\c$\EtalonR3\*\jboss\stage" -Include "*.readclaim.*","*.facade*","*.transfer*" -Recurse 
$T.Directory.FullName + " "|  Out-File C:\1\Center_VRX.txt -Append
}


$SERVERS = "fobo-vrq-ajb1","fobo-vrq-ajb2","fobo-vrq-ajb3","fobo-vrq-ajb4","fobo-vrq-ajb5","fobo-vrq-ajb6","fobo-vrq-ajb7"
foreach($S in $SERVERS){

$T = Get-ChildItem -Path "\\$S\c$\EtalonR3\*\jboss\stage" -Include "*.readclaim.*","*.facade*","*.transfer*" -Recurse 
$T.Directory.FullName + " "|  Out-File C:\1\Center_VRQ.txt -Append
}