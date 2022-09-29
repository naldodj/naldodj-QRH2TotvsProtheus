#https://stackoverflow.com/questions/32843563/grant-read-permission-for-msysobjects
#https://stackoverflow.com/questions/19971082/no-read-permission-on-msysobjects

$ScrUsr = $(whoami)
Write-Host $ScrUsr

$cmd = "GRANT SELECT ON MSysObjects TO Admin;"
Write-Host $cmd

Function Invoke-ADOCommand($Db, $SystemDb)
{
  $connection = New-Object -ComObject ADODB.Connection
  $ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$db;Jet OLEDB:System database=$SystemDb;"
  Write-Host $ConnectionString
  $connection.Open($ConnectionString)
  $discard = $connection.Execute($cmd)
  $connection.Close()
}

$Db = "C:\MiniGUI\SAMPLES\BASIC\solotica\SOLOTICA.qrh"
$SystemDb = "C:\Users\marin\AppData\Roaming\Microsoft\Access\System.mdw"

$dbEngine=New-Object -ComObject DBEngine

Invoke-ADOCommand -db $Db -SystemDb $SystemDb