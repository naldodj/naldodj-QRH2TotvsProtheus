$Files=dir .\7ZQRH2TOTVSProtheus.txt | % { [System.IO.File]::ReadAllLines($_.FullName) }
foreach ($filePath in $Files)
{
    $archParams += "`"" + $filePath + "`" "
}
$archParams = "a `".\QRH2TOTVSProtheus.zip`" " + $archParams

Start-Process "C:\Program Files\7-Zip\7z.exe" -Wait -ArgumentList $archParams