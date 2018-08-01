


$BackUpPath = "F:\Ricoh\backupfso.zip"
$Destination = "C:recovered"
Add-Type -assembly "system.io.compression.filesystem"

[io.compression.zipfile]::ExtractToDirectory($BackUpPath, $destination)

