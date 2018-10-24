$strQuery = "SELECT TOP 1 
 [High Jumper Data].Name ,[High Jumper Data].[Country] ,[High Jumper Data].[Personal Best] 
 FROM [High Jumper Data] 
 ORDER BY [Personal Best] DESC ;"

$strConn = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=$PWDHighJumperDatabase.mdb"
$oConn = New-Object System.Data.OleDb.OleDbConnection $strConn
$oCmd  = New-Object System.Data.OleDb.OleDbCommand($strQuery, $oConn)
$oConn.Open()
$oReader = $oCmd.ExecuteReader()
[void]$oReader.Read()
    $oJumper = New-Object PSObject
    $oJumper | Add-Member NoteProperty Name     $oReader[0]
    $oJumper | Add-Member NoteProperty Country  $oReader[1]
    $oJumper | Add-Member NoteProperty PBest    $oReader[2]
    $oJumper
$oReader.Close()
$oConn.Close()