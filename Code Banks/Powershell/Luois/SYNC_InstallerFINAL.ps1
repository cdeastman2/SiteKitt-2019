<#
.This PS1 Script Installs Smart SYNC Student 2011
##################################
Contact Louis.Razo@fresnounified.org for any questions.
##################################
#>

#########################################################
#This will run the powerhsell as Administrator
#########################################################
Function RequireAdmin {
	If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
		Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" $PSCommandArgs" -WorkingDirectory $pwd -Verb RunAs
		Exit
	}
}
RequireAdmin
#########################################################
$Script:showWindowAsync = Add-Type -MemberDefinition @"
[DllImport("user32.dll")]
public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
"@ -Name "Win32ShowWindowAsync" -Namespace Win32Functions -PassThru
#########################################################
#This will Hide the Powershell WIndow
#########################################################
Function Hide-Powershell()
{
$null = $showWindowAsync::ShowWindowAsync((Get-Process -Id $pid).MainWindowHandle, 2)
}
Hide-Powershell
#########################################################
#### This Turns Disables Sleep Mode in both Battery and AC
##########################################################
Function Set-NoPower {
powercfg -change -monitor-timeout-dc 0
powercfg -change -standby-timeout-dc 180
powercfg -change -monitor-timeout-ac 0
powercfg -change -standby-timeout-ac 180
powercfg -setacvalueindex 381b4222-f694-41f0-9685-ff5bb260df2e 4f971e89-eebd-4455-a8de-9e59040e7347 5ca83367-6e45-459f-a27b-476b1d01c936 0
powercfg -setdcvalueindex 381b4222-f694-41f0-9685-ff5bb260df2e 4f971e89-eebd-4455-a8de-9e59040e7347 5ca83367-6e45-459f-a27b-476b1d01c936 0
}
#########################################################
#### This sets the POWER settings back to normal
##########################################################
Function Set-Power {
powercfg -change -monitor-timeout-dc 10
powercfg -change -standby-timeout-dc 20
powercfg -change -monitor-timeout-ac 10
powercfg -change -standby-timeout-ac 20
powercfg -setacvalueindex 381b4222-f694-41f0-9685-ff5bb260df2e 4f971e89-eebd-4455-a8de-9e59040e7347 5ca83367-6e45-459f-a27b-476b1d01c936 1
powercfg -setdcvalueindex 381b4222-f694-41f0-9685-ff5bb260df2e 4f971e89-eebd-4455-a8de-9e59040e7347 5ca83367-6e45-459f-a27b-476b1d01c936 1
}
#########################################################
#### This is the Dialog Box
##########################################################
function Show-InputForm()
{
    #create input form
    $inputForm               = New-Object System.Windows.Forms.Form 
    $inputForm.Text          = $args[0]
    $inputForm.Size          = New-Object System.Drawing.Size(330, 100) 
    $inputForm.StartPosition = "CenterScreen"
    [System.Windows.Forms.Application]::EnableVisualStyles()

    #handle button click events
    $inputForm.KeyPreview = $true
    $inputForm.Add_KeyDown(
    {
        if ($_.KeyCode -eq "Enter")  
        {
            $inputForm.Close() 
        } 
    })
    $inputForm.Add_KeyDown(
    {
        if ($_.KeyCode -eq "Escape") 
        {
            $inputForm.Close() 
        } 
    
    })

    #create OK button
    $okButton          = New-Object System.Windows.Forms.Button
    $okButton.Size     = New-Object System.Drawing.Size(75, 23)
    $okButton.Text     = "OK" 
    $okButton.Add_Click(
    {
        $inputForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
    })
    $inputForm.Controls.Add($okButton)
    $inputForm.AcceptButton = $okButton

    #create Cancel button
    $cancelButton          = New-Object System.Windows.Forms.Button 
    $cancelButton.Size     = New-Object System.Drawing.Size(75,23)
    $cancelButton.Text     = "Cancel"
    $inputForm.Controls.Add($cancelButton)
    $inputForm.CancelButton = $cancelButton

    [System.Collections.Generic.List[System.Windows.Forms.TextBox]] $txtBoxes = New-Object System.Collections.Generic.List[System.Windows.Forms.TextBox]
    $y = -15;
    for($i=1;$i -lt $args.Count;$i++)
    {
        $y+=30
        $inputForm.Height += 30

        #create label
        $objLabel          = New-Object System.Windows.Forms.Label
        $objLabel.Location = New-Object System.Drawing.Size(10,  $y)
        $objLabel.Size     = New-Object System.Drawing.Size(280, 20)
        $objLabel.Text     = $args[$i] +":"
        $inputForm.Controls.Add($objLabel)
        $y+=20
        $inputForm.Height+=20
        
        #create TextBox
        $objTextBox          = New-Object System.Windows.Forms.TextBox 
        $objTextBox.Location = New-Object System.Drawing.Size(10,  $y)
        $objTextBox.Size     = New-Object System.Drawing.Size(290, 20) 
        $inputForm.Controls.Add($objTextBox)

        $txtBoxes.Add($objTextBox)

        $cancelButton.Location = New-Object System.Drawing.Size(165, (35+$y))
        $okButton.Location     = New-Object System.Drawing.Size(90, (35+$y))

        if ($args[$i].Contains("*"))
        {
            $objLabel.Text = ($objLabel.Text -replace '\*','')
            $objTextBox.UseSystemPasswordChar = $true 
        }
    }

    $inputForm.Topmost = $true 
    $inputForm.MinimizeBox = $false
    $inputForm.MaximizeBox = $false
    
    $inputForm.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $inputForm.SizeGripStyle = [System.Windows.Forms.SizeGripStyle]::Hide
    $inputForm.Add_Shown({$inputForm.Activate(); $txtBoxes[0].Focus()})
    if ($inputForm.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK)
    {
        exit
    }

    return ($txtBoxes | Select-Object {$_.Text} -ExpandProperty Text)
}
Add-Type -AssemblyName "system.windows.forms"
##########################################################
#Variables for the Input Form
##########################################################
$login    = Show-InputForm "MHS TECHS" "Enter Abbreviated Site Name" "Enter Room Number" "Enter Teacher Computer Name" "Enter SCCM Description"
$site   = $login[0]
$room     = $login[1]
$teacher = $login[2]
$description = $login[3]
##########################################################
#Creates Folder
##########################################################
Function CreateFOLDER {
if (!(Test-Path \\itdata\Shared\Sync\Site\$site\$Room\$Teacher))
    {
        New-Item -ItemType Directory -Path \\itdata\Shared\Sync\Site\$site\$Room\$Teacher
    }
#MD \\itdata\Shared\Sync\Site\$site\$Room\$Teacher
}

##########################################################
#Installs the Sync Software
##########################################################
Function Install-Sync {
Start-Process -FilePath msiexec.exe -ArgumentList '/i',"\\itdata\Shared\Sync\SmartSync2011Student\SMARTSyncStudent.msi /passive" -wait
#Start-Process -FilePath msiexec.exe -ArgumentList '/i',"\\itdata\Shared\Sync\SmartSync2011Student\SMARTSyncStudent.msi /passive" -wait -Verb "RunAs"
}
##########################################################
#Adds the Computer name into the Registry
##########################################################
###
###
##########################################################
#$Teacher = 'Louis'
##########################################################
###
#########################################################
#This Adds the Teacher IP Address
#########################################################
Function Set-Reg {
$RegLocations = @("HKLM:\Software\SMART Technologies\SMART Sync Student")

if([System.Environment]::Is64BitProcess){
    $RegLocations += "HKLM:\Software\WOW6432Node\SMART Technologies\SMART Sync Student"
}

foreach($Key in $RegLocations)
{
    if(-not(Test-Path $Key)){
        New-ItemProperty -Path $Key -Name "ConnectIP" -Value "$Teacher.staff.fusd.local" -ErrorAction SilentlyContinue -PropertyType String -Force
    }
    New-ItemProperty -Path $Key -Name "ConnectIP" -Value "$Teacher.staff.fusd.local" -ErrorAction SilentlyContinue -PropertyType String -Force
}
}
#########################################################
#This Hides the Sync from Students
#########################################################
Function Hide-Sync {
$RegLocations = @("HKLM:\Software\SMART Technologies\SMART Sync Student")

if([System.Environment]::Is64BitProcess){
    $RegLocations += "HKLM:\Software\WOW6432Node\SMART Technologies\SMART Sync Student"
}

foreach($Key in $RegLocations)
{
    if(-not(Test-Path $Key)){
        New-ItemProperty -Path $Key -Name "Visible" -Value 0 -ErrorAction SilentlyContinue -PropertyType DWord -Force
    }
    New-ItemProperty -Path $Key -Name "Visible" -Value 0 -ErrorAction SilentlyContinue -PropertyType DWord -Force
}
}
#########################################################
##Sets SYNC inline with Active Directory
#########################################################
Function AD-Sync {
$RegLocations = @("HKLM:\Software\SMART Technologies\SMART Sync Student")

if([System.Environment]::Is64BitProcess){
    $RegLocations += "HKLM:\Software\WOW6432Node\SMART Technologies\SMART Sync Student"
}

foreach($Key in $RegLocations)
{
    if(-not(Test-Path $Key)){
        New-ItemProperty -Path $Key -Name "ActiveDirStudentIdField" -Value cn -ErrorAction SilentlyContinue -PropertyType String -Force
    }
    New-ItemProperty -Path $Key -Name "ActiveDirStudentIdField" -Value cn -ErrorAction SilentlyContinue -PropertyType String -Force
}
}
#########################################################
#This Starts the SYNC Program
#########################################################
Function Start-Sync {
Start DAX64.exe
}
#########################################################
# This Kills the SYNC Process
#########################################################
Function Stop-Sync {
$SyncClient = Get-Process dax64 -ErrorAction SilentlyContinue
if ($SyncClient) {
  # try gracefully first
  $SyncClient.CloseMainWindow()
  # kill after five seconds
  Sleep 5
  if (!$SyncClient.HasExited) {
    $SyncClient | Stop-Process -Force
  }
}
Remove-Variable SyncClient
}
###
###
##########################################################
#
Function Add-Description {
$OSValues = Get-WmiObject -class Win32_OperatingSystem 
$OSValues.Description = "$description"
$OSValues.put()
}
#########################################################
Function Logs {
$ComputerSystem = Get-WmiObject -Class Win32_ComputerSystem | Select -Property Model , Description , PrimaryOwnerName , SystemType
$BootConfiguration = Get-WmiObject -Class Win32_BootConfiguration | Select -Property Name , ConfigurationPath 
$BIOS = Get-WmiObject -Class Win32_BIOS | Select -Property PSComputerName , Manufacturer , Version , SerialNumber
$OperatingSystem = Get-WmiObject -Class Win32_OperatingSystem | Select -Property Caption , CSDVersion , OSArchitecture , OSLanguage  
$report = New-Object psobject
$report | Add-Member -MemberType NoteProperty -name Manufacturer -Value $BIOS.Manufacturer
$report | Add-Member -MemberType NoteProperty -name Model -Value $ComputerSystem.Model
$report | Add-Member -MemberType NoteProperty -name ComputerName -Value $BIOS.PSComputerName
$report | Add-Member -MemberType NoteProperty -name SerialNumber -Value $BIOS.serialnumber
$all=@"
Site,Room,Teacher,Descripion,
$site,$room,$teacher,$description,
"@
$results = @()
$info = @{
Site = $site
Room = $room
Teacher = $teacher
Description = $Description
Bios = $bios.Manufacturer
Model = $ComputerSystem.Model
PCName = $BIOS.PSComputerName
SerialNumber = $BIOS.serialnumber
}
$results += New-Object psobject -Property $info
#md ".\Site\$site\$room\$description"
$csvfile = "\\itdata\Shared\Sync\Site\$site\$Room\$Teacher\Log.csv"
$results | Select-Object Site,Room,Teacher,Description,Bios,Model,PCName,SerialNumber | Export-CSV   -Path $csvfile -NoTypeInformation -Append
Get-Content $csvfile | Sort-Object -Unique | Set-Content $csvfile
}
###
###
###
###
###
###
###
$master = "\\itdata\Shared\Sync\REG_Master.ps1"
$file =  "\\itdata\Shared\Sync\Site\$site\$Room\$Teacher\Install.ps1"
###
##########################################################
Function Replace-n-Create1 {
(Get-Content $master | ForEach-Object { $_ -replace "Louis", "$Teacher" } ) | Set-Content $file
}
#
#
Function Replace-n-Create2 {
(Get-Content $file | ForEach-Object { $_ -replace "Razo", "$description" } ) | Set-Content $file
}
#
#
Function Replace-n-Create3 {
(Get-Content $file | ForEach-Object { $_ -replace "sitio", "$site" } ) | Set-Content $file
}
#
#
Function Replace-n-Create4 {
(Get-Content $file | ForEach-Object { $_ -replace "cuarto", "$room" } ) | Set-Content $file
}
#
#
##########################################################
#This will update the group policy
##########################################################
Function GPOUpdate {
write-output "Group Policy is updating"
(gpupdate) -join ""

}
##########################################################
#This Log off all User Accounts
##########################################################
#
#BLANK
#
##########################################################
#This will Delete all the User Profiles
##########################################################
Function DeleteUserProfile {
Write-output = "Deleting User Profiles"
$profiles = $null
$profiles = Get-WMIObject -class Win32_UserProfile | Where {((!$_.Special) -and ($_.LocalPath -ne "C:\Users\Administrator")  -and ($_.LocalPath -ne "C:\Users\UpdatusUser"))}

if ($profiles -ne $null) {
    $profiles | Remove-WmiObject
}
}
$source = "\\itdata\Shared\Sync\Site\LaunchLog.bat"
$destination = "\\itdata\Shared\Sync\Site\$site\$Room\$Teacher\LaunchLogRM$Room.bat"
Function CopyBAT {
copy-Item  -Recurse $source -Destination $destination
}
#########################################################
Set-NoPower
CreateFOLDER
#Install-Sync
#Stop-Sync
#Set-Reg
#AD-Sync
#Hide-Sync
#Start-Sync
#Add-Description
Replace-n-Create1
Replace-n-Create2
Replace-n-Create3
Replace-n-Create4
#Logs
CopyBAT
#LogOff
#DeleteUserProfile
#Set-Power
#GPOUpdate