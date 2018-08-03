$Script:showWindowAsync = Add-Type -MemberDefinition @"
[DllImport("user32.dll")]
public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
"@ -Name "Win32ShowWindowAsync" -Namespace Win32Functions -PassThru
Function Show-Powershell()
{
$null = $showWindowAsync::ShowWindowAsync((Get-Process -Id $pid).MainWindowHandle, 10)
}
Function Hide-Powershell()
{
$null = $showWindowAsync::ShowWindowAsync((Get-Process -Id $pid).MainWindowHandle, 2)
}

Hide-Powershell
md C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin

Import-Module ($ENV:systemdrive + "\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1")

#--------------------
# Begin Configuration
#--------------------
# Sometimes it's useful to filter collections by a common name prefix.  
# If you don't need to do any filtering, lave this an as empty string.
$collectionNamePrefix = "" 
# Sometimes it's useful to filter computers by a common name prefix.
# If you don't need to do any filtering, lave this an as empty string.
$computerNamePrefix = ""
# This is required.
$siteServer = "SC-CONFIG2.STAFF.FUSD.LOCAL"
# This is required.
$siteCode = "FS1"
$resourceExplorerPath = $ENV:systemdrive + "\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ResourceExplorer.exe"
$remoteControlToolPath = $ENV:systemdrive + "\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe"
#--------------------
# End Configuration
#--------------------

$sccmPsDrive = $siteCode + ":"
Set-Location $sccmPsDrive

Function Display-SuccessAlert {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $message,
        [string]
        $title="Alert"
    )

    [System.Windows.Forms.MessageBox]::Show($message, $title, [System.Windows.Forms.MessageBoxButtons]::Ok, [System.Windows.Forms.MessageBoxIcon]::Asterisk)
}

Function Display-ErrorAlert {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $message,
        [string]
        $title="Error"
    )

    [System.Windows.Forms.MessageBox]::Show($message, $title, [System.Windows.Forms.MessageBoxButtons]::Ok, [System.Windows.Forms.MessageBoxIcon]::Error)
}

Function Display-ConfirmationAlert {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $message,
        [string]
        $title="Question"
    )

    return [System.Windows.Forms.MessageBox]::Show($message, $title, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
}

Function Convert-WMIDateStringToDate {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)]
        [string]
        $dateString
    )

    return [System.Management.ManagementDateTimeConverter]::ToDateTime($dateString)
}

Function Invoke-ResourceExplorer {
    $statusBar.Text = "Looking up computer object in SCCM..."

    $computerName = $computerNameComboBox.SelectedItem
    $sccmComputer = Get-WmiObject -ComputerName $siteServer -Namespace "root\sms\site_$siteCode" -Query "Select * From SMS_R_System where Name='$computerName'"

    $namespace = "\\$siteServer\root\sms\site_$siteCode"
    $resourceExplorerArguments =  "-s -sms:ResourceID=$($sccmComputer.ResourceID) -sms:Connection=$namespace"

    Start-Process -FilePath $resourceExplorerPath -ArgumentList $resourceExplorerArguments

    $statusBar.Text = "Idle"
}

Function Get-LoggedOnUser {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)]
        [ValidateLength(1,15)]
        [ValidateNotNullOrEmpty()]
        [string]
        $computerName
    )
    $computer = Get-WmiObject -ComputerName $computerName -Class Win32_ComputerSystem
    $computer.UserName
}

Function Reboot-Computer {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $computerName,
        [ValidateNotNullOrEmpty()]
        [string]
        $rebootArgument = "r"
    )

    if($rebootArgument.ToLower().CompareTo("s") -eq 0) {
        $rebootMessage = "Your computer received a remote shutdown request from an administrator.  It will not come back up automatically."
    } else {
        $rebootArgument = "r"
        $rebootMessage = "Your computer received a remote reboot request from an administrator.  It should come back up automatically in a few minutes."
    }

    $process = [WMIClass]("\\$computerName\root\cimv2:Win32_Process")
    $result = $process.Create("shutdown.exe -$rebootArgument -t 30 -c `"$rebootMessage`" ", $null, $null)

    return $result.returnValue
}

Function Trigger-ClientAction {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [string]
        $computerName,
        [parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [string]
        $actionName
    )

    $scheduleIdList = @{
        ApplicationEvalCycle = "{00000000-0000-0000-0000-000000000121}"
        HardwareInventory = "{00000000-0000-0000-0000-000000000001}";
        SoftwareInventory = "{00000000-0000-0000-0000-000000000002}";
        DataDiscovery = "{00000000-0000-0000-0000-000000000003}";
        MachinePolicyRetrievalEvalCycle = "{00000000-0000-0000-0000-000000000021}";
        MachinePolicyEvaluation = "{00000000-0000-0000-0000-000000000022}";
        RefreshDefaultManagementPoint = "{00000000-0000-0000-0000-000000000023}";
        RefreshLocation = "{00000000-0000-0000-0000-000000000024}";
        SoftwareMeteringUsageReporting = "{00000000-0000-0000-0000-000000000031}";
        SourcelistUpdateCycle = "{00000000-0000-0000-0000-000000000032}";
        RefreshProxyManagementPoint = "{00000000-0000-0000-0000-000000000037}";
        CleanupPolicy = "{00000000-0000-0000-0000-000000000040}";
        ValidateAssignments = "{00000000-0000-0000-0000-000000000042}";
        CertificateMaintenance = "{00000000-0000-0000-0000-000000000051}";
        BranchDPScheduledMaintenance = "{00000000-0000-0000-0000-000000000061}";
        BranchDPProvisioningStatusReporting = "{ 00000000-0000-0000-0000-000000000062}";
        SoftwareUpdateDeployment = "{00000000-0000-0000-0000-000000000108}";
        StateMessageUpload = "{00000000-0000-0000-0000-000000000111}";
        StateMessageCacheCleanup = "{00000000-0000-0000-0000-000000000112}";
        SoftwareUpdateScan = "{00000000-0000-0000-0000-000000000113}";
        SoftwareUpdateDeploymentReEval = "{00000000-0000-0000-0000-000000000114}";
        OOBSDiscovery = "{00000000-0000-0000-0000-000000000120}";
    }

    $scheduleId = $scheduleIdList.$actionName
    $sccmClient = [WMIclass]"\\$computerName\root\ccm:SMS_Client" 

    try {
        $sccmClient.TriggerSchedule($scheduleId)
        Display-SuccessAlert "Successfully triggered action $actionName on $computerName."
    } catch {
        Display-ErrorAlert "Failed to trigger action $actionName on $computerName."
    }   
}

Function Setup-CollectionComboBox {
    $collectionComboBox.Items.Clear()

    $collectionNames = New-Object System.Windows.Forms.AutoCompleteStringCollection 
    $collections = Get-CMDeviceCollection -Name "$collectionNamePrefix*" | Sort-Object -property Name
 
    foreach($collection in $collections) {
        $collectionNames.Add([string]$($collection.Name))
    }

    $collectionComboBox.AutoCompleteMode = "SuggestAppend"
    $collectionComboBox.AutoCompleteSource = "CustomSource"
    $collectionComboBox.AutoCompleteCustomSource = $collectionNames  

    foreach($collection in $collections) {
        $collectionComboBox.Items.Add($collection.Name)
    }
    if($collectionComboBox.Items.Count -gt 0) {
        $collectionComboBox.SelectedIndex = 0
    }
}

Function Setup-ComputerNamesComboBox {
    $computerNameComboBox.Items.Clear()

    $computerNames = New-Object System.Windows.Forms.AutoCompleteStringCollection 
    $computers = Get-WmiObject -ComputerName $siteServer -Namespace "root\sms\site_$siteCode" -Query "Select * From SMS_R_System Where Name Like '$computerNamePrefix%'" | Sort-Object Name
    
    foreach($computer in $computers) {
        $computerNames.Add([string]$($computer.Name))
    }

    $computerNameComboBox.AutoCompleteMode = "SuggestAppend"
    $computerNameComboBox.AutoCompleteSource = "CustomSource"
    $computerNameComboBox.AutoCompleteCustomSource = $computerNames

    foreach($name in $computerNames) {
        $computerNameComboBox.Items.Add($name)
    }
    if($computerNameComboBox.Items.Count -gt 0) {
        $computerNameComboBox.SelectedIndex = 0
    }
}

Function Validate-ComputerName {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [string]
        $computerName
    )

    $invalidCharacters = @('\', '/', ':', '*', '?', '"', '<', '>', '|', '.', '+', '_')

    if (($computerName.Length -gt 15) -or ($computerName.Length -eq 5)) {
        return $false
    } else {
        foreach ($character in $invalidCharacters) {
            if($computerName.Contains($character)) {
                return $false
            }
        }
    }
    return $true
}

#Generated Form Function
function GenerateForm {
########################################################################
# Code Generated By: SAPIEN Technologies PrimalForms (Community Edition) v1.0.10.0
########################################################################

#region Import the Assemblies
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
#endregion

#region Generated Form Objects
$mainWindow = New-Object System.Windows.Forms.Form
$groupBox1 = New-Object System.Windows.Forms.GroupBox
$reloadCollectionsButton = New-Object System.Windows.Forms.Button
$collectionComboBox = New-Object System.Windows.Forms.ComboBox
$collectionLabel = New-Object System.Windows.Forms.Label
$addToCollectionButton = New-Object System.Windows.Forms.Button
$removeFromCollectionButton = New-Object System.Windows.Forms.Button
$userGroupBox = New-Object System.Windows.Forms.GroupBox
$loggedOnUserButton = New-Object System.Windows.Forms.Button
$reloadComputerNamesButton = New-Object System.Windows.Forms.Button
$computerNameComboBox = New-Object System.Windows.Forms.ComboBox
$statusBar = New-Object System.Windows.Forms.StatusBar
$networkGroupBox = New-Object System.Windows.Forms.GroupBox
$pingButton = New-Object System.Windows.Forms.Button
$resourcesGroupBox = New-Object System.Windows.Forms.GroupBox
$sccmLogsButton = New-Object System.Windows.Forms.Button
$resourceExplorerButton = New-Object System.Windows.Forms.Button
$remoteControlGroupBox = New-Object System.Windows.Forms.GroupBox
$remoteDesktopButton = New-Object System.Windows.Forms.Button
$remoteControlButton = New-Object System.Windows.Forms.Button
$shutdownToolsGroupBox = New-Object System.Windows.Forms.GroupBox
$shutdownButton = New-Object System.Windows.Forms.Button
$rebootButton = New-Object System.Windows.Forms.Button
$sccmClientActionsGroupBox = New-Object System.Windows.Forms.GroupBox
$applicationEvaluationCycleButton = New-Object System.Windows.Forms.Button
$updatesDeploymentEvalButton = New-Object System.Windows.Forms.Button
$updatesScanButton = New-Object System.Windows.Forms.Button
$softwareInventoryButton = New-Object System.Windows.Forms.Button
$hardwareInventoryButton = New-Object System.Windows.Forms.Button
$machinePolicyButton = New-Object System.Windows.Forms.Button
$discoveryCycleButton = New-Object System.Windows.Forms.Button
$computerNameLabel = New-Object System.Windows.Forms.Label
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
#endregion Generated Form Objects

#----------------------------------------------
#Generated Event Script Blocks
#----------------------------------------------
$handler_reloadNamesButton_Click = 
{
    $statusBar.Text = "Loading computer names..."
    Setup-ComputerNamesComboBox
    $statusBar.Text = "Idle"
}

$handler_rebootButton_Click = 
{
    $computerName = $computerNameComboBox.SelectedItem
    $result = Display-ConfirmationAlert "Are you sure you want to reboot the computer?" "Reboot Computer"
    if($result -eq [Windows.Forms.DialogResult]::Yes) {
        if((Reboot-Computer $computerName) -ne 0) {
            Display-ErrorAlert "There was a problem rebooting $computerName.  Check the machine name or network connection." "Reboot failed"
        } else {
            Display-SuccessAlert "Successfully rebooted $computerName." "Reboot successful"
        }
    }
}

$handler_shutdownButton_Click = 
{
    $computerName = $computerNameComboBox.SelectedItem
    $result = Display-ConfirmationAlert "Are you sure you want to shutdown the computer?" "Shutdown Computer"
    if($result -eq [Windows.Forms.DialogResult]::Yes) {
        if((Reboot-Computer $computerName "s") -ne 0) {
            Display-ErrorAlert "There was a problem shutting down $computerName.  Check the machine name or network connection." "Shutdown failed"
        } else {
            Display-SuccessAlert "Successfully shutdown $computerName." "Shutdown successful"
        }
    }
}  

$handler_remoteControlButton_Click = 
{
    $computerName = $computerNameComboBox.SelectedItem
    Start-Process -FilePath $remoteControlToolPath -ArgumentList "$computerName"
}

$handler_remoteDesktopButton_Click = 
{
    $computerName = $computerNameComboBox.SelectedItem
    $remoteDesktopPath = "mstsc.exe"
    Start-Process -FilePath $remoteDesktopPath -ArgumentList "/v:$computerName"
}

$handler_applicationEvaluationCycle_Click = {
    Trigger-ClientAction $computerNameComboBox.SelectedItem "ApplicationEvalCycle"
}

$handler_discoveryCycleButton_Click = 
{
    Trigger-ClientAction $computerNameComboBox.SelectedItem "DataDiscovery"
}

$handler_machinePolicyButton_Click = 
{
    Trigger-ClientAction $computerNameComboBox.SelectedItem "MachinePolicyRetrievalEvalCycle"
}

$handler_softwareInventoryButton_Click = 
{
    Trigger-ClientAction $computerNameComboBox.SelectedItem "SoftwareInventory"
}

$handler_hardwareInventoryButton_Click = 
{
    Trigger-ClientAction $computerNameComboBox.SelectedItem "HardwareInventory"
}

$handler_updatesScanButton_Click = 
{
    Trigger-ClientAction $computerNameComboBox.SelectedItem "SoftwareUpdateScan"
}

$handler_updatesDeploymentEvalButton_Click = 
{
    Trigger-ClientAction $computerNameComboBox.SelectedItem "SoftwareUpdateDeploymentReEval"
}

$handler_resourceExplorerButton_Click = 
{
    Invoke-ResourceExplorer
}

$handler_sccmLogsButton_Click = 
{
    Set-Location $ENV:systemdrive

    $computerName = $computerNameComboBox.SelectedItem
    $pathFor32BitMachines = "\\$computerName\c`$\Windows\CCM\Logs"
    $pathFor64BitMachines = "\\$computerName\c`$\Windows\CCM\Logs"
    $statusBar.Text = "Attempting to browse logs on $computerName"

    if(Test-Path $pathFor32BitMachines) {
        Invoke-Item $pathFor32BitMachines
    } elseif (Test-Path $pathFor64BitMachines) {
        Invoke-Item $pathFor64BitMachines
    } else {
        Display-ErrorAlert "Unable to browse logs on $computerName."
    }

    $statusBar.Text = "Idle"

    Set-Location $sccmPsDrive
}

$handler_loggedOnUserButton_Click = 
{
    $computerName = $computerNameComboBox.SelectedItem
    $statusBar.Text = "Checking logged on user on $computerName"

    try {
        if($computerName) {
            Test-Connection -ComputerName $computerName -ErrorAction Stop
            $loggedOnUser = Get-LoggedOnUser -computerName $computerName
            if($loggedOnUser) {
                Display-SuccessAlert "User currently logged on to computer `"$computerName`" is `"$loggedOnUser`"."
            } else {
                Display-SuccessAlert "No user is currently logged on to computer `"$computerName`"."
            }
        } else {
            Display-ErrorAlert "Invalid computer name `"$computerName`"."
        }
    } catch {
        Display-ErrorAlert "Failed to contact computer `"$computerName`"."
    }

    $statusBar.Text = "Idle"
}

$handler_reloadCollectionsButton_Click = 
{
    $statusBar.Text = "Loading collections..."
    Setup-CollectionComboBox
    $statusBar.Text = "Idle"
}

$handler_addToCollectionButton_Click = 
{
    $computerName = $computerNameComboBox.SelectedItem
    $collectionName = $collectionComboBox.SelectedItem
    $statusBar.Text = "Attempting to add computer $computerName to collection $collectionName..."

    $computer = Get-CMDevice -Name $computerName
    $collection = Get-CMDeviceCollection -Name $collectionName

    if((Validate-ComputerName $computerName) -and ($computer -ne $null)) {   
        Add-CMDeviceCollectionDirectMembershipRule -CollectionId $collection.CollectionID -ResourceId $computer.ResourceID                
        Display-SuccessAlert "Successfully added computer $computerName to $collectionName"
    } else {
        Display-ErrorAlert "Invalid computer name `"$computerName`""
    }

    $statusBar.Text = "Idle"
}

$handler_removeFromCollectionButton_Click = 
{
    $computerName = $computerNameComboBox.SelectedItem
    $collectionName = $collectionComboBox.SelectedItem
    $statusBar.Text = "Attempting to remove computer $computerName from collection $collectionName..."

    $computer = Get-CMDevice -Name $computerName
    $collection = Get-CMDeviceCollection -Name $collectionName

    if((Validate-ComputerName $computerName) -and ($computer -ne $null)) {   
        Remove-CMDeviceCollectionDirectMembershipRule -CollectionId $collection.CollectionID -ResourceId $computer.ResourceID -Force        
        Display-SuccessAlert "Successfully removed computer $computerName from $collectionName"
    } else {
        Display-ErrorAlert "Invalid computer name `"$computerName`""
    }

    $statusBar.Text = "Idle"
}


$handler_pingButton_Click = 
{
    $computerName = $computerNameComboBox.SelectedItem
    $hostIpAddress = $([System.Net.Dns]::GetHostAddresses("$computerName")).IPAddressToString
    $statusBar.Text = "Attempting to ping $computerName"

    try {
        $output = Test-Connection -ComputerName $computerName -ErrorAction Stop
        Display-SuccessAlert "Successfully pinged computer $computerName at $hostIpAddress."
    } catch {
        Display-ErrorAlert "Failed to ping computer $computerName at $hostIpAddress."
    }

    $statusBar.Text = "Idle"
}

$OnLoadForm_StateCorrection=
{
    # Correct the initial state of the form to prevent the .Net maximized form issue
    $mainWindow.WindowState = $InitialFormWindowState
    $mainWindow.Text = "$($mainWindow.Text)  - $siteServer - $siteCode"

    $statusBar.Text = "Loading collections..."
    Setup-CollectionComboBox
    $statusBar.Text = "Loading computer names..."
    Setup-ComputerNamesComboBox
    $statusBar.Text = "Idle"

    $mainWindow.Activate()
}

#----------------------------------------------
#region Generated Form Code
$mainWindow.AutoSizeMode = 0
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 486
$System_Drawing_Size.Width = 564
$mainWindow.ClientSize = $System_Drawing_Size
$mainWindow.DataBindings.DefaultDataSourceUpdateMode = 0
$mainWindow.FormBorderStyle = 1
$mainWindow.MaximizeBox = $False
$mainWindow.Name = "mainWindow"
$mainWindow.StartPosition = 1
$mainWindow.Text = "SCCM Quick Tools"


$groupBox1.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 374
$groupBox1.Location = $System_Drawing_Point
$groupBox1.Name = "groupBox1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 79
$System_Drawing_Size.Width = 538
$groupBox1.Size = $System_Drawing_Size
$groupBox1.TabIndex = 9
$groupBox1.TabStop = $False
$groupBox1.Text = "Collection Quick Add/Remove"

$mainWindow.Controls.Add($groupBox1)

$reloadCollectionsButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 407
$System_Drawing_Point.Y = 19
$reloadCollectionsButton.Location = $System_Drawing_Point
$reloadCollectionsButton.Name = "reloadCollectionsButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 125
$reloadCollectionsButton.Size = $System_Drawing_Size
$reloadCollectionsButton.TabIndex = 6
$reloadCollectionsButton.Text = "Reload Collections"
$reloadCollectionsButton.UseVisualStyleBackColor = $True
$reloadCollectionsButton.add_Click($handler_reloadCollectionsButton_Click)

$groupBox1.Controls.Add($reloadCollectionsButton)

$collectionComboBox.DataBindings.DefaultDataSourceUpdateMode = 0
$collectionComboBox.FormattingEnabled = $True
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 73
$System_Drawing_Point.Y = 19
$collectionComboBox.Location = $System_Drawing_Point
$collectionComboBox.Name = "collectionComboBox"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 21
$System_Drawing_Size.Width = 328
$collectionComboBox.Size = $System_Drawing_Size
$collectionComboBox.TabIndex = 2

$groupBox1.Controls.Add($collectionComboBox)

$collectionLabel.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 19
$collectionLabel.Location = $System_Drawing_Point
$collectionLabel.Name = "collectionLabel"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 61
$collectionLabel.Size = $System_Drawing_Size
$collectionLabel.TabIndex = 0
$collectionLabel.Text = "Collection:"
$collectionLabel.TextAlign = 64
$collectionLabel.add_Click($handler_collectionSoftwareDeploymentLabel_Click)

$groupBox1.Controls.Add($collectionLabel)


$addToCollectionButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 261
$System_Drawing_Point.Y = 46
$addToCollectionButton.Location = $System_Drawing_Point
$addToCollectionButton.Name = "addToCollectionButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 140
$addToCollectionButton.Size = $System_Drawing_Size
$addToCollectionButton.TabIndex = 4
$addToCollectionButton.Text = "Add to Collection"
$addToCollectionButton.UseVisualStyleBackColor = $True
$addToCollectionButton.add_Click($handler_addToCollectionButton_Click)

$groupBox1.Controls.Add($addToCollectionButton)


$removeFromCollectionButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 116
$System_Drawing_Point.Y = 46
$removeFromCollectionButton.Location = $System_Drawing_Point
$removeFromCollectionButton.Name = "removeFromCollectionButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 140
$removeFromCollectionButton.Size = $System_Drawing_Size
$removeFromCollectionButton.TabIndex = 4
$removeFromCollectionButton.Text = "Remove From Collection"
$removeFromCollectionButton.UseVisualStyleBackColor = $True
$removeFromCollectionButton.add_Click($handler_removeFromCollectionButton_Click)

$groupBox1.Controls.Add($removeFromCollectionButton)

$userGroupBox.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 283
$userGroupBox.Location = $System_Drawing_Point
$userGroupBox.Name = "userGroupBox"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 50
$System_Drawing_Size.Width = 260
$userGroupBox.Size = $System_Drawing_Size
$userGroupBox.TabIndex = 2
$userGroupBox.TabStop = $False
$userGroupBox.Text = "User"

$mainWindow.Controls.Add($userGroupBox)

$loggedOnUserButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 16
$loggedOnUserButton.Location = $System_Drawing_Point
$loggedOnUserButton.Name = "loggedOnUserButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 247
$loggedOnUserButton.Size = $System_Drawing_Size
$loggedOnUserButton.TabIndex = 0
$loggedOnUserButton.Text = "View Logged on User"
$loggedOnUserButton.UseVisualStyleBackColor = $True
$loggedOnUserButton.add_Click($handler_loggedOnUserButton_Click)

$userGroupBox.Controls.Add($loggedOnUserButton)



$reloadComputerNamesButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 417
$System_Drawing_Point.Y = 13
$reloadComputerNamesButton.Location = $System_Drawing_Point
$reloadComputerNamesButton.Name = "reloadComputerNamesButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 125
$reloadComputerNamesButton.Size = $System_Drawing_Size
$reloadComputerNamesButton.TabIndex = 8
$reloadComputerNamesButton.Text = "Reload Names"
$reloadComputerNamesButton.UseVisualStyleBackColor = $True
$reloadComputerNamesButton.add_Click($handler_reloadNamesButton_Click)

$mainWindow.Controls.Add($reloadComputerNamesButton)

$computerNameComboBox.DataBindings.DefaultDataSourceUpdateMode = 0
$computerNameComboBox.FormattingEnabled = $True
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 116
$System_Drawing_Point.Y = 14
$computerNameComboBox.Location = $System_Drawing_Point
$computerNameComboBox.MaxLength = 15
$computerNameComboBox.Name = "computerNameComboBox"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 21
$System_Drawing_Size.Width = 291
$computerNameComboBox.Size = $System_Drawing_Size
$computerNameComboBox.Sorted = $True
$computerNameComboBox.TabIndex = 0

$mainWindow.Controls.Add($computerNameComboBox)

$statusBar.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 0
$System_Drawing_Point.Y = 461
$statusBar.Location = $System_Drawing_Point
$statusBar.Name = "statusBar"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 25
$System_Drawing_Size.Width = 564
$statusBar.Size = $System_Drawing_Size
$statusBar.TabIndex = 1
$statusBar.Text = "Idle"

$mainWindow.Controls.Add($statusBar)


$networkGroupBox.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 290
$System_Drawing_Point.Y = 318
$networkGroupBox.Location = $System_Drawing_Point
$networkGroupBox.Name = "networkGroupBox"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 50
$System_Drawing_Size.Width = 260
$networkGroupBox.Size = $System_Drawing_Size
$networkGroupBox.TabIndex = 3
$networkGroupBox.TabStop = $False
$networkGroupBox.Text = "Network"

$mainWindow.Controls.Add($networkGroupBox)

$pingButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 20
$pingButton.Location = $System_Drawing_Point
$pingButton.Name = "pingButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 247
$pingButton.Size = $System_Drawing_Size
$pingButton.TabIndex = 0
$pingButton.Text = "Ping"
$pingButton.UseVisualStyleBackColor = $True
$pingButton.add_Click($handler_pingButton_Click)

$networkGroupBox.Controls.Add($pingButton)



$resourcesGroupBox.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 290
$System_Drawing_Point.Y = 230
$resourcesGroupBox.Location = $System_Drawing_Point
$resourcesGroupBox.Name = "resourcesGroupBox"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 82
$System_Drawing_Size.Width = 260
$resourcesGroupBox.Size = $System_Drawing_Size
$resourcesGroupBox.TabIndex = 7
$resourcesGroupBox.TabStop = $False
$resourcesGroupBox.Text = "Resources"

$mainWindow.Controls.Add($resourcesGroupBox)

$sccmLogsButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 50
$sccmLogsButton.Location = $System_Drawing_Point
$sccmLogsButton.Name = "sccmLogsButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 247
$sccmLogsButton.Size = $System_Drawing_Size
$sccmLogsButton.TabIndex = 1
$sccmLogsButton.Text = "Browse SCCM Logs"
$sccmLogsButton.UseVisualStyleBackColor = $True
$sccmLogsButton.add_Click($handler_sccmLogsButton_Click)

$resourcesGroupBox.Controls.Add($sccmLogsButton)


$resourceExplorerButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 20
$resourceExplorerButton.Location = $System_Drawing_Point
$resourceExplorerButton.Name = "resourceExplorerButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 247
$resourceExplorerButton.Size = $System_Drawing_Size
$resourceExplorerButton.TabIndex = 0
$resourceExplorerButton.Text = "Launch Resource Explorer"
$resourceExplorerButton.UseVisualStyleBackColor = $True
$resourceExplorerButton.add_Click($handler_resourceExplorerButton_Click)

$resourcesGroupBox.Controls.Add($resourceExplorerButton)



$remoteControlGroupBox.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 290
$System_Drawing_Point.Y = 54
$remoteControlGroupBox.Location = $System_Drawing_Point
$remoteControlGroupBox.Name = "remoteControlGroupBox"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 82
$System_Drawing_Size.Width = 260
$remoteControlGroupBox.Size = $System_Drawing_Size
$remoteControlGroupBox.TabIndex = 5
$remoteControlGroupBox.TabStop = $False
$remoteControlGroupBox.Text = "Remote Control"

$mainWindow.Controls.Add($remoteControlGroupBox)

$remoteDesktopButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 49
$remoteDesktopButton.Location = $System_Drawing_Point
$remoteDesktopButton.Name = "remoteDesktopButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 247
$remoteDesktopButton.Size = $System_Drawing_Size
$remoteDesktopButton.TabIndex = 1
$remoteDesktopButton.Text = "Remote Desktop"
$remoteDesktopButton.UseVisualStyleBackColor = $True
$remoteDesktopButton.add_Click($handler_remoteDesktopButton_Click)

$remoteControlGroupBox.Controls.Add($remoteDesktopButton)


$remoteControlButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 19
$remoteControlButton.Location = $System_Drawing_Point
$remoteControlButton.Name = "remoteControlButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 247
$remoteControlButton.Size = $System_Drawing_Size
$remoteControlButton.TabIndex = 0
$remoteControlButton.Text = "Remote Control"
$remoteControlButton.UseVisualStyleBackColor = $True
$remoteControlButton.add_Click($handler_remoteControlButton_Click)

$remoteControlGroupBox.Controls.Add($remoteControlButton)



$shutdownToolsGroupBox.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 290
$System_Drawing_Point.Y = 142
$shutdownToolsGroupBox.Location = $System_Drawing_Point
$shutdownToolsGroupBox.Name = "shutdownToolsGroupBox"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 82
$System_Drawing_Size.Width = 260
$shutdownToolsGroupBox.Size = $System_Drawing_Size
$shutdownToolsGroupBox.TabIndex = 6
$shutdownToolsGroupBox.TabStop = $False
$shutdownToolsGroupBox.Text = "Shutdown Tools"

$mainWindow.Controls.Add($shutdownToolsGroupBox)

$shutdownButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 49
$shutdownButton.Location = $System_Drawing_Point
$shutdownButton.Name = "shutdownButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 247
$shutdownButton.Size = $System_Drawing_Size
$shutdownButton.TabIndex = 1
$shutdownButton.Text = "Shutdown"
$shutdownButton.UseVisualStyleBackColor = $True
$shutdownButton.add_Click($handler_shutdownButton_Click)

$shutdownToolsGroupBox.Controls.Add($shutdownButton)


$rebootButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 20
$rebootButton.Location = $System_Drawing_Point
$rebootButton.Name = "rebootButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 247
$rebootButton.Size = $System_Drawing_Size
$rebootButton.TabIndex = 0
$rebootButton.Text = "Reboot"
$rebootButton.UseVisualStyleBackColor = $True
$rebootButton.add_Click($handler_rebootButton_Click)

$shutdownToolsGroupBox.Controls.Add($rebootButton)



$sccmClientActionsGroupBox.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 54
$sccmClientActionsGroupBox.Location = $System_Drawing_Point
$sccmClientActionsGroupBox.Name = "sccmClientActionsGroupBox"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 223
$System_Drawing_Size.Width = 260
$sccmClientActionsGroupBox.Size = $System_Drawing_Size
$sccmClientActionsGroupBox.TabIndex = 1
$sccmClientActionsGroupBox.TabStop = $False
$sccmClientActionsGroupBox.Text = "Trigger Client Actions"

$mainWindow.Controls.Add($sccmClientActionsGroupBox)

$applicationEvaluationCycleButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 8
$System_Drawing_Point.Y = 19
$applicationEvaluationCycleButton.Location = $System_Drawing_Point
$applicationEvaluationCycleButton.Name = "applicationEvaluationCycleButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 246
$applicationEvaluationCycleButton.Size = $System_Drawing_Size
$applicationEvaluationCycleButton.TabIndex = 6
$applicationEvaluationCycleButton.Text = "Application Evaluation Cycle"
$applicationEvaluationCycleButton.UseVisualStyleBackColor = $True
$applicationEvaluationCycleButton.add_Click($handler_applicationEvaluationCycle_Click)

$sccmClientActionsGroupBox.Controls.Add($applicationEvaluationCycleButton)


$updatesDeploymentEvalButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 192
$updatesDeploymentEvalButton.Location = $System_Drawing_Point
$updatesDeploymentEvalButton.Name = "updatesDeploymentEvalButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 247
$updatesDeploymentEvalButton.Size = $System_Drawing_Size
$updatesDeploymentEvalButton.TabIndex = 5
$updatesDeploymentEvalButton.Text = "Software Updates Deployment Eval Cycle"
$updatesDeploymentEvalButton.UseVisualStyleBackColor = $True
$updatesDeploymentEvalButton.add_Click($handler_updatesDeploymentEvalButton_Click)

$sccmClientActionsGroupBox.Controls.Add($updatesDeploymentEvalButton)


$updatesScanButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 163
$updatesScanButton.Location = $System_Drawing_Point
$updatesScanButton.Name = "updatesScanButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 248
$updatesScanButton.Size = $System_Drawing_Size
$updatesScanButton.TabIndex = 4
$updatesScanButton.Text = "Software Updates Scan Cycle"
$updatesScanButton.UseVisualStyleBackColor = $True
$updatesScanButton.add_Click($handler_updatesScanButton_Click)

$sccmClientActionsGroupBox.Controls.Add($updatesScanButton)


$softwareInventoryButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 134
$softwareInventoryButton.Location = $System_Drawing_Point
$softwareInventoryButton.Name = "softwareInventoryButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 248
$softwareInventoryButton.Size = $System_Drawing_Size
$softwareInventoryButton.TabIndex = 3
$softwareInventoryButton.Text = "Software Inventory Cycle"
$softwareInventoryButton.UseVisualStyleBackColor = $True
$softwareInventoryButton.add_Click($handler_softwareInventoryButton_Click)

$sccmClientActionsGroupBox.Controls.Add($softwareInventoryButton)


$hardwareInventoryButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 105
$hardwareInventoryButton.Location = $System_Drawing_Point
$hardwareInventoryButton.Name = "hardwareInventoryButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 247
$hardwareInventoryButton.Size = $System_Drawing_Size
$hardwareInventoryButton.TabIndex = 2
$hardwareInventoryButton.Text = "Hardware Inventory Cycle"
$hardwareInventoryButton.UseVisualStyleBackColor = $True
$hardwareInventoryButton.add_Click($handler_hardwareInventoryButton_Click)

$sccmClientActionsGroupBox.Controls.Add($hardwareInventoryButton)


$machinePolicyButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 76
$machinePolicyButton.Location = $System_Drawing_Point
$machinePolicyButton.Name = "machinePolicyButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 248
$machinePolicyButton.Size = $System_Drawing_Size
$machinePolicyButton.TabIndex = 1
$machinePolicyButton.Text = "Machine Policy Evaluation Cycle"
$machinePolicyButton.UseVisualStyleBackColor = $True
$machinePolicyButton.add_Click($handler_machinePolicyButton_Click)

$sccmClientActionsGroupBox.Controls.Add($machinePolicyButton)


$discoveryCycleButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 47
$discoveryCycleButton.Location = $System_Drawing_Point
$discoveryCycleButton.Name = "discoveryCycleButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 248
$discoveryCycleButton.Size = $System_Drawing_Size
$discoveryCycleButton.TabIndex = 0
$discoveryCycleButton.Text = "Discovery Data Collection Cycle"
$discoveryCycleButton.UseVisualStyleBackColor = $True
$discoveryCycleButton.add_Click($handler_discoveryCycleButton_Click)

$sccmClientActionsGroupBox.Controls.Add($discoveryCycleButton)


$computerNameLabel.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 13
$System_Drawing_Point.Y = 13
$computerNameLabel.Location = $System_Drawing_Point
$computerNameLabel.Name = "computerNameLabel"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 21
$System_Drawing_Size.Width = 97
$computerNameLabel.Size = $System_Drawing_Size
$computerNameLabel.TabIndex = 0
$computerNameLabel.Text = "Computer Name:"
$computerNameLabel.TextAlign = 64

$mainWindow.Controls.Add($computerNameLabel)

#endregion Generated Form Code

#Save the initial state of the form
$InitialFormWindowState = $mainWindow.WindowState
#Init the OnLoad event to correct the initial state of the form
$mainWindow.add_Load($OnLoadForm_StateCorrection)
#Show the Form
$mainWindow.ShowDialog()| Out-Null

} #End Function

#Call the Function
GenerateForm