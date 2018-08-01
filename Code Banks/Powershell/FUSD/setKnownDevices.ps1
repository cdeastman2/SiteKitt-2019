<# PS Script will check model and add the variable skipDeviceDriver with the value of No. Will not skip driver package and will ignore message prompt otherwise the opposite will happen #>

# Loads hardware model to a variable. 
$Model = Get-WmiObject -Class win32_computersystem | Select-Object -ExpandProperty model

# Our list of driver packages in SCCM. 
$Models = "ASM201",
# ASUS
    "E200HA",
    "E205SA",
    "ET2321I",
    "GR8 II",
    "TP500LA",
    "T100HAN",
    "TP500LAG",
    "TP500LAB",
    "TP501UA",
    "TP501UAM",
    "Z240IC-H170-GTX950",
# HP
    "870-280",
    "HP EliteBook 850 G3",
    "HP EliteBook 850 G4",
    "HP ProBook x360 11 G1 EE",
    "HP ProDesk 400 G2 MINI",
    "HP ProDesk 400 G3 DM",
    "HP Stream 11 Pro Notebook PC",
# Microsoft
    "Surface 3",
    "Surface Book",
    "Surface Laptop",
    "Surface Pro",
    "Surface Pro 3",
    "Surface Pro 4",
    "Surface Studio",
# Other
    "Virtual Machine"

# Functions that set skipDeviceDriver variable during OSD. 
function doNotDisplayMessage
{
    $TSEnv = New-Object -COMObject Microsoft.SMS.TSEnvironment 
    $TSEnv.Value("skipDeviceDriver") = "No"
}

function displayMessage
{
    $TSEnv = New-Object -COMObject Microsoft.SMS.TSEnvironment 
    $TSEnv.Value("skipDeviceDriver") = "Yes"
}


# Will displayMessage function will run first as a default value. 
displayMessage

# Will check each model in the array against the hardware model. If true the doNotDisplayMessage function will update the value from Yes, to No, allowing to check for a driver package and surpress unknown hardware message.  
Foreach ($CheckModel in $Models)
{
    if ($CheckModel -eq $Model)
        {
            doNotDisplayMessage
        }
}