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