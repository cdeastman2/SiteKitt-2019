#
#MHS Region
#
Function RequireAdmin {
	If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
		Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" $PSCommandArgs" -WorkingDirectory $pwd -Verb RunAs
		Exit
	}
}
RequireAdmin
Function ADD-REG {
$RegLocations = @("HKLM:\Software\MHS")

if([System.Environment]::Is64BitProcess){
    $RegLocations += "HKLM:\Software\WOW6432Node\MHS"
}

foreach($Key in $RegLocations)
{
    if(-not(Test-Path $Key)){
	    #New-Item -Path 'HKLM:\Software' -Name "MHS" -itemType "Directory" -Force
		New-ItemProperty -Path $Key -Name "NukeWasRanOn" -Value (Get-Date -Format 'M/d/yyyy hh:mm:ss tt') -ErrorAction SilentlyContinue -PropertyType String -Force | Out-Null
        #New-ItemProperty -Path $Key -Name "Nuke" -Value 1 -ErrorAction SilentlyContinue -PropertyType DWord -Force
    }
	#New-Item 'HKLM:\Software\WOW6432Node\' -Name "MHS" -itemType "Directory" -Force
	New-ItemProperty -Path $Key -Name "NukeWasRanOn" -Value (Get-Date -Format 'M/d/yyyy hh:mm:ss tt') -ErrorAction SilentlyContinue -PropertyType String -Force | Out-Null
    #New-ItemProperty -Path $Key -Name "Nuke" -Value 1 -ErrorAction SilentlyContinue -PropertyType DWord -Force
}
}
#requires -version 3
#Remove-Profile -DaysOld 90 -Remove
function Remove-Profile {
    param(
        [string[]]$ComputerName = $env:ComputerName,
        [pscredential]$Credential = $null,
        [string[]]$Name,
        [ValidateRange(0,365)][int]$DaysOld = 0,
        [string[]]$Exclude,
        [switch]$IgnoreLastUseTime,
        [switch]$Remove
    )

    $ComputerName | ForEach-Object {
    
        if(Test-Connection -ComputerName $_ -BufferSize 16 -Count 2 -Quiet) {
        
            $params = @{
                ComputerName = $_ 
                Namespace    = 'root\cimv2'
                Class        = 'Win32_UserProfile'
            }

            if($Credential -and (@($env:ComputerName,'localhost','127.0.0.1','::1','.') -notcontains $_)) { 
                $params.Add('Credential', $Credential) 
            }

            if($null -ne $Name) {
                if($Name.Count -gt 1) {
                    $params.Add('Filter', ($Name | % { "LocalPath = '{0}'" -f $_ }) -join ' OR ')
                } else {
                    $params.Add('Filter', "LocalPath LIKE '%{0}'" -f ($Name -replace '\*', '%'))
                }
            }

            Get-WmiObject @params | ForEach-Object {

                $WouldBeRemoved = $false
                if(($_.SID -notin @('S-1-5-18', 'S-1-5-19', 'S-1-5-20')) -and 
                    ((Split-Path -Path $_.LocalPath -Leaf) -notin $Exclude) -and (-not $_.Loaded) -and ($IgnoreLastUseTime -or (
                        ($_.LastUseTime) -and (([WMI]'').ConvertToDateTime($_.LastUseTime)) -lt (Get-Date).AddDays(-1*$DaysOld)))) {
                    $WouldBeRemoved = $true
                }

                $prf = [pscustomobject]@{
                    PSComputerName = $_.PSComputerName
                    Account = (New-Object System.Security.Principal.SecurityIdentifier($_.Sid)).Translate([System.Security.Principal.NTAccount]).Value
                    LocalPath = $_.LocalPath
                    LastUseTime = if($_.LastUseTime) { ([WMI]'').ConvertToDateTime($_.LastUseTime) } else { $null }
                    Loaded = $_.Loaded
                }

                if(-not $Remove) {
                    $prf | Select-Object -Property *, @{N='WouldBeRemoved'; E={$WouldBeRemoved}}
                }

                if($Remove -and $WouldBeRemoved) { 
                    try {
                        $_.Delete()
                        $Removed = $true
                    } catch {
                        $Removed = $false
                        Write-Error -Exception $_
                    }
                    finally {
                        $prf | Select-Object -Property *, @{N='Removed'; E={$Removed}}
                    }
                }
            }

        } else {
            Write-Warning -Message "Computer $_ is unavailable"
        }
    }
}



#This is where the Functions get called
ADD-REG
#if you want to adjust the days of how old to delete the profile, change the number after -DaysOld
Remove-Profile -DaysOld 30 -Remove