﻿#--------------- Set Variables ---------------
$Domain = "staff.fusd.local"
$OUPath = 'OU=Warehouse,OU=Staff Computers,DC=staff,DC=fusd,DC=local'
$UserCurrent = $env:ComputerName + "\" + $env:USERNAME
$UserAdmin = $env:ComputerName + "\Administrator"
$DomainCheck = (gwmi win32_computersystem).partofdomain

#--------------- Login Variables ---------------
$KeyFile = "C:\Windows\Temp\OSD-Scripts\OSD-DomainJoin\STA-DomainJoin\AES.key"
If($UserCurrent -eq $UserAdmin)
    {
        #Will restart computer if add-computer did not error.
        If($DomainCheck -eq "TRUE")
        Elseif ($error.count -eq 0)
            {
                Check-RunAsAdministrator
                Add-Computer -DomainName $Domain -OUPath $OUPath -Credential $Credential -ErrorAction SilentlyContinue
                If($error.count -eq 0)
                    {
                        deleteCheck
                        Write-Host -fore Cyan "Computer will join the $Domain domain and automatically restart."
                        Set-ExecutionPolicy Restricted -Force
                        Start-Sleep -s 10
                        Restart-Computer
                    }
        If($error.count -eq 1)
            {
                # Add to domain with no OU path. This will add the computer object back to the OU it's in.
                Write-Host -fore Red "A computer object already exist and will be joined to the same OU."
                Add-Computer -DomainName $Domain -Credential $Credential -InformationAction SilentlyContinue
                If($error.count -eq 1)
                    {
                        Write-Host -fore green "Computer succesfully joined the domain"
                        deleteCheck
                        Set-ExecutionPolicy Restricted -Force
                        Start-Sleep -s 10
                        Restart-Computer
                    }
                If($error.count -eq 2)
                    {
                        Write-Host -fore red "Computer failed to join the domain, check network connections and restart computer to try again."
                        pause
                    }
            }
            }

    }

# Join domain during OSD
        If($error.count -eq 0)
            {
                deleteCheck
            }

# If first add-computer failed, this will be the 2nd attempt and restart.
    If($error.count -eq 1)
            {
                # Add to domain with no OU path. This will add the computer object back to the OU it's in.
                Write-Host -fore Red "A computer object already exist and will be joined to the same OU."
                Add-Computer -DomainName $Domain -Credential $Credential -InformationAction SilentlyContinue
                If($error.count -eq 1)
                    {
                        Write-Host -fore green "Computer succesfully joined the domain"
                        deleteCheck
                        Start-Sleep -s 5
                    }
                If($error.count -eq 2)
                    {
                        Write-Host -fore red "Computer failed to join a domain, check network connections and try again."
                        Set-ExecutionPolicy Unrestricted -Force
                        Start-Sleep -s 10
                    }
            }