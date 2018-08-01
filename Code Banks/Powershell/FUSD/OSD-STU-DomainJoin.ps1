#--------------- Set Variables ---------------
$Domain = "students.fusd.local"
$OUPath = 'OU=Warehouse,OU=Schools,DC=students,DC=fusd,DC=local'
$UserCurrent = $env:ComputerName + "\" + $env:USERNAME
$UserAdmin = $env:ComputerName + "\Administrator"
$DomainCheck = (gwmi win32_computersystem).partofdomain

#--------------- Login Variables ---------------$PasswordFile = "C:\Windows\Temp\OSD-Scripts\OSD-DomainJoin\STU-DomainJoin\Password.txt"
$KeyFile = "C:\Windows\Temp\OSD-Scripts\OSD-DomainJoin\STU-DomainJoin\AES.key"$Key = Get-Content $KeyFile$UserName = "staff\DomainJoin"$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserName, (Get-Content $PasswordFile | ConvertTo-SecureString -Key $key)Function Check-RunAsAdministrator(){  #Get current user context  $CurrentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())    #Check user is running the script is member of Administrator Group  if($CurrentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator))  {       Write-host "Script is running with Administrator privileges!"  }  else    {       #Create a new Elevated process to Start PowerShell       $ElevatedProcess = New-Object System.Diagnostics.ProcessStartInfo "PowerShell";        # Specify the current script path and name as a parameter       $ElevatedProcess.Arguments = "& '" + $script:MyInvocation.MyCommand.Path + "'"        #Set the Process to elevated       $ElevatedProcess.Verb = "runas"        #Start the new elevated process       [System.Diagnostics.Process]::Start($ElevatedProcess)        #Exit from the current, unelevated, process       Exit     }}  # If computer was successfully joined to the domain, these files will be deleted on first login. function deleteCheck     {        #1709 delete registry key stopped working            #Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -Name "RunDomainJoinCheck"        #Will not work due to file being used            #Remove-Item C:\Windows\Temp\OSD-Scripts\OSD-DomainJoin\* -Recurse -Force        Get-ChildItem c:\windows\temp\OSD-Scripts\OSD-DomainJoin\ -Include *.* -File -Recurse | foreach {$_.Delete()}        Remove-Item C:\Windows\Temp\OSD-Scripts\OSD-DomainJoinCheck\* -Recurse -Force        Remove-Item C:\Windows\Temp\OSD-Scripts\OSD-SSID\* -Recurse    }function joinMessage    {        Write-Host -fore Cyan "Computer will join the $Domain domain and automatically restart."        Write-Host -fore Green "Joining domain..."    }#--------------- Main ---------------# Checks to see if computer is on the domain. Fail safe. If($DomainCheck -eq "TRUE")    {        deleteCheck        Write-Host -fore Green "Computer is already on the domain"        Set-ExecutionPolicy Restricted -Force        Start-Sleep -s 10        Exit    }# Checks to see if logged in as administrator. Prevents restart during OSD which will cause OSD to fail.
If($UserCurrent -eq $UserAdmin)
    {
        # Will restart computer if add-computer did not error.
        If($DomainCheck -eq "TRUE")            {                deleteCheck                Write-Host -fore Green "Computer is already on the domain"                Set-ExecutionPolicy Restricted -Force                Start-Sleep -s 10                Exit            }
        Elseif ($error.count -eq 0)
            {
                Check-RunAsAdministrator
                Add-Computer -DomainName $Domain -OUPath $OUPath -Credential $Credential -ErrorAction SilentlyContinue
                If($error.count -eq 0)
                    {
                        deleteCheck
                        Write-Host -fore Cyan "Computer will join the $Domain domain and automatically restart."                        Write-Host -fore Green "Joining domain..."
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

# Join domain during OSD    Add-Computer -DomainName $Domain -OUPath $OUPath -Credential $Credential -ErrorAction SilentlyContinue
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