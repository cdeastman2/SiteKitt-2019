Module SyncUSB

    '========================11================================
    ' LANG 					: VBScript
    ' NAME 					: SCSetup.vbs
    ' AUTHOR				: Clifford Dinel Eastman ii	
    ' Mail					: Clifford.Eastmanii@fresnounified.org
    ' VERSION 				: Alpah
    ' DATE 					: 12/2/2015
    ' Description			: install and setup Smartware Sync Client 2011 
    '						: 
    ' COMMENTS 				: 
    ' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
    ' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
    ' IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
    ' ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE
    ' LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
    ' CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
    ' SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
    ' INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
    ' CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
    ' ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
    ' POSSIBILITY OF SUCH DAMAGE.
    '==========================================================
    ' Settings
    Dim objArguments, NumberOfArguments, strPath
    Dim theGo, strComputer, MyOS, strProcessKill
    Dim strTeacherComputer

    Const HKEY_LOCAL_MACHINE = &H80000002
  




    
    Function PKiller(strComputer, strProcessToKill)
        Dim objWMIService, objProcess, colProcess
        objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
        colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name = " & strProcessToKill)

        For Each objProcess In colProcess
            objProcess.Terminate()
        Next

        objWMIService = Nothing
        colProcess = Nothing
        PKiller = "Done"
    End Function

    Sub Set_Sync_Defalut_Seting(strKeyPath)

        Dim strValueName
        Dim strValue
        Dim strTeacherComputer = "May-Day.staff.fusd.local"

        My.Computer.Registry.CurrentUser.CreateSubKey("TestKey")

        strValueName = "ConnectIP"
        strValue = strTeacherComputer
        Call WrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "RedrawHooks"
        strValue = &H3E8
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "LanguageID"
        strValue = &H0
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "MirrorDriver"
        strValue = &H3E8
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "MulticastTTL"
        strValue = &H20
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "UnicastNoDelay"
        strValue = &H1
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "StudentIDMode"
        strValue = &H0
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "PasswordHash"
        strValue = ""
        Call WrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "NICListLength"
        strValue = &H0
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "UseNamingServer"
        strValue = &H0
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "NTGroupListLength"
        strValue = &H0
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "Student Path"
        strValue = "C:\\Program Files\\SMART Technologies\\Sync Student\\"
        Call WrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "ActiveDirStudentIdField"
        strValue = "cn"
        Call WrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "CtrlAltDelSettings"
        strValue = &H2
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "EnableFileTransfer"
        strValue = &H1
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "Visible"
        strValue = &H1
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "StudentID"
        strValue = "AnonymousID"
        Call WrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "Build Version"
        strValue = "10.0.574.0"
        Call WrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "ConnectTeacherID"
        strValue = ""
        Call WrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)


        strValueName = "EnableHelp"
        sstrValue = &H0
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "DisplayExit"
        strValue = &H0
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "StoreFilesToMyDocs"
        strValue = &H1
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "AutoStart"
        strValue = &H1
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "CustomSharedFolder"
        strValue = ""
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "SecurityUsed"
        strValue = &H0
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "ConnectionUsed"
        strValue = &H3
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "EnableNICDefaultOrder"
        strValue = &H1
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "EnableQuestions"
        strValue = &H1
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "NamingServerLoc"
        sstrValue = ""
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "EnableChat"
        strValue = &H1
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "NamingServerPassedTest"
        strValue = &H0
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "HWGPUCapture"
        strValue = &H3E8
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "GPUBoolEnableTimings"
        strValue = &H0
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "GPUBoolMergeRegions"
        strValue = &H1
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "GPUTileCount"
        strValue = &H1
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "GPUScaleFactor"
        strValue = &H3
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)

        strValueName = "GPUThresholdPixelCorrection"
        strValue = &H3E8
        Call DwordWrightToReg(strComputer, HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)


    End Sub

    Function Check4file()
        Dim go
        Dim objFSO
        objFSO = CreateObject("Scripting.FileSystemObject")

        If objFSO.FileExists("C:\Program Files (x86)\SMART Technologies\Sync Student\SyncClient.exe") Then

            'wshShell.popup "The file does exist",10 'Debug
            go = 1
        Else
            'wshShell.popup "The file does not exist.",10 'Debug
            go = 0
        End If

        objFSO = Nothing
        Check4file = go
    End Function


End Module
