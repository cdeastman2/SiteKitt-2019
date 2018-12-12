Module Module1

    Function ReadIni(myFilePath, mySection, myKey)
        Const ForReading = 1
        REM Const ForWriting = 2
        REM Const ForAppending = 8

        Dim intEqualPos
        Dim objFSO, objIniFile
        Dim strFilePath, strKey, strLeftString, strLine, strSection

        objFSO = CreateObject("Scripting.FileSystemObject")

        ReadIni = ""
        strFilePath = Trim(myFilePath)
        strSection = Trim(mySection)
        strKey = Trim(myKey)
        'wshshell.popup "My Key = " & myKey   & ",  My Str = " & strKey & ",   " '<<< Hard Stop
        If objFSO.FileExists(strFilePath) Then
            objIniFile = objFSO.OpenTextFile(strFilePath, ForReading, False)
            Do While objIniFile.AtEndOfStream = False
                strLine = Trim(objIniFile.ReadLine)

                'wshshell.popup strLine

                ' Check if section is found in the current line
                If LCase(strLine) = "[" & LCase(strSection) & "]" Then
                    strLine = Trim(objIniFile.ReadLine)

                    ' Parse lines until the next section is reached
                    Do While Left(strLine, 1) <> "["
                        ' Find position of equal sign in the line
                        intEqualPos = InStr(1, strLine, "=", 1)
                        If intEqualPos > 0 Then
                            strLeftString = Trim(Left(strLine, intEqualPos - 1))
                            ' Check if item is found in the current line
                            If LCase(strLeftString) = LCase(strKey) Then
                                ReadIni = Trim(Mid(strLine, intEqualPos + 1))
                                ' In case the item exists but value is blank
                                If ReadIni = "" Then
                                    ReadIni = "*"
                                End If
                                ' Abort loop when item is found
                                Exit Do
                            End If
                        End If

                        ' Abort if the end of the INI file is reached
                        If objIniFile.AtEndOfStream Then Exit Do

                        ' Continue with next line
                        strLine = Trim(objIniFile.ReadLine)
                    Loop
                    Exit Do
                End If
            Loop
            objIniFile.Close()
        Else
            REM  WScript.Echo strFilePath & " doesn't exists. Exiting..."
            REM Wscript.Quit 1
        End If
    End Function
    '=======================================================
End Module
