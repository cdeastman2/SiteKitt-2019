Imports System.Collections.ObjectModel
Imports System.Management.Automation
Imports System.Management.Automation.Runspaces
Imports System.Text
Imports System.IO
Module PowerShellHelper
 
    Function RunScript(ByVal scriptText As String) As String
        ' Takes script text as input and runs it, then converts 
        ' the results to a string to return to the user 
       
	   ' create Powershell runspace 
        Dim MyRunSpace As Runspace = RunspaceFactory.CreateRunspace()

        ' open it 
        MyRunSpace.Open()

        ' create a pipeline and feed it the script text 
        Dim MyPipeline As Pipeline = MyRunSpace.CreatePipeline()

        MyPipeline.Commands.AddScript(scriptText)

        ' add an extra command to transform the script output objects into nicely formatted strings 
        ' remove this line to get the actual objects that the script returns. For example, the script 
        ' "Get-Process" returns a collection of System.Diagnostics.Process instances. 
        MyPipeline.Commands.Add("Out-String")

        ' execute the script 
        Dim results As Collection(Of PSObject) = MyPipeline.Invoke()

        ' close the runspace 
        MyRunSpace.Close()

        ' convert the script result into a single string 
        Dim MyStringBuilder As New StringBuilder()

        For Each obj As PSObject In results
            MyStringBuilder.AppendLine(obj.ToString())
        Next

        ' return the results of the script that has 
        ' now been converted to text 
        Return MyStringBuilder.ToString()

    End Function

    Function LoadScript(ByVal filename As String) As String
        ' helper method that takes your script path, loads up the script 
        ' into a variable, and passes the variable to the RunScript method 
        ' that will then execute the contents 
        Try

            ' Create an instance of StreamReader to read from our file. 
            ' The using statement also closes the StreamReader. 
            Dim sr As New StreamReader(filename)

            ' use a string builder to get all our lines from the file 
            Dim fileContents As New StringBuilder()

            ' string to hold the current line 
            Dim curLine As String = ""

            ' loop through our file and read each line into our 
            ' stringbuilder as we go along 
            Do
                ' read each line and MAKE SURE YOU ADD BACK THE 
                ' LINEFEED THAT IT THE ReadLine() METHOD STRIPS OFF 
                curLine = sr.ReadLine()
                fileContents.Append(curLine + vbCrLf)
            Loop Until curLine Is Nothing

            ' close our reader now that we are done 
            sr.Close()

            ' call RunScript and pass in our file contents 
            ' converted to a string 
            Return fileContents.ToString()

        Catch e As Exception
            ' Let the user know what went wrong. 
            Dim errorText As String = "The file could not be read:"
            errorText += e.Message + "\n"
            Return errorText
        End Try

    End Function


End Module
