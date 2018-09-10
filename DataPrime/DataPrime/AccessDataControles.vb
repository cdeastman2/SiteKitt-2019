Imports System.Data.OleDb
Public Class DBControle
    Public ModualName As String = "DBControle"

    REM Testing access data 
    Private Data_Connection As New OleDbConnection("Provider = Microsoft.ACE.OLEDB.12.0; data source = SightKitt2018.accdb")


    Private DataBace_Command As New OleDbCommand

    'Databace storage and handaling
    Public Data_Adaptor As OleDbDataAdapter
    Public Data_Table As DataTable

    'Query Parameters
    Public Params As New List(Of OleDbParameter)


    Public recordCounter As Integer
    Public Exception As String


    Public Sub ExecQuery(Query As String)
        recordCounter = 0
        Exception = ""

        Try
            Data_Connection.Open()

            DataBace_Command = New OleDbCommand(Query, Data_Connection)

            Params.ForEach(Sub(p) DataBace_Command.Parameters.Add(p))

            Params.Clear()

            Data_Table = New DataTable
            Data_Adaptor = New OleDbDataAdapter(DataBace_Command)
            recordCounter = Data_Adaptor.Fill(Data_Table)






        Catch ex As Exception
            Exception = ex.Message
        End Try
        If Data_Connection.State = ConnectionState.Open Then Data_Connection.Close()


    End Sub

    Public Sub AddParam(Name As String, Value As Object)
        Dim NewParam As New OleDbParameter(Name, Value)
        Params.Add(NewParam)

    End Sub
















End Class
