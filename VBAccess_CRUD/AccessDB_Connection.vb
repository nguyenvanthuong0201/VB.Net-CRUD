Option Explicit On
Option Strict On

Imports System.Data
Imports System.Data.OleDb

Module AccessDB_Connection
    Public Function GetConnnectionString() As String
        Dim strCon As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & Application.StartupPath
        strCon &= "\VB_SMART_CRUD.accdb;Persist Security Info=False;"
        Return strCon
    End Function

    Public Con As New OleDbConnection(GetConnnectionString())
    Public Cmd As OleDbCommand
    Public SQL As String = String.Empty

    Public Function PerformCURD(Com As OleDbCommand) As DataTable
        Dim da As OleDbDataAdapter
        Dim dt As New DataTable()
        Try
            da = New OleDbDataAdapter
            da.SelectCommand = Com
            da.Fill(dt)

            Return dt
        Catch ex As Exception
            MessageBox.Show("An error occurred:" & ex.Message, "Perform CURD OPERATION FAILER", MessageBoxButtons.OK, MessageBoxIcon.Error)
            dt = Nothing
        End Try
        Return dt

    End Function


End Module
