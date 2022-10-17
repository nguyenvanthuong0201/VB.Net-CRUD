Option Explicit On
Option Strict On

Imports System.Data.OleDb


Public Class Form1

    Private ID As String = ""
    Private intRow As Integer = 0

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ResetMe()
        LoadData()
    End Sub

    Private Sub LoadData(Optional keyword As String = "")
        SQL = "SELECT Auto_ID, First_Name, Last_Name, [First_Name] + ' ' + [Last_Name] AS Full_Name, Gender FROM TBL_SMART_CRUD " &
            "WHERE [First_Name] + ' ' + [Last_Name] LIKE @keyword1 OR Gender = @keyword2 ORDER BY Auto_ID ASC "

        Dim strKeyword As String = String.Format("%{0}%", keyword)

        Cmd = New OleDbCommand(SQL, Con)

        Cmd.Parameters.Clear()
        Cmd.Parameters.AddWithValue("keyword1", strKeyword)
        Cmd.Parameters.AddWithValue("keyword2", keyword)

        Dim dt As DataTable = PerformCURD(Cmd)
        If dt.Rows.Count > 0 Then
            intRow = Convert.ToInt32(dt.Rows.Count.ToString())
        Else
            intRow = 0
        End If

        With DataGridView1
            .MultiSelect = False
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .AutoGenerateColumns = True
            .DataSource = dt

            .Columns(0).HeaderText = "ID"
            .Columns(1).HeaderText = "First Name"
            .Columns(2).HeaderText = "Last Name"
            .Columns(3).HeaderText = "Full Name"
            .Columns(4).HeaderText = "Gender"

            .Columns(0).Width = 85
            .Columns(1).Width = 170
            .Columns(2).Width = 170
            .Columns(3).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            .Columns(4).Width = 150


        End With

        ToolStripStatusLabel1.Text = "Number of row(s): " & intRow.ToString()
    End Sub

    Private Sub ResetMe()

        Me.ID = ""

        FirstNameTextBox.Text = ""
        LastNameTextBox.Text = ""

        If GenderComboBox.Items.Count > 0 Then
            GenderComboBox.SelectedIndex = 0
        End If

        UpdateButton.Text = "UPDATE ()"
        DeleteButton.Text = "DELETE ()"

        With KeywordTextBox
            .Clear()
            .Select()
        End With
    End Sub

    Private Sub Execute(MySQL As String, Optional Parameter As String = "")
        Cmd = New OleDbCommand(MySQL, Con)
        AddParameter(Parameter)
        PerformCURD(Cmd)
    End Sub

    Private Sub AddParameter(str As String)
        Cmd.Parameters.Clear()
        If str = "DELETE" And Not String.IsNullOrEmpty(Me.ID) Then
            Cmd.Parameters.AddWithValue("ID", Me.ID)
        End If

        Cmd.Parameters.AddWithValue("FirstName", FirstNameTextBox.Text.Trim())
        Cmd.Parameters.AddWithValue("LastName", LastNameTextBox.Text.Trim())
        Cmd.Parameters.AddWithValue("Gender", GenderComboBox.SelectedItem.ToString())

        If str = "Update" And Not String.IsNullOrEmpty(Me.ID) Then
            Cmd.Parameters.AddWithValue("ID", Me.ID)
        End If
    End Sub

    Private Sub InsertButton_Click(sender As Object, e As EventArgs) Handles InsertButton.Click
        If String.IsNullOrEmpty(Me.FirstNameTextBox.Text.Trim()) Or String.IsNullOrEmpty(Me.LastNameTextBox.Text.Trim()) Then
            MessageBox.Show("Please input first name and last name.", "Access: Inssert data", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
        End If
        SQL = "INSERT INTO TBL_SMART_CRUD(First_Name, Last_Name, Gender) VALUES(@FirstName, @LastName, @Gender)"
        Execute(SQL, "Insert")
        MessageBox.Show("The record has been saved", "Access: Inssert data", MessageBoxButtons.OK, MessageBoxIcon.Information)

        LoadData()

        ResetMe()

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim dgv As DataGridView = DataGridView1

        If e.RowIndex <> -1 Then
            Me.ID = Convert.ToString(dgv.CurrentRow.Cells(0).Value).Trim()
            UpdateButton.Text = "UPDATE (" & Me.ID & ")"
            DeleteButton.Text = "DELETE (" & Me.ID & ")"

            FirstNameTextBox.Text = Convert.ToString(dgv.CurrentRow.Cells(1).Value).Trim()
            LastNameTextBox.Text = Convert.ToString(dgv.CurrentRow.Cells(0).Value).Trim()
            GenderComboBox.SelectedItem = Convert.ToString(dgv.CurrentRow.Cells(4).Value).Trim()

        End If
    End Sub
End Class
