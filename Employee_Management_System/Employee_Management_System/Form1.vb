Imports System.Data.OleDb

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim connString As String =
            "Provider=Microsoft.ACE.OLEDB.12.0;" &
            "Data Source=" & Application.StartupPath & "\Employee.accdb;" &
            "Persist Security Info=False;"
        Try
            Using con As New OleDbConnection(connString)
                con.Open()
                MessageBox.Show("Connection successful!",
                    "Database Connection",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information)
            End Using

        Catch ex As Exception
            MessageBox.Show("Connection failed!" & vbCrLf & vbCrLf &
                    ex.Message,
                    "Database Connection Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class