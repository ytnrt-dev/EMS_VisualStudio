Imports System.Data.OleDb

Public Class Form1
    Dim sex, position As String
    Dim lates As Integer = 0

    Dim connString As String =
            "Provider=Microsoft.ACE.OLEDB.12.0;" &
            "Data Source=" & Application.StartupPath & "\Payroll.accdb;" &
            "Persist Security Info=False;"

    Private Sub displayRECORDS()
        Using con As New OleDbConnection(connString)
            Dim sqlDisplay As String = "SELECT * FROM employee_info"
            Using adapter As New OleDbDataAdapter(sqlDisplay, con)
                Dim table As New DataTable()
                adapter.Fill(table)
                DataGridView1.DataSource = table
            End Using

        End Using
    End Sub

    Private Sub totalDays()
        Dim startDate As Date = DateTimePicker1.Value.Date
        Dim endDate As Date = DateTimePicker2.Value.Date

        If endDate >= startDate Then
            Dim totalDays As Integer = 0
            Dim currentDay As Date = startDate



            While currentDay <= endDate
                If currentDay.DayOfWeek <> DayOfWeek.Saturday AndAlso currentDay.DayOfWeek <> DayOfWeek.Sunday Then
                    totalDays += 1
                End If
                currentDay = currentDay.AddDays(1)
            End While
            TextBox8.Text = totalDays.ToString
        Else
            TextBox8.Text = "0"
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

        displayRECORDS()
        ComboBox2.Enabled = False
        TextBox23.Enabled = False
        TextBox21.Enabled = False

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Label22_Click(sender As Object, e As EventArgs) Handles Label22.Click

    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click

    End Sub

    Private Sub TextBox19_TextChanged(sender As Object, e As EventArgs) Handles TextBox19.TextChanged

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If RadioButton1.Checked = True Then
            sex = "Male"
        End If
        If RadioButton2.Checked = True Then
            sex = "Female"
        End If

        Using con As New OleDbConnection(connString)
            Dim sqlInsert As String = "INSERT INTO employee_info (ID, Lastname, Firstname, Middlename, Gender, [Position], Basic_salary) VALUES ('" & TextBox19.Text & "', '" & TextBox18.Text & "', '" & TextBox17.Text & "', '" & TextBox16.Text & "', '" & sex & "', '" & ComboBox1.SelectedItem & "', '" & TextBox20.Text & "')"
            Using cmd As New OleDbCommand(sqlInsert, con)
                con.Open()
                cmd.ExecuteNonQuery()
                MessageBox.Show("Record added successfully!")
                displayRECORDS()

                TextBox19.Clear()
                TextBox18.Clear()
                TextBox17.Clear()
                TextBox16.Clear()
                TextBox20.Clear()
            End Using
        End Using

    End Sub

    Private Sub TextBox18_TextChanged(sender As Object, e As EventArgs) Handles TextBox18.TextChanged

    End Sub

    ' =====================================================================
    '  ADDED: EMPTY HANDLERS FOR THE REMAINING CONTROLS (no code inside)
    ' =====================================================================

    ' ---- Employee Profile ----
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Dim sqlDisplay As String
        Using conn As New OleDbConnection(connString)
            If String.IsNullOrWhiteSpace(TextBox1.Text) Then
                sqlDisplay = "SELECT * FROM employee_info"
            Else
                sqlDisplay = "SELECT * FROM employee_info WHERE Lastname = '" & TextBox1.Text & "'"
            End If

            Using adapter As New OleDbDataAdapter(sqlDisplay, conn)
                Dim table As New DataTable()
                adapter.Fill(table)
                DataGridView1.DataSource = table


            End Using

        End Using
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Dim sqlDisplay As String
        Using conn As New OleDbConnection(connString)
            If String.IsNullOrWhiteSpace(TextBox2.Text) Then
                sqlDisplay = "SELECT * FROM employee_info"
            Else
                sqlDisplay = "SELECT * FROM employee_info WHERE ID = '" & TextBox2.Text & "'"
            End If

            Using adapter As New OleDbDataAdapter(sqlDisplay, conn)
                Dim table As New DataTable()
                adapter.Fill(table)
                DataGridView1.DataSource = table


            End Using

        End Using
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

    ' ---- Payroll ----
    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        Using conn As New OleDbConnection(connString)
            Dim sqlDisplay As String = "SELECT * FROM employee_info WHERE ID = '" & TextBox3.Text & "'"

            Using adapter As New OleDbDataAdapter(sqlDisplay, conn)
                Dim table As New DataTable
                adapter.Fill(table)

                If table.Rows.Count > 0 Then
                    TextBox4.Text = table.Rows(0)("Lastname").ToString
                    TextBox5.Text = table.Rows(0)("Firstname").ToString
                    TextBox6.Text = table.Rows(0)("Middlename").ToString
                    TextBox9.Text = table.Rows(0)("Basic_salary").ToString

                End If

            End Using

        End Using
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        TextBox7.Text = Val(TextBox9.Text) * Val(TextBox8.Text)
        TextBox11.Text = 150
        TextBox12.Text = 100
        TextBox10.Text = 150
        TextBox13.Text = 0
        TextBox14.Text = Val(TextBox11.Text) + (TextBox12.Text) + (TextBox10.Text) + (TextBox13.Text)
    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged

    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged

    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged

    End Sub

    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged
        Dim deduction As Integer
        deduction = Val(TextBox11.Text) + (TextBox12.Text) + (TextBox10.Text) + (TextBox13.Text)
        TextBox14.Text = deduction
        TextBox15.Text = Val(TextBox7.Text) - deduction


    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged

    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        totalDays()

    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        totalDays()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub

    ' ---- Add Employees ----
    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles TextBox16.TextChanged

    End Sub

    Private Sub TextBox17_TextChanged(sender As Object, e As EventArgs) Handles TextBox17.TextChanged

    End Sub

    Private Sub TextBox20_TextChanged(sender As Object, e As EventArgs) Handles TextBox20.TextChanged

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

    End Sub

    ' ---- Update Employees ----
    Private Sub TextBox21_TextChanged(sender As Object, e As EventArgs) Handles TextBox21.TextChanged

    End Sub

    Private Sub TextBox22_TextChanged(sender As Object, e As EventArgs) Handles TextBox22.TextChanged
        Using conn As New OleDbConnection(connString)
            Dim sqlDisplay As String = "SELECT * FROM employee_info WHERE ID = '" & TextBox22.Text & "'"
            Using adapter As New OleDbDataAdapter(sqlDisplay, conn)
                Dim table As New DataTable()
                adapter.Fill(table)
                DataGridView2.DataSource = table

                If table.Rows.Count > 0 Then
                    Dim Lastname As String = table.Rows(0)("Lastname").ToString()
                    Dim Firstname As String = table.Rows(0)("Firstname").ToString()
                    Dim Middlename As String = table.Rows(0)("Middlename").ToString()
                    Dim Position As String = table.Rows(0)("Position").ToString()
                    Dim Salary As String = table.Rows(0)("Basic_salary").ToString()

                    TextBox21.Text = Lastname & "," & " " & Firstname & " " & Middlename
                    ComboBox2.Text = Position
                    TextBox23.Text = Salary

                End If

            End Using
        End Using
    End Sub

    Private Sub TextBox23_TextChanged(sender As Object, e As EventArgs) Handles TextBox23.TextChanged

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Using con As New OleDbConnection(connString)
            Dim sqlUpdate As String = "UPDATE employee_info SET [Position] = '" & ComboBox2.Text & "', Basic_salary = '" & TextBox23.Text & "' WHERE ID = '" & TextBox22.Text & "'"
            Using cmd As New OleDbCommand(sqlUpdate, con)
                con.Open()
                cmd.ExecuteNonQuery()
                MessageBox.Show("Employee Record Updated Successfully!")
                displayRECORDS()

                Dim sqlDisplay As String = "SELECT * FROM employee_info WHERE ID = '" & TextBox22.Text & "'"
                Using adapter As New OleDbDataAdapter(sqlDisplay, con)
                    Dim table As New DataTable
                    adapter.Fill(table)
                    DataGridView2.DataSource = table

                End Using

            End Using
        End Using
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        ComboBox2.Enabled = True
        TextBox21.Enabled = True
        TextBox23.Enabled = True

    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

End Class