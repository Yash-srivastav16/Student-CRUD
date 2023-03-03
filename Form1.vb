Imports System.Data.OleDb
Imports System.Data
Imports System.Runtime.ExceptionServices
Imports System.Security.Policy

Public Class Form1
    Sub view()

        Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\ASUS\Documents\student.mdb")
        con.Open()
        Dim adp As New OleDbDataAdapter("select * from college", con)
        Dim ds As New DataSet
        adp.Fill(ds, "student")
        DataGridView1.DataSource = ds.Tables("student")
        con.Close()
    End Sub

    Public Sub ShowData(ByVal CurrentRow)
        Try
            TextBox1.Text = ds.Tables("student").Rows(CurrentRow)("ID")
            TextBox2.Text = ds.Tables("student").Rows(CurrentRow)("Sname")
            TextBox3.Text = ds.Tables("student").Rows(CurrentRow)("Address")
            TextBox4.Text = ds.Tables("student").Rows(CurrentRow)("Course")
        Catch ex As Exception
            MsgBox(ex.Message, "error")
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\ASUS\Documents\student.mdb")
        con.Open()
        Dim cmd As New OleDbCommand()
        cmd.Connection = con
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "insert into college values (" & TextBox1.Text & ", '" & TextBox2.Text & "','" & TextBox3.Text & "', '" & TextBox4.Text & "')"
        cmd.ExecuteNonQuery()
        con.Close()
        MsgBox("Record Saved Successfully")
        view()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        view()

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\ASUS\Documents\student.mdb")
        CurrentRow = 0
        con.Open()
        adp = New OleDbDataAdapter("SELECT * FROM college", con)
        adp.Fill(ds, "student")
        ShowData(CurrentRow)
        con.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\ASUS\Documents\student.mdb")
        con.Open()
        Dim cmd As New OleDbCommand()
        cmd.Connection = con
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "update college set Sname='" & TextBox2.Text & "', Address='" & TextBox3.Text & "', Course='" & TextBox4.Text & "' where id=" & TextBox1.Text & ""
        cmd.ExecuteNonQuery()
        con.Close()
        MsgBox("Record Updated Succesfully")
        view()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\ASUS\Documents\student.mdb")
        con.Open()
        Dim cmd As New OleDbCommand()
        cmd.Connection = con
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "delete from college where ID=" & TextBox1.Text & ""
        cmd.ExecuteNonQuery()
        con.Close()
        MsgBox("Record Deleted Succesfully")
        view()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If CurrentRow <> 0 Then
            CurrentRow -= 1
            ShowData(CurrentRow)
        Else
            MsgBox("First Record is Reached!", MsgBoxStyle.Exclamation)
        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        CurrentRow = 0
        ShowData(CurrentRow)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        CurrentRow = ds.Tables("student").Rows.Count - 1
        ShowData(CurrentRow)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If CurrentRow = ds.Tables("student").Rows.Count - 1 Then
            MsgBox("Last Record is Reached", MsgBoxStyle.Exclamation)
        Else
            CurrentRow += 1
            ShowData(CurrentRow)
        End If
    End Sub
End Class
