Imports System.Data.SQLite
Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        SQLiteConnection.CreateFile("Mydatabase.sqlite")

        Dim con As New SQLiteConnection("Data Source=Mydatabase.sqlite;Version=3;")
        con.Open()

        Dim sql As String = "create table employee (name varchar(20), salary int)"

        Dim cmd As New SQLiteCommand(sql, con)
        cmd.ExecuteNonQuery()

        sql = "insert into employee (name, salary) values ('A', 3000)"



        Dim cmd1 As New SQLiteCommand(sql, con)
        cmd1.ExecuteNonQuery()

        sql = "insert into employee (name, salary) values ('B', 4000)"

        Dim cmd2 As New SQLiteCommand(sql, con)
        cmd2.ExecuteNonQuery()

        sql = "insert into employee (name, salary) values ('C', 5000)"

        Dim cmd3 As New SQLiteCommand(sql, con)
        cmd3.ExecuteNonQuery()

        Dim ad As New SQLiteDataAdapter("Select * from employee", con)
        Dim dt As New DataSet()

        ad.Fill(dt)

        DataGridView1.DataSource = dt.Tables(0)

    End Sub
End Class
