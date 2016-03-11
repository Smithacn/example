Imports Excel = Microsoft.Office.Interop.Excel

Public Class Form1

    Dim xlpath1 As String = "C:\Printscreen_Project\Print_screen_Template.xls"
    Dim xlpath2 As String = "C:\Printscreen_Project\notes.pdf"
    Dim destfold As String = "C:\Printscreen_Project\"
    Dim restfold As String = "C:\Printscreen_Project\Results\"

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If (Not System.IO.Directory.Exists("C:\Printscreen_Project")) Then
            System.IO.Directory.CreateDirectory("C:\Printscreen_Project")
        End If

        If (Not System.IO.Directory.Exists("C:\Printscreen_Project\Results")) Then
            System.IO.Directory.CreateDirectory("C:\Printscreen_Project\Results")
        End If

        Dim filecopy = CreateObject("Scripting.FileSystemObject")

        If Not filecopy.FileExists(xlpath2) Then
            My.Computer.FileSystem.WriteAllBytes(destfold & "notes.pdf", My.Resources.notes, False)
        End If

        If Not filecopy.FileExists(xlpath1) Then
            My.Computer.FileSystem.WriteAllBytes(destfold & "Print_screen_Template.xls", My.Resources.Print_screen_Template, False)
        End If

        Dim j As Integer
        j = TextBox1.Text


        For i = 1 To j

                Dim xlapp As New Excel.Application
                Dim proc As Process = System.Diagnostics.Process.Start(xlpath2)
                proc.WaitForInputIdle()
                Dim title As String = proc.MainWindowTitle
                AppActivate(title)

                SendKeys.SendWait("%({PRTSC})")


                Dim xlwb As Excel.Workbook = xlapp.Workbooks.Open(xlpath1)

                Dim xlsheet As Excel.Worksheet
                xlsheet = xlwb.Worksheets("Sheet1")

            xlapp.Visible = True

            xlsheet.Activate()

            xlsheet.Cells(1, 1).Activate

            SendKeys.SendWait("^(v)")

            xlwb.SaveAs(restfold & "file" & i & ".xls")

                xlwb.Close()
                proc.CloseMainWindow()
                xlapp.Quit()
            Next
            MessageBox.Show("All Files have been created")
        Process.Start(restfold)

    End Sub
End Class
