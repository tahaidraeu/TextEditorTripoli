Imports System.IO

Module FileHandling
    Public Sub OpenFile(filePath As String)
        Try
            Dim content As String = File.ReadAllText(filePath)
            MainForm.TextBox1.Text = content
            MainForm.Text = "Text Editor - " & Path.GetFileName(filePath)
        Catch ex As Exception
            MessageBox.Show("Error opening the file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub SaveFile(filePath As String, content As String)
        Try
            File.WriteAllText(filePath, content)
            MainForm.Text = "Text Editor - " & Path.GetFileName(filePath)
            MessageBox.Show("File saved successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Error saving the file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Module

