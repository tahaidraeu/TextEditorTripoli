Module TextManipulation
    Public Sub Cut()
        Try
            MainForm.TextBox1.Cut()
        Catch ex As Exception
            MessageBox.Show("Error cutting text: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub Copy()
        Try
            MainForm.TextBox1.Copy()
        Catch ex As Exception
            MessageBox.Show("Error copying text: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub Paste()
        Try
            MainForm.TextBox1.Paste()
        Catch ex As Exception
            MessageBox.Show("Error pasting text: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub Undo()
        Try
            MainForm.TextBox1.Undo()
        Catch ex As Exception
            MessageBox.Show("Error undoing action: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub Redo()
        Try
            MainForm.TextBox1.Redo()
        Catch ex As Exception
            MessageBox.Show("Error redoing action: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub SelectAll()
        Try
            MainForm.TextBox1.SelectAll()
        Catch ex As Exception
            MessageBox.Show("Error selecting all text: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub ClearAll()
        Try
            MainForm.TextBox1.Clear()
        Catch ex As Exception
            MessageBox.Show("Error clearing all text: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Module
