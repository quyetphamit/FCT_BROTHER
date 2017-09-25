Public Class Form1

    Private Sub Form1_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Hide()
        Control.Show()
        Control.Labelresult.BackColor = Color.Blue
        Control.Labelresult.ForeColor = Color.White
        Control.Labelresult.Text = "Wait"
        Control.textserial.Enabled = True
        Control.textserial.Text = ""
        Control.textserial.Focus()
    End Sub


End Class