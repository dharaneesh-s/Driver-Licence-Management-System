Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If (TextBox1.Text = "lumbini" And TextBox2.Text = "students") Then
            Me.Hide()
            Form2.ShowDialog()
        ElseIf (TextBox1.Text = "" Or TextBox2.Text = "") Then
            MsgBox("Any Field Should Not Be Blank", MsgBoxStyle.Information, "DLMS-Blank Data")
            TextBox1.Clear()
            TextBox2.Clear()
            TextBox1.Focus()
        Else
            MsgBox("Enter Correct Username And Password", MsgBoxStyle.Information, "DLMS-Error")
            TextBox1.Clear()
            TextBox2.Clear()
            TextBox1.Focus()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        End
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
