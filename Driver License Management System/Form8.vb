Public Class Form8

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        Form4.Show()
        Me.Hide()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        Form7.Show()
    End Sub

    Private Sub Form8_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        End
    End Sub

    Private Sub Form8_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        MenuStrip1.BackColor = Color.Beige
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        End
    End Sub

    Private Sub AddVehicleToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddVehicleToolStripMenuItem.Click
        Form6.Show()
    End Sub

    Private Sub RenewLicenseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RenewLicenseToolStripMenuItem.Click
        Form5.Show()
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        Form5.Show()
    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        Form6.Show()
    End Sub

    Private Sub RecordInTableToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RecordInTableToolStripMenuItem.Click
        Form3.Show()
        Me.Hide()
    End Sub

    Private Sub IssueToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IssueToolStripMenuItem.Click
        Form4.Show()
        Me.Hide()

    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Form4.Show()
        Me.Hide()

    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Form4.Show()
        Me.Hide()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Form4.Show()
        Me.Hide()
    End Sub
End Class