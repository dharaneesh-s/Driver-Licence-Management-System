Imports System.Data.OleDb

Public Class Form6
    Dim cn As OleDbConnection
    Dim cmd As OleDbCommand
    Dim da As OleDbDataAdapter
    Dim ds As DataSet
    Dim STrPath = System.IO.Directory.GetCurrentDirectory & "\dlms1.mdb"

    Private Sub Form6_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

    End Sub
    Private Sub Form6_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Clear()
        Label4.Text = ""
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub CheckedListBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBox1.LostFocus
        Dim strdata As String
       
        For Each strdata In CheckedListBox1.CheckedItems
            TextBox1.Text &= strdata & " , "
        Next
    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckedListBox1.SelectedIndexChanged

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox2.Text = "" Then
            Exit Sub
        End If


        Try
            cn.Open()
            Dim sql As String = "update [dlms] set [License Category]='" & TextBox1.Text & "' where [License No]='" & TextBox2.Text & "'"
            Dim ct As OleDbCommand
            ct = New OleDbCommand(sql, cn)


            Dim temp As Integer = ct.ExecuteNonQuery()
            If temp > 0 Then
                MsgBox("Congratulation Vehicle Added", MsgBoxStyle.Information)
            Else
                MsgBox("Try Again,Failure")
            End If

        Catch ex As Exception

        End Try
        cn.Close()
    End Sub

    Private Sub TextBox2_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.Enter

      
    End Sub

    Private Sub TextBox2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.LostFocus
       

    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TextBox1.Text = ""
       
        If TextBox2.Text = "" Then
            MessageBox.Show("No Data Entry Enter Valid License No", "Error")
            TextBox2.Focus()
        End If
        TextBox1.DataBindings.Clear()
        Try
            cn = New OleDbConnection("Provider=microsoft.jet.oledb.4.0; data source='" & STrPath & "';")
            cmd = New OleDbCommand("select * from dlms order by License No asc")
            cn.Open()

            cmd.Connection = cn
            da = New OleDbDataAdapter("select [License Category] from dlms where [License No]='" & TextBox2.Text & "'", cn)
            ds = New DataSet("ms")
            da.Fill(ds, "ms")

            TextBox1.DataBindings.Add("text", ds, "ms.License Category")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "")
            End
        End Try
        cn.Close()
        If TextBox1.Text <> "" Then
            Button1.Enabled = True
            Label4.Text = "Valid License"
        Else
            Button1.Enabled = False
            Label4.Text = "Invalid License"
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub
End Class