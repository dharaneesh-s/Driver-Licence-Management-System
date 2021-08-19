Imports System.Data.OleDb
Public Class Form5
    Dim cn As OleDbConnection
    Dim cmd As OleDbCommand
    Dim da As OleDbDataAdapter
    Dim ds As DataSet
    Dim STrPath = System.IO.Directory.GetCurrentDirectory & "\dlms1.mdb"
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Form5_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

    End Sub

    Private Sub Form5_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim d As DateTime = Now

        DateTimePicker3.Text = d.AddYears(5).ToLongDateString()
        DateTimePicker1.Value = DateTimePicker1.MinDate
       


      

    End Sub

    Private Sub DateTimePicker2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker2.ValueChanged





    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim expiredate As Date
        expiredate = (DateTimePicker1.Text)
        Dim currentdate As Date = Now
        Dim renewontime As Date
        Dim renewunder1yr As Date
        Dim renewunder2yr As Date
        renewontime = expiredate.AddMonths(3)
        renewunder1yr = renewontime.AddYears(1)
        renewunder2yr = renewunder1yr.AddYears(1)
        Dim datestate As Integer = Date.Compare(currentdate, expiredate)
        
        If datestate > 0 Then
            Dim a As Integer = Date.Compare(renewontime, currentdate)
            Dim b As Integer = Date.Compare(renewunder1yr, currentdate)
            Dim c As Integer = Date.Compare(renewunder2yr, currentdate)
            If a >= 0 Then
                MsgBox("Your License is Renewed,Please pay charge for renew License")
            ElseIf b >= 0 Then
                MsgBox("Your License is Renewed,pay renew charge plus 1 year fine")
            ElseIf c >= 0 Then
                MsgBox("Your License is Renewed,pay renew charge plus 1 year fine")
            Else
                MsgBox("License Expired,Not Renewable")
                Exit Sub
            End If
        Else
            MsgBox("License Not Expired")
            Exit Sub
        End If
        renew()


       
    End Sub
    Sub renew()

        cn.Open()
        Try
            Dim sql As String = "update [dlms] set [Expire Date]='" & DateTimePicker3.Text & "' where [License No]='" & TextBox1.Text & "'"
            Dim ct As OleDbCommand
            ct = New OleDbCommand(sql, cn)
            ct.ExecuteNonQuery()

        Catch ex As Exception
        End Try
        cn.Close()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Button1.Enabled = True
        DateTimePicker1.DataBindings.Clear()
        DateTimePicker1.Value = DateTimePicker1.MinDate

        Try
            cn = New OleDbConnection("Provider=microsoft.jet.oledb.4.0; data source='" & STrPath & "';")
            cmd = New OleDbCommand("select * from dlms order by License No asc")
            cn.Open()

            cmd.Connection = cn
            da = New OleDbDataAdapter("select [Expire Date] from [dlms] where [License No]='" & TextBox1.Text & "'", cn)
            ds = New DataSet("ms")
            da.Fill(ds, "ms")

            DateTimePicker1.DataBindings.Add("text", ds, "ms.Expire Date")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "")
            End
        End Try
        cn.Close()
        If DateTimePicker1.Value = DateTimePicker1.MinDate Then
            Button1.Enabled = False
            Label5.Text = "Invalid Licence"
        Else
            Label5.Text = "Valid Licence"

        End If
    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged

    End Sub

    Private Sub DateTimePicker3_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker3.ValueChanged

    End Sub
End Class