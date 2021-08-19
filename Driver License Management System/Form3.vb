Imports System.Data.OleDb
Public Class Form3
    Dim cn As OleDbConnection
    Dim cmd As OleDbCommand
    Dim da As OleDbDataAdapter
    Dim ds As New DataSet
    Friend WithEvents db As System.Windows.Forms.BindingSource

    Private Sub Form3_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        Form8.Show()
    End Sub




    
    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'TODO: This line of code loads data into the 'Dlms1DataSet.dlms' table. You can move, or remove it, as needed.
     

        Dim STrPath = System.IO.Directory.GetCurrentDirectory & "\dlms1.mdb"

        Try
            cn = New OleDbConnection("Provider=microsoft.jet.oledb.4.0; data source='" & STrPath & "';")
            cmd = New OleDbCommand("Select * from dlms ")

            cn.Open()


            cmd.Connection = cn
            da = New OleDb.OleDbDataAdapter("SELECT * FROM dlms", cn)

            ds = New DataSet("MS")
            da.Fill(ds, "MS")

            DlmsDataGridView.DataSource = ds
            DlmsDataGridView.DataMember = "MS"
            Dim x = ds.Tables("ms").Rows.Count()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "")
            End
        End Try
        cn.Close()
      

    End Sub
        sub con()

        Dim STrPath = System.IO.Directory.GetCurrentDirectory & "\dlms1.mdb"

        Try
            cn = New OleDbConnection("Provider=microsoft.jet.oledb.4.0; data source='" & STrPath & "';")
            cmd = New OleDbCommand("Select * from dlms ")

            cn.Open()
       

            cmd.Connection = cn
            If ComboBox1.SelectedItem = "License No" Then
                da = New OleDb.OleDbDataAdapter("SELECT * FROM [dlms] Where [License No] ='" & (TextBox1.Text) & "' ", cn)
            ElseIf ComboBox1.SelectedItem = "Citizenship No" Then
                da = New OleDb.OleDbDataAdapter("SELECT * FROM [dlms] Where [Citizenship No] ='" & (TextBox1.Text) & "' ", cn)
            Else
                MsgBox("select search item")
                Exit Sub
            End If


            ds = New DataSet("MS")
            da.Fill(ds, "MS")

            DlmsDataGridView.DataSource = ds
            DlmsDataGridView.DataMember = "MS"
            Dim x = ds.Tables("ms").Rows.Count()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "")
            End
        End Try
        cn.Close()
    End Sub

   


   

   

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.Close()
        Form8.Show()

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

       


        If ComboBox1.SelectedItem = "License No" And TextBox1.Text <> "" Then
            con()
            Exit Sub
        ElseIf ComboBox1.SelectedItem = "Citizenship No" And TextBox1.Text <> "" Then
            con()
            Exit Sub
        ElseIf ComboBox1.SelectedItem = Nothing Then
            MsgBox("Choose Your Search Category")
            TextBox1.Clear()
            ComboBox1.Focus()
        End If
        If ComboBox1.SelectedItem = "License No" Or ComboBox1.SelectedItem = "Citizenship No" And TextBox1.Text = Nothing Then
            If ComboBox1.SelectedItem = "License No" Then
                MsgBox("Enter License No")
                TextBox1.Focus()
            End If
            If ComboBox1.SelectedItem = "Citizenship No" Then
                MsgBox("Enter Citizenship No")
                TextBox1.Focus()
            End If
        End If
        


    End Sub

    Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
       
      

        Dim STrPath = System.IO.Directory.GetCurrentDirectory & "\dlms1.mdb"

        Try
            cn = New OleDbConnection("Provider=microsoft.jet.oledb.4.0; data source='" & STrPath & "';")
            cmd = New OleDbCommand("Select * from dlms ")

            cn.Open()


            cmd.Connection = cn
            da = New OleDb.OleDbDataAdapter("SELECT * FROM dlms", cn)

            ds = New DataSet("MS")
            da.Fill(ds, "MS")

            DlmsDataGridView.DataSource = ds
            DlmsDataGridView.DataMember = "MS"
           
            Dim x = ds.Tables("ms").Rows.Count()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "")
            End
        End Try
        cn.Close()
    End Sub
End Class