Imports System.Data.OleDb
Public Class Form4
    Dim cn As OleDbConnection
    Dim cmd As OleDbCommand
    Dim da As OleDbDataAdapter
    Dim ds As DataSet
   

    Dim STrPath = System.IO.Directory.GetCurrentDirectory & "\dlms1.mdb"

    Private Sub Form4_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        Form8.Show()
    End Sub

    Private Sub Form4_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CheckedListBox1.Visible = False
        Textcategory.Location = New System.Drawing.Point(634, 477)
        Textcategory.Size = New System.Drawing.Point(137, 75)

        jdcomplex()

        Me.BindingContext(ds, "ms").Position = 0
        Label21.Text = (Me.BindingContext(ds, "ms").Position + _
       1).ToString & " of " & Me.BindingContext(ds, _
       "ms").Count.ToString
        check()
    End Sub

    Sub jdcomplex()

        Try
            cn = New OleDbConnection("Provider=microsoft.jet.oledb.4.0; data source='" & STrPath & "';")
            cmd = New OleDbCommand("select * from dlms order by License No asc")
            cn.Open()

            cmd.Connection = cn
            da = New OleDbDataAdapter("select * from dlms", cn)
            ds = New DataSet("ms")
            da.Fill(ds, "ms")
            BindingSource1.DataSource = ds
            BindingSource1.DataMember = "ms"
            kamal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "")
            End
        End Try
        cn.Close()

    End Sub
    Sub kamal()

        Textlicense.DataBindings.Add("text", ds, "ms.License No")
        Texthollo.DataBindings.Add("text", ds, "ms.Hollogram No")
        Textownersname.DataBindings.Add("text", ds, "ms.Owners Name")
        dissue.DataBindings.Add("text", ds, "ms.Issue Date")
        dexpire.DataBindings.Add("text", ds, "ms.Expire Date")
        Textcardno.DataBindings.Add("text", ds, "ms.Card No")
        Textage.DataBindings.Add("text", ds, "ms.Age")
        Textbloodgroup.DataBindings.Add("text", ds, "ms.Blood Group")
        Texthusband.DataBindings.Add("text", ds, "ms.Father/Husband Name")
        Combolicence.DataBindings.Add("text", ds, "ms.Zone")
        Combodistrict.DataBindings.Add("text", ds, "ms.District")
        Textmunci.DataBindings.Add("text", ds, "ms.Muncipality/Vdc")
        Textwardno.DataBindings.Add("text", ds, "ms.Ward No")
        Texttole.DataBindings.Add("text", ds, "ms.Tole")
        Textphone.DataBindings.Add("text", ds, "ms.Phone No")
        Textblockno.DataBindings.Add("text", ds, "ms.Block No")
        Textcontact.DataBindings.Add("text", ds, "ms.Contact Address/Office")
        Texttempadd.DataBindings.Add("text", ds, "ms.Temporary Address")
        Textoccupation.DataBindings.Add("text", ds, "ms.Occupation")
        Textqualification.DataBindings.Add("text", ds, "ms.Qualification")
        Textcitizen.DataBindings.Add("text", ds, "ms.Citizenship No")
        Combocitizen.DataBindings.Add("text", ds, "ms.Citizenship District")
        Textpassport.DataBindings.Add("text", ds, "ms.Passport No")
        Textdatarec.DataBindings.Add("text", ds, "ms.Data Recorder")
        Textlprovider.DataBindings.Add("text", ds, "ms.License provider")
        Textcategory.DataBindings.Add("text", ds, "ms.License Category")

    End Sub
    Sub clear()

        Textlicense.DataBindings.Clear()
        Texthollo.DataBindings.Clear()
        Textownersname.DataBindings.Clear()
        dissue.DataBindings.Clear()
        dexpire.DataBindings.Clear()
        Textcardno.DataBindings.Clear()
        Textage.DataBindings.Clear()
        Textbloodgroup.DataBindings.Clear()
        Texthusband.DataBindings.Clear()
        Combolicence.DataBindings.Clear()
        Combodistrict.DataBindings.Clear()
        Textmunci.DataBindings.Clear()
        Textwardno.DataBindings.Clear()
        Texttole.DataBindings.Clear()
        Textphone.DataBindings.Clear()
        Textblockno.DataBindings.Clear()
        Textcontact.DataBindings.Clear()
        Texttempadd.DataBindings.Clear()
        Textoccupation.DataBindings.Clear()
        Textqualification.DataBindings.Clear()
        Textcitizen.DataBindings.Clear()
        Combocitizen.DataBindings.Clear()
        Textpassport.DataBindings.Clear()
        Textdatarec.DataBindings.Clear()
        Textlprovider.DataBindings.Clear()
        Textcategory.DataBindings.Clear()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Me.BindingContext(ds, "ms").Position = 0
        Label21.Text = (Me.BindingContext(ds, "ms").Position + _
        1).ToString & " of " & Me.BindingContext(ds, _
        "ms").Count.ToString
        check()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Me.BindingContext(ds, "ms").Position = Me.BindingContext(ds, "ms").Position + 1
        Label21.Text = (Me.BindingContext(ds, "ms").Position + _
       1).ToString & " of " & Me.BindingContext(ds, _
       "ms").Count.ToString
        check()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Me.BindingContext(ds, "ms").Position = Me.BindingContext(ds, "ms").Position - 1
        Label21.Text = (Me.BindingContext(ds, "ms").Position + _
       1).ToString & " of " & Me.BindingContext(ds, _
       "ms").Count.ToString
        check()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Me.BindingContext(ds, "ms").Position = Me.BindingContext(ds, "ms").Count - 1
        Label21.Text = (Me.BindingContext(ds, "ms").Position + _
       1).ToString & " of " & Me.BindingContext(ds, _
       "ms").Count.ToString
        check()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Textcategory.TabIndex = 200
        status.Text = ""
        textclear()
        Dim d As DateTime = Now

        dexpire.Text = d.AddYears(5).ToLongDateString()
        CheckedListBox1.Visible = True
        CheckedListBox1.Location = New System.Drawing.Point(634, 477)
        CheckedListBox1.Size = New System.Drawing.Size(137, 49)
        Textcategory.Location = New System.Drawing.Point(634, 533)
        Textcategory.Size = New System.Drawing.Size(137, 49)

        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button6.Enabled = False
        Button7.Enabled = False
        Button8.Enabled = True
        Label21.Text = "Adding"


    End Sub
    Sub textclear()
        Textlicense.Clear()

        Texthollo.Clear()

        Textownersname.Clear()
        dissue.Text = Now

        dexpire.Text = Now

        Textcardno.Clear()
        Textage.Clear()
        Textbloodgroup.Text = ""
        Texthusband.Clear()
        Combolicence.Text = ""
        Combodistrict.Text = ""
        Textmunci.Clear()
        Textwardno.Clear()
        Texttole.Clear()
        Textphone.Clear()
        Textblockno.Clear()
        Textcontact.Clear()
        Texttempadd.Clear()
        Textoccupation.Clear()
        Textqualification.Clear()
        Textcitizen.Clear()
        Combocitizen.Text = ""
        Textpassport.Clear()

        Textdatarec.Clear()
        Textlprovider.Clear()
        Textcategory.Text = ""
        Textlicense.Focus()

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
      
        Try
            cn.Open()

            Dim sql As String = "update [dlms] set [Hollogram No]=" & CInt(Texthollo.Text) & ",[Owners Name]='" & Textownersname.Text & "',[Issue Date]='" & CDate(dissue.Text) & "',[Expire Date]='" & CDate(dexpire.Text) & "',[Card No]=" & CInt(Textcardno.Text) & ",[Age]='" & Textage.Text & "',[Blood Group]='" & Textbloodgroup.Text & "',[Father/Husband Name]='" & Texthusband.Text & "',[Zone]='" & Combolicence.Text & "',[District]='" & Combodistrict.Text & "',[Muncipality/Vdc]='" & Textmunci.Text & "',[Ward No]='" & (Textwardno.Text) & "',[Tole]='" & _
            Texttole.Text & "',[Contact Address/Office]='" & Textcontact.Text & "',[Phone No]='" & Textphone.Text & "',[Block No]='" & Textblockno.Text & "',[Occupation]='" & Textoccupation.Text & "',[Qualification]='" & Textqualification.Text & "',[Citizenship No]='" & Textcitizen.Text & "',[Citizenship District]='" & Combocitizen.Text & "',[Passport No]='" & Textpassport.Text & "',[License Category]='" & Textcategory.Text & "',[Data Recorder]='" & Textdatarec.Text & "',[License Provider]='" & Textlprovider.Text & "',[Temporary Address]='" & Texttempadd.Text & "' Where [License No]='" & Textlicense.Text & "'"
            Dim ct As OleDbCommand
            ct = New OleDbCommand(sql, cn)
            ct.ExecuteNonQuery()

        Catch ex As Exception
        End Try
        cn.Close()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim ans = MsgBox("Do you really want to delete this record?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "")
        If ans = vbNo Then
            Exit Sub
        End If
        Try

            cn = New OleDbConnection("Provider=microsoft.jet.oledb.4.0; data source='" & STrPath & "';")
            cn.Open()
            Dim str As String = "delete from [dlms] where [License No]='" & Textlicense.Text & "'"
            'string stores the command and CInt is used to convert number to string
            cmd = New OleDbCommand(str, cn)
            Dim temp As Integer = cmd.ExecuteNonQuery

            'displays number of records inserted
            If temp > 0 Then
                MsgBox("Data Are deleted Successfully")
                clear()
                jdcomplex()

            Else
                MsgBox("Try Again")

            End If

        Catch ex As Exception

        End Try
        cn.Close()

        Label21.Text = (Me.BindingContext(ds, "ms").Position + _
       1).ToString & " of " & Me.BindingContext(ds, _
       "ms").Count.ToString
        check()
        CheckedListBox1.Visible = False
        Textcategory.Location = New System.Drawing.Point(634, 477)
        Textcategory.Size = New System.Drawing.Point(137, 75)

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click

        ' If Texttempadd.Text = "" Or Textpassport.Text = "" Or Textblockno.Text = "" Or Textcontact.Text = "" Or Textwardno.Text = "" Then
        'Texttempadd.Text = ""
        'Textpassport.Text = ""
        'Textblockno.Text = ""
        'Textcontact.Text = ""
        'Textwardno.Text = ""
        'End If
        save()
        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
        Button6.Enabled = True
        Button7.Enabled = True
        CheckedListBox1.Visible = False
        Textcategory.Location = New System.Drawing.Point(634, 477)
        Textcategory.Size = New System.Drawing.Point(137, 75)

    End Sub

    Sub save()
        Try

            cn = New OleDbConnection("Provider=microsoft.jet.oledb.4.0; data source='" & STrPath & "';")
            cn.Open()
            Dim str As String = "insert into dlms values('" & Textlicense.Text & "'," & CInt(Texthollo.Text) & ",'" & _
    Textownersname.Text & "','" & CDate(dissue.Text) & "','" & CDate(dexpire.Text) & "'," & CInt(Textcardno.Text) & ",'" & _
    Textage.Text & "','" & Textbloodgroup.Text & "','" & Texthusband.Text & "','" & Combolicence.Text & "','" & Combodistrict.Text & "','" & _
    Textmunci.Text & "','" & Textwardno.Text & "','" & Texttole.Text & "','" & Textcontact.Text & "','" & Textphone.Text & "','" & Textblockno.Text & "','" & _
    Textoccupation.Text & "','" & Textqualification.Text & "','" & Textcitizen.Text & "','" & Combocitizen.Text & "','" & (Textpassport.Text) & "','" & Textcategory.Text & "','" & _
    Textdatarec.Text & "','" & Textlprovider.Text & "','" & Texttempadd.Text & "')"
            'string stores the command and CInt is used to convert number to string
            cmd = New OleDbCommand(str, cn)
            Dim temp As Integer = cmd.ExecuteNonQuery

            'displays number of records inserted
            If temp > 0 Then
                MsgBox("Data Are Added Successfully")
                clear()
                jdcomplex()
            Else
                MsgBox("Try Again")

            End If

        Catch ex As Exception
        End Try

        cn.Close()
        Label21.Text = (Me.BindingContext(ds, "ms").Position + _
        1).ToString & " of " & Me.BindingContext(ds, _
        "ms").Count.ToString
        check()

    End Sub
   



    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Me.Close()
        Form8.Show()

    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub





    Private Sub MaskedTextBox2_MaskInputRejected(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MaskInputRejectedEventArgs)

    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox18_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Textcardno_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Textcardno.MouseClick

    End Sub

    Private Sub Textcardno_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Textcardno.TextChanged

    End Sub

    Private Sub Button10_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        clear()
        jdcomplex()
        Label21.Text = (Me.BindingContext(ds, "ms").Position + _
      1).ToString & " of " & Me.BindingContext(ds, _
      "ms").Count.ToString
        check()
        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
        Button6.Enabled = True
        Button7.Enabled = True
        Button8.Enabled = True
        CheckedListBox1.Visible = False
        Textcategory.Location = New System.Drawing.Point(634, 477)
        Textcategory.Size = New System.Drawing.Point(137, 75)

    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button8.Enabled = False
        clear()
        textclear()
        Try

            cn = New OleDbConnection("Provider=microsoft.jet.oledb.4.0; data source='" & STrPath & "';")
            cmd = New OleDbCommand("select * from dlms ")
            cn.Open()

            cmd.Connection = cn
            If ComboBox1.SelectedItem = "License No" Then
                da = New OleDb.OleDbDataAdapter("SELECT * FROM [dlms] Where [License No] ='" & (TextBox1.Text) & "' ", cn)
            ElseIf ComboBox1.SelectedItem = "Citizenship No" Then
                da = New OleDb.OleDbDataAdapter("SELECT * FROM [dlms] Where [Citizenship No] ='" & (TextBox1.Text) & "' ", cn)
            Else
                MsgBox("select search item")
                ComboBox1.Focus()
                Exit Sub
            End If
            ds = New DataSet("ms")
            da.Fill(ds, "ms")
            BindingSource1.DataSource = ds
            BindingSource1.DataMember = "ms"

            kamal()
            If Textlicense.Text = "" Then
                clear()
                Label21.Text = "No Record"
            Else
                Label21.Text = "Record Found"
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "")

        End Try
        cn.Close()
        check()
    End Sub

   

    Private Sub Textcategory_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Textcategory.TextChanged

    End Sub
    Sub check()
        Dim expiredate As Date
        expiredate = (dexpire.Text)
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
            If a >= 0 Or b >= 0 Or c >= 0 Then
               
                status.Text = "Renew In Progress"
            Else
                status.Text = "Expired"
            End If
        Else
            status.Text = "Valid"

        End If
    End Sub

    Private Sub Textwardno_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Textwardno.TextChanged

    End Sub

    Private Sub CheckedListBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBox1.LostFocus
        Dim strdata As String

        For Each strdata In CheckedListBox1.CheckedItems
            Textcategory.Text &= strdata & " , "
        Next
    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckedListBox1.SelectedIndexChanged

    End Sub
End Class