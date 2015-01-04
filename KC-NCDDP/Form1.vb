Imports System.Data.OleDb
Public Class main_form
    Dim cnn As New OleDb.OleDbConnection
    Dim cmd As New OleDb.OleDbCommand
    Dim adapter As New OleDb.OleDbDataAdapter
    Dim dataTbl As New DataTable

    Dim curRecId As Integer
    Dim curRecName As String
    Dim Sql As String

    Dim fourPsStat As String = ""
    Dim fourPsSpec As String = ""
    Dim slpStat As String = ""
    Dim familyHead As String = ""

    Dim valid As Boolean = False

    Private Sub main_form_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'KalahiDataSet.resident' table. You can move, or remove it, as needed.
        Me.ResidentTableAdapter.Fill(Me.KalahiDataSet.resident)
        conn()
        Dim da As New OleDb.OleDbCommand("SELECT * FROM resident", cnn)

        curRecName = ""
        ck4psNo.Checked = True
        ckSlpNo.Checked = True
        txt4ps.Enabled = False

        showStatistics()
    End Sub

    Private Sub conn()
        'TODO: This line of code loads data into the 'KalahiDataSet1.resident' table. You can move, or remove it, as needed.
        'Me.ResidentTableAdapter.Fill(Me.KalahiDataSet1.resident)
        cnn = New OleDb.OleDbConnection
        cnn.ConnectionString = "Provider=Microsoft.Jet.oledb.4.0; Data source = " & Application.StartupPath & "\kalahi.mdb"

    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        If txtSearch.Text = "" Then
            MsgBox("Please enter name to search!")
        Else
            displayRecord()
        End If
    End Sub

    Private Sub displayRecord()
        conn()
        cnn.Close()
        dataTbl.Clear()

        ' Split string with spaces
        If txtSearch.Text.Contains(" "c) Then
            Dim words As String() = txtSearch.Text.Split(New Char() {" "c})
            Dim w1 As String = words(0)
            Dim w2 As String = words(1)
            ' Define the SQL to grab data from table.
            Sql = "SELECT * " & _
                "FROM(resident)" & _
                "WHERE ( first_name LIKE '%" & w1 & "%' AND last_name LIKE '%" & w2 & "%' )" & _
                "OR ( last_name LIKE '%" & w2 & "%' AND first_name LIKE '%" & w1 & "%' ) "
        Else
            Sql = "SELECT * " & _
                "FROM(resident)" & _
                "WHERE first_name LIKE '%" & txtSearch.Text & "%' OR last_name LIKE '%" & txtSearch.Text & "%' "
        End If
        ' Try, Catch, Finally
        Try
            cnn.Open()
            cmd.Connection = cnn
            cmd.CommandText = Sql
            adapter.SelectCommand = cmd
            adapter.Fill(dataTbl)

            If dataTbl.Rows.Count > 0 Then

                clearAllFields()

                curRecId = dataTbl.Rows(0).Item("res_id")
                curRecName = dataTbl.Rows(0).Item("first_name") & " " & dataTbl.Rows(0).Item("last_name")

                lblNotif.Text = curRecName
                txtLastName.Text = dataTbl.Rows(0).Item("last_name")
                txtMidName.Text = dataTbl.Rows(0).Item("middle_name")
                txtFirstName.Text = dataTbl.Rows(0).Item("first_name")
                dtBday.Text = dataTbl.Rows(0).Item("birthday")
                cbGender.Text = dataTbl.Rows(0).Item("gender")
                txtWork.Text = dataTbl.Rows(0).Item("res_work")
                txtPurok.Text = dataTbl.Rows(0).Item("purok")
                cbBgy.Text = dataTbl.Rows(0).Item("barangay")
                If dataTbl.Rows(0).Item("4ps") = "Yes" Then
                    ck4psYes.Checked = True
                    txt4ps.Text = dataTbl.Rows(0).Item("4ps_specified")
                Else
                    ck4psNo.Checked = True
                End If
                If dataTbl.Rows(0).Item("slp") = "Yes" Then
                    ckSlpYes.Checked = True
                Else
                    ckSlpNo.Checked = True
                End If
                cbRemarks.Text = dataTbl.Rows(0).Item("remarks")
                If dataTbl.Rows(0).Item("family_head") Is "Yes" Then
                    ckFamilyHead.Checked = False
                End If

                disableFields()
            Else
                MessageBox.Show("Record not found!")
            End If
        Catch myerror As OleDbException
            MessageBox.Show("Cannot connect to database: " & myerror.Message)
        Finally
            cnn.Close()
            cnn.Dispose()
        End Try
    End Sub

    Private Sub disableFields()
        txtFirstName.Enabled = False
        txtLastName.Enabled = False
        txtMidName.Enabled = False
        dtBday.Enabled = False
        cbGender.Enabled = False
        txtWork.Enabled = False
        txtPurok.Enabled = False
        cbBgy.Enabled = False
        ck4psYes.Enabled = False
        ck4psNo.Enabled = False
        txt4ps.Enabled = False
        ckSlpYes.Enabled = False
        ckSlpNo.Enabled = False
        cbRemarks.Enabled = False
        ckFamilyHead.Enabled = False

        btnUpdate.Hide()
        btnSubmit.Hide()
        btnCancel.Hide()
    End Sub

    Private Sub enableFields()
        txtFirstName.Enabled = True
        txtLastName.Enabled = True
        txtMidName.Enabled = True
        dtBday.Enabled = True
        cbGender.Enabled = True
        txtWork.Enabled = True
        txtPurok.Enabled = True
        cbBgy.Enabled = True
        ck4psYes.Enabled = True
        ck4psNo.Enabled = True
        txt4ps.Enabled = True
        ckSlpYes.Enabled = True
        ckSlpNo.Enabled = True
        cbRemarks.Enabled = True
        ckFamilyHead.Enabled = True

        btnUpdate.Show()
        btnSubmit.Show()
        btnCancel.Show()
    End Sub

    Private Sub clearAllFields()
        btnSearch.Text = ""
        txtFirstName.Text = ""
        txtLastName.Text = ""
        txtMidName.Text = ""
        dtBday.Text = ""
        cbGender.Text = ""
        txtWork.Text = ""
        txtPurok.Text = ""
        cbBgy.Text = ""
        ck4psYes.Checked = False
        ck4psNo.Checked = True
        txt4ps.Text = ""
        ckSlpYes.Checked = False
        ckSlpNo.Checked = True
        cbRemarks.Text = ""
        ckFamilyHead.Checked = False
        lblNotif.Text = ""
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        clearAllFields()
        enableFields()

        btnUpdate.Hide()
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        If curRecName.Equals("") Then
            MsgBox("Please specify record to modify.")
        Else
            enableFields()
            btnSubmit.Hide()
        End If
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim mb As String

        If curRecName.Equals("") Then
            MsgBox("Please specify record to delete.")
        Else
            mb = MsgBox("Are you sure you want to delete " & curRecName & "s Record?", MsgBoxStyle.YesNo).ToString()

            If mb.Equals("Yes") Then
                Sql = "DELETE FROM resident WHERE res_id=" & curRecId
                addUpdateDeleteData(Sql, "Successfully Deleted Record")
                clearAllFields()
                enableFields()
                btnUpdate.Hide()
            End If
        End If
    End Sub

    Private Sub btnSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        validateData()
        If valid = True Then
            Sql = "INSERT INTO resident " & _
                "(last_name, middle_name, first_name, birthday, gender, res_work, purok, barangay, 4ps, 4ps_specified, slp, remarks, family_head )" & _
                "VALUES('" & txtLastName.Text & "', '" & txtMidName.Text & "', '" & txtFirstName.Text & "'," & _
                "'" & dtBday.Text & "', '" & cbGender.Text & "',  '" & txtWork.Text & "', '" & txtPurok.Text & "'," & _
            "'" & cbBgy.Text & "', '" & fourPsStat & "', '" & fourPsSpec & "', '" & slpStat & "'," & _
            "'" & cbRemarks.Text & "', '" & familyHead & "')"

            addUpdateDeleteData(Sql, "Successfully Added Record")
            valid = False

            clearAllFields()
            enableFields()
            btnUpdate.Hide()
        End If
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        validateData()
        If valid = True Then
            Sql = "UPDATE resident " & _
                "SET last_name='" & txtLastName.Text & "', middle_name='" & txtMidName.Text & "', first_name='" & txtFirstName.Text & "'," & _
                "birthday='" & dtBday.Text & "', gender='" & cbGender.Text & "', res_work='" & txtWork.Text & "', purok='" & txtPurok.Text & "'," & _
                "barangay='" & cbBgy.Text & "', 4ps='" & fourPsStat & "', 4ps_specified='" & fourPsSpec & "', slp='" & slpStat & "'," & _
                "remarks='" & cbRemarks.Text & "', family_head='" & familyHead & "' WHERE res_id = " & curRecId

            addUpdateDeleteData(Sql, "Successfully Modified Record!")
            valid = False
            displayRecord()
        End If
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        clearAllFields()
        displayRecord()
    End Sub
    Private Function validateData()
        If cbBgy.Text = "" Then
            MessageBox.Show("Barangay is empty! Please fill in.")
        ElseIf txtPurok.Text = "" Then
            MessageBox.Show("Purok is empty! Please fill in.")
        ElseIf txtFirstName.Text = "" Then
            MessageBox.Show("First Name is empty! Please fill in.")
        ElseIf txtMidName.Text = "" Then
            MessageBox.Show("Middle Name is empty! Please fill in.")
        ElseIf txtLastName.Text = "" Then
            MessageBox.Show("Last Name is empty! Please fill in.")
        ElseIf cbRemarks.Text = "" Then
            MessageBox.Show("Remarks is empty! Please fill in.")
        ElseIf cbGender.Text = "" Then
            MessageBox.Show("Gender is empty! Please fill in.")
        ElseIf txtWork.Text = "" Then
            MessageBox.Show("Work is empty! Please fill in.")
        ElseIf ck4psYes.Checked And txt4ps.Text = "" Then
            MessageBox.Show("Please specify 4Ps involvement.")
        Else
            ' 4P's
            If ck4psYes.Checked Then
                fourPsStat = "Yes"
                fourPsSpec = txt4ps.Text
            Else
                fourPsStat = "No"
                fourPsSpec = "N/A"
            End If
            ' SLP
            If ckSlpYes.Checked Then
                slpStat = "Yes"
            Else
                slpStat = "No"
            End If
            ' Family Head
            If ckFamilyHead.Checked Then
                familyHead = "Yes"
            Else
                familyHead = "No"
            End If

            valid = True
        End If
        Return valid
    End Function

    Private Function addUpdateDeleteData(ByVal Sql, ByVal msg)
        conn()
        cnn.Close()
        dataTbl.Clear()

        Try
            cnn.Open()
            cmd.Connection = cnn
            cmd.CommandText = Sql
            adapter.SelectCommand = cmd
            adapter.Fill(dataTbl)

            MessageBox.Show(msg)
            showStatistics()

        Catch myerror As OleDbException
            MessageBox.Show("An error occured. Contact Dina Fajardo to fix this..." & myerror.Message)
        Finally
            cnn.Close()
            cnn.Dispose()
        End Try
        Return True
    End Function

    Private Function retrieveData(ByVal Sql)
        Dim val As String
        val = ""
        conn()
        dataTbl.Clear()

        Try
            cnn.Open()
            cmd.Connection = cnn
            cmd.CommandText = Sql
            adapter.SelectCommand = cmd
            adapter.Fill(dataTbl)

            val = dataTbl.Rows(0).Item(0)

        Catch myerror As OleDbException
            MessageBox.Show("Cannot connect to database: " & myerror.Message)
        Finally
            cnn.Close()
            cnn.Dispose()
        End Try

        Return val
    End Function

    Private Sub showStatistics()
        lblNoHouseholds1.Text = "No. of Households: " & retrieveData("SELECT count(*) FROM resident WHERE remarks = 'Household Head' AND barangay='Chico Island'")
        lblNoHouseholds2.Text = "No. of Households: " & retrieveData("SELECT count(*) FROM resident WHERE remarks = 'Household Head' AND barangay='Divisoria'")
        lblNoHouseholds3.Text = "No. of Households: " & retrieveData("SELECT count(*) FROM resident WHERE remarks = 'Household Head' AND barangay='San Vicente'")

        lblNoFamilies1.Text = "No. of Families: " & retrieveData("SELECT count(*) FROM resident WHERE family_head = 'Yes' AND barangay='Chico Island'")
        lblNoFamilies2.Text = "No. of Families: " & retrieveData("SELECT count(*) FROM resident WHERE family_head = 'Yes' AND barangay='Divisoria'")
        lblNoFamilies3.Text = "No. of Families: " & retrieveData("SELECT count(*) FROM resident WHERE family_head = 'Yes' AND barangay='San Vicente'")

        lblNo4ps1.Text = "4P's: " & retrieveData("SELECT count(*) FROM resident WHERE [4ps] = 'Yes' AND barangay='Chico Island'")
        lblNo4ps2.Text = "4P's: " & retrieveData("SELECT count(*) FROM resident WHERE [4ps] = 'Yes' AND barangay='Divisoria'")
        lblNo4ps3.Text = "4P's: " & retrieveData("SELECT count(*) FROM resident WHERE [4ps] = 'Yes' AND barangay='San Vicente'")

        lblNoSlp1.Text = "SLP: " & retrieveData("SELECT count(*) FROM resident WHERE slp = 'Yes' AND barangay='Chico Island'")
        lblNoSlp2.Text = "SLP: " & retrieveData("SELECT count(*) FROM resident WHERE slp = 'Yes' AND barangay='Divisoria'")
        lblNoSlp3.Text = "SLP: " & retrieveData("SELECT count(*) FROM resident WHERE slp = 'Yes' AND barangay='San Vicente'")

        lblNoMale1.Text = "Male: " & retrieveData("SELECT count(*) FROM resident WHERE gender = 'Male' AND barangay='Chico Island'")
        lblNoMale2.Text = "Male: " & retrieveData("SELECT count(*) FROM resident WHERE gender = 'Male' AND barangay='Divisoria'")
        lblNoMale3.Text = "Male: " & retrieveData("SELECT count(*) FROM resident WHERE gender = 'Male' AND barangay='San Vicente'")

        lblNoFemale1.Text = "Female: " & retrieveData("SELECT count(*) FROM resident WHERE gender = 'Female' AND barangay='Chico Island'")
        lblNoFemale2.Text = "Female: " & retrieveData("SELECT count(*) FROM resident WHERE gender = 'Female' AND barangay='Divisoria'")
        lblNoFemale3.Text = "Female: " & retrieveData("SELECT count(*) FROM resident WHERE gender = 'Female' AND barangay='San Vicente'")

    End Sub

    Private Sub ck4psYes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ck4psYes.CheckedChanged
        If ck4psYes.Checked Then
            txt4ps.Enabled = True
        End If
    End Sub

    Private Sub ck4psNo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ck4psNo.CheckedChanged
        If ck4psNo.Checked Then
            txt4ps.Text = ""
            txt4ps.Enabled = False
        End If
    End Sub

    Private Sub dtBday_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtBday.ValueChanged
        Dim bdayYear As String() = dtBday.Text.Split(New Char() {" "c})
        Dim by As String = bdayYear(3)

        If Not dtBday.Text = "" And IsNumeric(by) Then
            txtAge.Text = Date.Today.Year - by
        End If
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'PrintDocument1.Print()
    End Sub

    Private Sub FillBy_San_Vicente_ToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Me.ResidentTableAdapter.FillBy_San_Vicente_(Me.KalahiDataSet.resident)
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub FillBy_San_Vicente_ToolStripButton_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Me.ResidentTableAdapter.FillBy_San_Vicente_(Me.KalahiDataSet.resident)
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub cbDispRec_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbDispRec.SelectedIndexChanged
        If cbDispRec.Text = "Chico Island" Then
            Try
                Me.ResidentTableAdapter.FillBy_Chico_Island(Me.KalahiDataSet.resident)
            Catch ex As System.Exception
                System.Windows.Forms.MessageBox.Show(ex.Message)
            End Try
        ElseIf cbDispRec.Text = "Divisoria" Then
            Try
                Me.ResidentTableAdapter.FillBy3(Me.KalahiDataSet.resident)
            Catch ex As System.Exception
                System.Windows.Forms.MessageBox.Show(ex.Message)
            End Try
        ElseIf cbDispRec.Text = "San Vicente" Then
            Try
                Me.ResidentTableAdapter.FillBy_San_Vicente_(Me.KalahiDataSet.resident)
            Catch ex As System.Exception
                System.Windows.Forms.MessageBox.Show(ex.Message)
            End Try
        ElseIf cbDispRec.Text = "ALL" Then
            Try
                Me.ResidentTableAdapter.Fill(Me.KalahiDataSet.resident)
            Catch ex As System.Exception
                System.Windows.Forms.MessageBox.Show(ex.Message)
            End Try
        ElseIf cbDispRec.Text = "4P's - Chico Island" Then
            Try
                Me.ResidentTableAdapter.Fill4ps_Chico_Island(Me.KalahiDataSet.resident)
            Catch ex As System.Exception
                System.Windows.Forms.MessageBox.Show(ex.Message)
            End Try
        ElseIf cbDispRec.Text = "4P's - Divisoria" Then
            Try
                Me.ResidentTableAdapter.Fill4ps_Divisoria(Me.KalahiDataSet.resident)
            Catch ex As System.Exception
                System.Windows.Forms.MessageBox.Show(ex.Message)
            End Try
        ElseIf cbDispRec.Text = "4P's - San Vicente" Then
            Try
                Me.ResidentTableAdapter.Fill4ps_San_Vicente(Me.KalahiDataSet.resident)
            Catch ex As System.Exception
                System.Windows.Forms.MessageBox.Show(ex.Message)
            End Try
        ElseIf cbDispRec.Text = "SLP - Chico Island" Then
            Try
                Me.ResidentTableAdapter.FillSlp_Chico_Island(Me.KalahiDataSet.resident)
            Catch ex As System.Exception
                System.Windows.Forms.MessageBox.Show(ex.Message)
            End Try
        ElseIf cbDispRec.Text = "SLP - Divisoria" Then
            Try
                Me.ResidentTableAdapter.FillBy4(Me.KalahiDataSet.resident)
            Catch ex As System.Exception
                System.Windows.Forms.MessageBox.Show(ex.Message)
            End Try
        ElseIf cbDispRec.Text = "SLP - San Vicente" Then
            Try
                Me.ResidentTableAdapter.FillSlp_San_Vicente(Me.KalahiDataSet.resident)
            Catch ex As System.Exception
                System.Windows.Forms.MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub


    Private Sub btnPrint_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        PrintPreviewDialog1.Document = PrintDocument1
        PrintPreviewDialog1.ShowDialog()
    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim bm As New Bitmap(Me.dgvRecord.Width, Me.dgvRecord.Height)
        dgvRecord.DrawToBitmap(bm, New Rectangle(0, 0, dgvRecord.Width, Me.dgvRecord.Height))
        e.Graphics.DrawImage(bm, 0, 0)
    End Sub

    Private Sub Fill4ps_Chico_IslandToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub
End Class
