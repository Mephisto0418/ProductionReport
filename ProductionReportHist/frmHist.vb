Imports System.Data.SqlClient

Public Class frmHist
    Private intQueryMode As Integer = 1
    '1204 新增參數
    Dim DbProcParameter As String = "[H3_Systematic].[dbo].[H3_Production_ProcParameter]" '欄位設定Config DB
    Dim DbProcParameterRule As String = "[H3_Systematic].[dbo].[H3_Production_ProcParameter_Rule]" '欄位帶入設定Config DB
    Dim DbLog As String = "[H3_Systematic].[dbo].[H3_ProductionLog]" '固定欄位資料紀錄 DB
    Dim DbLogParameter As String = "[H3_Systematic].[dbo].[H3_ProductionParameter]" '客製化欄位資料紀錄 DB
    Dim DbMachine As String = "[UTCHFACMRPT_REAL].[acme].[dbo].[PDL_Machine]" '愉進機台資料表
    Dim DbRemark As String = "[H3_Systematic].[dbo].[ProductionReport_Remark]"
    Dim DbRemark_Type As String = "[H3_Systematic].[dbo].[ProductionReport_Remark_Type]"

    Dim DbVersion As String = "[Datamation_H3].[dbo].[H3_Leo_Program_Version]" '版本卡控DB
    Dim DbProc As String = "[H3_Systematic].[dbo].[H3_Proc]" '報表設定Config DB
    Dim DbHist As String = "[UTCHFACMRPT_REAL].[report].[dbo].[view_Hist]"
    Dim DbPartInfo As String = "[UTCHFACMRPT_REAL].[acme].[dbo].[PartInfo_ACME]"


    Private Sub frmHist_Load(sender As Object, e As EventArgs) Handles Me.Load
        dtpStartDateTime.MaxDate = Now
        dtpEndDateTime.MaxDate = Now
        Dim start_date As DateTime
        start_date = CType(Now, Date).AddDays(-1)
        dtpStartDateTime.Value = start_date.Year & "-" & start_date.Month & "-" & start_date.Day & " 07:20:00"
        dtpEndDateTime.Value = Now.ToString("yyyy-MM-dd HH:mm:ss")

        dtpEndDateTime.MinDate = dtpStartDateTime.Value

        dgvResult.ReadOnly = True
        Dim strQuerySection As String = "SELECT DISTINCT [Section]
                                         FROM " & DbProc & " WITH(NOLOCK)"

        Dim dtStation As New DataTable
        dtStation = SQL_Select(strQuerySection, sqlconnMQL03)
        clbSection.Items.Add("ALL")
        For i As Integer = 0 To dtStation.Rows.Count - 1
            clbSection.Items.Add(dtStation.Rows(i).Item(0))
        Next

    End Sub

    Private Sub btnQueryByLotNum_Click(sender As Object, e As EventArgs) Handles btnQueryByLotNum.Click
        If btnQueryByLotNum.BackColor <> Color.Blue Then
            txtLotNo.Visible = True
            Label4.Visible = True
            clbLayer.Visible = True
            btnClear.Visible = True
            txtLotNumAll.Visible = True
            Label6.Visible = True
            Label1.Visible = False
            Label2.Visible = False
            dtpEndDateTime.Visible = False
            dtpStartDateTime.Visible = False
            btnQueryByTime.BackColor = SystemColors.Control
            btnQueryByTime.ForeColor = Color.Black
            btnQueryByLotNum.BackColor = Color.Blue
            btnQueryByLotNum.ForeColor = Color.White
        End If
    End Sub

    Private Sub btnQueryByTime_Click(sender As Object, e As EventArgs) Handles btnQueryByTime.Click
        If btnQueryByTime.BackColor <> Color.Blue Then
            txtLotNo.Visible = False
            Label4.Visible = False
            clbLayer.Visible = False
            btnClear.Visible = False
            txtLotNumAll.Visible = False
            Label6.Visible = False
            Label1.Visible = True
            Label2.Visible = True
            dtpEndDateTime.Visible = True
            dtpStartDateTime.Visible = True
            btnQueryByLotNum.BackColor = SystemColors.Control
            btnQueryByLotNum.ForeColor = Color.Black
            btnQueryByTime.BackColor = Color.Blue
            btnQueryByTime.ForeColor = Color.White
        End If
    End Sub

    Private Sub dtpStartDateTime_ValueChanged(sender As System.Object, e As System.EventArgs) Handles dtpStartDateTime.ValueChanged
        dtpEndDateTime.MinDate = dtpStartDateTime.Value
    End Sub

    Private Sub dtpEndDateTime_ValueChanged(sender As System.Object, e As System.EventArgs) Handles dtpEndDateTime.ValueChanged
        dtpEndDateTime.MinDate = dtpStartDateTime.Value
    End Sub

    Private Sub clbSection_SelectedIndexChanged(sender As Object, e As EventArgs) Handles clbSection.SelectedIndexChanged
        clbStation.Items.Clear()
        If clbSection.SelectedItem.ToString = "ALL" And clbSection.GetItemCheckState(0) = CheckState.Checked Then
            For i As Integer = 1 To clbSection.Items.Count - 1
                If i <> clbSection.SelectedIndex Then
                    clbSection.SetItemChecked(i, System.Windows.Forms.CheckState.Checked)
                End If
            Next

            ' Uncheck "ALL" will uncheck all items
        ElseIf clbSection.SelectedItem.ToString = "ALL" And clbSection.GetItemCheckState(0) = CheckState.Unchecked Then
            For i As Integer = 1 To clbSection.Items.Count - 1
                If i <> clbSection.SelectedIndex Then
                    clbSection.SetItemChecked(i, System.Windows.Forms.CheckState.Unchecked)
                End If
            Next
        End If

        ' If one of the items in checkedlistbox is unchecked, uncheck "ALL"
        If clbSection.CheckedItems.Count < clbSection.Items.Count Then
            clbSection.SetItemChecked(0, System.Windows.Forms.CheckState.Unchecked)
        End If

        If clbSection.CheckedItems.Count > 0 Then
            Dim strQueryStation As String = "SELECT DISTINCT [ProcName]
                                             FROM " & DbProc & " WITH(NOLOCK)
                                             WHERE [Section] IN ("

            For i As Integer = 1 To clbSection.Items.Count - 1
                If clbSection.GetItemCheckState(i) = CheckState.Checked Then
                    strQueryStation = strQueryStation & "'" & clbSection.Items(i).ToString & "',"
                End If
            Next
            strQueryStation = Mid(strQueryStation, 1, strQueryStation.Length - 1)
            strQueryStation = strQueryStation & ")"

            Dim dtStation As New DataTable
            dtStation = SQL_Select(strQueryStation, sqlconnMQL03)
            clbStation.Items.Add("ALL")
            For i As Integer = 0 To dtStation.Rows.Count - 1
                clbStation.Items.Add(dtStation.Rows(i).Item(0))
            Next

        End If
    End Sub

    Private Sub clbStation_SelectedIndexChanged(sender As Object, e As EventArgs) Handles clbStation.SelectedIndexChanged
        If clbStation.SelectedItem.ToString = "ALL" And clbStation.GetItemCheckState(0) = CheckState.Checked Then
            For i As Integer = 1 To clbStation.Items.Count - 1
                If i <> clbStation.SelectedIndex Then
                    clbStation.SetItemChecked(i, System.Windows.Forms.CheckState.Checked)
                End If
            Next

            ' Uncheck "ALL" will uncheck all items
        ElseIf clbStation.SelectedItem.ToString = "ALL" And clbStation.GetItemCheckState(0) = CheckState.Unchecked Then
            For i As Integer = 1 To clbStation.Items.Count - 1
                If i <> clbStation.SelectedIndex Then
                    clbStation.SetItemChecked(i, System.Windows.Forms.CheckState.Unchecked)
                End If
            Next
        End If

        ' If one of the items in checkedlistbox is unchecked, uncheck "ALL"
        If clbStation.CheckedItems.Count < clbStation.Items.Count Then
            clbStation.SetItemChecked(0, System.Windows.Forms.CheckState.Unchecked)
        End If
    End Sub

    Private Sub txtLotNo_KeyUp(sender As Object, e As KeyEventArgs) Handles txtLotNo.KeyUp
        If e.KeyData = Keys.Enter Then
            If txtLotNo.Text.Trim() <> "" Then
                Dim strQueryPartNum As String = "SELECT TOP 1 RTRIM([partnum]) + RTRIM([revision]) 
                                             FROM  " & DbHist & "  WHERE [lotnum] = @LotNum"
                Dim dtPartNum As New DataTable
                Dim sqldaQueryPartNum As New SqlDataAdapter
                sqldaQueryPartNum.SelectCommand = sqlconnMQL03.CreateCommand
                sqldaQueryPartNum.SelectCommand.CommandText = strQueryPartNum
                sqldaQueryPartNum.SelectCommand.Parameters.AddWithValue("@LotNum", Mid(txtLotNo.Text, 1, 14))
                If sqlconnMQL03.State = ConnectionState.Closed Then sqlconnMQL03.Open()
                sqldaQueryPartNum.Fill(dtPartNum)
                If sqlconnMQL03.State = ConnectionState.Open Then sqlconnMQL03.Close()
                If dtPartNum.Rows.Count = 0 Then
                    MsgBox("批號查無資訊")
                    Exit Sub
                End If


                Dim strPartNum As String = dtPartNum.Rows(0).Item(0).ToString().Trim()

                Dim strQueryLayerName As String = "SELECT CASE WHEN [PressCount] = 0 THEN 'Core' ELSE 'BU' + CONVERT(nvarchar,[PressCount]) END AS [BU]
                                               FROM " & DbPartInfo & "
                                               WHERE RTRIM([PartNum]) + RTRIM([Revision]) = '" & strPartNum & "'
                                               ORDER BY [BU]"
                Dim dtLayer As New DataTable
                dtLayer = SQL_Select(strQueryLayerName, sqlconnMQL03)

                If clbLayer.Items.Count = 0 Then
                    clbLayer.Items.Add("ALL")
                End If
                For i As Integer = 0 To dtLayer.Rows.Count - 1
                    Dim booAdd As Boolean = True
                    For j As Integer = 0 To clbLayer.Items.Count - 1
                        If clbLayer.Items(j).ToString().Trim() = dtLayer.Rows(i).Item(0).ToString().Trim() Then
                            booAdd = False
                            Exit For
                        End If
                    Next
                    If booAdd = True Then
                        clbLayer.Items.Add(dtLayer.Rows(i).Item(0).ToString().Trim())
                    End If
                Next
                txtLotNumAll.Text = txtLotNumAll.Text & Mid(txtLotNo.Text, 1, 14) & vbCrLf
                txtLotNo.Text = ""
            End If
        End If
    End Sub

    Private Sub clbLayer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles clbLayer.SelectedIndexChanged
        If clbLayer.SelectedItem.ToString = "ALL" And clbLayer.GetItemCheckState(0) = CheckState.Checked Then
            For i As Integer = 1 To clbLayer.Items.Count - 1
                If i <> clbLayer.SelectedIndex Then
                    clbLayer.SetItemChecked(i, System.Windows.Forms.CheckState.Checked)
                End If
            Next

            ' Uncheck "ALL" will uncheck all items
        ElseIf clbLayer.SelectedItem.ToString = "ALL" And clbLayer.GetItemCheckState(0) = CheckState.Unchecked Then
            For i As Integer = 1 To clbLayer.Items.Count - 1
                If i <> clbLayer.SelectedIndex Then
                    clbLayer.SetItemChecked(i, System.Windows.Forms.CheckState.Unchecked)
                End If
            Next
        End If
    End Sub

    Private Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        Dim strQuery As String = "EXEC [H3_ProductuinReport_Hist] "
        Dim dtShow As New DataTable
        If btnQueryByLotNum.BackColor = Color.Blue Then
            If txtLotNumAll.Text.Trim() = "" Then
                MsgBox("請輸入批號")
                Exit Sub
            End If

            If clbStation.SelectedItems.Count = 0 Then
                MsgBox("請選擇站別")
                Exit Sub
            End If


            strQuery = strQuery & "1, @LotNum, @EQID, '" & dtpStartDateTime.Value.ToString("yyyy-MM-dd HH:mm:ss") & "', '" & dtpEndDateTime.Value.ToString("yyyy-MM-dd HH:mm:ss") & "'"
            If sqlconnMQL03.State = ConnectionState.Closed Then sqlconnMQL03.Open()

            Dim sqlcmdQueryByLot As New SqlCommand(strQuery, sqlconnMQL03)
            Dim arrstrLotNum As String() = txtLotNumAll.Text.Trim().Split(vbCrLf)
            Dim strQueryLotNums As String = ""

            For i As Integer = 0 To arrstrLotNum.Length - 1
                If clbStation.GetItemCheckState(i) = CheckState.Checked Then
                    strQueryLotNums = strQueryLotNums & "'" & arrstrLotNum(i).Trim() & "',"
                End If
            Next
            strQueryLotNums = Mid(strQueryLotNums, 1, strQueryLotNums.Length - 1)

            Dim sqlparmLot As SqlParameter = sqlcmdQueryByLot.Parameters.AddWithValue("@LotNum", strQueryLotNums)
            Dim strEQID As String = ""
            For i As Integer = 1 To clbStation.Items.Count - 1
                If clbStation.GetItemCheckState(i) = CheckState.Checked Then
                    strEQID = strEQID & "'" & clbStation.Items(i).ToString & "',"
                End If
            Next
            strEQID = Mid(strEQID, 1, strEQID.Length - 1)
            Dim sqlparmEQID As SqlParameter = sqlcmdQueryByLot.Parameters.AddWithValue("@EQID", strEQID)
            Dim sqlrdQueryByLot As SqlDataReader = sqlcmdQueryByLot.ExecuteReader()

            dtShow.Load(sqlrdQueryByLot)
            dtShow.Columns.Remove("BU")
            dtShow.Columns.Remove("PID")
            dgvResult.DataSource = dtShow


            If sqlconnMQL03.State = ConnectionState.Open Then sqlconnMQL03.Close()
        ElseIf btnQueryByTime.BackColor = Color.Blue Then
            If clbStation.SelectedItems.Count = 0 Then
                MsgBox("請選擇站別")
                Exit Sub
            End If
            strQuery = strQuery & "2, '', @EQID, '" & dtpStartDateTime.Value.ToString("yyyy-MM-dd HH:mm:ss") & "', '" & dtpEndDateTime.Value.ToString("yyyy-MM-dd HH:mm:ss") & "'"
            If sqlconnMQL03.State = ConnectionState.Closed Then sqlconnMQL03.Open()
            Dim strEQID As String = ""
            For i As Integer = 1 To clbStation.Items.Count - 1
                If clbStation.GetItemCheckState(i) = CheckState.Checked Then
                    strEQID = strEQID & "'" & clbStation.Items(i).ToString & "',"
                End If
            Next
            strEQID = Mid(strEQID, 1, strEQID.Length - 1)
            Dim sqlcmdQueryByTime As New SqlCommand(strQuery, sqlconnMQL03)
            Dim sqlparmEQID As SqlParameter = sqlcmdQueryByTime.Parameters.AddWithValue("@EQID", strEQID)
            Dim sqlrdQueryByTime As SqlDataReader = sqlcmdQueryByTime.ExecuteReader()

            dtShow.Load(sqlrdQueryByTime)
            dtShow.Columns.Remove("BU")
            dtShow.Columns.Remove("PID")
            dgvResult.DataSource = dtShow
        End If

        For i As Integer = 0 To dtShow.Rows.Count - 1
            If dtShow.Rows(i).Item("備註").ToString().Trim() <> "" Then
                dgvResult.Rows(i).Cells("備註").Style.BackColor = Color.Wheat
            End If
        Next

        dgvResult.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        txtLotNo.Text = ""
        txtLotNumAll.Text = ""
        clbLayer.Items.Clear()
    End Sub

    Private Sub btnExportToExcel_Click(sender As Object, e As EventArgs) Handles btnExportToExcel.Click

    End Sub
End Class
