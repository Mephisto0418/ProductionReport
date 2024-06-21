Imports System.IO
Imports Microsoft.VisualBasic.FileIO



Public Class ProductionReport_Setting

    ' 檔案名稱 : 生產報表欄位設定介面
    ' 檔案簡短描述 : 
    '
    ' 檔案詳細描述, 內容包含哪些函式或類別, 實作了哪些功能, 
    '
    ' 作者名稱 : Boris_Li@unimicron.com
    ' 更新日期 : 2023/07/25
    Dim Version As String = "1.0.24.05.30.1"
    Dim Program As String = "ProductionReport_Setting"
    Dim isException As Boolean = False
    Dim ProcPkey As String = ""
    Dim ProcPassword As New Dictionary(Of String, String)
    Dim ParaPkey As String = ""
    Dim ParaSort As Integer = 0
    Dim QID As String
    Dim QueryCommand As String = ""
    Dim ServerCommand As New Dictionary(Of String, String)
    Dim SelectCommand As New Dictionary(Of String, String)
    Dim DefaulfFilter As New Dictionary(Of String, String)
    Dim SPCGroup As New Dictionary(Of String, String)
    Dim isError As Boolean = False
    Dim ParaColumn As String
    Dim F_SaveType As String
    Dim AreaID As String = ""
    '2024/05/29新增
    Dim F_IPQC_Proc As New Dictionary(Of String, String)


    '-----------------------------------DB參數----------------------------------------
    Dim DbProc As String = "[H3_Systematic].[dbo].[H3_Proc]" '報表設定Config DB
    Dim DbProcParameter As String = "[H3_Systematic].[dbo].[H3_Production_ProcParameter]" '欄位設定Config DB
    Dim DbProcParameterRule As String = "[H3_Systematic].[dbo].[H3_Production_ProcParameter_Rule]" '欄位帶入設定Config DB
    Dim FxSpilt As String = "[H3_Systematic].[dbo].[SplitString]" 'SQL用的Spiltf函數
    '以下Db在使用時請注意Schema不一定和KM廠一樣，請確認欄位名稱是否一致
    Dim DbVersion As String = "[Datamation_H3].[dbo].[H3_Leo_Program_Version]" '版本卡控DB，若不需要可以拿掉，或是把版本管控與法改為各廠慣用語法
    Dim DbMachine As String = "[utchfacmrpt].[acme].[dbo].[PDL_Machine]" '愉進機台資料表
    Dim DbOP As String = "[utchfacmrpt].[report].[dbo].[view_OPParam]" '愉進OP參數表
    Dim DbSPCData As String = "[10.44.65.110].[SPC_PPT].[dbo].[Var_Data]" 'SPC Raw Data
    Dim DbSPCDataGroup As String = "[10.44.65.110].[SPC_PPT].[dbo].[Var_DataGroup]" 'SPC 資料群組
    Dim DbSPCCtrl As String = "[10.44.65.110].[SPC_PPT].[dbo].[Var_Ctrl]" 'SPC 管制項目 SPC第四階
    Dim DbSPCFile As String = "[10.44.65.110].[SPC_PPT].[dbo].[Var_File]" 'SPC 檔案 SPC第三階
    Dim DbSPCFileGroup As String = "[10.44.65.110].[SPC_PPT].[dbo].[Var_FileGroup]" 'SPC第一、二階
    Dim DbHist As String = "[utchfacmrpt].[report].[dbo].[view_HIST]" '愉進過帳資料View表
    Dim DbIPQC_Panel As String = "[MES-C].[IPQC_Mapping_H3].[dbo].[RawData_Panel]" 'IPQC_Panel檢驗資訊
    Dim DbIPQC_Pcs As String = "[MES-C].[IPQC_Mapping_H3].[dbo].[RawData_Pcs]" 'IPQC_Pcs缺點資料




    Private Sub ProductionReport_Setting_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '標題初始化
        Me.Text = "生產報表欄位設定程式 v" & Version & " [維運 : 李博軒]"

        TabMain.SelectedIndex = 0
        BtnProc_Refresh.PerformClick()
        Command_Initialize()


        Using dt As DataTable = SQL_Query(SQL_Conn_MQL03, "SELECT  [OnlineVersion],CASE WHEN ISNULL([TestVersion],'') = '' THEN '0' END AS [TestVersion] FROM " & DbVersion & " WHERE [Program] = '" & Program & "'") '搜尋資料庫內的所有站點和密碼
            If Not (CInt(Version.Replace(".", "")) > CInt(dt(0)(0).ToString.Replace(".", "")) OrElse CInt(Version.Replace(".", "")) = CInt(dt(0)(1).ToString.Replace(".", ""))) Then
                MessageBox.Show("請確認是否使用最新版本")
                Environment.Exit(Environment.ExitCode)
                Application.Exit()
            End If
        End Using

    End Sub

    Private Sub TabMain_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabMain.SelectedIndexChanged
        Try
            If TabMain.SelectedIndex = 0 Then
                'BtnProc_Refresh.PerformClick()
            ElseIf TabMain.SelectedIndex = 1 Then

            ElseIf TabMain.SelectedIndex = 2 Then

            End If
        Catch ex As Exception
            WriteLog(ex, "TabMain_SelectedIndexChanged")
        End Try
    End Sub
#Region "站點設定"

    'Private Sub DgvProc_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DgvProc.CellFormatting
    '    Try
    '        If e.ColumnIndex = 4 Then
    '            e.Value = "****"
    '        End If
    '    Catch ex As Exception
    '        WriteLog(ex, "DgvProc_CellFormatting")
    '    End Try
    'End Sub

    Private Sub BtnProc_Refresh_Click(sender As Object, e As EventArgs) Handles BtnProc_Refresh.Click
        Try
            'DgvProc初始化
            DgvProc.DataSource = Nothing
            DgvProc.Rows.Clear()
            DgvProc.Columns.Clear()
            '查詢目前站點資料

            Dim cmd As String = "SELECT [Pkey],[Module] AS [模組] ,[Section] AS [課別] ,[Area] AS [區域名稱] ,[ProcName] AS [站點] ,[hasFace] ,[Password] AS [密碼] ,ISNULL([Location],'') AS [Location] ,ISNULL([MachineNo],'') AS [機台愉進編號] ,[Machine] AS [機台愉進名稱] 
                                                      FROM " & DbProc & " WITH(NOLOCK) 
                                                      ORDER BY [Pkey] DESC"
            'Dim cmd As String = "SELECT [Pkey],[Module] AS [模組] ,[Section] AS [課別] ,[Area] AS [區域名稱] ,[ProcName] AS [站點]  ,'****' AS [密碼] ,ISNULL([Location],'') AS [Location] ,ISNULL([MachineNo],'') AS [機台愉進編號] ,[Machine] AS [機台愉進名稱] 
            '                                          FROM " & DbProc & " WITH(NOLOCK) 
            '                                          ORDER BY [Pkey] DESC"
            Dim ProcData As DataTable = SQL_Query(SQL_Conn_MQL03, cmd)
            ProcPassword.Clear()
            For Each row As DataRow In ProcData.Rows
                ProcPassword.Add(row("區域名稱").ToString, row("密碼").ToString)
            Next
            DgvProc.DataSource = ProcData
            'DgvProc.Columns("密碼").Visible = False
            '將Pkey欄位轉成Button

            Proc_PkeyToButton()

            Dim buttonColumn As New DataGridViewButtonColumn()
            buttonColumn.HeaderText = "欄位設定"
            buttonColumn.Text = "編輯"
            buttonColumn.Name = "欄位設定"
            buttonColumn.FlatStyle = FlatStyle.Flat
            buttonColumn.UseColumnTextForButtonValue = True
            DgvProc.Columns.Insert(0, buttonColumn)

            TxtProc_Module.Text = ""
            TxtProc_Section.Text = ""
            TxtProc_Area.Text = ""
            TxtProc_ProcName.Text = ""
            TxtProc_Password.Text = ""
            TxtProc_Machine.Text = ""
            TxtProc_Pkey.Text = ""
            TxtProc_ProcQuery.Text = ""
            CboProc_Location.Items.Clear()
            CboProc_Location.Enabled = False
            BtnProc_Add.Enabled = True
            BtnProc_Update.Enabled = False

        Catch ex As Exception
            WriteLog(ex, "BtnProc_Refresh_Click")
        End Try
    End Sub

    Private Sub TxtProc_ProcName_KeyUp(sender As Object, e As KeyEventArgs) Handles TxtProc_ProcName.KeyUp
        Try
            If e.KeyData = Keys.Enter Then
                Dim cmd As String = "SELECT DISTINCT ISNULL([MachinePosition],'') AS [MachinePosition] FROM " & DbMachine & " WHERE [EnId] = 29 AND [GTID] IN (SELECT ISNULL([Value],'') AS [Value] FROM " & FxSpilt & "('" & TxtProc_ProcName.Text & "',','))"
                Dim dt As DataTable = SQL_Query(SQL_Conn_MQL03, cmd)

                CboProc_Location.Items.Clear()
                CboProc_Location.Enabled = True
                CboProc_Location.Items.Add("")
                For Each row As DataRow In dt.Rows
                    CboProc_Location.Items.Add(row(0))
                Next
                CboProc_Location.SelectedIndex = 0
            End If
        Catch ex As Exception
            WriteLog(ex, "TxtProc_ProcName_KeyUp")
        End Try
    End Sub

    Private Sub TxtProc_ProcName_LostFocus(sender As Object, e As EventArgs) Handles TxtProc_ProcName.LostFocus
        Try
            If TxtProc_ProcName.Text <> "" Then
                Dim args As New KeyEventArgs(Keys.Enter)
                TxtProc_ProcName_KeyUp(sender, args)
            End If
        Catch ex As Exception
            WriteLog(ex, "TxtProc_ProcName_LostFocus")
        End Try
    End Sub

    Private Sub TxtProc_Password_Leave(sender As Object, e As EventArgs) Handles TxtProc_Password.Leave
        Try
            If TxtProc_Password.Text.Length > 20 OrElse Not IsNumeric(TxtProc_Password.Text) Then
                TxtProc_Password.Text = ""
                MessageBox.Show("請輸入20位以內數字")
            End If
        Catch ex As Exception
            WriteLog(ex, "TxtProc_Password_Leave")
        End Try
    End Sub
    Private Sub TxtProc_Module_Leave(sender As Object, e As EventArgs) Handles TxtProc_Module.Leave
        Try
            If TxtProc_Module.Text.Length > 1 OrElse Not IsNumeric(TxtProc_Module.Text) Then
                TxtProc_Module.Text = ""
                MessageBox.Show("請輸入1位以內數字")
            End If
        Catch ex As Exception
            WriteLog(ex, "TxtProc_Password_Leave")
        End Try
    End Sub
    Private Sub TxtProc_Section_Leave(sender As Object, e As EventArgs) Handles TxtProc_Section.Leave
        Try
            If TxtProc_Section.Text.Length > 3 Then
                TxtProc_Section.Text = ""
                MessageBox.Show("請輸入課別3碼")
            End If
        Catch ex As Exception
            WriteLog(ex, "TxtProc_Section_Leave")
        End Try
    End Sub

    Private Sub BtnProc_Add_Click(sender As Object, e As EventArgs) Handles BtnProc_Add.Click
        Try
            If Proc_Check() Then
                isException = True
                Return
            End If
            For Each row As DataGridViewRow In DgvProc.Rows
                If row.Cells("區域名稱").Value IsNot Nothing AndAlso row.Cells("區域名稱").Value.ToString = TxtProc_Area.Text Then
                    MessageBox.Show("已有相同名稱的的區域，請確認是否重複建立")
                    Return
                End If
            Next
            Dim cmd As String = "DECLARE @ID INT
                                                 INSERT INTO " & DbProc & " 
                                                 ([Module],[Section],[Area],[ProcName],[Password],[Location],[MachineNo])
                                                 VALUES ('" & TxtProc_Module.Text & "','" & TxtProc_Section.Text & "','" & Trim(TxtProc_Area.Text) & "','" & Trim(TxtProc_ProcName.Text) & "','" & TxtProc_Password.Text & "','" & CboProc_Location.SelectedText & "','" & TxtProc_Machine.Text & "')
                                                 SET @ID = (SELECT [Pkey] FROM " & DbProc & "  where [Area] = '" & TxtProc_Area.Text & "')
                                                 SELECT @ID"
            Dim dt As DataTable = SQL_Query(SQL_Conn_MQL03, cmd)
            ProcPkey = dt(0)(0).ToString

            isException = False
        Catch ex As Exception
            WriteLog(ex, "BtnProc_Add_Click")
            isException = True
        Finally
            Dim Status As String = ""
            If isException Then
                Status = "Fail"
                ChangeLog("站點設定", ProcPkey, TxtProc_Area.Text, "Add", Status)
            Else
                Status = "Success"
                ChangeLog("站點設定", ProcPkey, TxtProc_Area.Text, "Add", Status)
                BtnProc_Refresh.PerformClick()
            End If

        End Try
    End Sub

    Private Sub BtnProc_Update_Click(sender As Object, e As EventArgs) Handles BtnProc_Update.Click
        Try
            If Proc_Check() Then
                isException = True
                Return
            End If

            Dim cmd As String = "UPDATE " & DbProc & "
                                                 SET [Module] = '" & TxtProc_Module.Text & "',[Section] = '" & TxtProc_Section.Text & "',[Area] = '" & TxtProc_Area.Text & "',[ProcName] = '" & TxtProc_ProcName.Text & "',[Password] = '" & TxtProc_Password.Text & "',[Location] = '" & CboProc_Location.SelectedText & "',[MachineNo] = '" & TxtProc_Machine.Text & "'
                                                 WHERE [Pkey] = '" & TxtProc_Pkey.Text & "'"
            SQL_Query(SQL_Conn_MQL03, cmd)
            ProcPkey = TxtProc_Pkey.Text
            isException = False
        Catch ex As Exception
            WriteLog(ex, "BtnProc_Update_Click")
            isException = True
        Finally
            Dim Status As String = ""
            If isException Then
                Status = "Fail"
                ChangeLog("站點設定", ProcPkey, TxtProc_Area.Text, "Update", Status)
            Else
                Status = "Success"
                ChangeLog("站點設定", ProcPkey, TxtProc_Area.Text, "Update", Status)
                BtnProc_Refresh.PerformClick()
            End If

        End Try
    End Sub

    Private Sub TxtProc_ProcQuery_KeyUp(sender As Object, e As KeyEventArgs) Handles TxtProc_ProcQuery.KeyUp
        If e.KeyCode = Keys.Enter Then
            BtnProc_ProcQuery.PerformClick()
        End If
    End Sub

    Private Sub BtnProc_ProcQuery_Click(sender As Object, e As EventArgs) Handles BtnProc_ProcQuery.Click
        Dim cmd As String = ""
        If TxtProc_ProcQuery.Text = "" OrElse TxtProc_ProcQuery.Text = "*" Then
            cmd = "SELECT [Pkey],[Module] AS [模組] ,[Section] AS [課別] ,[Area] AS [區域名稱] ,[ProcName] AS [站點]  ,[Password] AS [密碼] ,ISNULL([Location],'') AS [Location] ,ISNULL([MachineNo],'') AS [機台愉進編號] ,[Machine] AS [機台愉進名稱] FROM " & DbProc & " WITH(NOLOCK) ORDER BY [Pkey] DESC"
        Else
            cmd = "SELECT [Pkey],[Module] AS [模組] ,[Section] AS [課別] ,[Area] AS [區域名稱] ,[ProcName] AS [站點]  ,[Password] AS [密碼] ,ISNULL([Location],'') AS [Location] ,ISNULL([MachineNo],'') AS [機台愉進編號] ,[Machine] AS [機台愉進名稱] FROM " & DbProc & " WITH(NOLOCK) WHERE [ProcName] LIKE '%" & TxtProc_ProcQuery.Text & "%'  ORDER BY [Pkey] DESC"
        End If
        Dim dt As DataTable = SQL_Query(SQL_Conn_MQL03, cmd)

        DgvProc.DataSource = Nothing
        DgvProc.DataSource = dt
        Proc_PkeyToButton()
    End Sub

    Private Sub DgvProc_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DgvProc.CellContentClick
        Try
            If e.ColumnIndex = DgvProc.Columns("Pkey").Index AndAlso e.RowIndex >= 0 Then
                Dim pw As String = InputBox("請輸入密碼：", "輸入密碼", "")
                If pw <> ProcPassword(DgvProc.Rows(e.RowIndex).Cells("區域名稱").Value.ToString) Then
                    MessageBox.Show("密碼輸入錯誤")
                    Return
                End If

                '現有資料填入欄位
                TxtProc_Module.Text = DgvProc.Rows(e.RowIndex).Cells("模組").Value.ToString
                TxtProc_Section.Text = DgvProc.Rows(e.RowIndex).Cells("課別").Value.ToString
                TxtProc_Area.Text = DgvProc.Rows(e.RowIndex).Cells("區域名稱").Value.ToString
                TxtProc_ProcName.Text = DgvProc.Rows(e.RowIndex).Cells("站點").Value.ToString
                Dim args As New KeyEventArgs(Keys.Enter)
                TxtProc_ProcName_KeyUp(sender, args)
                TxtProc_Password.Text = DgvProc.Rows(e.RowIndex).Cells("密碼").Value.ToString
                TxtProc_Machine.Text = DgvProc.Rows(e.RowIndex).Cells("機台愉進編號").Value.ToString
                TxtProc_Pkey.Text = DgvProc.Rows(e.RowIndex).Cells("Pkey").Value.ToString
                If DgvProc.Rows(e.RowIndex).Cells("Location").Value.ToString = "" Then
                    CboProc_Location.SelectedIndex = 0
                Else
                    CboProc_Location.SelectedItem = DgvProc.Rows(e.RowIndex).Cells("Location").Value.ToString
                End If

                BtnProc_Update.Enabled = True
                BtnProc_Add.Enabled = False
            ElseIf e.ColumnIndex = DgvProc.Columns("欄位設定").Index AndAlso e.RowIndex >= 0 Then
                TabMain.SelectedIndex = 1
                TxtPara_Area.Text = DgvProc.Rows(e.RowIndex).Cells("區域名稱").Value.ToString
                TxtPara_AreaID.Text = DgvProc.Rows(e.RowIndex).Cells("Pkey").Value.ToString
                BtnPara_Refresh.PerformClick()
                TxtF_Pkey.Text = ""
                TxtF_QID.Text = ""
                CboF_Type.Items.Clear()
                CboF_Value.Items.Clear()
                CboF_Value.Enabled = False
                ComboBox_Invisible()
                TxtF_QueryCommand.Text = ""
            End If
        Catch ex As Exception
            WriteLog(ex, "DgvProc_CellContentClick")
        End Try
    End Sub

    Private Sub DgvProc_Sorted(sender As Object, e As EventArgs) Handles DgvProc.Sorted
        Proc_PkeyToButton()
    End Sub

    Private Function Proc_Check() As Boolean
        Dim sender As Object
        Dim e As New EventArgs
        TxtProc_Password_Leave(sender, e)
        TxtProc_Module_Leave(sender, e)
        TxtProc_Section_Leave(sender, e)


        If TxtProc_Module.Text = "" OrElse TxtProc_Section.Text = "" OrElse TxtProc_Area.Text = "" OrElse TxtProc_ProcName.Text = "" OrElse TxtProc_Password.Text = "" Then
            MessageBox.Show("請確認必填欄位輸入完畢")
            Return True
        End If

        Return False
    End Function

    Private Sub Proc_PkeyToButton()
        For Each row As DataGridViewRow In DgvProc.Rows
            Dim ButtonCell As New DataGridViewButtonCell
            ButtonCell.Value = row.Cells("Pkey").Value.ToString
            ButtonCell.FlatStyle = FlatStyle.Flat
            'ButtonCell.Style.BackColor = SystemColors.ControlLight
            row.Cells("Pkey") = ButtonCell
        Next
    End Sub

#End Region

#Region "欄位設定"
    Private Sub BtnPara_Refresh_Click(sender As Object, e As EventArgs) Handles BtnPara_Refresh.Click
        Try
            'DgvPara初始化
            DgvPara.DataSource = Nothing
            DgvPara.Rows.Clear()
            DgvPara.Columns.Clear()

            '查詢目前站點資料
            Dim cmd As String = "SELECT [Pkey],[PID] AS [順序],[Area] AS [區域名稱],[ParameterName] AS [欄位名稱],[EnglishName] AS [英文名稱],ISNULL([isQuery],'') AS [是否自動代入數值],ISNULL([QID],'') AS [QID],ISNULL([isRequire],'') AS [是否為必填欄位],ISNULL([DefaultValues],'') AS [欄位預設值]
                                                 FROM " & DbProcParameter & "
                                                 WHERE [AreaID] = '" & TxtPara_AreaID.Text & "'
                                                 ORDER BY [PID] DESC"
            Dim ParaData As DataTable = SQL_Query(SQL_Conn_MQL03, cmd)
            DgvPara.DataSource = ParaData
            CboPara_Sort.Items.Clear()
            BtnPara_Add.Enabled = True
            BtnPara_Update.Enabled = False

            For i = 1 To DgvPara.Rows.Count + 1
                CboPara_Sort.Items.Add(i)
            Next

            Para_PkeyToButton()


            Dim buttonColumn As New DataGridViewButtonColumn()
            buttonColumn.HeaderText = "查詢(運算)設定"
            buttonColumn.Text = "編輯"
            buttonColumn.Name = "編輯"
            buttonColumn.FlatStyle = FlatStyle.Flat
            buttonColumn.UseColumnTextForButtonValue = True
            DgvPara.Columns.Insert(0, buttonColumn)

            TxtPara_ParaName.Text = ""
            TxtPara_EngName.Text = ""
            ChkPara_isQuery.Checked = False
            TxtPara_QID.Text = ""
            TxtPara_Pkey.Text = ""
            ChkPara_isRequire.Checked = True
            TxtPara_Default.Text = ""

        Catch ex As Exception
            WriteLog(ex, "BtnPara_Refresh_Click")
        End Try
    End Sub

    Private Sub BtnPara_Add_Click(sender As Object, e As EventArgs) Handles BtnPara_Add.Click
        Try
            If Para_Check() Then
                isException = True
                Return
            End If
            For Each row As DataGridViewRow In DgvPara.Rows
                If row.Cells("欄位名稱").Value IsNot Nothing AndAlso row.Cells("欄位名稱").Value.ToString = TxtPara_ParaName.Text Then
                    MessageBox.Show("已有相同名稱的的欄位，請確認是否重複建立")
                    Return
                End If
            Next
            Dim ChBoolean As New Dictionary(Of String, String)
            ChBoolean.Add("否", "0")
            ChBoolean.Add("是", "1")
            Dim cmd As String
            If CboPara_Sort.Text <> "" Then
                cmd = "UPDATE " & DbProcParameter & "
                               SET [PID] = [PID] +1
                               WHERE [AreaID] = " & TxtPara_AreaID.Text & " AND [PID] >= " & CboPara_Sort.Text & "
                               DECLARE @ID INT
                               INSERT INTO " & DbProcParameter & " 
                               ([AreaID],[PID],[Area],[ParameterName],[EnglishName],[isQuery],[QID],[isRequire],[DefaultValues])
                               VALUES(" & TxtPara_AreaID.Text & ",'" & CboPara_Sort.Text & "','" & TxtPara_Area.Text & "','" & Trim(TxtPara_ParaName.Text) & "','" & Trim(TxtPara_EngName.Text) & "','" & Convert.ToInt32(ChkPara_isQuery.Checked).ToString & "','0','" & Convert.ToInt32(ChkPara_isRequire.Checked).ToString & "','" & Trim(TxtPara_Default.Text) & "')
                               SET @ID = (SELECT [Pkey] FROM " & DbProcParameter & "  where [Area] = '" & TxtPara_Area.Text & "' AND [ParameterName] = '" & TxtPara_ParaName.Text & "')
                               SELECT @ID"
            Else
                cmd = "DECLARE @ID INT
                               INSERT INTO " & DbProcParameter & " 
                               ([AreaID],[PID],[Area],[ParameterName],[EnglishName],[isQuery],[QID],[isRequire],[DefaultValues])
                               VALUES(" & TxtPara_AreaID.Text & ",'" & CboPara_Sort.Text & "','" & TxtPara_Area.Text & "','" & Trim(TxtPara_ParaName.Text) & "','" & Trim(TxtPara_EngName.Text) & "','" & Convert.ToInt32(ChkPara_isQuery.Checked).ToString & "','0','" & Convert.ToInt32(ChkPara_isRequire.Checked).ToString & "','" & Trim(TxtPara_Default.Text) & "')
                               SET @ID = (SELECT [Pkey] FROM " & DbProcParameter & "  where [Area] = '" & TxtPara_Area.Text & "' AND [ParameterName] = '" & TxtPara_ParaName.Text & "')
                               SELECT @ID"
            End If
            Dim dt As DataTable = SQL_Query(SQL_Conn_MQL03, cmd)
            ParaPkey = dt(0)(0).ToString

            isException = False
        Catch ex As Exception
            WriteLog(ex, "BtnPara_Add_Click")
            isException = True
        Finally
            Dim Status As String = ""
            If isException Then
                Status = "Fail"
                ChangeLog("欄位設定", ParaPkey, TxtPara_ParaName.Text, "Add", Status)
            Else
                Status = "Success"
                ChangeLog("欄位設定", ParaPkey, TxtPara_ParaName.Text, "Add", Status)
                BtnPara_Refresh.PerformClick()
            End If

        End Try
    End Sub

    Private Sub BtnPara_Update_Click(sender As Object, e As EventArgs) Handles BtnPara_Update.Click
        Try
            If Para_Check() Then
                isException = True
                Return
            End If

            Dim ChBoolean As New Dictionary(Of String, String)
            ChBoolean.Add("否", "0")
            ChBoolean.Add("是", "1")

            Dim cmd As String = ""

            If CInt(CboPara_Sort.Text) = ParaSort Then '順序不變
                cmd = "UPDATE " & DbProcParameter & " 
                               SET [PID] = '" & CboPara_Sort.Text & "',[ParameterName] = '" & Trim(TxtPara_ParaName.Text) & "',[EnglishName] = '" & Trim(TxtPara_EngName.Text) & "',[isQuery] = '" & Convert.ToInt32(ChkPara_isQuery.Checked).ToString & "',[isRequire] = '" & Convert.ToInt32(ChkPara_isRequire.Checked).ToString & "',[DefaultValues] = '" & Trim(TxtPara_Default.Text) & "'
                               WHERE [Pkey] = '" & TxtPara_Pkey.Text & "'"
            ElseIf CInt(CboPara_Sort.Text) > ParaSort Then '新的順序編號>舊的順序編號
                cmd = "UPDATE " & DbProcParameter & "
                               SET [PID] = [PID] - 1
                               WHERE [AreaID] = " & TxtPara_AreaID.Text & " AND [PID] BETWEEN " & ParaSort & " AND " & CboPara_Sort.Text & " 
                               UPDATE " & DbProcParameter & " 
                               SET [PID] = '" & CboPara_Sort.Text & "',[ParameterName] = '" & Trim(TxtPara_ParaName.Text) & "',[EnglishName] = '" & Trim(TxtPara_EngName.Text) & "',[isQuery] = '" & Convert.ToInt32(ChkPara_isQuery.Checked).ToString & "',[isRequire] = '" & Convert.ToInt32(ChkPara_isRequire.Checked).ToString & "',[DefaultValues] = '" & Trim(TxtPara_Default.Text) & "'
                               WHERE [Pkey] = '" & TxtPara_Pkey.Text & "'
                               UPDATE t1 
                               SET t1.[PID] = t2.[NEW_PID]
                               FROM " & DbProcParameter & " AS t1
                               INNER JOIN (SELECT TOP (1000) [Pkey]
                                     ,[AreaID]
                                     ,[PID]
                                     ,ROW_NUMBER() OVER (ORDER BY PID) AS [NEW_PID]
                                 FROM " & DbProcParameter & "
                                 WHERE AreaID = " & TxtPara_AreaID.Text & "
                                 ORDER BY PID
                               ) AS t2 ON t1.[Pkey] = t2.[Pkey]"
            ElseIf CInt(CboPara_Sort.Text) < ParaSort Then '新的順序編號<舊的順序編號
                cmd = "UPDATE " & DbProcParameter & "
                               SET [PID] = [PID] + 1
                               WHERE [AreaID] = " & TxtPara_AreaID.Text & " AND [PID] BETWEEN " & CboPara_Sort.Text & " AND " & ParaSort & " 
                               UPDATE " & DbProcParameter & " 
                               SET [PID] = '" & CboPara_Sort.Text & "',[ParameterName] = '" & Trim(TxtPara_ParaName.Text) & "',[EnglishName] = '" & Trim(TxtPara_EngName.Text) & "',[isQuery] = '" & Convert.ToInt32(ChkPara_isQuery.Checked).ToString & "',[isRequire] = '" & Convert.ToInt32(ChkPara_isRequire.Checked).ToString & "',[DefaultValues] = '" & Trim(TxtPara_Default.Text) & "'
                               WHERE [Pkey] = '" & TxtPara_Pkey.Text & "'
                               UPDATE t1 
                               SET t1.[PID] = t2.[NEW_PID]
                               FROM " & DbProcParameter & " AS t1
                               INNER JOIN (SELECT TOP (1000) [Pkey]
                                     ,[AreaID]
                                     ,[PID]
                                     ,ROW_NUMBER() OVER (ORDER BY PID) AS [NEW_PID]
                                 FROM " & DbProcParameter & "
                                 WHERE AreaID = " & TxtPara_AreaID.Text & "
                                 ORDER BY PID
                               ) AS t2 ON t1.[Pkey] = t2.[Pkey]"
            End If

            SQL_Query(SQL_Conn_MQL03, cmd)
            ParaPkey = TxtPara_Pkey.Text
            isException = False
        Catch ex As Exception
            WriteLog(ex, "BtnPara_Update_Click")
            isException = True
        Finally
            Dim Status As String = ""
            If isException Then
                Status = "Fail"
                ChangeLog("欄位設定", ParaPkey, TxtPara_ParaName.Text, "Update", Status)
            Else
                Status = "Success"
                ChangeLog("欄位設定", ParaPkey, TxtPara_ParaName.Text, "Update", Status)
                BtnPara_Refresh.PerformClick()
            End If

        End Try
    End Sub

    Private Sub DgvPara_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DgvPara.CellContentClick
        Try
            If e.ColumnIndex = DgvPara.Columns("Pkey").Index AndAlso e.RowIndex >= 0 Then
                '現有資料填入欄位
                TxtPara_Pkey.Text = DgvPara.Rows(e.RowIndex).Cells("Pkey").Value.ToString
                If CInt(DgvPara.Rows(e.RowIndex).Cells("順序").Value.ToString) > CboPara_Sort.Items.Count Then
                    CboPara_Sort.Text = "1"
                Else
                    CboPara_Sort.Text = DgvPara.Rows(e.RowIndex).Cells("順序").Value.ToString
                End If
                TxtPara_ParaName.Text = DgvPara.Rows(e.RowIndex).Cells("欄位名稱").Value.ToString
                TxtPara_EngName.Text = DgvPara.Rows(e.RowIndex).Cells("英文名稱").Value.ToString
                ChkPara_isQuery.Checked = Convert.ToBoolean(DgvPara.Rows(e.RowIndex).Cells("是否自動代入數值").Value)
                TxtPara_QID.Text = DgvPara.Rows(e.RowIndex).Cells("QID").Value.ToString
                ChkPara_isRequire.Checked = Convert.ToBoolean(DgvPara.Rows(e.RowIndex).Cells("是否為必填欄位").Value)
                TxtPara_Default.Text = DgvPara.Rows(e.RowIndex).Cells("欄位預設值").Value.ToString
                If CboPara_Sort.Text = "" Then
                    ParaSort = 0
                Else
                    ParaSort = CboPara_Sort.Text
                End If
                BtnPara_Update.Enabled = True
                BtnPara_Add.Enabled = False
            ElseIf e.ColumnIndex = DgvPara.Columns("編輯").Index AndAlso e.RowIndex >= 0 Then
                TabMain.SelectedIndex = 2
                If DgvPara.Rows(e.RowIndex).Cells("QID").Value.ToString = "0" Then TxtF_QID.Text = "" Else TxtF_QID.Text = DgvPara.Rows(e.RowIndex).Cells("QID").Value.ToString
                TxtF_Pkey.Text = DgvPara.Rows(e.RowIndex).Cells("Pkey").Value.ToString
                ParaColumn = "入料片數,出料片數,料號,"
                For Each pararow As DataGridViewRow In DgvPara.Rows
                    ParaColumn += pararow.Cells("欄位名稱").Value.ToString + ","
                Next
                ParaColumn = ParaColumn.Substring(0, ParaColumn.Length - 1)
                BtnF_Refresh.PerformClick()
            End If
        Catch ex As Exception
            WriteLog(ex, "DgvPara_CellContentClick")
        End Try
    End Sub

    Private Sub ChkPara_isQuery_CheckedChanged(sender As Object, e As EventArgs) Handles ChkPara_isQuery.CheckedChanged
        If ChkPara_isQuery.Checked Then
            TxtPara_Default.Enabled = False
        Else
            TxtPara_Default.Enabled = True
        End If
    End Sub

    Private Sub DgvPara_Sorted(sender As Object, e As EventArgs) Handles DgvPara.Sorted
        Para_PkeyToButton()
    End Sub

    Private Function Para_Check() As Boolean
        Dim sender As Object
        Dim e As New EventArgs


        If TxtPara_ParaName.Text = "" Then
            MessageBox.Show("請確認必填欄位輸入完畢")
            Return True
        End If

        Return False
    End Function

    Private Sub Para_PkeyToButton()
        '將Pkey欄位轉成Button
        For Each row As DataGridViewRow In DgvPara.Rows
            Dim ButtonCell As New DataGridViewButtonCell
            ButtonCell.Value = row.Cells("Pkey").Value.ToString
            ButtonCell.FlatStyle = FlatStyle.Flat
            'ButtonCell.Style.BackColor = SystemColors.ControlLight
            row.Cells("Pkey") = ButtonCell
        Next
    End Sub
#End Region

#Region "查詢(運算)設定"

    Private Sub BtnF_Refresh_Click(sender As Object, e As EventArgs) Handles BtnF_Refresh.Click
        Try
            txtF_QID_Copy.Text = ""
            QueryCommand = ""
            DgvF.AllowUserToAddRows = False
            ComboBox_Invisible()
            DgvF_Operator.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            FunctionParameterLoad()
            AddFunctionType(CboF_Type)
            CboF_Value.Items.Clear()
            CboF_Value.Enabled = False
            If TxtF_QID.Text <> "" Then
                Dim cmd As String = "SELECT [QueryCommand],[QueryType],[QueryValue],[Filter1_operator],[Filter1],[Filter2_operator],[Filter2],[Filter3_operator],[Filter3],[Filter4_operator],[Filter4],[Filter5_operator],[Filter5]
                                                 FROM " & DbProcParameterRule & "
                                                 WHERE [QID] = " & TxtF_QID.Text & ""
                Dim dt As DataTable = SQL_Query(SQL_Conn_MQL03, cmd)

                TxtF_QueryCommand.Text = dt(0)(0).ToString
                CboF_Type.Text = dt(0)(1).ToString
                'Call CboF_Type_SelectedIndexChanged(Nothing, e)
                CboF_Value.Text = dt(0)(2).ToString
                Dim index As Integer = 0
                For Each row As DataGridViewRow In DgvF.Rows
                    index += 1
                    row.Cells(DgvF_Operator.Index).Value = dt(0)(2 + index * 2 - 1)
                    row.Cells(DgvF_Values.Index).Value = dt(0)(2 + index * 2)
                Next
                CboF_Condition.Text = dt(0)(11)
                CboF_Group.Text = dt(0)(12)
            End If

            If CboF_Type.SelectedItem = "其他" Then TxtF_QueryCommand.ReadOnly = False
            If TxtF_QID.Text = "" Then TxtF_QueryCommand.Text = ""

        Catch ex As Exception
            WriteLog(ex, "BtnF_Refresh_Click")
        End Try
    End Sub

    Private Sub CboF_Type_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboF_Type.SelectedIndexChanged
        DgvF.ReadOnly = False

        Select Case CboF_Type.SelectedItem
            Case "OP參數"
                ComboBox_Invisible()
                DgvF_Operator.ReadOnly = False
                TxtF_QueryCommand.ReadOnly = True
                TxtF_QueryCommand.Text = ""
                DgvF.Rows.Add("製程代碼(8碼)", "[ProcName]", "=", "", "若留空白則查詢當站資訊 Ex. SMK1CLN1")
                DgvF.Rows.Add("OP參數名稱", "[PARAMETER_DESC]", "=", "", "Ex. 基材銅厚中值")
                DgvTest.Rows.Add("層別", "VarLayer", "")
                DgvTest.Rows.Add("料號", "VarPart", "")
                DgvTest.Rows.Add("版序", "VarRev", "")
                CboF_Value.Items.Clear()
                CboF_Value.Enabled = True
                CboF_Value.Items.Add("參數中值")
                CboF_Value.Items.Add("參數上限")
                CboF_Value.Items.Add("參數下限")
                CboF_Value.SelectedIndex = 0
            Case "SPC"
                DgvF.Rows.Clear()
                CboF_Condition.Visible = True
                LblF_Condition.Visible = True
                CboF_Group.Visible = True
                LblF_Group.Visible = True
                DgvF_Operator.ReadOnly = False
                TxtF_QueryCommand.ReadOnly = True
                TxtF_QueryCommand.Text = ""
                DgvF.Rows.Add("一階", "d.[FileFactory]", "=", "", "Ex. 13新豐廠H3")
                DgvF.Rows.Add("二階", "e.[FileGroupName]", "=", "", "Ex. 15MDL_機械鑽孔")
                DgvF.Rows.Add("三階", "d.[FileName]", "=", "", "Ex. 15MDL_SUEP_ETCHINIG AMOUNT_EQP")
                DgvF.Rows.Add("四階", "c.[CtrlName]", "=", "", "Ex. Etching Amount")
                DgvTest.Rows.Clear()
                DgvTest.Rows.Add("批號", "VarLot", "")
                CboF_Value.Items.Clear()
                CboF_Value.Enabled = True
                CboF_Value.Items.Add("AVG")
                CboF_Value.Items.Add("MAX")
                CboF_Value.Items.Add("MIN")
                CboF_Value.Items.Add("CL")
                CboF_Value.Items.Add("SL")
                CboF_Value.Items.Add("AL")
                CboF_Value.SelectedIndex = 0
                CboF_Condition.Items.Clear()
                CboF_Condition.Items.Add("批號")
                CboF_Condition.Items.Add("最近一筆")
                CboF_Condition.SelectedIndex = 0
            Case "愉進系統"
                ComboBox_Invisible()
                DgvF_Operator.ReadOnly = False
                TxtF_QueryCommand.ReadOnly = True
                TxtF_QueryCommand.Text = ""
                DgvF.Rows.Add("製程代碼(8碼)", "[ProcName]", "=", "", "若留空白則查詢當站資訊 Ex. SMK1CLN1")
                DgvF.Rows.Add("愉進過帳狀態", "[AftStatus]", "=", "", "MoveIn Or CheckIn Or CheckOut Or MoveOut")
                DgvTest.Rows.Add("批號", "VarLot", "")
                DgvTest.Rows.Add("層別", "VarLayer", "")
                CboF_Value.Items.Clear()
                CboF_Value.Enabled = True
                CboF_Value.Items.Add("過帳時間")
                CboF_Value.Items.Add("PCS數量")
                CboF_Value.SelectedIndex = 0
            Case "欄位間計算"
                ComboBox_Invisible()
                DgvF_Operator.ReadOnly = True
                TxtF_QueryCommand.ReadOnly = True
                TxtF_QueryCommand.Text = ""
                DgvF.Rows.Add("計算欄位", "Var", "", "", "請按照順序填寫，複數欄位請以逗號("","")分隔 Ex. " & ParaColumn & " ")
                DgvF.Rows.Add("計算公式", "Formula", "", "", "填寫公式時請按照上面填寫的順序將變數命名為var1、var2... Ex. (var1/var2)*100")
                CboF_Value.Items.Clear()
                CboF_Value.Items.Add("數字")
                CboF_Value.Items.Add("文字")
                CboF_Value.Items.Add("其他")
                CboF_Value.Enabled = True

                '建立語法
                QueryCommand_Create()
            Case "IPQC" 'H3客製
                ComboBox_Invisible()
                DgvF_Operator.ReadOnly = False
                TxtF_QueryCommand.ReadOnly = True
                TxtF_QueryCommand.Text = ""
                Dim cmdProc As String = "SELECT [ProcName],[ProcNo]FROM [MES-C].[IPQC_Mapping_H3].[dbo].[DefineProcess] WITH(NOLOCK)"
                Dim dtProc As DataTable = SQL_Query(SQL_Conn_MQL03, cmdProc)
                Dim dgvcboProc As New DataGridViewComboBoxCell
                F_IPQC_Proc.Clear()
                For Each row As DataRow In dtProc.Rows
                    F_IPQC_Proc.Add(row("ProcName").ToString, row("ProcNo").ToString)
                    dgvcboProc.Items.Add(row("ProcName"))
                Next
                DgvF.Rows.Add("Process", "ProcNo", "=", "", "請填入IPQC佔點 Ex. SM Cure UV")
                DgvF.Rows(0).Cells(DgvF_Values.Index) = dgvcboProc
                DgvF.Rows.Add("缺點名稱", "ItemName", "=", "", "Ex. 其他主缺 (Other Major)")
                DgvTest.Rows.Add("批號", "VarLot", "")
                DgvTest.Rows.Add("面次", "VarFace", "")
                DgvTest.Rows.Add("站點8碼", "VarProc", "")
                CboF_Value.Items.Clear()
                CboF_Value.Enabled = True
                CboF_Value.Items.Add("缺點數量")
                CboF_Condition.Visible = True
                CboF_Condition.Items.Clear()
                CboF_Condition.Items.Add("需篩選層別")
                CboF_Condition.Items.Add("不需篩選層別")
                CboF_Condition.SelectedIndex = 0
                CboF_Value.SelectedIndex = 0
            Case "其他"
                ComboBox_Invisible()
                TxtF_QueryCommand.ReadOnly = False
                CboF_Value.Items.Clear()
                CboF_Value.Enabled = False
        End Select
    End Sub

    Private Sub CboF_Value_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboF_Value.SelectedIndexChanged
        QueryCommand_Create()
    End Sub

    Private Sub CboF_SPC_Condition_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboF_Condition.SelectedIndexChanged
        Select Case CboF_Condition.SelectedItem
            Case "批號"
                CboF_Group.Items.Clear()
                CboF_Group.Items.Add("一階")
                CboF_Group.Items.Add("二階")
                CboF_Group.Items.Add("三階")
                CboF_Group.Items.Add("四階")
                CboF_Group.Items.Add("樣本平均")
                CboF_Group.SelectedIndex = 0
                DgvTest.Rows.Clear()
                DgvTest.Rows.Add("批號", "VarLot", "")
            Case "最近一筆"
                DgvTest.Rows.Clear()
            Case "需篩選層別"
                DgvTest.Rows.Clear()
                DgvTest.Rows.Add("批號", "VarLot", "")
                DgvTest.Rows.Add("層別", "VarLayer", "")
                DgvTest.Rows.Add("面次", "VarFace", "")
                DgvTest.Rows.Add("站點8碼", "VarProc", "")
            Case "不需篩選層別"
                DgvTest.Rows.Clear()
                DgvTest.Rows.Add("批號", "VarLot", "")
                DgvTest.Rows.Add("面次", "VarFace", "")
                DgvTest.Rows.Add("站點8碼", "VarProc", "")
        End Select


        QueryCommand_Create()
    End Sub
    Private Sub CboF_SPC_Group_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboF_Group.SelectedIndexChanged
        QueryCommand_Create()
    End Sub

    Private Sub BtnF_Save_Click(sender As Object, e As EventArgs) Handles BtnF_Save.Click
        Try
            F_SaveType = ""
            If F_Check() Then
                isException = True
                Return
            ElseIf isError Then
                MessageBox.Show("請確認各欄位資料格式正確")
                isException = True
                Return
            End If
            Dim cmd As String = ""

            If TxtF_QID.Text IsNot Nothing AndAlso TxtF_QID.Text <> "" Then
                F_SaveType = "Update"
                cmd = "UPDATE " & DbProcParameterRule & "
                             SET [QueryCommand] = '" & TxtF_QueryCommand.Text.Replace("'", "''") & "'
                             ,[QueryType] = '" & CboF_Type.SelectedItem & "'
                             ,[QueryValue] = '" & CboF_Value.SelectedItem & "'
                             ,[Filter1] = 'f1'
                             ,[Filter1_operator] = 'ff1'
                             ,[Filter2] = 'f2'
                             ,[Filter2_operator] = 'ff2'
                             ,[Filter3] = 'f3'
                             ,[Filter3_operator] = 'ff3'
                             ,[Filter4] = 'f4'
                             ,[Filter4_operator] = 'ff4'
                             ,[Filter5] = 'f5'
                             ,[Filter5_operator] = 'ff5'
                             WHERE [QID] = '" & TxtF_QID.Text & "'"
                Dim index As Integer = 0
                For Each row As DataGridViewRow In DgvF.Rows
                    index += 1
                    Dim ope As String
                    Dim val As String
                    If row.Cells(DgvF_Operator.Index).Value IsNot Nothing Then
                        ope = row.Cells(DgvF_Operator.Index).Value.ToString
                    Else
                        ope = ""
                    End If
                    If row.Cells(DgvF_Values.Index).Value IsNot Nothing Then
                        val = row.Cells(DgvF_Values.Index).Value.ToString
                    Else
                        val = ""
                    End If
                    cmd = cmd.Replace("ff" + index.ToString, ope).Replace("f" + index.ToString, val)
                Next
                For i = index + 1 To 5
                    If i = 5 Then
                        cmd = cmd.Replace("ff5", CboF_Condition.Text).Replace("f5", CboF_Group.Text)
                    Else
                        cmd = cmd.Replace("ff" + i.ToString, "").Replace("f" + i.ToString, "")
                    End If
                Next
            Else
                F_SaveType = "Add"
                cmd = "DECLARE @ID TABLE (id INT)
                             INSERT INTO " & DbProcParameterRule & "
                             ([PID],[QueryCommand],[QueryType],[QueryValue],[Filter1],[Filter1_operator],[Filter2],[Filter2_operator],[Filter3],[Filter3_operator],[Filter4],[Filter4_operator],[Filter5],[Filter5_operator])
                             OUTPUT INSERTED.[QID] INTO @ID 
                             VALUES('" & TxtF_Pkey.Text & "','" & TxtF_QueryCommand.Text.Replace("'", "''") & "','" & CboF_Type.SelectedItem & "','" & CboF_Value.SelectedItem & "','f1','ff1','f2','ff2','f3','ff3','f4','ff4','f5','ff5')
                             UPDATE " & DbProcParameter & "
                             SET [QID] = (SELECT id FROM @ID)
                             WHERE [Pkey] = '" & TxtF_Pkey.Text & "'
                             SELECT [id] FROM @ID"
                Dim index As Integer = 0
                For Each row As DataGridViewRow In DgvF.Rows
                    index += 1
                    Dim ope As String
                    Dim val As String
                    If row.Cells(DgvF_Operator.Index).Value IsNot Nothing Then
                        ope = row.Cells(DgvF_Operator.Index).Value.ToString
                    Else
                        ope = ""
                    End If
                    If row.Cells(DgvF_Values.Index).Value IsNot Nothing Then
                        val = row.Cells(DgvF_Values.Index).Value.ToString
                    Else
                        val = ""
                    End If
                    cmd = cmd.Replace("ff" + index.ToString, ope).Replace("f" + index.ToString, val)
                Next
                For i = index + 1 To 5
                    If i = 5 Then
                        cmd = cmd.Replace("ff5", CboF_Condition.Text).Replace("f5", CboF_Group.Text)
                    Else
                        cmd = cmd.Replace("ff" + i.ToString, "").Replace("f" + i.ToString, "")
                    End If
                Next
                'For Each row As DataGridViewRow In DgvF.Rows
                '    index += 1
                '    cmd2 += ",'" & row.Cells(DgvF_Values.Index).Value.ToString & "','" & row.Cells(DgvF_Operator.Index).Value.ToString & "'"
                'Next
                'For i = 0 To 4 - index
                '    cmd2 += ",'',''"
                'Next

            End If

            ' SQL_Query(SQL_Conn_MQL03, cmd)
            If F_SaveType = "Add" Then
                Dim dt As DataTable = SQL_Query(SQL_Conn_MQL03, cmd)
                TxtF_QID.Text = dt(0)(0)
            Else
                SQL_Query(SQL_Conn_MQL03, cmd)
            End If
            isException = False
            MessageBox.Show("儲存成功")
        Catch ex As Exception
            WriteLog(ex, "BtnF_Save_Click")
            isException = True
        Finally
            Dim Status As String = ""
            If isException Then
                Status = "Fail"
                ChangeLog("邏輯設定", TxtF_QID.Text, CboF_Type.Text, F_SaveType, Status)
            Else
                Status = "Success"
                ChangeLog("邏輯設定", TxtF_QID.Text, CboF_Type.Text, F_SaveType, Status)
                BtnF_Refresh.PerformClick()
                TabMain.SelectedIndex = 1
                BtnPara_Refresh_Click(sender, e)
            End If

        End Try
    End Sub


    Private Function F_Check() As Boolean
        Dim Required As String() = {"AftStatus", "Var", "Formula", "PARAMETER_DESC", "FileFactory", "FileGroupName", "ProcNo", "ItemName"}
        If CboF_Type.SelectedItem = "" Then
            MessageBox.Show("請確認必填欄位輸入完畢")
            Return True
        Else
            For Each row As DataGridViewRow In DgvF.Rows
                If Required.Contains(row.Cells(DgvF_ColumnName.Index).Value.ToString) AndAlso (row.Cells(DgvF_Values.Index).Value Is Nothing OrElse row.Cells(DgvF_Values.Index).Value.ToString = "") Then
                    MessageBox.Show("請確認必填欄位輸入完畢")
                    Return True
                End If
            Next
        End If
        Return False
    End Function

    Private Sub DgvF_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DgvF.CellValueChanged
        Try
            If TxtF_Pkey.Text <> "" Then
                Dim AlarmMessages As String = ""
                For Each row As DataGridViewRow In DgvF.Rows
                    If row.Cells(DgvF_ColumnName.Index).Value.ToString.Contains("ProcName") Then
                        If row.Cells(DgvF_Values.Index).Value IsNot Nothing AndAlso row.Cells(DgvF_Values.Index).Value.ToString <> "" Then
                            Dim procs As String() = row.Cells(DgvF_Values.Index).Value.ToString.Split(",")
                            For Each proc In procs
                                If Trim(proc) <> "" AndAlso Trim(proc).Length <> 8 Then
                                    AlarmMessages += "請確認輸入的製程代碼為8位字元" + vbCrLf
                                End If
                            Next
                        End If
                    ElseIf row.Cells(DgvF_ColumnName.Index).Value.ToString.Contains("AftStatus") Then
                        If row.Cells(DgvF_Values.Index).Value IsNot Nothing AndAlso row.Cells(DgvF_Values.Index).Value.ToString <> "" Then
                            Dim AftStatus As String = ""
                            Dim TrimStr As String = ""
                            If Not IsNumeric(AftStatus) Then AftStatus = row.Cells(DgvF_Values.Index).Value.ToString.ToLower
                            AftStatus = AftStatus.Replace("movein", "MoveIn").Replace("checkin", "CheckIn").Replace("checkout", "CheckOut").Replace("moveout", "MoveOut")
                            Dim Status As String() = AftStatus.Split(",")
                            For Each item In Status
                                If Trim(item) <> "" AndAlso Not (Trim(item) = "MoveIn" OrElse Trim(item) = "CheckIn" OrElse Trim(item) = "CheckOut" OrElse Trim(item) = "MoveOut") Then
                                    AlarmMessages += "請確認輸入正確的過帳狀態(MoveIn 、 CheckIn 、 CheckOut 、 MoveOut)" + vbCrLf
                                    GoTo AftError
                                Else
                                    TrimStr += Trim(item) + ","
                                End If
                            Next
                            row.Cells(DgvF_Values.Index).Value = TrimStr.Substring(0, TrimStr.Length - 1)
AftError:
                        End If
                    ElseIf row.Cells(DgvF_ColumnName.Index).Value.ToString.Contains("Var") Then
                        If row.Cells(DgvF_Values.Index).Value IsNot Nothing AndAlso row.Cells(DgvF_Values.Index).Value.ToString <> "" Then
                            Dim Vars As String() = row.Cells(DgvF_Values.Index).Value.ToString.Split(",")
                            DgvTest.Rows.Clear()
                            Dim index As Integer = 0
                            For Each Var In Vars
                                index += 1
                                DgvTest.Rows.Add(Var, "var" & index.ToString, "")
                                For Each para In ParaColumn.Split(",")
                                    If Trim(Var) <> para Then
                                        AlarmMessages = "請輸入正確的欄位名稱"
                                    Else
                                        AlarmMessages = ""
                                        Exit For
                                    End If
                                Next
                            Next
                        End If
                    End If
                Next
                If AlarmMessages <> "" Then
                    MessageBox.Show(AlarmMessages)
                    isError = True
                    Exit Sub
                Else
                    isError = False
                End If
                QueryCommand_Create()
            End If
        Catch ex As Exception
            WriteLog(ex, "DgvF_CellValueChanged")
        End Try
    End Sub

    Private Sub ComboBox_Invisible()
        DgvF.Rows.Clear()
        DgvTest.Rows.Clear()
        LblF_Condition.Visible = False
        CboF_Condition.Visible = False
        CboF_Condition.Items.Clear()
        LblF_Group.Visible = False
        CboF_Group.Visible = False
        CboF_Group.Items.Clear()
    End Sub

    Private Sub QueryCommand_Create()
        Try
            Select Case CboF_Type.SelectedItem
                Case "其他"
                Case "欄位間計算"
                    QueryCommand = ServerCommand(CboF_Type.SelectedItem.ToString)
                    If DgvF.Rows(0).Cells(DgvF_Values.Index).Value IsNot Nothing AndAlso DgvF.Rows(1).Cells(DgvF_Values.Index).Value IsNot Nothing AndAlso DgvF.Rows(0).Cells(DgvF_Values.Index).Value.ToString <> "" AndAlso DgvF.Rows(1).Cells(DgvF_Values.Index).Value.ToString <> "" Then
                        Dim FormulaCommand As String = ""
                        Dim FormulaColumn As String()
                        FormulaColumn = DgvF.Rows(0).Cells(DgvF_Values.Index).Value.ToString.Split(",")
                        FormulaCommand = DgvF.Rows(1).Cells(DgvF_Values.Index).Value.ToString
                        FormulaCommand = FormulaCommand.Replace("var", "VAR").Replace("Var", "VAR")
                        If CboF_Value.SelectedIndex = 0 Then
                            For i = FormulaColumn.Count To 1 Step -1
                                FormulaCommand = FormulaCommand.Replace("VAR" + i.ToString + "", "CDbl(""var" + i.ToString + """)")
                            Next
                            QueryCommand = QueryCommand.Replace("0.0", "(" & FormulaCommand & ").ToString")
                        ElseIf CboF_Value.SelectedIndex = 1 Then
                            For i = FormulaColumn.Count To 1 Step -1
                                FormulaCommand = FormulaCommand.Replace("VAR" + i.ToString + "", """var" + i.ToString + """")
                            Next
                            QueryCommand = QueryCommand.Replace("0.0", FormulaCommand)
                        ElseIf CboF_Value.SelectedIndex = 2 Then
                            For i = FormulaColumn.Count To 1 Step -1
                                FormulaCommand = FormulaCommand.Replace("VAR" + i.ToString + "", """var" + i.ToString + """")
                            Next
                            QueryCommand = QueryCommand.Replace("result = 0.0", FormulaCommand)
                        Else
                            For i = FormulaColumn.Count To 1 Step -1
                                FormulaCommand = FormulaCommand.Replace("VAR" + i.ToString + "", "CDbl(""var" + i.ToString + """)")
                            Next
                            QueryCommand = QueryCommand.Replace("0.0", FormulaCommand)
                        End If
                    End If
                Case "IPQC" 'H3客製
                    QueryCommand = ServerCommand(CboF_Type.SelectedItem.ToString)
                    For Each row As DataGridViewRow In DgvF.Rows
                        '每個篩選條件名稱
                        Dim column As String = row.Cells(DgvF_ColumnName.Index).Value.ToString
                        '每個篩選條件運算子
                        Dim op As String = ""
                        If row.Cells(DgvF_Operator.Index).Value IsNot Nothing AndAlso row.Cells(DgvF_Operator.Index).Value.ToString <> "" Then
                            op = row.Cells(DgvF_Operator.Index).Value.ToString
                        End If
                        '每個篩選條件
                        Dim val As String = ""
                        If row.Cells(DgvF_Values.Index).Value IsNot Nothing AndAlso row.Cells(DgvF_Values.Index).Value.ToString <> "" Then
                            val = Trim(row.Cells(DgvF_Values.Index).Value.ToString)
                        End If

                        If op <> "" And val <> "" Then
                            If column = "ProcNo" Then val = F_IPQC_Proc(val)
                            QueryCommand = QueryCommand.Replace("= '" & column & "'", Operator_Values_Check(op, val))
                        End If
                    Next

                    If CboF_Condition.SelectedItem = "需篩選層別" Then

                        QueryCommand = QueryCommand.Replace("= @Face", "= @Face
AND [LayerName] = 'VarLayer'")

                    End If
                Case Else
                    QueryCommand = SelectCommand(CboF_Value.SelectedItem.ToString) + vbCrLf + ServerCommand(CboF_Type.SelectedItem.ToString)
                    For Each row As DataGridViewRow In DgvF.Rows
                        '每個篩選條件名稱
                        Dim column As String = row.Cells(DgvF_ColumnName.Index).Value.ToString
                        '每個篩選條件運算子
                        Dim op As String = ""
                        If row.Cells(DgvF_Operator.Index).Value IsNot Nothing AndAlso row.Cells(DgvF_Operator.Index).Value.ToString <> "" Then
                            op = row.Cells(DgvF_Operator.Index).Value.ToString
                        End If
                        '每個篩選條件
                        Dim val As String = ""
                        If row.Cells(DgvF_Values.Index).Value IsNot Nothing AndAlso row.Cells(DgvF_Values.Index).Value.ToString <> "" Then
                            val = Trim(row.Cells(DgvF_Values.Index).Value.ToString)
                        End If

                        If op <> "" And val <> "" Then
                            QueryCommand += vbCrLf + "AND " + column + Space(1) + Operator_Values_Check(op, val)
                        ElseIf DefaulfFilter.ContainsKey(column) Then
                            QueryCommand += vbCrLf + "AND " + DefaulfFilter(column)
                        End If
                    Next

                    If CboF_Condition.SelectedIndex = 0 Then
                        QueryCommand += vbCrLf + "AND b.[L2] LIKE '%VarLot%'" + vbCrLf + SPCGroup(CboF_Group.SelectedItem)
                        If CboF_Value.SelectedItem = "CL" Then
                            QueryCommand += ",c.[TopCL],c.[TopUCL],c.[TopLCL]"
                        ElseIf CboF_Value.SelectedItem = "SL" Then
                            QueryCommand += ",c.[SL],c.[USL],c.[LSL]"
                        ElseIf CboF_Value.SelectedItem = "AL" Then
                            QueryCommand += ",c.[AL],c.[UAL],c.[LAL]"
                        End If
                    ElseIf CboF_Type.SelectedIndex = 1 Then
                        QueryCommand += vbCrLf + "GROUP BY a.[MeasTime]"
                        If CboF_Value.SelectedItem = "CL" Then
                            QueryCommand += ",c.[TopCL],c.[TopUCL],c.[TopLCL]"
                        ElseIf CboF_Value.SelectedItem = "SL" Then
                            QueryCommand += ",c.[SL],c.[USL],c.[LSL]"
                        ElseIf CboF_Value.SelectedItem = "AL" Then
                            QueryCommand += ",c.[AL],c.[UAL],c.[LAL]"
                        End If
                        QueryCommand += vbCrLf + "ORDER BY a.[MeasTime] DESC"
                    End If
            End Select

            TxtF_QueryCommand.Text = QueryCommand
        Catch ex As Exception
            WriteLog(ex, "QueryCommand_Create")
        End Try

    End Sub

    Private Sub BtnF_Test_Click(sender As Object, e As EventArgs) Handles BtnF_Test.Click
        Try
            If F_Check() Then
                Return
            ElseIf isError Then
                MessageBox.Show("請確認各欄位資料格式正確")
                Return
            End If

            Dim cmd As String = ""

            cmd = TxtF_QueryCommand.Text
            For i = DgvTest.Rows.Count - 1 To 0 Step -1
                If DgvTest.Rows(i).Cells(DgvTest_Values.Index).Value IsNot Nothing AndAlso DgvTest.Rows(i).Cells(DgvTest_Values.Index).Value.ToString <> "" Then
                    cmd = cmd.Replace(DgvTest.Rows(i).Cells(DgvTest_ColumnName.Index).Value.ToString, DgvTest.Rows(i).Cells(DgvTest_Values.Index).Value.ToString)
                Else
                    MessageBox.Show("請確認測試參數已填寫完畢")
                    Exit Sub
                End If
            Next

            If CboF_Type.SelectedItem <> "欄位間計算" AndAlso CboF_Type.SelectedItem <> "其他" Then

                Dim dt As DataTable = SQL_Query(SQL_Conn_MQL03, cmd)
                If dt.Rows.Count > 0 Then
                    MessageBox.Show("連線 : 成功 " + vbCrLf + "結果 : " + dt(0)(0).ToString)
                Else
                    MessageBox.Show("連線 : 成功 " + vbCrLf + "結果 : 查無資料")
                End If
            ElseIf CboF_Type.SelectedIndex = 3 Then
                Dim result As String = " "
                result = ReCoding(cmd)
                If result <> " " Then
                    MessageBox.Show("測試 : 成功 " + vbCrLf + "結果 : " + result)
                Else
                    MessageBox.Show("測試 : 失敗 ")
                End If
            End If
        Catch ex As Exception
            WriteLog(ex, "BtnF_Test_Click")
        End Try
    End Sub

    Private Sub txtF_QID_Copy_TextChanged(sender As Object, e As EventArgs) Handles txtF_QID_Copy.TextChanged
        Try
            If txtF_QID_Copy.Text <> "" AndAlso Not IsNumeric(txtF_QID_Copy.Text) Then
                txtF_QID_Copy.Text = ""
                MessageBox.Show("請輸入數字")
            End If
        Catch ex As Exception
            WriteLog(ex, "txtF_QID_Copy_TextChanged")
        End Try
    End Sub

    Private Sub txtF_QID_Copy_KeyUp(sender As Object, e As KeyEventArgs) Handles txtF_QID_Copy.KeyUp
        Try
            If e.KeyData = Keys.Enter Then
                Call btnF_QID_Copy_Click(sender, e)
            End If
        Catch ex As Exception
            WriteLog(ex, "txtF_QID_Copy_KeyUp")
        End Try
    End Sub

    Private Sub btnF_QID_Copy_Click(sender As Object, e As EventArgs) Handles btnF_QID_Copy.Click
        Try
            If txtF_QID_Copy.Text <> "" Then
                QueryCommand = ""
                DgvF.AllowUserToAddRows = False
                ComboBox_Invisible()
                DgvF_Operator.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                AddFunctionType(CboF_Type)
                CboF_Value.Items.Clear()
                CboF_Value.Enabled = False

                Dim cmd As String = "SELECT [QueryCommand],[QueryType],[QueryValue],[Filter1_operator],[Filter1],[Filter2_operator],[Filter2],[Filter3_operator],[Filter3],[Filter4_operator],[Filter4],[Filter5_operator],[Filter5]
                                                 FROM " & DbProcParameterRule & "
                                                 WHERE [QID] = " & txtF_QID_Copy.Text & ""
                Dim dt As DataTable = SQL_Query(SQL_Conn_MQL03, cmd)

                TxtF_QueryCommand.Text = dt(0)(0).ToString
                CboF_Type.Text = dt(0)(1).ToString
                'Call CboF_Type_SelectedIndexChanged(Nothing, e)
                CboF_Value.Text = dt(0)(2).ToString
                Dim index As Integer = 0
                For Each row As DataGridViewRow In DgvF.Rows
                    index += 1
                    row.Cells(DgvF_Operator.Index).Value = dt(0)(2 + index * 2 - 1)
                    row.Cells(DgvF_Values.Index).Value = dt(0)(2 + index * 2)
                Next
                CboF_Condition.Text = dt(0)(11)
                CboF_Group.Text = dt(0)(12)
                TxtF_QueryCommand.ReadOnly = True
            End If
        Catch ex As Exception
            WriteLog(ex, "btnF_QID_Copy_Click")
        End Try
    End Sub
#End Region

    Private Function Operator_Values_Check(ByVal op As String, ByVal val As String) As String
        Dim str As String = ""
        val = val.Replace("'", "''")
        If op = "=" OrElse op = "<>" Then
            val = "'" + val + "'"
            str = op + Space(1) + val
        ElseIf op = "IN" Then
            Dim trimval As String = ""
            For Each item In val.Split(",")
                trimval += Trim(item) + ","
            Next
            val = trimval.Substring(0, trimval.Length - 1)
            val = "('" + val + "')"
            val = val.Replace(",", "','")
            str = op + Space(1) + val
        ElseIf op.Contains("LIKE") Then
            val = "'%" + val + "%'"
            str = op + Space(1) + val
        ElseIf op.Contains("NULL") Then
            str = op
        End If
        Return str
    End Function

    Private Sub Command_Initialize()

        '查詢種類
        ServerCommand.Add("OP參數", "FROM " & DbOP & " WITH (NOLOCK)
WHERE [LayerName] = 'VarLayer'
AND [PartNum] = 'VarPart'
AND [Revision] = 'VarRev'")
        ServerCommand.Add("SPC", "FROM " & DbSPCData & " AS a WITH (NOLOCK)
	  INNER JOIN " & DbSPCDataGroup & " AS b WITH (NOLOCK) ON a.[DataGroupID] = b.[DataGroupID] 
	  INNER JOIN " & DbSPCCtrl & " AS c WITH (NOLOCK) ON b.[CtrlID] = c.[CtrlID]
	  INNER JOIN " & DbSPCFile & " AS d WITH (NOLOCK) ON c.[FileID] = d.[FileID]
	  INNER JOIN " & DbSPCFileGroup & " AS e WITH (NOLOCK) ON d.[FileGroupID] = e.[FileGroupID]
WHERE 1=1")
        ServerCommand.Add("愉進系統", "FROM " & DbHist & " WITH(NOLOCK)
WHERE [lotnum] = 'VarLot'
AND [LayerName] = 'VarLayer'")
        ServerCommand.Add("欄位間計算", "Imports System
  Public Class Calculator
    Public Function Calculate() as String
    Dim result As String = """"
    result = 0.0
    Return result
  End Function
End Class")
        ServerCommand.Add("IPQC", "DECLARE @Face AS varchar(5)
SET @Face = CASE WHEN 'VarFace' = '1' THEN 1 WHEN 'VarFace' = '2' THEN 0 END

SELECT 
    CASE 
        WHEN EXISTS (
SELECT TOP 1 [ChangeTIme] FROM " & DbHist & " WITH(NOLOCK) 
WHERE [lotnum] = 'VarLot'
AND [LayerName] = 'VarLayer'
AND [ProcName] = 'VarProc'
        )
        THEN (
SELECT TOP 1 CAST(ISNULL(SUM([DefectQty]),0) AS varchar) AS [Qty] FROM " & DbIPQC_Panel & " AS pnl WITH(NOLOCK) 
LEFT JOIN " & DbIPQC_Pcs & " AS pcs WITH(NOLOCK) ON pnl.[RecNo] = pcs.[RecNo] 
WHERE pnl.[LotID] + pnl.[LotNo] = 'VarLot'
AND pnl.[Side] = @Face
AND pnl.[ProcNo] = 'ProcNo'
AND TRIM(pcs.[ItemName]) = 'ItemName'
        )
        ELSE ''
    END")


        '結果查詢
        SelectCommand.Add("參數中值", "SELECT TOP 1 [PARAMETER_VALUE]")
        SelectCommand.Add("參數上限", "SELECT TOP 1 ISNULL([PARAMETER_VALUE_MAX],'') AS [PARAMETER_VALUE_MAX]")
        SelectCommand.Add("參數下限", "SELECT TOP 1 ISNULL([PARAMETER_VALUE_Min],'') AS [PARAMETER_VALUE_Min]")
        SelectCommand.Add("AVG", "SELECT TOP 1 ISNULL(ROUND(AVG(CONVERT(float,a.[MeasData])),3),'') as 'AVG'")
        SelectCommand.Add("MAX", "SELECT TOP 1 ISNULL(ROUND(MAX(CONVERT(float,a.[MeasData])),3),'') as 'MAX'")
        SelectCommand.Add("MIN", "SELECT TOP 1 ISNULL(ROUND(MIN(CONVERT(float,a.[MeasData])),3),'') as 'MIN'")
        SelectCommand.Add("CL", "SELECT TOP 1 ISNULL(CAST(ROUND(c.[TopCL],4) AS nvarchar) + ' / ' + CAST(ROUND(c.[TopUCL],4) AS nvarchar) + ' / ' + CAST(ROUND(c.[TopLCL],4) AS nvarchar),'') as 'CL'")
        SelectCommand.Add("SL", "SELECT TOP 1 ISNULL(CAST(ROUND(c.[SL],4) AS nvarchar) + ' / ' + CAST(ROUND(c.[USL],4) AS nvarchar) + ' / ' + CAST(ROUND(c.[LSL],4) AS nvarchar),'') as 'SL'")
        SelectCommand.Add("AL", "SELECT TOP 1 ISNULL(CAST(ROUND(c.[AL],4) AS nvarchar) + ' / ' + CAST(ROUND(c.[UAL],4) AS nvarchar) + ' / ' + CAST(ROUND(c.[LAL],4) AS nvarchar),'') as 'AL'")
        SelectCommand.Add("過帳時間", "SELECT TOP 1 [ChangeTime]")
        SelectCommand.Add("PCS數量", "SELECT TOP 1 [Qnty_S]")

        '預設查詢條件
        DefaulfFilter.Add("[ProcName]", "[ProcName] = 'VarProc'")
        DefaulfFilter.Add("[Lotnum]", "[Lotnum] = 'VarLot'")
        DefaulfFilter.Add("[LayerName]", "[LayerName] = 'VarLayer'")
        DefaulfFilter.Add("[PartNum]", "[PartNum] = 'VarPart'")
        DefaulfFilter.Add("[Revision]", "[Revision] = 'VarRev'")
        DefaulfFilter.Add("b.[L2]", "b.[L2] = 'VarLot'")

        'SPC分組
        SPCGroup.Add("一階", "GROUP BY d.[FileFactory]")
        SPCGroup.Add("二階", "GROUP BY e.[FileGroupName]")
        SPCGroup.Add("三階", "GROUP BY d.[FileName]")
        SPCGroup.Add("四階", "GROUP BY c.[CtrlName]")
        SPCGroup.Add("樣本平均", "GROUP BY a.[DataGroupID]")

    End Sub

End Class
