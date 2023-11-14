Imports System.IO
Imports System.Linq
Imports Microsoft.VisualBasic.FileIO
Imports System.Threading
Imports System.Reflection
Imports System.Data.SqlClient

'時間            修改人員    修改內容
'20231016 Boris            新增功能 : 1. 必填欄位及選填欄位 2. 留空會上傳預設值 3. 英文名稱
'20231017 Boris            修正預存程序查詢語法、取消人員輸入欄位的ReadOnly
'20231019 Boris            修正欄位計算的語法、資料庫查詢語法修正
'20231030 Boris            建立Table & SP名稱的變數

Public Class ReportUI
    Dim Version As String = "2.0.23.11.06.1"
    Dim Program As String = "ProductionReport"
    Public Area As String = ""
    Public AreaID As String = ""
    'Public TempFolder As String = "C:\ProductionReportLog\Temp\"
    'Public UploadFolder As String = "C:\ProductionReportLog\Uploaded\"
    Public ProcInfo As New Dictionary(Of String, String()) 'Key:報表名稱 ,{密碼,站點,Location,報表ID}
    Dim Requirement As New Dictionary(Of String, String()) 'Key:欄位名稱,{選必填,預設值}
    Dim isRefresh As Boolean = False '判定是否為是Timer刷新
    Dim Column_All As New List(Of String) '浮動欄位
    Dim Column As New List(Of String) '手動輸入欄位
    Dim Column_Query As New List(Of String) '自動帶入欄位
    Dim Column_Formula As New List(Of String) '運算欄位
    Dim Column_Formula_All As String '運算欄位
    Dim Column_CCF As String() = {"Trace Width(TOP)", "Trace Width(BOTTOM)", "Trace Space(TOP)", "Trace Space(BOTTOM)", "CuPad(TOP)", "CuPad(BOTTOM)"} '內層線路特殊欄位
    Dim Machine As String()
    Dim Cmd_Param As String = ""
    Dim Cmd_Formula As New List(Of String())
    Dim HistPath As String = "ProductionReportHist\ProductionReportHist.exe"
    Dim asr As New ProductionReport.My.MySettings
    Dim MFmodule As String = asr.Item("MFmodule")
    '11/06新增參數
    Dim CID As New Dictionary(Of String, String)

    '-----------------------------------DB參數----------------------------------------
    Dim DbVersion As String = "[Datamation_H3].[dbo].[H3_Leo_Program_Version]" '版本卡控DB
    Dim DbProc As String = "[H3_Systematic].[dbo].[H3_Proc]" '報表設定Config DB
    Dim DbProcParameter As String = "[H3_Systematic].[dbo].[H3_Production_ProcParameter]" '欄位設定Config DB
    Dim DbProcParameterRule As String = "[H3_Systematic].[dbo].[H3_Production_ProcParameter_Rule]" '欄位帶入設定Config DB
    'Dim SpFreeColumn As String = "[H3_Systematic].[dbo].[ProductionQuery_Columns_Insert]" '客製欄位建立&更新SP
    'Dim SpQuery As String = "[H3_Systematic].[dbo].[ProductionQuery_NEW]" '未完成&完成未超過12小時資料查詢SP
    Dim SpFixecColumnNew As String = "[H3_Systematic].[dbo].[ProductionQuery_Insert_New]" '固定欄位建立&更新SP(新的，只需要AreaID)
    Dim SpFreeColumn As String = "[H3_Systematic].[dbo].[ProductionQuery_Columns_Insert_New]" '客製欄位建立&更新SP
    Dim SpQuery As String = "[H3_Systematic].[dbo].[ProductionQuery]" '未完成&完成未超過12小時資料查詢SP
    Dim DbLog As String = "[H3_Systematic].[dbo].[H3_ProductionLog]" '固定欄位資料紀錄 DB
    Dim DbLogParameter As String = "[H3_Systematic].[dbo].[H3_ProductionParameter]" '客製化欄位資料紀錄 DB
    'Dim DbLog As String = "[H3_Systematic].[dbo].[H3_ProductionLog_NEW]" '固定欄位資料紀錄 DB
    'Dim DbLogParameter As String = "[H3_Systematic].[dbo].[H3_ProductionParameter_NEW]" '客製化欄位資料紀錄 DB

    Dim DbWIP As String = "[utchfacmrpt].[report].[dbo].[view_WIP]" '愉進WIP View表 ， 請注意各廠WIP表的欄位名稱是否一致，如果不同記得修改

    Private Sub ReportUI_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '標題初始化
        Me.Text = "ReportUI v" & Version & " [維運 : 李博軒]"

        '建立Log資料夾
        If Not Directory.Exists(LogFolder) Then
            Directory.CreateDirectory(LogFolder)
        End If
        'If Not Directory.Exists(TempFolder) Then
        '    Directory.CreateDirectory(TempFolder)
        'End If
        'If Not Directory.Exists(UploadFolder) Then
        '    Directory.CreateDirectory(UploadFolder)
        'End If
        '暫停刷新資料直到選擇站點
        TimerRefresh.Stop()
        '站點名稱搜尋
        AreaName.Items.Clear()

        Using dt As DataTable = SQL_Query("SELECT  [OnlineVersion],ISNULL([TestVersion],0) AS [TestVersion] FROM " & DbVersion & " WHERE [Program] = '" & Program & "'")
            If Not (CInt(Version.Replace(".", "")) > CInt(dt(0)(0).ToString.Replace(".", "")) OrElse CInt(Version.Replace(".", "")) = CInt(dt(0)(1).ToString.Replace(".", ""))) Then
                MessageBox.Show("請確認是否使用最新版本")
                Environment.Exit(Environment.ExitCode)
                Application.Exit()
            End If
        End Using

        Using dt As DataTable = SQL_Query("SELECT DISTINCT [Area],[Password],[ProcName],ISNULL([Location],'') AS [Location],[Pkey] FROM " & DbProc & " WITH (NOLOCK) WHERE [Module] = " & MFmodule & "　ORDER BY [Area]") '搜尋資料庫內的所有站點和密碼
            For Each row As DataRow In dt.Rows
                AreaName.Items.Add(row("Area")) '在站點選項添加對應項目
                Dim procinfo As String() = {row("Password"), row("ProcName"), row("Location"), row("Pkey")}
                Me.ProcInfo.Add(row("Area"), procinfo) '新增與站點對應密碼
            Next
        End Using

        'dgv雙重緩衝
        Dim type As Type = ReportUI_DataGridView.GetType()
        Dim pi As PropertyInfo = type.GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.NonPublic)
        pi.SetValue(ReportUI_DataGridView, True, Nothing)
    End Sub
    Private Sub AreaName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles AreaName.SelectedIndexChanged
        Try
            '清空資料
            If AreaName.SelectedItem <> Nothing AndAlso Area = "" Or Area <> AreaName.SelectedItem Then
                Dim pw As String = InputBox("請輸入密碼：", "輸入密碼", "")
                If pw <> ProcInfo(AreaName.SelectedItem.ToString)(0) Then
                    MessageBox.Show("密碼輸入錯誤")
                    AreaName.SelectedItem = Area
                    Return
                End If
                'isRefresh = True
                'DataGridView初始化
                ReportUI_DataGridView.DataSource = Nothing
                ReportUI_DataGridView.Rows.Clear()
                ReportUI_DataGridView.Columns.Clear()
                Column_All.Clear()
                Column.Clear()
                Column_Query.Clear()
                Column_Formula.Clear()
                Column_Formula_All = ""
                Cmd_Param = ""
                Cmd_Formula.Clear()
                Requirement.Clear()

                '查詢客製化欄位設定
                '20231016 新增查詢[EnglishName],[isRequire],[DefaultValues]
                Dim cmdProc As String = "SELECT pa.[Pkey] AS [CID]
                                                     ,pa.[PID] AS [Sort]
                                                     ,pa.[ParameterName]
                                                     ,pa.[EnglishName]
                                                     ,ISNULL(pa.[isQuery],'') AS [isQuery]
                                                     ,ISNULL(r.[QueryCommand],'') AS [QueryCommand]
                                                     ,ISNULL(r.[QueryType],'') AS [QueryType]
                                                     ,CASE WHEN [QueryType] = '欄位間計算' THEN ISNULL(r.[Filter1],'') ELSE '' END AS [FormulaColumn]
                                                     ,ISNULL(pr.[Machine],'') AS [Machine]
													 ,ISNULL(pa.[isRequire],1) AS [isRequire]
                                                     ,ISNULL(pa.[DefaultValues],'') AS [DefaultValues]
                                                     ,ISNULL(pr.[hasFace],'') AS [hasFace]
                                                     FROM " & DbProcParameter & " AS pa
                                                     LEFT JOIN " & DbProc & " AS pr WITH (NOLOCK) ON pr.[Pkey] = pa.[AreaID]
                                                     LEFT JOIN " & DbProcParameterRule & " AS r ON pa.[QID] = r.[QID]
                                                     WHERE pr.[Area] = '" & AreaName.SelectedItem.ToString & "'
                                                     ORDER BY [Sort]"
                Dim dtProc As DataTable = SQL_Query(cmdProc)

                '設定初始欄位
                Dim ColsOrigin() As String = {"日期; SystemTime", "班別; Class", "前站結束時間; Previous CheckOut", "開始時間; (CheckIn)", "結束時間; CheckOut", "料號; PartNo.", "批號; LotNo.", "層別; LayerName", "站點; ProcName", "機台; Machine", "面次; PF/PB", "產品類型; IType", "入料片數; Qnty_In", "出料片數; Qnty_Out", "過帳工號; WID", "過帳人員; User", "操作員; OP", "備註; Remark"}
                '生成模組通用欄位
                For Each cols As String In ColsOrigin
                    Dim col As String() = cols.Split(";")
                    ReportUI_DataGridView.Columns.Add(col(0), col(0) & vbCrLf & Trim(col(1)))
                Next
                For Each dgvCol As DataGridViewColumn In ReportUI_DataGridView.Columns
                    If dgvCol.Name <> "操作員" AndAlso dgvCol.Name <> "備註" Then
                        dgvCol.ReadOnly = True
                        dgvCol.DefaultCellStyle.BackColor = SystemColors.ControlLight
                    End If
                    dgvCol.SortMode = DataGridViewColumnSortMode.NotSortable '變更欄位不可排序
                    If dgvCol.Name = "班別" Then dgvCol.Width = 50
                    If dgvCol.Name = "前站結束時間" OrElse dgvCol.Name = "開始時間" OrElse dgvCol.Name = "結束時間" OrElse dgvCol.Name = "機台" Then dgvCol.Width = 110
                    If dgvCol.Name = "日期" Then
                        dgvCol.Width = 110
                        dgvCol.SortMode = DataGridViewColumnSortMode.Automatic
                    End If
                    If dgvCol.Name = "批號" Then dgvCol.Width = 100
                    If dgvCol.Name = "料號" Then dgvCol.Width = 80
                    If dgvCol.Name = "層別" Then dgvCol.Width = 70
                    If dgvCol.Name = "站點" Then dgvCol.Width = 70
                    If dgvCol.Name = "過帳工號" OrElse dgvCol.Name = "過帳人員" OrElse dgvCol.Name = "入料片數" OrElse dgvCol.Name = "出料片數" Then dgvCol.Width = 65
                    If dgvCol.Name = "操作員" Then dgvCol.Width = 60
                    If dgvCol.Name = "面次" Then dgvCol.Width = 60
                Next

                '取得設定篩選機台
                If dtProc.Rows.Count = 0 Then
                    Dim cmdMachine As String = "SELECT [Machine] 
                                     FROM " & DbProc & " WITH (NOLOCK)
                                     WHERE [Area] = '" & AreaName.SelectedItem.ToString & "'"
                    Dim dtMachine As DataTable = SQL_Query(cmdMachine)
                    Dim newMachine As String() = dtMachine(0)("Machine").ToString.Split(",")
                    ReDim Machine(newMachine.Length - 1)
                    Machine = newMachine
                Else
                    Dim newMachine As String() = dtProc(0)("Machine").ToString.Split(",")
                    ReDim Machine(newMachine.Length - 1)
                    Machine = newMachine
                End If


                '將手動和自動帶入參數分開儲存、建立
                For Each row As DataRow In dtProc.Rows
                    CID.Add(row("ParameterName").ToString, row("CID").ToString)
                    If Convert.ToBoolean(row("isQuery")) AndAlso row("QueryType") = "欄位間計算" Then '計算欄位
                        Column_Formula.Add(row("ParameterName").ToString) '儲存參數名稱
                        Dim strFormula() As String = {row("QueryCommand").ToString, row("FormulaColumn").ToString}
                        Cmd_Formula.Add(strFormula) '儲存計算欄位和公式
                        Column_Formula_All += row("FormulaColumn").ToString
                        Requirement.Add(row("ParameterName").ToString, {row("isRequire").ToString, row("DefaultValues").ToString}) '儲存欄位選必填&預設值資訊
                        '建立欄位
                        ReportUI_DataGridView.Columns.Add(row("ParameterName").ToString, row("ParameterName").ToString & vbCrLf & row("EnglishName").ToString)
                        ReportUI_DataGridView.Columns(row("ParameterName").ToString).ReadOnly = True
                        ReportUI_DataGridView.Columns(row("ParameterName").ToString).DefaultCellStyle.BackColor = SystemColors.ControlLight
                        Column_All.Add(row("ParameterName").ToString)
                    ElseIf Convert.ToBoolean(row("isQuery")) Then '自動代入欄位
                        Column_Query.Add(row("ParameterName").ToString)  '儲存參數名稱
                        Cmd_Param = Cmd_Param + row("QueryCommand").ToString + vbCrLf  'SQL語法串接
                        Requirement.Add(row("ParameterName").ToString, {row("isRequire").ToString, row("DefaultValues").ToString}) '儲存欄位選必填&預設值資訊
                        '建立欄位
                        ReportUI_DataGridView.Columns.Add(row("ParameterName").ToString, row("ParameterName").ToString & vbCrLf & row("EnglishName").ToString)
                        ReportUI_DataGridView.Columns(row("ParameterName").ToString).ReadOnly = True
                        ReportUI_DataGridView.Columns(row("ParameterName").ToString).DefaultCellStyle.BackColor = SystemColors.ControlLight
                        Column_All.Add(row("ParameterName").ToString)
                    Else '手動輸入欄位
                        Column.Add(row("ParameterName"))
                        Requirement.Add(row("ParameterName").ToString, {row("isRequire").ToString, row("DefaultValues").ToString}) '儲存欄位選必填&預設值資訊
                        '建立欄位
                        ReportUI_DataGridView.Columns.Add(row("ParameterName").ToString, row("ParameterName").ToString & vbCrLf & row("EnglishName").ToString)
                        Column_All.Add(row("ParameterName").ToString)
                    End If
                Next

                Dim btnCol As New DataGridViewButtonColumn()
                btnCol.HeaderText = "修改" & vbCrLf & "(Fix)"
                btnCol.Name = "btnModify"
                btnCol.Text = "修改"
                Dim bc As New DataGridViewButtonCell()
                bc.FlatStyle = FlatStyle.Flat
                bc.Style.BackColor = Color.Cyan
                btnCol.CellTemplate = bc
                ReportUI_DataGridView.Columns.Add(btnCol)
                ReportUI_DataGridView.Columns.Add("完成", "完成" & vbCrLf & "(Compelete)")
                ReportUI_DataGridView.Columns.Add("LogID", "LogID")
                ReportUI_DataGridView.Columns("btnModify").Width = 40
                ReportUI_DataGridView.Columns("完成").Width = 60
                ReportUI_DataGridView.Columns("完成").ReadOnly = True
                ReportUI_DataGridView.Columns("LogID").Width = 60
                ReportUI_DataGridView.Columns("LogID").ReadOnly = True
                ReportUI_DataGridView.Columns("LogID").Visible = False
                'For Each dgvCol As DataGridViewColumn In ReportUI_DataGridView.Columns
                '    'dgvCol.SortMode = DataGridViewColumnSortMode.NotSortable
                '    If dgvCol.HeaderText = "修改" Then dgvCol.Width = 40
                '    If dgvCol.HeaderText = "完成" Then
                '        dgvCol.Width = 60
                '        dgvCol.ReadOnly = True
                '    End If
                'Next

                Area = AreaName.SelectedItem '報表名稱變數
                AreaID = ProcInfo(AreaName.SelectedItem.ToString)(3) '報表ID變數
            Else
                Return
            End If
            ReportUI_DataGridView.Columns(10).Frozen = True '凍結欄位
            SAP_CheckID(AreaID)

            TimerRefresh.Start()
            TimerRefresh_Tick(sender, e)
        Catch ex As Exception
            WriteLog(ex, LogFilePath, "AreaName_SelectedIndexChanged")
        End Try
    End Sub

    '20231017 Boris 修正預存程序查詢語法
    Private Sub TimerRefresh_Tick(sender As Object, e As EventArgs) Handles TimerRefresh.Tick
        Try
            'isRefresh = True
            TimerRefresh.Stop()
            'Dim proc As String() = ProcInfo(AreaName.SelectedItem.ToString)(1).Split(",")
            Dim StrMachine As String = ""
            For Each mac In Machine
                StrMachine += mac + ","
            Next
            StrMachine = StrMachine.Substring(0, StrMachine.Length - 1)
            Dim Parameters As New List(Of String())
            'For i = 0 To proc.Length - 1
            '    Try
            '        'Dim cmdproc As String = "SET ARITHABORT ON EXECUTE [H3_Systematic].[dbo].[ProductionQuery_Insert] @ProcName = '" & proc(i).Substring(0, 3) & "%" + proc(i).Substring(3, 3) & "%', @Location = '" & ProcInfo(AreaName.SelectedItem.ToString)(2) & "%', @AreaID = " & AreaID & ", @Machine = '" & StrMachine & "'"
            '        'SQL_Query(cmdproc)
            '        Parameters.Clear()
            '        Parameters.Add({"@ProcName", proc(i).Substring(0, 3) & "%" + proc(i).Substring(3, 3) & "%"})
            '        Parameters.Add({"@Location", ProcInfo(AreaName.SelectedItem.ToString)(2) & "%"})
            '        Parameters.Add({"@AreaID", AreaID})
            '        Parameters.Add({"@Machine", StrMachine})
            '        SQL_StoredProcedure(SpFixedColumn, Parameters)
            '    Catch ex As SqlException
            '    End Try
            'Next

            'Dim cmdcol As String = "EXECUTE [H3_Systematic].[dbo].[ProductionQuery_Columns_Insert] @AreaID = '" & AreaID & "'"
            Parameters.Clear()
            Parameters.Add({"@AreaID", AreaID})

            Try
                SQL_StoredProcedure(SpFixecColumnNew, Parameters)
            Catch ex As SqlException
                SQL_StoredProcedure(SpFixecColumnNew, Parameters)
            End Try

            Try
                'SQL_Query(cmdcol)
                SQL_StoredProcedure(SpFreeColumn, Parameters)
            Catch ex As SqlException
                'SQL_Query(cmdcol)
                SQL_StoredProcedure(SpFreeColumn, Parameters)
            End Try
            'Dim cmdquery As String = "EXECUTE [H3_Systematic].[dbo].[ProductionQuery_NEW] @AreaID = '" & AreaID & "'"
            'Dim new_dt As DataTable = SQL_Query(cmdquery)
            Dim new_dt As DataTable = SQL_StoredProcedure(SpQuery, Parameters)
            Dim lots As New List(Of String)



            ' 將新查詢到的資料與舊的資料合併，並只新增不重複的資料
            For Each new_dr As DataRow In new_dt.Rows
                'If Machine(0) <> "" AndAlso new_dr("機台").ToString <> "" AndAlso Not Machine.Contains(new_dr("機台").ToString) Then
                '    Continue For
                'End If
                lots.Add(new_dr("批號").ToString)
                Dim RIndex As Integer
                Dim is_duplicate As Boolean = False
                For Each dr As DataGridViewRow In ReportUI_DataGridView.Rows
                    If new_dr("批號") = dr.Cells("批號").Value.ToString AndAlso new_dr("站點") = dr.Cells("站點").Value.ToString AndAlso new_dr("層別") = dr.Cells("層別").Value.ToString AndAlso new_dr("面次") = dr.Cells("面次").Value.ToString Then
                        is_duplicate = True
                        RIndex = dr.Index
                        Exit For
                    End If
                Next
                If Not is_duplicate Then
                    Dim MoveInTime As String = ""
                    Dim LastEndTime As String = ""
                    Dim CheckInTime As String = ""
                    Dim CheckOutTime As String = ""
                    If CType(new_dr("日期"), DateTime) > DateTime.Parse("2000-01-01 00:00:00") Then MoveInTime = CType(new_dr("日期"), DateTime).ToString("yyyy-MM-dd HH:mm:ss")
                    If CType(new_dr("前站結束時間"), DateTime) > DateTime.Parse("1900-01-03 00:00:00") Then
                        LastEndTime = CType(new_dr("前站結束時間"), DateTime).ToString("yyyy-MM-dd HH:mm:ss")
                    ElseIf CType(new_dr("前站結束時間"), DateTime) > DateTime.Parse("1900-01-01 23:00:00") Then
                        LastEndTime = "首件"
                    End If
                    If CType(new_dr("開始時間"), DateTime) > DateTime.Parse("2000-01-01 00:00:00") Then
                        CheckInTime = CType(new_dr("開始時間"), DateTime).ToString("yyyy-MM-dd HH:mm:ss")
                    End If
                    If CType(new_dr("結束時間"), DateTime) > DateTime.Parse("2000-01-01 00:00:00") Then CheckOutTime = CType(new_dr("結束時間"), DateTime).ToString("yyyy-MM-dd HH:mm:ss")
                    Dim para As New List(Of Object)
                    Dim ParaArray() As Object = {MoveInTime, new_dr("班別"), LastEndTime, CheckInTime, CheckOutTime, new_dr("料號"), new_dr("批號"), new_dr("層別"), new_dr("站點"), new_dr("機台"), new_dr("面次"), new_dr("產品類型"), new_dr("入料片數"), new_dr("出料片數"), new_dr("過帳工號"), new_dr("過帳人員"), new_dr("操作員").ToString, new_dr("備註").ToString}
                    para.AddRange(ParaArray)
                    For Each col As String In Column_All
                        para.Add(new_dr(col))
                    Next
                    para.Add("")
                    If new_dr("已上傳").ToString = "False" Then
                        para.Add("")
                    Else
                        para.Add(CType(new_dr("完成時間"), DateTime).ToString("yyyy-MM-dd HH:mm:ss"))
                    End If
                    para.Add(new_dr("LogID"))

                    ReportUI_DataGridView.Rows.Insert(0, para.ToArray)
                    If new_dr("面次").ToString = "PB" Then
                        ReportUI_DataGridView.Rows(0).Cells("操作員").ReadOnly = True
                        ReportUI_DataGridView.Rows(0).Cells("操作員").Style.BackColor = SystemColors.ControlLight
                    End If
                    'If new_dr("操作員").ToString <> "" Then
                    '    ReportUI_DataGridView.Rows(0).Cells("操作員").ReadOnly = True
                    '    ReportUI_DataGridView.Rows(0).Cells("操作員").Style.BackColor = SystemColors.ControlLight
                    'End If
                    If new_dr("已上傳").ToString = "True" Then
                        ReportUI_DataGridView.Rows(0).Cells("完成").Style.BackColor = Color.Lime
                    End If
                Else
                    Dim MoveInTime As String = ""
                    Dim LastEndTime As String = ""
                    Dim CheckInTime As String = ""
                    Dim CheckOutTime As String = ""
                    If CType(new_dr("日期"), DateTime) > DateTime.Parse("2000-01-01 00:00:00") Then MoveInTime = CType(new_dr("日期"), DateTime).ToString("yyyy-MM-dd HH:mm:ss")
                    If CType(new_dr("前站結束時間"), DateTime) > DateTime.Parse("1900-01-03 00:00:00") Then
                        LastEndTime = CType(new_dr("前站結束時間"), DateTime).ToString("yyyy-MM-dd HH:mm:ss")
                    ElseIf CType(new_dr("前站結束時間"), DateTime) > DateTime.Parse("1900-01-01 23:00:00") Then
                        LastEndTime = "首件"
                    End If
                    If CType(new_dr("開始時間"), DateTime) > DateTime.Parse("2000-01-01 00:00:00") Then
                        CheckInTime = CType(new_dr("開始時間"), DateTime).ToString("yyyy-MM-dd HH:mm:ss")
                    End If
                    If CType(new_dr("結束時間"), DateTime) > DateTime.Parse("2000-01-01 00:00:00") Then CheckOutTime = CType(new_dr("結束時間"), DateTime).ToString("yyyy-MM-dd HH:mm:ss")
                    Dim para As New List(Of Object)
                    Dim ParaArray() As Object = {MoveInTime, new_dr("班別"), LastEndTime, CheckInTime, CheckOutTime, new_dr("料號"), new_dr("批號"), new_dr("層別"), new_dr("站點"), new_dr("機台"), new_dr("面次"), new_dr("產品類型"), new_dr("入料片數"), new_dr("出料片數"), new_dr("過帳工號"), new_dr("過帳人員"), new_dr("操作員").ToString, new_dr("備註").ToString}
                    para.AddRange(ParaArray)
                    For Each col As String In Column_All
                        para.Add(new_dr(col))
                    Next
                    para.Add("")
                    If new_dr("已上傳").ToString = "False" Then
                        para.Add("")
                    Else
                        para.Add(CType(new_dr("完成時間"), DateTime).ToString("yyyy-MM-dd HH:mm:ss"))
                    End If
                    para.Add(new_dr("LogID"))

                    ReportUI_DataGridView.Rows.RemoveAt(RIndex)
                    ReportUI_DataGridView.Rows.Insert(RIndex, para.ToArray)
                    If new_dr("面次").ToString = "PB" Then
                        ReportUI_DataGridView.Rows(RIndex).Cells("操作員").ReadOnly = True
                        ReportUI_DataGridView.Rows(RIndex).Cells("操作員").Style.BackColor = SystemColors.ControlLight
                    End If
                    'If new_dr("操作員").ToString <> "" Then
                    '    ReportUI_DataGridView.Rows(RIndex).Cells("操作員").ReadOnly = True
                    '    ReportUI_DataGridView.Rows(RIndex).Cells("操作員").Style.BackColor = SystemColors.ControlLight
                    'End If
                    If new_dr("已上傳").ToString = "True" Then
                        ReportUI_DataGridView.Rows(RIndex).Cells("完成").Style.BackColor = Color.Lime
                    End If
                End If
            Next
            DataGridViewRefresh(new_dt, lots)
            new_dt.Dispose()
            'ExportDataGridViewToDataTable(ReportUI_DataGridView)
            ' 恢復 DataGridView 的更新
            'isRefresh = False
            CheckColumn()
            ReportUI_DataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader
            ReportUI_DataGridView.ResumeLayout()
            ReportUI_DataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            TimerRefresh.Start()
        Catch ex As Exception
            WriteLog(ex, LogFilePath, "TimerRefresh")
            'isRefresh = False
            TimerRefresh.Start()
        End Try
    End Sub

    Private Sub ModifyButton_Click(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs)
        If ReportUI_DataGridView.Rows(e.RowIndex).Cells("完成").Value IsNot Nothing AndAlso ReportUI_DataGridView.Rows(e.RowIndex).Cells("完成").Value.ToString <> "" Then
            Return
        End If

        ' 將該行特定欄位 ReadOnly 屬性改為 False
        For i = ReportUI_DataGridView.Columns("備註").Index To ReportUI_DataGridView.Columns("btnModify").Index - 1
            '判斷是否為手動欄位或是H3內層線路特別欄位
            If Column.Contains(ReportUI_DataGridView.Columns(i).Name) OrElse Column_CCF.Contains(ReportUI_DataGridView.Columns(i).Name) OrElse (ReportUI_DataGridView.Rows(e.RowIndex).Cells(i).Value IsNot Nothing AndAlso ReportUI_DataGridView.Rows(e.RowIndex).Cells(i).Value.ToString = "") Then
                ReportUI_DataGridView.Rows(e.RowIndex).Cells(i).ReadOnly = False
                ReportUI_DataGridView.Rows(e.RowIndex).Cells(i).Style.BackColor = SystemColors.ControlLightLight
            End If
        Next
        ReportUI_DataGridView.Rows(e.RowIndex).Cells("操作員").ReadOnly = False
        ReportUI_DataGridView.Rows(e.RowIndex).Cells("操作員").Style.BackColor = SystemColors.ControlLightLight
    End Sub

    Private Sub CheckParaUploaded(ByVal row As DataGridViewRow)
        For i = 0 To ReportUI_DataGridView.Columns("完成").Index - 1
            row.Cells(i).ReadOnly = True
            row.Cells(i).Style.BackColor = Color.Cyan
        Next
    End Sub

    '20231017 Boris 取消ReadOnly
    Private Sub CheckColumn()
        For Each row As DataGridViewRow In ReportUI_DataGridView.Rows
            If row.Cells("完成").Value IsNot Nothing AndAlso row.Cells("完成").Value.ToString <> "" Then
                CheckParaUploaded(row)
            Else
                'For i = ReportUI_DataGridView.Columns("面次").Index + 1 To ReportUI_DataGridView.Columns("btnModify").Index - 1
                '    If row.Cells(i).Value IsNot Nothing AndAlso row.Cells(i).Value.ToString <> "" Then
                '        row.Cells(i).ReadOnly = True
                '        row.Cells(i).Style.BackColor = SystemColors.ControlLight
                '    End If
                'Next
            End If
        Next
    End Sub

    Private Sub Btn_TEST_Click(sender As Object, e As EventArgs) Handles Btn_TEST.Click
        TimerRefresh_Tick(sender, e)
    End Sub

    '20231017 Boris 取消ReadOnly
    Private Sub ReportUI_DataGridView_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles ReportUI_DataGridView.CellValueChanged
        Try
            Dim row As DataGridViewRow = ReportUI_DataGridView.Rows(e.RowIndex)
            If row.Cells("完成").Value IsNot Nothing AndAlso row.Cells("完成").Value <> "" Then
                Exit Sub
            End If
            If row.Cells("站點").Value IsNot Nothing AndAlso row.Cells("站點").Value.ToString() <> "" AndAlso row.Cells("批號").Value IsNot Nothing AndAlso row.Cells("批號").Value.ToString() <> "" AndAlso row.Cells("層別").Value IsNot Nothing AndAlso row.Cells("層別").Value.ToString() <> "" AndAlso row.Cells("料號").Value IsNot Nothing AndAlso row.Cells("料號").Value.ToString() <> "" Then
                Dim part As String = row.Cells("料號").Value.ToString
                Dim proc As String = row.Cells("站點").Value.ToString
                Dim lot As String = row.Cells("批號").Value.ToString
                Dim layer As String = row.Cells("層別").Value.ToString

                Dim ParaName As String = ReportUI_DataGridView.Columns(e.ColumnIndex).Name

                If ParaName = "站點" OrElse ParaName = "料號" OrElse ParaName = "批號" OrElse ParaName = "層別" AndAlso row.Cells("前站結束時間").Value.ToString = "" Then
                    SAP_First_Upload(row, AreaID)
                    Return
                End If

                Dim PID As String = row.Cells("LogID").Value.ToString

                If Not isRefresh Then
                    Dim cell As DataGridViewCell = row.Cells(e.ColumnIndex)
                    Dim count As String
                    If row.Cells("前站結束時間").Value.ToString = "首件" Then
                        count = "1"
                    Else
                        count = "0"
                    End If

                    If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
                        If Cmd_Formula.Count > 0 AndAlso Column_Formula_All.Contains(ReportUI_DataGridView.Columns(cell.ColumnIndex).Name) Then
                            ColumnFormula(row, cell)
                        End If
                    End If
                    SAP_CheckPnl(row, e, AreaID)
                    If cell.Value IsNot Nothing AndAlso cell.Value.ToString <> "" AndAlso cell.ColumnIndex >= ReportUI_DataGridView.Columns("操作員").Index Then

                        If cell.ColumnIndex = ReportUI_DataGridView.Columns("操作員").Index Then
                            ' 設定資料格為唯讀
                            'cell.ReadOnly = True
                            'cell.Style.BackColor = SystemColors.ControlLight
                            'Dim cmd As String = "UPDATE " & DbLog & "
                            '                                      SET [OP] = '" + Trim(cell.Value.ToString) + "'
                            '                                      WHERE [AreaID] = " + AreaID + " AND [ProcName] = '" + proc + "' AND [Lotnum] = '" + lot + "' AND [Layer] = '" + layer + "' AND [Count] = " + count
                            Dim cmd As String = "UPDATE " & DbLog & "
                                                                  SET [OP] = '" + Trim(cell.Value.ToString) + "'
                                                                  WHERE [Pkey] = " + PID

                            SQL_Query(cmd)
                        ElseIf cell.ColumnIndex <> ReportUI_DataGridView.Columns("備註").Index Then
                            ' 設定資料格為唯讀
                            'cell.ReadOnly = True
                            'cell.Style.BackColor = SystemColors.ControlLight

                            '確認面次
                            Dim face As String = ""
                            If ReportUI_DataGridView.Rows(cell.RowIndex).Cells("面次").Value.ToString = "PF" Then
                                face = "1"
                            ElseIf ReportUI_DataGridView.Rows(cell.RowIndex).Cells("面次").Value.ToString = "PB" Then
                                face = "2"
                            Else
                                face = "0"
                            End If
                            '20231012 修改可覆寫資料
                            Dim cmd As String = " UPDATE " & DbLogParameter & "
                                                                  SET [UploadTime] = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',
                                                                          [ParameterVaules] = '" + Trim(cell.Value.ToString) + "'
                                                                  WHERE [PID] = " + PID + " AND [CID] = '" + CID(ParaName) + "' AND [Face] = '" + face + "' AND [Count] = " + count + ""
                            'Dim cmd As String = " UPDATE " & DbLogParameter & "
                            '                                      SET [UploadTime] = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',
                            '                                              [ParameterVaules] = '" + Trim(cell.Value.ToString) + "'
                            '                                      WHERE [AreaID] = " + AreaID + " AND [ProcName] = '" + proc + "' AND [Lotnum] = '" + lot + "' AND [LayerName] = '" + layer + "' AND [ParameterName] = '" + ParaName + "' AND [Count] = " + count + ""

                            'Dim cmd As String = "DECLARE @Val nvarchar(MAX)
                            '                                      SET @Val = (SELECT ISNULL([ParameterVaules],'') AS Val FROM " & DbLogParameter & " WITH (NOLOCK) WHERE [AreaID] = " + AreaID + " AND [ProcName] = '" + proc + "' AND [Lotnum] = '" + lot + "' AND [LayerName] = '" + layer + "' AND [ParameterName] = '" + ParaName + "'  AND [Count] = " + count + ")
                            '                                      IF @Val IS NULL OR @Val = ''
                            '                                      BEGIN
                            '                                      UPDATE " & DbLogParameter & "
                            '                                      SET [UploadTime] = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',
                            '                                              [ParameterVaules] = '" + cell.Value.ToString + "'
                            '                                      WHERE [AreaID] = " + AreaID + " AND [ProcName] = '" + proc + "' AND [Lotnum] = '" + lot + "' AND [LayerName] = '" + layer + "' AND [ParameterName] = '" + ParaName + "' AND [Count] = " + count + "
                            '                                      END"
                            SQL_Query(cmd)
                            Else
                                'Dim cmd As String = "UPDATE " & DbLog & "
                                '                                      SET [Remark] = '" + Trim(cell.Value.ToString) + "'
                                '                                      WHERE [AreaID] = " + AreaID + " AND [ProcName] = '" + proc + "' AND [Lotnum] = '" + lot + "' AND [Layer] = '" + layer + "' AND [Count] = " + count
                                Dim cmd As String = "UPDATE " & DbLog & "
                                                                  SET [Remark] = '" + Trim(cell.Value.ToString) + "'
                                                                  WHERE [Pkey] = " + PID

                            SQL_Query(cmd)
                        End If
                    End If
                End If
            Else
                If ReportUI_DataGridView.Columns(e.ColumnIndex).Name <> "料號" AndAlso ReportUI_DataGridView.Columns(e.ColumnIndex).Name <> "批號" AndAlso ReportUI_DataGridView.Columns(e.ColumnIndex).Name <> "層別" AndAlso ReportUI_DataGridView.Columns(e.ColumnIndex).Name <> "站點" Then
                    MessageBox.Show("請將""料號""、""批號""、""層別""、""站點""填寫完畢")
                End If
            End If
        Catch ex As Exception
            WriteLog(ex, LogFilePath, "ReportUI_DataGridView_CellValueChanged_" + ReportUI_DataGridView.Columns(e.ColumnIndex).Name + "_" + ReportUI_DataGridView.Rows(e.RowIndex).Cells("批號").Value.ToString)
        End Try
    End Sub

    Private Sub RrportUI_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        e.Cancel = True ' 取消視窗關閉操作
    End Sub

    Private Sub DataGridViewRefresh(ByVal new_dt As DataTable, ByVal lots As List(Of String))
        Try
            '查詢WIP
            ' Dim proc As String() = ProcInfo(AreaName.SelectedItem.ToString)(1).Split(",")
            Dim wip_dt As New DataTable
            'For i = 0 To proc.Length - 1
            Dim cmdWIP As String = "  SELECT [CurrProcName] AS [站點],[LotNum] AS [批號],[LayerName] AS [層別]
                                                                    FROM " & DbWIP & " AS w WITH(NOLOCK),
                                                                    " & DbProc & " AS p WITH(NOLOCK)
                                                                    WHERE [LineID] = 29 AND p.[Pkey] = " + AreaID + " AND [ProcName] LIKE SUBSTRING([CurrProcName],1,3) + '%' + SUBSTRING([CurrProcName],5,3)"
            wip_dt = SQL_Query(cmdWIP)
            'Next

            ' 比較 DataTable 中的每一行，更新 DataGridView 中的值
            For Each row As DataGridViewRow In ReportUI_DataGridView.Rows
                If row.Cells("站點").Value IsNot Nothing AndAlso row.Cells("站點").Value.ToString() <> "" AndAlso row.Cells("批號").Value IsNot Nothing AndAlso row.Cells("批號").Value.ToString() <> "" AndAlso row.Cells("層別").Value IsNot Nothing AndAlso row.Cells("層別").Value.ToString() <> "" AndAlso row.Cells("料號").Value IsNot Nothing AndAlso row.Cells("料號").Value.ToString() <> "" Then
                    Dim proc_dgv As String = row.Cells("站點").Value.ToString()
                    Dim lot_dgv As String = row.Cells("批號").Value.ToString()
                    Dim layer_dgv As String = row.Cells("層別").Value.ToString()

                    '確認每一筆DataGridView上的資料是否包含在最新查詢出來的Table
                    If lots.Contains(lot_dgv) Then
                        '確認物料是否還在WIP
                        Dim ExistsWIP As Boolean = False
                        For Each wipRow As DataRow In wip_dt.Rows
                            If proc_dgv = wipRow("站點").ToString AndAlso lot_dgv = wipRow("批號").ToString AndAlso layer_dgv = wipRow("層別").ToString Then
                                ExistsWIP = True
                                Exit For
                            End If
                        Next

                        If (row.Cells("完成").Value Is Nothing OrElse row.Cells("完成").Value.ToString = "") AndAlso Cmd_Formula.Count > 0 Then
                            For i = 0 To Column_Formula.Count - 1
                                If row.Cells(Column_Formula(i)).Value Is Nothing OrElse row.Cells(Column_Formula(i)).Value.ToString = "" Then
                                    Dim formulacol As String() = Cmd_Formula(i)(1).Split(",")
                                    Dim Notyet As Boolean = False

                                    For Each col In formulacol
                                        If row.Cells(col).Value Is Nothing OrElse String.IsNullOrEmpty(row.Cells(col).Value.ToString()) OrElse (col = "入料片數" And row.Cells(col).Value.ToString = 0) Then
                                            Notyet = True
                                            Exit For
                                        Else
                                            Notyet = False
                                        End If
                                    Next
                                    If Notyet Then
                                        Exit For
                                    Else
                                        Dim cmd As String = Cmd_Formula(i)(0)
                                        For j = formulacol.Count To 1 Step -1
                                            cmd = cmd.Replace("var" + j.ToString, row.Cells(formulacol(j - 1)).Value.ToString)
                                        Next

                                        Dim result As Double
                                        result = Math.Round(ReCoding(cmd), 4)
                                        row.Cells(Column_Formula(i)).Value = result
                                    End If
                                End If
                            Next
                        End If

                        '若不在WIP則確認資料是否完整
                        If Not ExistsWIP Then
                            Dim isUpload As Boolean = False
                            isUpload = DataUpload(row)
                            '資料如果完整則Show提示字元
                            If isUpload Then
                                Notice.Text = "上傳完成,批號: " & lot_dgv & ", 時間: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                            End If
                        End If
                    Else
                        '不包含在新的查詢結果則刪除此行
                        ReportUI_DataGridView.Rows.Remove(row)
                        Continue For
                    End If
                End If
            Next

        Catch ex As Exception
            WriteLog(ex, LogFilePath, "DataGridViewRefresh")
        End Try
    End Sub

    Private Function DataUpload(row As DataGridViewRow) As Boolean
        Try
            If Not (row.Cells("完成").Value IsNot Nothing AndAlso row.Cells("完成").Value.ToString() <> "") Then
                If row.Cells("站點").Value IsNot Nothing AndAlso row.Cells("站點").Value.ToString() <> "" AndAlso row.Cells("批號").Value IsNot Nothing AndAlso row.Cells("批號").Value.ToString() <> "" AndAlso row.Cells("層別").Value IsNot Nothing AndAlso row.Cells("層別").Value.ToString() <> "" AndAlso row.Cells("料號").Value IsNot Nothing AndAlso row.Cells("料號").Value.ToString() <> "" Then
                    Dim lot As String = row.Cells("批號").Value.ToString()
                    Dim proc As String = row.Cells("站點").Value.ToString()
                    Dim layer As String = row.Cells("層別").Value.ToString()
                    Dim PID As String = row.Cells("LogID").Value.ToString()
                    Dim Face As String = row.Cells("面次").Value.ToString()
                    Dim NextFace As String = ""
                    If Face = "PF" Then
                        NextFace = "PB"
                    ElseIf Face = "PB" Then
                        NextFace = "PF"
                    End If
                        Dim isFullData As Boolean = False
                    Dim now As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                    If row.Cells("操作員").Value IsNot Nothing AndAlso row.Cells("操作員").Value.ToString() <> "" Then
                        For i = ReportUI_DataGridView.Columns("備註").Index + 1 To ReportUI_DataGridView.Columns("btnModify").Index - 1 Step 1
                            If Convert.ToBoolean(Requirement(ReportUI_DataGridView.Columns(i).Name)(0)) Then
                                If row.Cells(i).Value IsNot Nothing AndAlso row.Cells(i).Value.ToString() <> "" Then
                                    isFullData = True
                                Else
                                    isFullData = False
                                    Exit For
                                End If
                            End If
                        Next
                    Else
                        isFullData = False
                    End If
                    '確認正反面資料
                    If isFullData AndAlso row.Cells("結束時間").Value.ToString <> "" AndAlso Face.Contains("P") Then
                        Dim row2 As New DataGridViewRow
                        For Each ro As DataGridViewRow In ReportUI_DataGridView.Rows
                            If ro.Cells("LogID").Value.ToString = PID AndAlso ro.Cells("面次").Value.ToString = NextFace Then
                                row2 = ro
                                Exit For
                            End If
                        Next
                        For i = ReportUI_DataGridView.Columns("備註").Index + 1 To ReportUI_DataGridView.Columns("btnModify").Index - 1 Step 1
                            If Convert.ToBoolean(Requirement(ReportUI_DataGridView.Columns(i).Name)(0)) Then
                                If row2.Cells(i).Value IsNot Nothing AndAlso row2.Cells(i).Value.ToString() <> "" Then
                                    isFullData = True
                                Else
                                    isFullData = False
                                    Exit For
                                End If
                            End If
                        Next
                    End If

                    '確認所有可填欄位接已填寫後變更資料Flag
                    If isFullData AndAlso row.Cells("結束時間").Value.ToString <> "" Then
                        Dim remark As String = ""
                        If row.Cells("備註").Value <> Nothing Then remark = row.Cells("備註").Value.ToString
                        Dim cmd As String = "UPDATE " & DbLog & "
                                                          SET [UploadTime] = '" + now + "',
                                                              [Flag] = 1,
                                                              [Version] = '" + Version + "'
                                                          WHERE [Pkey] = " + PID
                        SQL_Query(cmd)

                        row.Cells("完成").Value = now
                        For i = 0 To ReportUI_DataGridView.Columns("完成").Index - 1
                            row.Cells(i).ReadOnly = True
                            row.Cells(i).Style.BackColor = Color.Cyan
                        Next
                        row.Cells("完成").Style.BackColor = Color.Lime
                        Return True
                    Else
                        Return False
                    End If
                End If
            End If
            Return False
        Catch ex As Exception
            WriteLog(ex, LogFilePath, "DataUpload")
            Return False
        End Try
    End Function

    Private Sub ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Query.Click
        Dim check() As Process = Process.GetProcessesByName("ProductionReportHist")
        If check.Length = 0 Then
            Dim ProcessInfo As New ProcessStartInfo(HistPath)
            Process.Start(ProcessInfo)
        End If
    End Sub

    Private Sub ColumnFormula(ByVal row As DataGridViewRow, ByVal Cell As DataGridViewCell)
        Try
            'isRefresh = False
            Dim Notyet As Boolean = False
            Dim tarcol As String = ""
            Dim formulacol As String()
            Dim cmd As String
            Dim index As Integer = 0
            For Each cols In Cmd_Formula
                If cols(1).Contains(ReportUI_DataGridView.Columns(Cell.ColumnIndex).Name) Then
                    formulacol = cols(1).Split(",")
                    For Each col In formulacol
                        If row.Cells(col).Value Is Nothing OrElse String.IsNullOrEmpty(row.Cells(col).Value.ToString()) OrElse (col = "入料片數" And row.Cells(col).Value.ToString = 0) Then
                            Notyet = True
                            Exit For
                        Else
                            Notyet = False
                        End If
                    Next
                    If Notyet Then
                        Exit For
                    Else
                        tarcol = Column_Formula(index)
                        cmd = cols(0)
                        For i = formulacol.Count To 1 Step -1
                            cmd = cmd.Replace("var" + i.ToString, row.Cells(formulacol(i - 1)).Value.ToString)
                        Next

                        Dim result As Double
                        result = Math.Round(ReCoding(cmd), 4)
                        row.Cells(tarcol).Value = result
                        'ColumnFormula(row.Cells(tarcol))
                        Exit For
                    End If

                    index += 1
                Else
                    index += 1
                End If
            Next
        Catch ex As Exception
            WriteLog(ex, LogFilePath, "ColumnFormula")
        End Try
    End Sub

    Private Sub Btn_First_Click(sender As Object, e As EventArgs) Handles Btn_First.Click
        Try
            Dim count As Integer = ReportUI_DataGridView.Rows.Count
            ReportUI_DataGridView.Rows.Add("", "", "首件", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            For i = 0 To ReportUI_DataGridView.Columns("備註").Index - 2
                If ReportUI_DataGridView.Columns(i).Name <> "前站結束時間" Then
                    ReportUI_DataGridView.Rows(count).Cells(i).ReadOnly = False
                    ReportUI_DataGridView.Rows(count).Cells(i).Style.BackColor = SystemColors.ControlLightLight
                End If
            Next
        Catch ex As Exception
            WriteLog(ex, LogFilePath, "Btn_First_Click")
        End Try
    End Sub

    Private Sub ReportUI_DataGridView_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles ReportUI_DataGridView.CellClick
        Try
            If e.RowIndex >= 0 And e.ColumnIndex >= 0 Then
                ' 確認是在 btncol 欄位點擊
                If e.ColumnIndex = ReportUI_DataGridView.Columns("btnModify").Index Then
                    ' 執行 ModifyButton_Click 事件
                    ModifyButton_Click(sender, e)
                Else
                    SAP_CheckPnl(ReportUI_DataGridView.Rows(e.RowIndex), e, AreaID)
                End If
            End If
        Catch ex As Exception
            WriteLog(ex, LogFilePath, "ReportUI_DataGridView_CellClick")
        End Try
    End Sub

    Private Sub TxtLot_KeyUp(sender As Object, e As KeyEventArgs) Handles TxtLot.KeyUp
        Try
            Dim isMatch As Boolean = False
            Dim MatchCount As Integer = 0
            If e.KeyData = Keys.Enter Then
                If TxtLot.Text.Length >= 14 AndAlso TxtLot.Text.Length <= 18 Then
                    TimerRefresh.Stop()
                    For Each row As DataGridViewRow In ReportUI_DataGridView.Rows
                        If row.Cells("批號").Value IsNot Nothing AndAlso row.Cells("批號").Value.ToString Like "*" + Trim(TxtLot.Text.Substring(0, 14)) + "*" AndAlso Not isMatch Then
                            ReportUI_DataGridView.Rows.RemoveAt(row.Index)
                            ReportUI_DataGridView.Rows.Insert(0, row)
                            isMatch = True
                            MatchCount += 1
                        Else
                            isMatch = False
                        End If
                    Next
                    If MatchCount = 0 Then
                        MessageBox.Show("未偵測到目標批號 : " + TxtLot.Text)
                    Else
                        MessageBox.Show("已將批號" + TxtLot.Text + "資料 " + MatchCount.ToString + " 筆移至資料列頂端")
                    End If
                Else
                    TxtLot.Text = ""
                End If
            End If
            TimerRefresh.Start()
        Catch ex As Exception
            WriteLog(ex, LogFilePath, "TxtLot_KeyUp")
            TimerRefresh.Start()
        End Try
    End Sub


    'Private Sub ReportUI_DataGridView_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles ReportUI_DataGridView.CellEndEdit
    '    Try
    '        Dim row As DataGridViewRow = ReportUI_DataGridView.Rows(e.RowIndex)
    '        If row.Cells("站點").Value IsNot Nothing AndAlso row.Cells("站點").Value.ToString() <> "" AndAlso row.Cells("批號").Value IsNot Nothing AndAlso row.Cells("批號").Value.ToString() <> "" AndAlso row.Cells("層別").Value IsNot Nothing AndAlso row.Cells("層別").Value.ToString() <> "" AndAlso row.Cells("料號").Value IsNot Nothing AndAlso row.Cells("料號").Value.ToString() <> "" Then
    '            Dim part As String = row.Cells("料號").Value.ToString
    '            Dim proc As String = row.Cells("站點").Value.ToString
    '            Dim lot As String = row.Cells("批號").Value.ToString
    '            Dim layer As String = row.Cells("層別").Value.ToString
    '            Dim ParaName As String = ReportUI_DataGridView.Columns(e.ColumnIndex).Name

    '            If ParaName = "站點" OrElse ParaName = "料號" OrElse ParaName = "批號" OrElse ParaName = "層別" AndAlso row.Cells("前站結束時間").Value.ToString = "" Then
    '                SAP_First_Upload(row, AreaID)
    '                Return
    '            End If

    '            If Not isRefresh Then
    '                Dim cell As DataGridViewCell = row.Cells(e.ColumnIndex)
    '                Dim count As String
    '                If row.Cells("前站結束時間").Value.ToString = "首件" Then
    '                    count = "1"
    '                Else
    '                    count = "0"
    '                End If

    '                If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
    '                    If Cmd_Formula.Count > 0 AndAlso Column_Formula_All.Contains(ReportUI_DataGridView.Columns(cell.ColumnIndex).Name) Then
    '                        ColumnFormula(cell)
    '                    End If
    '                End If
    '                SAP_CheckPnl(row, e, AreaID)
    '                If cell.Value IsNot Nothing AndAlso cell.Value.ToString <> "" AndAlso cell.ColumnIndex >= ReportUI_DataGridView.Columns("操作員").Index Then

    '                    If cell.ColumnIndex = ReportUI_DataGridView.Columns("操作員").Index Then
    '                        ' 設定資料格為唯讀
    '                        'cell.ReadOnly = True
    '                        cell.Style.BackColor = SystemColors.ControlLight
    '                        Dim cmd As String = "UPDATE " & DbLog & "
    '                                                              SET [OP] = '" + Trim(cell.Value.ToString) + "'
    '                                                              WHERE [AreaID] = " + AreaID + " AND [ProcName] = '" + proc + "' AND [Lotnum] = '" + lot + "' AND [Layer] = '" + layer + "' AND [Count] = " + count
    '                        SQL_Query(cmd)
    '                    ElseIf cell.ColumnIndex <> ReportUI_DataGridView.Columns("備註").Index Then
    '                        ' 設定資料格為唯讀
    '                        'cell.ReadOnly = True
    '                        cell.Style.BackColor = SystemColors.ControlLight
    '                        '20231012 修改可覆寫資料
    '                        Dim cmd As String = " UPDATE " & DbLogParameter & "
    '                                                              SET [UploadTime] = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',
    '                                                                      [ParameterVaules] = '" + Trim(cell.Value.ToString) + "'
    '                                                              WHERE [AreaID] = " + AreaID + " AND [ProcName] = '" + proc + "' AND [Lotnum] = '" + lot + "' AND [LayerName] = '" + layer + "' AND [ParameterName] = '" + ParaName + "' AND [Count] = " + count + ""
    '                        'Dim cmd As String = "DECLARE @Val nvarchar(MAX)
    '                        '                                      SET @Val = (SELECT ISNULL([ParameterVaules],'') AS Val FROM " & DbLogParameter & " WITH (NOLOCK) WHERE [AreaID] = " + AreaID + " AND [ProcName] = '" + proc + "' AND [Lotnum] = '" + lot + "' AND [LayerName] = '" + layer + "' AND [ParameterName] = '" + ParaName + "'  AND [Count] = " + count + ")
    '                        '                                      IF @Val IS NULL OR @Val = ''
    '                        '                                      BEGIN
    '                        '                                      UPDATE " & DbLogParameter & "
    '                        '                                      SET [UploadTime] = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',
    '                        '                                              [ParameterVaules] = '" + cell.Value.ToString + "'
    '                        '                                      WHERE [AreaID] = " + AreaID + " AND [ProcName] = '" + proc + "' AND [Lotnum] = '" + lot + "' AND [LayerName] = '" + layer + "' AND [ParameterName] = '" + ParaName + "' AND [Count] = " + count + "
    '                        '                                      END"
    '                        SQL_Query(cmd)
    '                    Else
    '                        Dim cmd As String = "UPDATE " & DbLog & "
    '                                                              SET [Remark] = '" + Trim(cell.Value.ToString) + "'
    '                                                              WHERE [AreaID] = " + AreaID + " AND [ProcName] = '" + proc + "' AND [Lotnum] = '" + lot + "' AND [Layer] = '" + layer + "' AND [Count] = " + count
    '                        SQL_Query(cmd)
    '                    End If
    '                End If
    '            End If
    '        Else
    '            If ReportUI_DataGridView.Columns(e.ColumnIndex).Name <> "料號" AndAlso ReportUI_DataGridView.Columns(e.ColumnIndex).Name <> "批號" AndAlso ReportUI_DataGridView.Columns(e.ColumnIndex).Name <> "層別" AndAlso ReportUI_DataGridView.Columns(e.ColumnIndex).Name <> "站點" Then
    '                MessageBox.Show("請將""料號""、""批號""、""層別""、""站點""填寫完畢")
    '            End If
    '        End If
    '    Catch ex As Exception
    '        WriteLog(ex, LogFilePath, "ReportUI_DataGridView_CellEndEdit_" + ReportUI_DataGridView.Columns(e.ColumnIndex).Name + "_" + ReportUI_DataGridView.Rows(e.RowIndex).Cells("批號").Value.ToString)
    '    End Try
    'End Sub

End Class
