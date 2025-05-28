Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports log4net
Imports log4net.Config

Module Proc_Update_WIP

    '-----------------------------------DB參數----------------------------------------
    ReadOnly DbProc As String = "[H3_Systematic].[dbo].[H3_Proc]" '報表設定Config DB
    ReadOnly SpFixecColumnNew As String = "[H3_Systematic].[dbo].[ProductionQuery_Insert_New]" '固定欄位建立&更新SP(新的，只需要AreaID)
    ReadOnly SpFreeColumn As String = "[H3_Systematic].[dbo].[ProductionQuery_Columns_Insert_New]" '客製欄位建立&更新SP

    '-------------------log參數---------------
    Public ReadOnly log_Local As ILog = LogManager.GetLogger(GetType(Proc_Update_WIP))
    Public ReadOnly log_SQL As ILog = LogManager.GetLogger("SQL")
    Public ReadOnly log_Error As ILog = LogManager.GetLogger("Error")

    Sub Main()

        Dim StartTime = Now()
        'Log文件配置初始化
        Dim log As New Mylog4net With {
            .LogFilePath = "C:\ProductionReport\Update_Log\",
            .LogFileName = "WIP_Update"
        }
        log.ConfigInit()

        log_Local.Debug("開始撈取 " & DbProc & " 資料")
        Using dt_AreaID As DataTable = SQL_Query("SELECT [Pkey] FROM " & DbProc & " WITH (NOLOCK) WHERE [Enable] = 1 AND [Module] <> '0'  ORDER BY [Pkey] ") '搜尋資料庫內的所有站點和密碼
            log_Local.Debug("完成 " & DbProc & " 資料撈取")
            For Each row As DataRow In dt_AreaID.Rows
                Try
                    'log_Local.Info("開始報表ID : " & row("Pkey").ToString & " 資料建立")
                    Dim Parameters As New List(Of String())
                    Parameters.Clear()
                    log_Local.Debug("開始變更 AreaID")
                    Parameters.Add({"@AreaID", row("Pkey").ToString})
                    Mylog4net.SetProperty.AreaID = row("Pkey").ToString
                    log_Local.Debug("完成 AreaID 變更")

                    Try
                        '抓取WIP資料
                        log_SQL.Info("開始 WIP 資料建立")
                        SQL_StoredProcedure(SpFixecColumnNew, Parameters)
                        log_SQL.Info("完成 WIP 資料建立")
                    Catch ex As Exception
                        log_Error.Error("WIP 資料建立失敗", ex)
                    End Try

                    'Try
                    '    '建立浮動欄位 & 抓取資料
                    '    log_SQL.Info("開始浮動欄位資料建立&撈取")
                    '    SQL_StoredProcedure(SpFreeColumn, Parameters)
                    '    log_SQL.Info("完成浮動欄位資料建立&撈取")
                    'Catch ex As Exception
                    '    log_Error.Error("浮動欄位資料建立或撈取失敗")
                    'End Try
                    'log_Local.Info("完成報表ID :  " & row("Pkey").ToString & " 資料建立")
                Catch ex As Exception
                    'log_Error.Error("報表ID :  " & row("Pkey").ToString & " 資料建立失敗", ex)
                    Continue For
                End Try
            Next
        End Using
        GlobalContext.Properties("AreaID") = "999999"
        log_SQL.Info("完成 WIP 資料建立，耗時 : " & Now.Subtract(StartTime).TotalMinutes.ToString & " 分鐘")
    End Sub
End Module
