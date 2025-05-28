Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports log4net
Imports log4net.Config

Module Proc_Update

    '-----------------------------------DB參數----------------------------------------
    ReadOnly DbProc As String = "[H3_Systematic].[dbo].[H3_Proc]" '報表設定Config DB
    ReadOnly SpFixecColumnNew As String = "[H3_Systematic].[dbo].[ProductionQuery_Insert_New]" '固定欄位建立&更新SP(新的，只需要AreaID)
    ReadOnly SpFreeColumn As String = "[H3_Systematic].[dbo].[ProductionQuery_Columns_Insert_New]" '客製欄位建立&更新SP

    '-------------------log參數---------------
    Public ReadOnly log_Local As ILog = LogManager.GetLogger(GetType(Proc_Update))
    Public ReadOnly log_SQL As ILog = LogManager.GetLogger("SQL")
    Public ReadOnly log_Error As ILog = LogManager.GetLogger("Error")

    Sub Main()

        Dim StartTime = Now()
        Dim log As New Mylog4net With {
            .LogFilePath = "C:\ProductionReport\Update_Log\",
            .LogFileName = "CustomColumn_Update"
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

                    'Try
                    '    '抓取WIP資料
                    '    log_SQL.Info("開始 WIP 資料建立")
                    '    SQL_StoredProcedure(SpFixecColumnNew, Parameters)
                    '    log_SQL.Info("完成 WIP 資料建立")
                    'Catch ex As Exception
                    '    log_Error.Error("WIP 資料建立失敗")
                    'End Try

                    Try
                        '建立浮動欄位 & 抓取資料
                        log_SQL.Info("開始浮動欄位資料建立&撈取")
                        SQL_StoredProcedure(SpFreeColumn, Parameters)
                        log_SQL.Info("完成浮動欄位資料建立&撈取")
                    Catch ex As Exception
                        log_Error.Error("浮動欄位資料建立或撈取失敗", ex)
                    End Try
                    'log_Local.Info("完成報表ID :  " & row("Pkey").ToString & " 資料建立")
                Catch ex As Exception
                    'log_Error.Error("報表ID :  " & row("Pkey").ToString & " 資料建立失敗", ex)
                    Continue For
                End Try
            Next
        End Using
        GlobalContext.Properties("AreaID") = "999999"
        log_SQL.Info("完成浮動欄位資料建立或撈取，耗時 : " & Now.Subtract(StartTime).TotalMinutes.ToString & " 分鐘")
        'temp()
    End Sub

    '    Sub temp()
    '        Dim cmd As String = ""
    '        Dim machine As String() = {"LDLCOL1023", "LDLCOL1033", "LDLCOL1043", "LDLCOL1053", "LDLCOL1063", "LDLCOL1073", "LDLCOL1083", "LDLCOL1093", "LDLCOL1103", "LDLCOL1113", "LDLCOL1123", "LDLCOL1133", "LDLCOL1143", "LDLCOL1153", "LDLCOL1163", "LDLCOL1173", "LDLCOL1233", "LDLCOL1243", "LDLCOL1253", "LDLCOL1263", "LDLCOL1273", "LDLCOL1283", "LDLCOL1293", "LDLCOL1303", "LDLCOL1313", "LDLCOL1323", "LDLCOL1333", "LDLCOL1343", "LDLCOL1353", "LDLCOL1363", "LDLCOL1373", "LDLCOL1383", "LDLCOL1393", "LDLCOL1453", "LDLCOL1463", "LDLCOL1193", "LDLCOL1203", "LDLCOL1423", "LDLCOL1413", "LDLCOL1403", "LDLCOL1473", "LDLCOL1483", "LDLCOL1493", "LDLCOL1503", "LDLCOL1513", "LDLCOL1523", "LDLCOL1533", "LDLCOL1543", "LDLCOL1553", "LDLCOL1563"}
    '        Dim GTF5 As String = ""

    '        cmd += "SELECT * INTO ##LDLCOL_Temp FROM
    '(
    'SELECT [MachineNo]
    '      ,[MachineName]
    '      ,REPLACE(CONVERT(varchar(13),[TimeStamp],120),' ',':') AS [Time]
    '      ,MAX([VisionNo]) AS [VisionNo]
    '      ,[STAGE]
    '      ,ROUND(AVG(CAST([Score] AS float)),4) AS [Score]
    '      ,ROUND(AVG(CAST([SIZE] AS float)),2) AS [SIZE]
    '       FROM(SELECT [MachineNo]
    '      ,[MachineName]
    '      ,[Description]
    '      ,[Value]
    '      ,[TimeStamp]
    '  FROM [Datamation_EIS_H3].[dbo].[H3LDLCOL1013_Data_HD] nolock
    '  WHERE EQName = 'H3 CO2雷射機台#01_aaa'
    '  AND [TimeStamp] > DATEADD(day , -7,GETDATE())
    '  ) a

    '  PIVOT(
    '    	MAX([Value])
    '    	FOR [Description] IN([VisionNo],[STAGE],[Score],[SIZE])

    '    	)p
    'WHERE VisionNo = 2 AND
    '[STAGE] IN ('RA','RB','LA','LB')
    'GROUP BY [MachineNo],[MachineName],[STAGE],CONVERT(varchar(13),[TimeStamp],120)
    ')b
    'UNPIVOT
    '(
    '[Values] For [Type] IN([SIZE],[SCORE])
    ')up

    'INSERT INTO ##LDLCOL_Temp
    '"

    '        For i As Integer = 2 To 51
    '            If i >= 16 AndAlso GTF5 = "" Then GTF5 = "(GTF5)"
    '            cmd += "SELECT * FROM
    '(
    'SELECT TOP (10000) [MachineNo]
    '      ,[MachineName]
    '      ,REPLACE(CONVERT(varchar(13),[TimeStamp],120),' ',':') AS [Time]
    '      ,MAX([VisionNo]) AS [VisionNo]
    '      ,[STAGE]
    '      ,ROUND(AVG(CAST([Score] AS float)),4) AS [Score]
    '      ,ROUND(AVG(CAST([SIZE] AS float)),2) AS [SIZE]
    '       FROM(SELECT [MachineNo]
    '      ,[MachineName]
    '      ,[Description]
    '      ,[Value]
    '      ,[TimeStamp]
    '  FROM [Datamation_EIS_H3].[dbo].[H3" & machine(i - 2) & "_Data_HD] nolock
    '  WHERE EQName = 'H3 CO2雷射機台#" & i.ToString("D2") & GTF5 & "_aaa'
    '  AND [TimeStamp] > DATEADD(day , -7,GETDATE())
    '  ) a

    '  PIVOT(
    '    	MAX([Value])
    '    	FOR [Description] IN([VisionNo],[STAGE],[Score],[SIZE])

    '    	)p
    'WHERE VisionNo = 2 AND
    '[STAGE] IN ('RA','RB','LA','LB')
    'GROUP BY [MachineNo],[MachineName],[STAGE],CONVERT(varchar(13),[TimeStamp],120)
    ')b
    'UNPIVOT
    '(
    '[Values] For [Type] IN([SIZE],[SCORE])
    ')up
    '  "
    '            If i < 51 Then cmd += vbCrLf & "INSERT INTO ##LDLCOL_Temp" & vbCrLf
    '        Next

    '        Dim pause As String = ""

    '    End Sub

End Module
