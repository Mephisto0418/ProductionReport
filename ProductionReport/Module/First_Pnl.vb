Module First_Pnl

    Sub First_Upload(ByVal row As DataGridViewRow, ByVal AreaID As String, ByVal lot As String, ByVal proc As String, ByVal layer As String, ByVal face As String)
        Try
            Dim msg As String
            Dim title As String
            Dim style As MsgBoxStyle
            Dim response As MsgBoxResult
            msg = "請確認""料號""、""批號""、""站點""、""層別""是否正確?"
            style = MsgBoxStyle.DefaultButton2 Or MsgBoxStyle.Question Or MsgBoxStyle.YesNo
            title = "內容確認"
            response = MsgBox(msg, style, title)
            If response = MsgBoxResult.No Then
                Return
            End If

            Dim cmdCount As String = "SELECT ISNULL(MAX([Count]),0) AS [Count]
                                                                  FROM [H3_Systematic].[dbo].[H3_ProductionLog] WITH(NOLOCK)
                                                                  WHERE [AreaID] = " + AreaID + " AND [Lotnum] = '" + lot + "' AND [Layer] = '" + layer + "' AND [ProcName] = '" + proc + "'"
            Dim dtCount As DataTable = SQL_Query(cmdCount)
            Dim Count As String = ""
            If dtCount.Rows.Count = 0 Then
                Count = "1"
            Else
                Count = (CInt(dtCount(0)("Count").ToString) + 1).ToString
            End If

            Dim cmdUpload As String = "INSERT INTO [H3_Systematic].[dbo].[H3_ProductionLog]
                                                     ([AreaID],[SystemTime],[Class],[LastEndTime],[StartTime],[EndTime],[Partnum],[Lotnum],[Layer],[ProcName],[EQID],[IType],[Inputpcs],[Outputpcs],[WID],[User],[OP],[Remark],[Flag],[Count])
                                                     VALUES
                                                     (" + AreaID + ",GETDATE(),'" + row.Cells("班別").Value.ToString + "','" + row.Cells("前站結束時間").Value.ToString + "','" + row.Cells("開始時間").Value.ToString + "','" + row.Cells("結束時間").Value.ToString + "','" + row.Cells("料號").Value.ToString + "','" + lot + "','" + layer + "','" + proc + "','" + row.Cells("機台").Value.ToString + "','" + row.Cells("產品類型").Value.ToString + "','" + row.Cells("入料片數").Value.ToString + "','" + row.Cells("出料片數").Value.ToString + "','" + row.Cells("過帳工號").Value.ToString + "','" + row.Cells("過帳人員").Value.ToString + "','" + row.Cells("操作員").Value.ToString + "','" + row.Cells("備註").Value.ToString + "',0," + Count + ")
                                                     DECLARE @PID AS TABLE (
                                                     [PID] int,
                                                     [SystemTime] datetime
                                                     )
                                                     INSERT INTO @PID
                                                     SELECT [Pkey] AS [PID], [SystemTime] FROM [H3_Systematic].[dbo].[H3_ProductionLog] WITH(NOLOCK) WHERE [AreaID] = " + AreaID + " AND [Lotnum] = '" + lot + "' AND [Layer] = '" + layer + "' AND [ProcName] = '" + proc + "' AND [Count] = " + Count + "
                                                     SELECT * FROM @PID
                                                     DECLARE @Temp AS TABLE (
                                                     [PID] int,
                                                     [CID] int,
                                                     [AreaID] int,
                                                     [站點] nvarchar(MAX),
                                                     [料號] nvarchar(MAX),
                                                     [批號] nvarchar(MAX),
                                                     [層別] nvarchar(MAX),
                                                     [Sort] int,
                                                     [ParameterName] nvarchar(MAX),
                                                     [isQuery] bit,
                                                     [QueryCommand] nvarchar(MAX),
                                                     [QueryType] nvarchar(MAX),
                                                     [FormulaColumn] nvarchar(MAX),
                                                     [Machine] nvarchar(MAX),
                                                     [欄位紀錄名稱] nvarchar(MAX),
                                                     [欄位紀錄] nvarchar(MAX),
                                                     [Count] int,
                                                     [Face] nvarchar(50)
                                                     )
                                                     INSERT INTO @Temp
                                                     SELECT l.[Pkey],pa.[pkey],l.[AreaID],l.[ProcName] AS [站點],l.[Partnum] AS [料號],l.[Lotnum] AS [批號],l.[Layer] AS [層別],pa.[PID] AS [Sort]
                                                     ,pa.[ParameterName]
                                                     ,ISNULL(pa.[isQuery],'') AS [isQuery]
                                                     ,ISNULL(r.[QueryCommand],'') AS [QueryCommand]
                                                     ,ISNULL(r.[QueryType],'') AS [QueryType]
                                                     ,CASE WHEN [QueryType] = '欄位間計算' THEN ISNULL(r.[Filter1],'') ELSE '' END AS [FormulaColumn]
                                                     ,ISNULL(pr.[Machine],'') AS [Machine]
                                                     ,ISNULL(p.[ParameterName],'') AS [欄位紀錄名稱]
                                                     ,CASE WHEN pa.isQuery = 0 AND pa.[DefaultValues] <> '' AND (p.[ParameterVaules] = '' OR p.[ParameterVaules] IS NULL) THEN pa.[DefaultValues] ELSE ISNULL(p.[ParameterVaules],'') END AS [欄位紀錄]
                                                     ,l.[Count],'" & face & "' AS [Face]
                                                     FROM [H3_Systematic].[dbo].[H3_ProductionLog] AS l WITH (NOLOCK)
                                                     LEFT JOIN [H3_Systematic].[dbo].[H3_Proc] AS pr WITH (NOLOCK) ON pr.[Pkey] = l.[AreaID] 
                                                     LEFT JOIN [H3_Systematic].[dbo].[H3_Production_ProcParameter] AS pa WITH (NOLOCK) ON pr.[Pkey] = pa.[AreaID]
                                                     LEFT JOIN [H3_Systematic].[dbo].[H3_Production_ProcParameter_Rule] AS r WITH (NOLOCK) ON pa.[QID] = r.[QID]
                                                     LEFT JOIN [H3_Systematic].[dbo].[H3_ProductionParameter] AS p WITH (NOLOCK) ON p.[AreaID] = l.[AreaID] AND p.[ProcName] = l.[ProcName] AND p.[Lotnum] = l.[Lotnum] AND p.[LayerName] = l.[Layer] AND p.[ParameterName] = pa.[ParameterName] AND l.[Count] = p.[Count]
                                                     WHERE pr.[Pkey] = " & AreaID & " AND (l.[EQID] = '' OR pr.[Machine] LIKE '%' + l.[EQID] + '%' OR pr.[Machine] = '')  
                                                     AND l.[Flag] = 0 AND l.[Pkey] = (SELECT [PID] FROM @PID)
                                                     ORDER BY l.[Pkey],[Sort]
                                                      
                                                     INSERT [H3_Systematic].[dbo].[H3_ProductionParameter]
                                                     ([PID],[CID],[AreaID],[Lotnum],[ProcName],[ParameterName],[ParameterVaules],[SystemTime],[LayerName],[Count],[Face])
                                                     SELECT [PID],[CID]," & AreaID & " AS [AreaID],[批號],[站點],[ParameterName],[欄位紀錄],GETDATE() AS [SystemTime],[層別],[Count],[Face]
                                                     FROM @Temp 
                                                     WHERE [欄位紀錄名稱] = '' AND [ParameterName] IS NOT NULL ORDER BY [PID],[Sort]"
            Dim dtPID As DataTable = SQL_Query(cmdUpload)
            ReportUI.ChangeValueIgnore = True
            row.Cells("LogID").Value = dtPID(0)("PID")
            Dim result As DateTime
            If DateTime.TryParse(dtPID(0)("SystemTime"), result) Then
                row.Cells("日期").Value = result.ToString("yyyy-MM-dd HH:mm:ss")
            End If
            ReportUI.ChangeValueIgnore = False
            For i = 0 To row.DataGridView.Columns("備註").Index - 2
                'If row.Cells(i).Value IsNot Nothing AndAlso row.Cells(i).Value.ToString <> "" Then
                '    row.Cells(i).ReadOnly = True
                '    row.Cells(i).Style.BackColor = SystemColors.ControlLight
                'End If
                If {"班別", "料號", "批號", "層別", "站點", "機台", "日期", "前站結束時間", "產品類型", "面次"}.Contains(ReportUI.dgvReport.Columns(i).Name) Then
                    row.Cells(i).ReadOnly = True
                    row.Cells(i).Style.BackColor = SystemColors.ControlLight
                End If
            Next
        Catch ex As Exception
            WriteLog(ex, LogFilePath, "First_Upload")
        End Try
    End Sub


End Module
