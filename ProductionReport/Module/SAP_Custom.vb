Module SAP_Custom
    Private SAP_AreaID As String() = {"74", "75", "76", "125"}
    Public SAP_Pnlqty As String = ""

    Sub SAP_CheckID(ByVal AreaID As String)
        Try
            If AreaID >= 74 And AreaID <= 81 Then
                ReportUI.Btn_First.Visible = True
            Else
                ReportUI.Btn_First.Visible = False
            End If
        Catch ex As Exception
            WriteLog(ex, LogFilePath, "SAP_CheckID")
        End Try
    End Sub

    Sub SAP_CheckPnl(ByVal row As DataGridViewRow, ByVal e As DataGridViewCellEventArgs, ByVal AreaID As String)
        Dim Cell As DataGridViewCell = row.Cells(e.ColumnIndex)
        Try
            If Cell.RowIndex >= 0 AndAlso SAP_AreaID.Contains(AreaID) Then
                If ReportUI.dgvReport.Columns(Cell.ColumnIndex).Name = "第一片基板數量" AndAlso Cell.Value IsNot Nothing AndAlso Cell.Value.ToString <> "" Then
                    If IsNumeric(Cell.Value.ToString) AndAlso (CInt(Cell.Value.ToString) >= 1 And CInt(Cell.Value.ToString) <= 58) Then
                        SAP_Pnlqty = Cell.Value.ToString
                        SAP_PnlKeyIn.Show()
                        SAP_PnlKeyIn.Focus()
                        SAP_PnlKeyIn.dgv_PnlKeyIn.Rows.Clear()
                        For i = 1 To CInt(Cell.Value)
                            SAP_PnlKeyIn.dgv_PnlKeyIn.Rows.Add(i.ToString, "")
                        Next
                        SAP_PnlKeyIn.Cell = Cell
                        SAP_PnlKeyIn.tarCell = row.Cells(e.ColumnIndex + 1)
                    Else
                        Cell.Value = ""
                        MessageBox.Show("請輸入數字1~58")
                    End If
                End If
            End If
        Catch ex As Exception
            WriteLog(ex, LogFilePath, "SAP_CheckPnl")
        End Try
    End Sub

    Sub SAP_First_Upload(ByVal row As DataGridViewRow, ByVal AreaID As String)
        Try
            Dim cmd As String = "INSERT INTO [H3_Systematic].[dbo].[H3_ProductionLog]
                                                      ([AreaID],[SystemTime],[Class],[StartTime],[EndTime],[Partnum],[Lotnum],[Layer],[ProcName],[EQID],[IType],[Inputpcs],[Outputpcs],[WID],[User],[OP],[Remark],[Flag],[Count])
                                                      VALUES
                                                      (" + AreaID + ",GETDATE(),'" + row.Cells("班別").Value.ToString + "','" + row.Cells("開始時間").Value.ToString + "','" + row.Cells("結束時間").Value.ToString + "','" + row.Cells("料號").Value.ToString + "','" + row.Cells("批號").Value.ToString + "','" + row.Cells("層別").Value.ToString + "','" + row.Cells("站點").Value.ToString + "','" + row.Cells("機台").Value.ToString + "','" + row.Cells("產品類型").Value.ToString + "','" + row.Cells("入料片數").Value.ToString + "','" + row.Cells("出料片數").Value.ToString + "','" + row.Cells("過帳工號").Value.ToString + "','" + row.Cells("過帳人員").Value.ToString + "','" + row.Cells("操作員").Value.ToString + "','" + row.Cells("備註").Value.ToString + "',0,1)
                                                      SELECT [Pkey] AS [PID] FROM [H3_Systematic].[dbo].[H3_ProductionLog] WITH(NOLOCK) WHERE [AreaID] = " + AreaID + " AND [Lotnum] = '" + row.Cells("批號").Value.ToString + "' AND [Layer] = '" + row.Cells("層別").Value.ToString + "' AND [ProcName] = '" + row.Cells("站點").Value.ToString + "' AND [Count] = 1
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
                                                     ,l.[Count],face.[FNo] AS [Face]
                                                     FROM [H3_Systematic].[dbo].[H3_ProductionLog] AS l WITH (NOLOCK)
                                                     LEFT JOIN [H3_Systematic].[dbo].[H3_Proc] AS pr WITH (NOLOCK) ON pr.[Pkey] = l.[AreaID] 
                                                     LEFT JOIN [H3_Systematic].[dbo].[H3_Production_ProcParameter] AS pa WITH (NOLOCK) ON pr.[Pkey] = pa.[AreaID]
                                                     LEFT JOIN [H3_Systematic].[dbo].[H3_Production_ProcParameter_Rule] AS r WITH (NOLOCK) ON pa.[QID] = r.[QID]
                                                     LEFT JOIN [H3_Systematic].[dbo].[H3_ProductionParameter] AS p WITH (NOLOCK) ON p.[AreaID] = l.[AreaID] AND p.[ProcName] = l.[ProcName] AND p.[Lotnum] = l.[Lotnum] AND p.[LayerName] = l.[Layer] AND p.[ParameterName] = pa.[ParameterName] AND l.[Count] = p.[Count]
                                                     LEFT JOIN (SELECT [Pkey] AS [AreaID] , [FNo]
                                                     FROM(
                                                     SELECT [Pkey], CASE WHEN [hasFace] = 1 THEN 1 ELSE 0 END AS [PF],CASE WHEN [hasFace] = 1 THEN 2 ELSE 0 END AS [PB] FROM [H3_Systematic].[dbo].[H3_Proc] WITH (NOLOCK) WHERE [Pkey] = " & AreaID & "
                                                     ) s
                                                     UNPIVOT
                                                     (
                                                     [FNo] FOR [Face]IN([PF],[PB])
                                                     )upv
                                                     GROUP BY [Pkey],[FNo]) AS face ON face.[AreaID] = pa.[AreaID] AND (face.[FNo] = p.[Face] OR p.[Face] IS NULL)
                                                     WHERE pr.[Pkey] = " & AreaID & " AND (l.[EQID] = '' OR pr.[Machine] LIKE '%' + l.[EQID] + '%' OR pr.[Machine] = '')  
                                                     AND l.[Flag] = 0
                                                     ORDER BY l.[Pkey],[Sort]
                                                      
                                                      INSERT [H3_Systematic].[dbo].[H3_ProductionParameter]
                                                      ([PID],[CID],[AreaID],[Lotnum],[ProcName],[ParameterName],[ParameterVaules],[SystemTime],[LayerName],[Count],[Face])
                                                      SELECT [PID],[CID]," & AreaID & " AS [AreaID],[批號],[站點],[ParameterName],[欄位紀錄],GETDATE() AS [SystemTime],[層別],[Count],[Face]
                                                      FROM @Temp 
                                                      WHERE [欄位紀錄名稱] = '' AND [ParameterName] IS NOT NULL ORDER BY [PID],[Sort]"
            Dim dtPID As DataTable = SQL_Query(cmd)
            row.Cells("LogID").Value = dtPID(0)("PID")
            For i = 0 To row.DataGridView.Columns("備註").Index - 2
                row.Cells(i).ReadOnly = True
                row.Cells(i).Style.BackColor = SystemColors.ControlLight
            Next
        Catch ex As Exception
            WriteLog(ex, LogFilePath, "SAP_First_Upload")
        End Try
    End Sub

End Module
