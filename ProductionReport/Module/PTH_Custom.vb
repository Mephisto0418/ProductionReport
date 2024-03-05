Module PTH_Custom
    Public PTH_AreaID As String() = {"49", "61", "62", "101", "102", "103"}
    Private DeletePw As String = "0000"

    Sub PTH_CheckID(ByVal AreaID As String)
        Try
            If PTH_AreaID.Contains(AreaID) Then
                '建立分批按鈕
                Dim btnCol As New DataGridViewButtonColumn()
                btnCol.HeaderText = "分批" & vbCrLf & "(Split)"
                btnCol.Name = "btnSplit"
                btnCol.Text = "分批"

                Dim bc As New DataGridViewButtonCell()
                bc.FlatStyle = FlatStyle.Flat
                bc.Style.BackColor = Color.Orange
                btnCol.CellTemplate = bc
                ReportUI.dgvReport.Columns.Add(btnCol)
                ReportUI.dgvReport.Columns("btnSplit").Width = 40

                '建立刪除按鈕
                Dim btndel As New DataGridViewButtonColumn()
                btndel.HeaderText = "刪除" & vbCrLf & "(Delete)"
                btndel.Name = "btnDelete"
                btndel.Text = "刪除"

                Dim bc2 As New DataGridViewButtonCell()
                bc2.FlatStyle = FlatStyle.Flat
                bc2.Style.BackColor = Color.Red
                bc2.Value = "Delete"
                btndel.CellTemplate = bc2
                ReportUI.dgvReport.Columns.Add(btndel)
                ReportUI.dgvReport.Columns("btnDelete").Width = 40
            End If
        Catch ex As Exception
            WriteLog(ex, LogFilePath, "PTH_CheckID")
        End Try
    End Sub

    Sub PTH_SplitClick(ByVal e As DataGridViewCellEventArgs, ByRef ChangeValueIgnore As Boolean)
        Try
            '判斷此筆資料是否已上傳
            If ReportUI.dgvReport.Rows(e.RowIndex).Cells("完成").Value IsNot Nothing AndAlso ReportUI.dgvReport.Rows(e.RowIndex).Cells("完成").Value.ToString <> "" Then
                MessageBox.Show("此筆資料已上傳完畢")
                Return
            End If
            ''請使用者輸入分批數量，必須輸入整數1~4
            'Dim Num As String = InputBox("請輸入分批數量：", "輸入分批數量", " ")
            'If Num = " " Then
            '    Return
            'ElseIf Not IsNumeric(Num) OrElse Int(Num) < 1 OrElse Int(Num) > 4 Then
            '    MessageBox.Show("請輸入1~4的整數")
            '    Return
            'End If
            Dim RowNum As Integer = e.RowIndex + 1

            ChangeValueIgnore = True
            Dim paralist As New List(Of String)
            For i As Integer = 0 To ReportUI.dgvReport.Columns("備註").Index
                If ReportUI.dgvReport.Columns(i).Name = "班別" Then
                    paralist.Add("分批")
                ElseIf {"料號", "批號", "層別", "站點", "機台", "前站結束時間", "產品類型"}.Contains(ReportUI.dgvReport.Columns(i).Name) Then
                    paralist.Add(ReportUI.dgvReport.Rows(e.RowIndex).Cells(i).Value.ToString)
                Else
                    paralist.Add("")
                End If
            Next
            ReportUI.dgvReport.Rows.Insert(RowNum, paralist.ToArray)
            Dim dgvcbocFace As New DataGridViewComboBoxCell
            dgvcbocFace.Items.Add("N/A")
            dgvcbocFace.Items.Add("PF")
            dgvcbocFace.Items.Add("PB")

            ReportUI.dgvReport.Rows(RowNum).Cells("面次") = dgvcbocFace

            ReportUI.dgvReport.Rows(RowNum).Cells("面次").Value = "N/A"

            ChangeValueIgnore = False
            For i = 0 To ReportUI.dgvReport.Columns("備註").Index - 2
                If Not {"班別", "料號", "批號", "層別", "站點", "機台", "日期", "前站結束時間", "產品類型"}.Contains(ReportUI.dgvReport.Columns(i).Name) Then
                    ReportUI.dgvReport.Rows(RowNum).Cells(i).ReadOnly = False
                    ReportUI.dgvReport.Rows(RowNum).Cells(i).Style.BackColor = SystemColors.ControlLightLight
                End If
            Next

        Catch ex As Exception
            WriteLog(ex, LogFilePath, "PTH_SplitClick")
            ChangeValueIgnore = False
        End Try

    End Sub

    Sub PTH_DeleteClick(ByVal e As DataGridViewCellEventArgs)
        Try
            '請使用者確認是否刪除
            Dim pw As String = InputBox("請輸入刪除密碼：", "輸入刪除密碼", " ")
            If pw = " " Then
                Return
            ElseIf pw <> DeletePw Then
                MessageBox.Show("密碼輸入錯誤")
                Return
            End If

            If ReportUI.dgvReport.Rows(e.RowIndex).Cells("班別").Value.ToString = "分批" AndAlso ReportUI.dgvReport.Rows(e.RowIndex).Cells("LogID").Value IsNot Nothing AndAlso ReportUI.dgvReport.Rows(e.RowIndex).Cells("LogID").Value.ToString <> "" Then
                Dim cmd As String = "
                                                         DELETE [H3_Systematic].[dbo].[H3_ProductionParameter]
                                                         WHERE [PID] = " & ReportUI.dgvReport.Rows(e.RowIndex).Cells("LogID").Value.ToString & "
                                                         DELETE [H3_Systematic].[dbo].[H3_ProductionLog]
                                                         WHERE [Pkey] = " & ReportUI.dgvReport.Rows(e.RowIndex).Cells("LogID").Value.ToString
                SQL_Query(cmd)
                ReportUI.dgvReport.Rows.RemoveAt(e.RowIndex)
                MessageBox.Show("刪除完畢(Deleted compelete.)")
            ElseIf ReportUI.dgvReport.Rows(e.RowIndex).Cells("班別").Value.ToString = "分批" Then
                ReportUI.dgvReport.Rows.RemoveAt(e.RowIndex)
            Else
                MessageBox.Show("請勿刪除原始資料")
            End If

        Catch ex As Exception
            WriteLog(ex, LogFilePath, "PTH_DeleteClick")
        End Try

    End Sub

End Module
