Module PTH_Custom
    Public PTH_AreaID As String() = {"49", "61", "62", "83", "101", "102", "103"}
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

                ReportUI.txtSpiltNum.Visible = True
                ReportUI.lblSpiltNum.Visible = True
            Else
                ReportUI.txtSpiltNum.Visible = False
                ReportUI.lblSpiltNum.Visible = False
            End If
        Catch ex As Exception
            WriteLog(ex, LogFilePath, "PTH_CheckID")
        End Try
    End Sub

    Sub PTH_SplitClick(ByVal e As DataGridViewCellEventArgs, ByRef ChangeValueIgnore As Boolean)
        Try
            ReportUI.TimerRefresh.Stop()
            '判斷此筆資料是否已上傳
            If ReportUI.dgvReport.Rows(e.RowIndex).Cells("完成").Value IsNot Nothing AndAlso ReportUI.dgvReport.Rows(e.RowIndex).Cells("完成").Value.ToString <> "" Then
                MessageBox.Show("此筆資料已上傳完畢")
                Return
            End If

            If ReportUI.txtSpiltNum.Text <> "" AndAlso IsNumeric(ReportUI.txtSpiltNum.Text) AndAlso (CInt(ReportUI.txtSpiltNum.Text) > 0 AndAlso CInt(ReportUI.txtSpiltNum.Text) <= 48) Then
                ReportUI.txtSpiltNum.Text = CInt(ReportUI.txtSpiltNum.Text).ToString
            Else
                MessageBox.Show("請在分批數量的欄位內輸入1~48的數字")
                ReportUI.txtSpiltNum.Focus()
                ReportUI.txtSpiltNum.Text = "1"
                Return
            End If

            Dim paralist As New List(Of String)
            Dim face As String = ""
            For i As Integer = 0 To ReportUI.dgvReport.Columns("備註").Index
                If ReportUI.dgvReport.Columns(i).Name = "班別" Then
                    paralist.Add("分批")
                ElseIf ReportUI.AreaID = "83" AndAlso {"日期", "班別", "前站結束時間", "開始時間", "結束時間", "料號", "批號", "層別", "站點", "機台", "面次", "產品類型", "入料片數", "出料片數", "過帳工號", "過帳人員"}.Contains(ReportUI.dgvReport.Columns(i).Name) Then
                    If ReportUI.dgvReport.Columns(i).Name = "面次" Then
                        face = ReportUI.dgvReport.Rows(e.RowIndex).Cells(i).Value.ToString
                    End If
                    paralist.Add(ReportUI.dgvReport.Rows(e.RowIndex).Cells(i).Value.ToString)
                ElseIf {"料號", "批號", "層別", "站點", "機台", "前站結束時間", "產品類型"}.Contains(ReportUI.dgvReport.Columns(i).Name) Then
                    paralist.Add(ReportUI.dgvReport.Rows(e.RowIndex).Cells(i).Value.ToString)
                Else
                    paralist.Add("")
                End If
            Next

            Dim RowNum As Integer = e.RowIndex + 1

            For j As Integer = 0 To CInt(ReportUI.txtSpiltNum.Text) - 1

                Dim dgvcbocFace As New DataGridViewComboBoxCell
                dgvcbocFace.Items.Add("N/A")
                dgvcbocFace.Items.Add("PF")
                dgvcbocFace.Items.Add("PB")

                ChangeValueIgnore = True

                ReportUI.dgvReport.Rows.Insert(RowNum, paralist.ToArray)

                ReportUI.dgvReport.Rows(RowNum).Cells("面次") = dgvcbocFace
                If face = "" OrElse face = "N/A" Then
                    ReportUI.dgvReport.Rows(RowNum).Cells("面次").Value = "N/A"
                Else
                    ReportUI.dgvReport.Rows(RowNum).Cells("面次").Value = face
                End If



                For k = 0 To ReportUI.dgvReport.Columns("備註").Index - 2
                    If Not {"班別", "料號", "批號", "層別", "站點", "機台", "日期", "前站結束時間", "產品類型"}.Contains(ReportUI.dgvReport.Columns(k).Name) Then
                        ReportUI.dgvReport.Rows(RowNum).Cells(k).ReadOnly = False
                        ReportUI.dgvReport.Rows(RowNum).Cells(k).Style.BackColor = SystemColors.ControlLightLight
                    End If
                Next
            Next
            face = ""
            ChangeValueIgnore = False
            ReportUI.TimerRefresh.Start()
        Catch ex As Exception
            WriteLog(ex, LogFilePath, "PTH_SplitClick")
            ChangeValueIgnore = False
            ReportUI.TimerRefresh.Start()
        End Try

    End Sub

    Sub PTH_DeleteClick(ByVal e As DataGridViewCellEventArgs)
        Try
            '請使用者確認是否刪除
            If ReportUI.dgvReport.Rows(e.RowIndex).Cells("LogID").Value IsNot Nothing AndAlso ReportUI.dgvReport.Rows(e.RowIndex).Cells("LogID").Value.ToString <> "" Then
                Dim pw As String = InputBox("請輸入刪除密碼：", "輸入刪除密碼", " ")
                If pw = " " Then
                    Return
                ElseIf pw <> DeletePw Then
                    MessageBox.Show("密碼輸入錯誤")
                    Return
                End If
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
