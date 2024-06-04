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
            Dim row As DataGridViewRow = ReportUI.dgvReport.Rows(e.RowIndex)

            If row.Cells("完成").Value IsNot Nothing AndAlso row.Cells("完成").Value.ToString <> "" Then
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
            Dim lot As String = row.Cells("批號").Value.ToString
            Dim layer As String = row.Cells("層別").Value.ToString
            Dim proc As String = row.Cells("站點").Value.ToString
            Dim face As String = ""
            For i As Integer = 0 To ReportUI.dgvReport.Columns("備註").Index
                If ReportUI.dgvReport.Columns(i).Name = "班別" Then
                    paralist.Add("分批")
                ElseIf ReportUI.AreaID = "83" AndAlso {"日期", "班別", "前站結束時間", "開始時間", "結束時間", "料號", "批號", "層別", "站點", "機台", "面次", "產品類型", "入料片數", "出料片數", "過帳工號", "過帳人員"}.Contains(ReportUI.dgvReport.Columns(i).Name) Then
                    If ReportUI.dgvReport.Columns(i).Name = "面次" Then
                        face = row.Cells(i).Value.ToString
                    End If
                    paralist.Add(row.Cells(i).Value.ToString)
                ElseIf {"料號", "批號", "層別", "站點", "機台", "前站結束時間", "產品類型"}.Contains(ReportUI.dgvReport.Columns(i).Name) Then
                    paralist.Add(row.Cells(i).Value.ToString)
                Else
                    paralist.Add("")
                End If
            Next

            Dim RowNum As Integer = e.RowIndex + 1
            Dim MaxVal As Integer = 0
            Dim CurrentLotMax As Double = 0
            If ReportUI.AreaID <> "83" Then
                For Each rowmax As DataGridViewRow In ReportUI.dgvReport.Rows
                    If IsNumeric(rowmax.Cells("排序").Value.ToString) Then
                        If rowmax.Cells("排序").Value > MaxVal Then
                            MaxVal = rowmax.Cells("排序").Value
                        End If
                        If rowmax.Cells("排序").Value > CurrentLotMax AndAlso rowmax.Cells("批號").Value.ToString.Contains(lot) AndAlso rowmax.Cells("層別").Value.ToString.Contains(layer) AndAlso rowmax.Cells("站點").Value.ToString.Contains(proc) Then
                            CurrentLotMax = rowmax.Cells("排序").Value
                        End If
                    End If
                Next
                If row.Cells("排序").Value Is Nothing OrElse row.Cells("排序").Value.ToString = "" OrElse row.Cells("排序").Value.ToString = "NA" Then
                    If row.Cells("班別").Value.ToString = "分批" Then
                        CurrentLotMax += 0.1
                        row.Cells("排序").Value = CurrentLotMax.ToString.Substring(0, CurrentLotMax.ToString.Length - 2).PadLeft(10, "0") + CurrentLotMax.ToString.Substring(CurrentLotMax.ToString.Length - 2, 2)
                    ElseIf row.Cells("班別").Value.ToString = "D" OrElse row.Cells("班別").Value.ToString = "N" Then
                        MaxVal += 1
                        row.Cells("排序").Value = MaxVal.ToString.PadLeft(10, "0")
                        CurrentLotMax = MaxVal
                    End If
                End If

            End If

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
                ChangeValueIgnore = False
                Dim facenum As String = ""

                If ReportUI.AreaID <> "83" Then
                    Select Case face
                        Case "PB"
                            facenum = "2"
                        Case Else
                            facenum = "1"
                    End Select

                    First_Upload(ReportUI.dgvReport.Rows(RowNum), ReportUI.AreaID, lot, proc, layer, facenum, True)

                    If ReportUI.dgvReport.Rows(RowNum).Cells("排序").Value Is Nothing OrElse ReportUI.dgvReport.Rows(RowNum).Cells("排序").Value.ToString = "" OrElse ReportUI.dgvReport.Rows(RowNum).Cells("排序").Value.ToString = "NA" Then
                        If ReportUI.dgvReport.Rows(RowNum).Cells("班別").Value.ToString = "分批" Then
                            CurrentLotMax += 0.1
                            ReportUI.dgvReport.Rows(RowNum).Cells("排序").Value = CurrentLotMax.ToString.Substring(0, CurrentLotMax.ToString.Length - 2).PadLeft(10, "0") + CurrentLotMax.ToString.Substring(CurrentLotMax.ToString.Length - 2, 2)
                        End If
                    End If
                End If
            Next
            face = ""
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
            Dim row As DataGridViewRow = ReportUI.dgvReport.Rows(e.RowIndex)

            If row.Cells("LogID").Value IsNot Nothing AndAlso row.Cells("LogID").Value.ToString <> "" Then
                Dim pw As String = InputBox("請輸入刪除密碼：", "輸入刪除密碼", " ")
                If pw = " " Then
                    Return
                ElseIf pw <> DeletePw Then
                    MessageBox.Show("密碼輸入錯誤")
                    Return
                End If
            End If

            If row.Cells("班別").Value.ToString = "分批" AndAlso row.Cells("LogID").Value IsNot Nothing AndAlso row.Cells("LogID").Value.ToString <> "" Then
                Dim cmd As String = "
                                                         DELETE [H3_Systematic].[dbo].[H3_ProductionParameter]
                                                         WHERE [PID] = " & row.Cells("LogID").Value.ToString & "
                                                         DELETE [H3_Systematic].[dbo].[H3_ProductionLog]
                                                         WHERE [Pkey] = " & row.Cells("LogID").Value.ToString
                SQL_Query(cmd)
                ReportUI.dgvReport.Rows.RemoveAt(e.RowIndex)
                MessageBox.Show("刪除完畢(Deleted compelete.)")
            ElseIf row.Cells("班別").Value.ToString = "分批" Then
                ReportUI.dgvReport.Rows.RemoveAt(e.RowIndex)
            Else
                MessageBox.Show("請勿刪除原始資料")
            End If

        Catch ex As Exception
            WriteLog(ex, LogFilePath, "PTH_DeleteClick")
        End Try

    End Sub

End Module
