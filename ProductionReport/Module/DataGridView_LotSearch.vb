Module DataGridView_LotSearch
    Public CurrentFindLot As New Search_Lot("", "", "")
    Public CurrentFindIndex As Integer = 0
    Sub LotSearch(ByVal dgv As DataGridView, ByRef MatchCount As Integer, ByVal Lot As String, Optional ByVal Layer As String = "", Optional ByVal Face As String = "")
        Try
            For Each row As DataGridViewRow In dgv.Rows
                If row.Cells("批號").Value IsNot Nothing AndAlso row.Cells("批號").Value.ToString Like "*" + Lot + "*" AndAlso row.Cells("層別").Value.ToString Like "*" + Layer + "*" AndAlso row.Cells("面次").Value.ToString Like "*" + Face + "*" Then
                    dgv.Rows.RemoveAt(row.Index)
                    dgv.Rows.Insert(0, row)
                    MatchCount += 1
                End If
            Next
        Catch ex As Exception
            WriteLog(ex, LogFilePath, "LotSearch_Lot")
        End Try
    End Sub

    Sub LotSearch_Focus(ByRef dgv As DataGridView, ByRef MatchCount As Integer, ByVal Lot As String, Optional ByVal Layer As String = "", Optional ByVal Face As String = "")
        Try
            For Each row As DataGridViewRow In dgv.Rows
                If CurrentFindIndex > row.Index Then
                    row.Selected = False
                    Continue For
                ElseIf row.Index + 1 >= dgv.Rows.Count AndAlso CurrentFindIndex > 0 Then
                    MessageBox.Show("已到最後一行")
                    Return
                End If
                If row.Cells("批號").Value IsNot Nothing AndAlso row.Cells("批號").Value.ToString.Contains(Lot) AndAlso row.Cells("層別").Value.ToString.Contains(Layer) AndAlso row.Cells("面次").Value.ToString.Contains(Face) Then
                    ' 選取找到的列
                    row.Selected = True
                    ' 將滾動條調整至該列可見
                    dgv.FirstDisplayedScrollingRowIndex = row.Index
                    ' 更新下一次尋找的起始索引
                    CurrentFindIndex = row.Index + 1
                    MatchCount += 1

                    Dim SelectRow As DataGridViewRow = ReportUI.dgvReport.Rows(row.Index)
                    If PTH_AreaID.Contains(ReportUI.AreaID) AndAlso ReportUI.AreaID <> "83" AndAlso (SelectRow.Cells("完成").Value Is Nothing OrElse SelectRow.Cells("完成").Value.ToString = "") Then
                        Dim MaxVal As Integer = 0
                        Dim CurrentLotMax As Double = 0
                        For Each rowmax As DataGridViewRow In ReportUI.dgvReport.Rows
                            If IsNumeric(rowmax.Cells("排序").Value.ToString) Then
                                If rowmax.Cells("排序").Value > MaxVal Then
                                    MaxVal = rowmax.Cells("排序").Value
                                End If
                                If rowmax.Cells("排序").Value > CurrentLotMax AndAlso rowmax.Cells("批號").Value.ToString.Contains(Lot) AndAlso rowmax.Cells("層別").Value.ToString.Contains(Layer) AndAlso rowmax.Cells("面次").Value.ToString.Contains(Face) Then
                                    CurrentLotMax = rowmax.Cells("排序").Value
                                End If
                            End If
                        Next
                        If SelectRow.Cells("排序").Value Is Nothing OrElse SelectRow.Cells("排序").Value.ToString = "" OrElse SelectRow.Cells("排序").Value.ToString = "NA" Then
                            If SelectRow.Cells("班別").Value.ToString = "分批" Then
                                CurrentLotMax += 0.1
                                SelectRow.Cells("排序").Value = CurrentLotMax.ToString.Substring(0, CurrentLotMax.ToString.Length - 2).PadLeft(10, "0") + CurrentLotMax.ToString.Substring(CurrentLotMax.ToString.Length - 2, 2)
                            ElseIf SelectRow.Cells("班別").Value.ToString = "D" OrElse SelectRow.Cells("班別").Value.ToString = "N" Then
                                MaxVal += 1
                                SelectRow.Cells("排序").Value = MaxVal.ToString.PadLeft(10, "0")
                            End If
                        End If
                    End If
                    Return
                End If
            Next
        Catch ex As Exception
            WriteLog(ex, LogFilePath, "LotSearch_Lot")
        End Try
    End Sub
End Module
