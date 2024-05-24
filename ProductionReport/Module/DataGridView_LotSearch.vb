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

    Sub LotSearch_Focus(ByVal dgv As DataGridView, ByRef MatchCount As Integer, ByVal Lot As String, Optional ByVal Layer As String = "", Optional ByVal Face As String = "")
        Try
            For Each row As DataGridViewRow In dgv.Rows
                If CurrentFindIndex > row.Index Then
                    row.Selected = False
                    Continue For
                ElseIf row.Index + 1 >= dgv.Rows.Count AndAlso CurrentFindIndex > 0 Then
                    MessageBox.Show("已到最後一行")
                    CurrentFindIndex = 0
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
                    Return
                End If
            Next
        Catch ex As Exception
            WriteLog(ex, LogFilePath, "LotSearch_Lot")
        End Try
    End Sub
End Module
