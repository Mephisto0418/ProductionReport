Public Class SAP_PnlKeyIn
    Public Cell As DataGridViewCell
    Public tarCell As DataGridViewCell
    Dim isClosed As Boolean = False

    Private Sub Btn_Confirm_Click(sender As Object, e As EventArgs) Handles Btn_Confirm.Click
        Try
            Dim pnl As String = ""
            For Each row As DataGridViewRow In dgv_PnlKeyIn.Rows
                If row.Cells(1).Value IsNot Nothing AndAlso row.Cells(1).Value.ToString <> "" Then
                    pnl += row.Cells(1).Value.ToString + ","
                Else
                    MessageBox.Show("請將所有欄位填寫完畢")
                    Return
                End If
            Next
            tarCell.Value = pnl.Substring(0, pnl.Length - 1)
            isClosed = True
            Me.Close()
        Catch ex As Exception
            WriteLog(ex, LogFilePath, "Btn_Confirm_Click")
            isClosed = False
        End Try
    End Sub

    Private Sub dgv_PnlKeyIn_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_PnlKeyIn.CellValueChanged

        Try
            If e.RowIndex > 0 Then
                Dim cell As DataGridViewCell = dgv_PnlKeyIn.Rows(e.RowIndex).Cells(e.ColumnIndex)
                If ReportUI.AreaID = "93" AndAlso CInt(Cell.Value.ToString) <= 0 AndAlso CInt(Cell.Value.ToString) > CInt(SAP_Pnlqty) AndAlso (Not IsNumeric(Cell.Value.ToString) And UCase(Cell.Value.ToString) <> "X") Then
                    Cell.Value = ""
                    MessageBox.Show("請輸入1~" + SAP_Pnlqty + "，或請輸入""X""")
                ElseIf ReportUI.AreaID = "93" And UCase(Cell.Value.ToString) = "X" Then
                    Cell.Value = "X"
                ElseIf Not (IsNumeric(Cell.Value.ToString)) OrElse Cell.Value.ToString = "" Then
                    Cell.Value = ""
                    MessageBox.Show("請輸入數字")
                ElseIf CInt(Cell.Value.ToString) <= 0 And CInt(Cell.Value.ToString) > CInt(SAP_Pnlqty) Then
                    Cell.Value = ""
                    MessageBox.Show("請輸入1~" + SAP_Pnlqty)
                End If
            End If
        Catch ex As Exception
            WriteLog(ex, LogFilePath, "dgv_PnlKeyIn_CellValueChanged")
        End Try
    End Sub


End Class