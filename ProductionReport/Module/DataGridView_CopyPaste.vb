Module DataGridView_CopyPaste
    Public IgnoreColumns As String() = {"btnModify", "btnSplit", "btnDelete"}

    ''' <summary>
    ''' 複製選取的DataGridViewCells內容並將其存儲到剪貼板
    ''' </summary>
    ''' <param name="dgv">DataGridView，包含要複製的選取內容</param>
    ''' <remarks>
    ''' 此複製方法會包含DataGridView的儲存格資料
    ''' </remarks>
    Sub CopyCells(ByRef dgv As DataGridView)
        Clipboard.SetDataObject(dgv.GetClipboardContent)
    End Sub

    ''' <summary>
    ''' 在選取的範圍貼上剪貼板內容
    ''' </summary>
    ''' <param name="dgv">DataGridView，包含要貼上資料的目標範圍</param>
    ''' <param name="IgnoreReadOnly">Boolean，True表示不會將資料貼到ReadOnly的儲存格</param>
    ''' <param name="IgnoreCol">String集合，非必填項目，不會將資料貼到與集合內字串相同的DataGridViewColumns，填寫範例: {"ColumnsName1", "ColumnsName2", "ColumnsName3"}。</param>
    ''' <remarks>
    ''' 
    ''' </remarks>
    Sub PasteCells(ByRef dgv As DataGridView, ByVal IgnoreReadOnly As Boolean, Optional IgnoreCol As String() = Nothing)
        'IgnoreCol若是沒有填寫，指定一個空集合給他
        If IgnoreCol Is Nothing Then IgnoreCol = {}
        '參數
        Dim s_Row As String() = Clipboard.GetText.Split({ControlChars.CrLf}, StringSplitOptions.None) '將剪貼簿內容轉字串並依換行符號分割
        Dim colMax = dgv.Columns.Count '目標dgv的最大ColumnIndex
        Dim rowMax = dgv.Rows.Count '目標dgv的最大RowIndex
        Dim SelectRowCount = GetSelectedRowCount(dgv.SelectedCells) '選取範圍列數計算
        Dim StartRow = GetSelectedMinRowIndex(dgv.SelectedCells) '選取範圍的初始RowIndex
        Dim StartCol = GetSelectedMinColumnIndex(dgv.SelectedCells) '選取範圍的初始ColumnIndex

        '判斷剪貼簿內容是否大於一列
        '若是只有一列則可支援貼上複數目標列，否則會根據剪貼簿內容貼上固定列數
        If s_Row.Count = 1 Then
            For i As Integer = 0 To SelectRowCount - 1
                Dim values As String() = s_Row(0).Split({ControlChars.Tab}, StringSplitOptions.None)
                For k As Integer = 0 To values.Count - 1
                    If StartRow + i < rowMax AndAlso StartCol + k < colMax Then
                        If dgv(StartCol + k, StartRow + i).ReadOnly = IgnoreReadOnly OrElse IgnoreCol.Contains(dgv.Columns(StartCol + k).Name) Then Continue For
                        dgv.Rows(StartRow + i).Cells(StartCol + k).Value = values(k)
                    End If
                Next
            Next

        ElseIf s_Row.Count > 1 Then
            Dim CurrentRow = StartRow
            For Each r In s_Row
                Dim CurrentCol = StartCol
                For Each c In r.Split({ControlChars.Tab}, StringSplitOptions.None)
                    If CurrentCol >= colMax Then Exit For
                    If dgv(CurrentCol, CurrentRow).ReadOnly = IgnoreReadOnly OrElse IgnoreCol.Contains(dgv.Columns(CurrentCol).Name) Then Continue For
                    dgv(CurrentCol, CurrentRow).Value = c
                    CurrentCol += 1
                Next
                CurrentRow += 1
                If CurrentRow >= rowMax Then Exit For
            Next
        End If

    End Sub
    ''' <summary>
    ''' 將選取範圍儲存格內容以空白取代
    ''' </summary>
    ''' <param name="dgv">DataGridView，包含要刪除資料的目標範圍</param>
    Sub DeleteCells(ByRef dgv As DataGridView)
        For Each cell As DataGridViewCell In dgv.SelectedCells
            If cell.ReadOnly = True Then Continue For
            cell.Value = ""
        Next
    End Sub

    ''' <summary>
    ''' 計算目標dgv被選取的儲存格包含幾個Row
    ''' </summary>
    ''' <param name="cells">DataGridViewSelectedCellCollection，目標dgv被選取的儲存格</param>
    ''' <returns>選取範圍的Row數量</returns>
    Public Function GetSelectedRowCount(ByVal cells As DataGridViewSelectedCellCollection) As Integer
        Dim selectedRows As New HashSet(Of Integer)()
        For Each cell As DataGridViewCell In cells
            selectedRows.Add(cell.RowIndex)
        Next

        Return selectedRows.Count

    End Function

    ''' <summary>
    ''' 計算目標dgv被選取的儲存格最小RowIndex
    ''' </summary>
    ''' <param name="cells">DataGridViewSelectedCellCollection，目標dgv被選取的儲存格</param>
    ''' <returns>選取範圍的最小RowIndex</returns>
    Public Function GetSelectedMinRowIndex(ByVal cells As DataGridViewSelectedCellCollection) As Integer
        Dim selectedRows As New HashSet(Of Integer)()
        For Each cell As DataGridViewCell In cells
            selectedRows.Add(cell.RowIndex)
        Next

        Return selectedRows.Min

    End Function

    ''' <summary>
    ''' 計算目標dgv被選取的儲存格包含幾個Column
    ''' </summary>
    ''' <param name="cells">DataGridViewSelectedCellCollection，目標dgv被選取的儲存格</param>
    ''' <returns>選取範圍的Column數量</returns>
    Public Function GetSelectedColumnCount(ByVal cells As DataGridViewSelectedCellCollection) As Integer
        Dim selectedcols As New HashSet(Of Integer)()
        For Each cell As DataGridViewCell In cells
            selectedcols.Add(cell.RowIndex)
        Next

        Return selectedcols.Count

    End Function

    ''' <summary>
    ''' 計算目標dgv被選取的儲存格最小ColumnIndex
    ''' </summary>
    ''' <param name="cells">DataGridViewSelectedCellCollection，目標dgv被選取的儲存格</param>
    ''' <returns>選取範圍的最小ColumnIndex</returns>
    Public Function GetSelectedMinColumnIndex(ByVal cells As DataGridViewSelectedCellCollection) As Integer
        Dim selectedcols As New HashSet(Of Integer)()
        For Each cell As DataGridViewCell In cells
            selectedcols.Add(cell.RowIndex)
        Next

        Return selectedcols.Min

    End Function


End Module
