Imports System.IO
Imports System.Text

Module DataGridViewExport

    Public LogFolder As String = "C:\ProductionReportLog\Log\"

    Sub ExportToCSV(data As DataTable, logfolder As String, lot As String, layer As String, proc As String, index As String)
        Try
            Dim fileName As String = String.Format("{0}{1}_{2}_{3}{4}.csv", logfolder, index, proc, lot, layer)
            Dim sb As New StringBuilder()
            ' 寫入資料
            For Each row As DataRow In data.Rows
                For Each column As DataColumn In data.Columns
                    sb.Append(row(column).ToString()).Append(",")
                Next
                sb.Remove(sb.Length - 1, 1) ' 刪除最後一個逗號
                sb.AppendLine()
            Next

            ' 將 CSV 字串寫入檔案
            File.WriteAllText(fileName, sb.ToString(), Encoding.UTF8)
        Catch ex As Exception
            WriteLog(ex, LogFilePath, "ExportToCSV")
        End Try
    End Sub
    Sub ExportDataGridViewToDataTable(dataGridView As DataGridView)
        Try
            Dim Procs As String() = ReportUI.ProcInfo(ReportUI.AreaName.SelectedItem.ToString)(1).Split(",")
            Dim index As Integer = 0
            For Each Proc In Procs
                '刪除原檔案
                Dim fileQuery As String = "*" + Proc.Substring(0, 3) + "*" + Proc.Substring(3, 3)
                Dim FileToDelete As String() = Directory.GetFiles(LogFolder, fileQuery + "*.csv")
                For Each Files In FileToDelete
                    File.Delete(Files)
                Next
            Next

            For i = dataGridView.Rows.Count - 1 To 0 Step -1
                If Not dataGridView.Rows(i).IsNewRow Then
                    If dataGridView.Rows(i).Cells("上傳").Value IsNot Nothing AndAlso dataGridView.Rows(i).Cells("上傳").Value.ToString = "已上傳" Then
                        Continue For
                    End If
                    index += 1
                    Dim lot As String = dataGridView.Rows(i).Cells("批號").Value.ToString()
                    Dim proc As String = dataGridView.Rows(i).Cells("站點").Value.ToString()
                    Dim layer As String = dataGridView.Rows(i).Cells("層別").Value.ToString()
                    'Dim data As DataTable = CType(row.DataBoundItem, DataRowView).Row.Table
                    Dim data As New DataTable()

                    'Adding the Columns.
                    For Each column As DataGridViewColumn In dataGridView.Columns
                        If column.Index = dataGridView.Columns("btnModify").Index And column.Index = dataGridView.Columns("上傳").Index Then
                            Continue For
                        End If
                        data.Columns.Add(column.HeaderText)
                    Next

                    'Adding the Rows.

                    data.Rows.Add()
                    For Each cell As DataGridViewCell In dataGridView.Rows(i).Cells
                        If cell.ColumnIndex = dataGridView.Columns("btnModify").Index And cell.ColumnIndex = dataGridView.Columns("上傳").Index Then
                            Continue For
                        End If
                        If TypeOf cell.Value Is DBNull OrElse cell.Value Is Nothing OrElse String.IsNullOrEmpty(cell.Value.ToString()) Then
                            ' 當值為 DBNull、Nothing 或空字串時
                            data.Rows(data.Rows.Count - 1)(cell.ColumnIndex) = ""
                        Else
                            ' 當值不是 DBNull、Nothing 且不是空字串時
                            data.Rows(data.Rows.Count - 1)(cell.ColumnIndex) = cell.Value.ToString()
                        End If
                    Next

                    ExportToCSV(data, LogFolder, lot, layer, proc, index.ToString("000"))
                End If
            Next
        Catch ex As Exception
            WriteLog(ex, LogFilePath, "ExportDataGridViewToDataTable")
        End Try
    End Sub

    Sub ExportDataGridViewRowToDataTable(dgvrow As DataGridViewRow)
        Try
            If Not dgvrow.IsNewRow Then
                If dgvrow.Cells("上傳").Value IsNot Nothing AndAlso dgvrow.Cells("上傳").Value.ToString = "已上傳" Then
                    Return
                End If
                Dim lot As String = dgvrow.Cells("批號").Value.ToString()
                Dim proc As String = dgvrow.Cells("站點").Value.ToString()
                Dim layer As String = dgvrow.Cells("層別").Value.ToString()
                'Dim data As DataTable = CType(DataGridViewRow.DataBoundItem, DataRowView).Row.Table
                '獲取原檔案名稱
                Dim fileQuery As String = "*" + proc + "*" + lot + "*" + layer
                Dim FileName As String = Directory.GetFiles(LogFolder, fileQuery + "*.csv")(0)

                Dim data As New DataTable()

                'Adding the Columns.
                For Each column As DataGridViewColumn In dgvrow.DataGridView.Columns
                    If column.Index = dgvrow.DataGridView.Columns("btnModify").Index And column.Index = dgvrow.DataGridView.Columns("上傳").Index Then
                        Continue For
                    End If
                    data.Columns.Add(column.HeaderText)
                Next

                'Adding the Rows.

                data.Rows.Add()
                For Each cell As DataGridViewCell In dgvrow.Cells
                    If cell.ColumnIndex = dgvrow.DataGridView.Columns("btnModify").Index And cell.ColumnIndex = dgvrow.DataGridView.Columns("上傳").Index Then
                        Continue For
                    End If
                    If cell.Value = Nothing Then
                        data.Rows(data.Rows.Count - 1)(cell.ColumnIndex) = ""
                    Else
                        data.Rows(data.Rows.Count - 1)(cell.ColumnIndex) = cell.Value.ToString()
                    End If

                Next

                ExportToCSV(data, LogFolder, lot, layer, proc, CInt(FileName.Split("\")(3).Split("_")(0)).ToString("000"))
            End If
        Catch ex As Exception
            WriteLog(ex, LogFilePath, "ExportDataGridViewRowToDataTable")
        End Try
    End Sub
End Module
