Imports System.Data.SqlClient

Module SQL
    Public SQL_Conn_MQL03 As String = "Data Source=10.44.66.105;Initial Catalog=master;User ID=pptcim;Password=cim945"

    Friend Function SQL_Query(ByVal cmd As String)
        Dim conn As New SqlConnection(SQL_Conn_MQL03)
        Dim Command As SqlCommand = New SqlCommand(cmd, conn) With {
            .CommandTimeout = 300
        }
        conn.Open()
        Dim rs As New SqlDataAdapter(Command)
        Dim ds As New DataSet()
        Dim dt As New DataTable()
        rs.Fill(ds)
        conn.Close()
        If ds.Tables.Count > 1 Then
            Return ds
        ElseIf ds.Tables.Count = 1 Then
            dt = ds.Tables(0)
            Return dt
        Else
            Return dt
        End If
    End Function

    '20231017 Boris 新增預存程序專用語法

    Friend Function SQL_StoredProcedure(ByVal StoredProcedure As String, ByVal Parameters As List(Of String()))
        ' 建立變數
        Dim ds As New DataSet()
        Dim dt As New DataTable()

        ' 建立連線

        Using Connection As New SqlConnection(SQL_Conn_MQL03)
            Connection.Open()

            Using cmdArithAbort As New SqlCommand("SET ARITHABORT ON", Connection)
                cmdArithAbort.ExecuteNonQuery()
            End Using

            ' 建立命令物件，並指定預存程序名稱
            Using Command As New SqlCommand(StoredProcedure, Connection)

                Command.CommandType = CommandType.StoredProcedure
                Command.CommandTimeout = 300

                ' 如果預存程序有輸入參數，請添加它們
                For Each Parameter In Parameters
                    Command.Parameters.Add(New SqlParameter(Parameter(0), SqlDbType.NVarChar) With {.Value = Parameter(1)})
                Next


                ' 建立 SqlDataAdapter 來填充 DataSet
                Dim adapter As New SqlDataAdapter(Command)

                ' 使用 SqlDataAdapter 執行預存程序並填充 DataSet
                adapter.Fill(ds)
                If ds.Tables.Count > 1 Then
                    Return ds
                ElseIf ds.Tables.Count = 1 Then
                    dt = ds.Tables(0)
                    Return dt
                Else
                    Return dt
                End If

                ' 現在您可以處理資料或顯示結果
            End Using
        End Using
    End Function

End Module
