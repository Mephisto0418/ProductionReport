Imports System.IO

Module ErrorLog
    Public LogFilePath As String = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) & "\ErrorLog\"
    Public Sub WriteLog(ByVal ex As Exception, ByVal ErrorArea As String)
        Dim Path As String = ""

        '取得行數
        Dim stackTrace As New StackTrace(ex, True)
        Dim frame As StackFrame = stackTrace.GetFrame(0)
        Dim lineNumber As Integer = frame.GetFileLineNumber()

        If Not Directory.Exists(LogFilePath) Then
            Directory.CreateDirectory(LogFilePath)
        End If

        Dim currentProcess As Process = Process.GetCurrentProcess()
        Dim cpuUsage As Integer = CInt(currentProcess.TotalProcessorTime.TotalMilliseconds / Environment.ProcessorCount / Environment.TickCount * 100)
        Dim ErrorMessages As String = "==============================" & vbCrLf &
                                      "Error Time: " & DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") & vbCrLf &
                                      "Error Location: " & ex.TargetSite.ToString & vbCrLf &
                                      "Error Line: " & lineNumber.ToString & vbCrLf &
                                      "Error Area : " & ErrorArea & vbCrLf &
                                      "CPU Usage: " & cpuUsage.ToString & vbCrLf &
                                      "Memory Usage: " & My.Computer.Info.AvailablePhysicalMemory.ToString & vbCrLf &
                                      "Error Message: " & ex.Message & vbCrLf & ex.ToString & vbCrLf
        Try
            Path = LogFilePath & "ErrorLog_" & DateTime.Now.ToString("yyyy-MM-dd") & ".txt"
            Using writer As New StreamWriter(Path, True)
                writer.WriteLine(ErrorMessages)
            End Using
        Catch ex2 As Exception

        Finally
            MessageBox.Show(ErrorMessages)
        End Try
    End Sub

    Public Sub ChangeLog(ByVal Targrt As String, ByVal ID As String, ByVal Name As String, ByVal EventType As String, ByVal Status As String)
        Try
            Dim PCName As String = My.Computer.Name  '電腦名稱
            Dim IP As String = ""
            Dim Address() As System.Net.IPAddress
            Address = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName()).AddressList
            Dim i As Integer
            For i = 0 To UBound(Address)
                If Address(i).ToString.Contains(".") Then
                    IP = Address(i).ToString()
                End If
            Next
            Dim cmd As String = "INSERT INTO [H3_Systematic].[dbo].[H3_Production_SettingRecord]
                                             ([SystemTime],[IP],[PCName],[SettingTarget],[SettingID],[SettingName],[Event],[Status])
                                             VALUES('" & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") & "','" & IP & "','" & PCName & "','" & Targrt & "','" & ID & "','" & Name & "','" & EventType & "','" & Status & "')"
            SQL_Query(SQL_Conn_MQL03, cmd)
        Catch ex As Exception
            WriteLog(ex, "ChangeLog")
        End Try
    End Sub
End Module
