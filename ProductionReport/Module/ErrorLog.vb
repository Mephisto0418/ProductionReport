Imports System.IO

Module ErrorLog
    'Public LogFilePath As String = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) & "\Log\"
    Public LogFilePath As String = "C:\ProductionReport\Log\"
    Sub WriteLog(ByVal ex As Exception, ByVal LogFilePath As String, ByVal ErrorArea As String)
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
            LogFilePath = LogFilePath & "ErrorLog_" & DateTime.Now.ToString("yyyy-MM-dd") & ".txt"
            Using writer As New StreamWriter(LogFilePath, True)
                writer.WriteLine(ErrorMessages)
            End Using
        Catch ex2 As Exception

        Finally
            MessageBox.Show(ErrorMessages)
        End Try
    End Sub
End Module
