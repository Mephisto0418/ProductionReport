Imports System.IO
Imports System.Net
Imports System.Diagnostics
Imports System.Data
Imports System.Data.SqlClient
Imports log4net
Imports log4net.Config
Imports log4net.Appender
Imports log4net.Repository.Hierarchy


Public Class Mylog4net

    Shared ReadOnly workingDirectory As String = AppDomain.CurrentDomain.BaseDirectory
    Shared ReadOnly configFilePath As String = workingDirectory & "log4net.config" 'log4net配置文件名稱

    Public Property LogFilePath As String = "C:\ProductionReport\Update_Log\" 'Log存放資料夾
    Public Property LogFileName As String = ""

    ''' <summary>
    ''' 請先設定[LogFilePath]以及[LogFileName]再進行ConfigInit
    ''' </summary>
    Sub New()
        '建立Log資料夾
        If Not Directory.Exists(LogFilePath) Then
            Directory.CreateDirectory(LogFilePath)
        End If

        ' 記錄未處理的例外
        AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf UnhandledExceptionTrapper

    End Sub

    ''' <summary>
    ''' 請確認已設定[LogFilePath]以及[LogFileName]
    ''' </summary>
    Sub ConfigInit()
        ' 檢查配置文件是否存在
        If Not File.Exists(configFilePath) Then
            ' 如果配置文件不存在，則生成一份新的配置文件
            GenerateLog4netConfig(configFilePath, LogFilePath, LogFileName)
        End If

        'Log4net配置文件導入
        Dim logConfigFile As New FileInfo(configFilePath)
        XmlConfigurator.ConfigureAndWatch(logConfigFile)

        ' 動態設置 AdoNetAppender 的連接字符串
        Dim hierarchy As Hierarchy = CType(LogManager.GetRepository(), Hierarchy)
        Dim adoNetAppender As AdoNetAppender = CType(hierarchy.GetAppenders().FirstOrDefault(Function(appender) TypeOf appender Is AdoNetAppender), AdoNetAppender)
        If adoNetAppender IsNot Nothing Then
            adoNetAppender.ConnectionString = SQL_Conn_MQL03
            adoNetAppender.ActivateOptions()
        End If

        '設定Log參數初始化
        GlobalContext.Properties("HostName") = Dns.GetHostName()
        GlobalContext.Properties("Process") = Process.GetCurrentProcess().ProcessName
        GlobalContext.Properties("AreaID") = Nothing
        GlobalContext.Properties("IP") = GetLocalIPAddress()

    End Sub

    Sub GenerateLog4netConfig(FilePath As String, LogPath As String, LogName As String)
        '配置文件內容設定

        Dim configContent As String = "<?xml version=""1.0"" encoding=""utf-8"" ?>" & vbCrLf &
                                      "<configuration>" & vbCrLf &
                                      "  <log4net>" & vbCrLf &
                                      "    <appender name=""DefaultAppender"" type=""log4net.Appender.RollingFileAppender"">" & vbCrLf &
                                      "      <file type=""log4net.Util.PatternString"" value=""" & LogPath & LogName & "Log.log"" />    <!--存放log路徑-->" & vbCrLf &
                                      "      <staticLogFileName value=""true""/>" & vbCrLf &
                                      "      <appendToFile value=""true""/>" & vbCrLf &
                                      "      <rollingStyle value=""Size""/>" & vbCrLf &
                                      "      <maxSizeRollBackups value=""5""/>" & vbCrLf &
                                      "      <maximumFileSize value=""10MB""/>" & vbCrLf &
                                      "      <layout type=""log4net.Layout.PatternLayout"">" & vbCrLf &
                                      "        <conversionPattern value=""%-5p [%thread] %date{yyyy/MM/dd HH:mm:ss} %-20c{1} %-20M ID : %property{AreaID} %m%n"" />    <!--內容格式-->" & vbCrLf &
                                      "      </layout>" & vbCrLf &
                                      "    </appender>" & vbCrLf &
                                      vbCrLf &
                                      "    <appender name=""ExceptionAppender"" type=""log4net.Appender.RollingFileAppender"">" & vbCrLf &
                                      "      <file type=""log4net.Util.PatternString"" value=""" & LogPath & LogName & "ErrorLog.log"" />    <!--存放log路徑-->" & vbCrLf &
                                      "      <staticLogFileName value=""true""/>" & vbCrLf &
                                      "      <appendToFile value=""true""/>" & vbCrLf &
                                      "      <rollingStyle value=""Size""/>" & vbCrLf &
                                      "      <maxSizeRollBackups value=""5""/>" & vbCrLf &
                                      "      <maximumFileSize value=""10MB""/>" & vbCrLf &
                                      "      <layout type=""log4net.Layout.PatternLayout"">" & vbCrLf &
                                      "        <conversionPattern value=""%-5p [%thread] %date{yyyy/MM/dd HH:mm:ss} %-20c{1} %-20M %file:%line  ID : %property{AreaID} %m%n"" />    <!--內容格式-->" & vbCrLf &
                                      "      </layout>" & vbCrLf &
                                      "    </appender>" & vbCrLf &
                                      vbCrLf &
                                      "    <appender name=""AdoNetAppender"" type=""log4net.Appender.AdoNetAppender"">" & vbCrLf &
                                      "      <bufferSize value=""1"" />" & vbCrLf &
                                      "      <connectionType value=""System.Data.SqlClient.SqlConnection, System.Data, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"" />" & vbCrLf &
                                      "      <connectionString value=""Data Source=10.44.66.105;Initial Catalog=master;User ID=pptcim;Password=cim945"" />" & vbCrLf &
                                      "      <commandText value=""INSERT INTO [H3_Systematic].[dbo].[H3_Production_Record] ([SystemTime],[IP],[HostName],[Process],[LogLevel],[FileName],[Methods],[Line],[AreaID],[Exception],[Messages]) VALUES (@log_date, @ip_address, @host_name, @process, @log_level, @file, @method, @line, @areaID, @exception, @message)"" />" & vbCrLf &
                                      "      <parameter>" & vbCrLf &
                                      "        <parameterName value=""@log_date"" />" & vbCrLf &
                                      "        <dbType value=""DateTime"" />" & vbCrLf &
                                      "        <layout type=""log4net.Layout.RawTimeStampLayout"" />	" & vbCrLf &
                                      "      </parameter>" & vbCrLf &
                                      "      <parameter>" & vbCrLf &
                                      "        <parameterName value=""@ip_address"" />" & vbCrLf &
                                      "        <dbType value=""String"" />" & vbCrLf &
                                      "        <size value=""20"" />" & vbCrLf &
                                      "        <layout type=""log4net.Layout.PatternLayout"">" & vbCrLf &
                                      "          <conversionPattern value=""%property{IP}"" />" & vbCrLf &
                                      "        </layout>" & vbCrLf &
                                      "      </parameter>" & vbCrLf &
                                      "      <parameter>" & vbCrLf &
                                      "        <parameterName value=""@host_name"" />" & vbCrLf &
                                      "        <dbType value=""String"" />" & vbCrLf &
                                      "        <size value=""20"" />" & vbCrLf &
                                      "        <layout type=""log4net.Layout.PatternLayout"">" & vbCrLf &
                                      "          <conversionPattern value=""%property{HostName}"" />" & vbCrLf &
                                      "        </layout>" & vbCrLf &
                                      "      </parameter>" & vbCrLf &
                                      "      <parameter>" & vbCrLf &
                                      "        <parameterName value=""@process"" />" & vbCrLf &
                                      "        <dbType value=""String"" />" & vbCrLf &
                                      "        <size value=""50"" />" & vbCrLf &
                                      "        <layout type=""log4net.Layout.PatternLayout"">" & vbCrLf &
                                      "          <conversionPattern value=""%property{Process}"" />" & vbCrLf &
                                      "        </layout>" & vbCrLf &
                                      "      </parameter>" & vbCrLf &
                                      "      <parameter>" & vbCrLf &
                                      "        <parameterName value=""@log_level"" />" & vbCrLf &
                                      "        <dbType value=""String"" />" & vbCrLf &
                                      "        <size value=""10"" />" & vbCrLf &
                                      "        <layout type=""log4net.Layout.PatternLayout"">" & vbCrLf &
                                      "          <conversionPattern value=""%level"" />" & vbCrLf &
                                      "        </layout>" & vbCrLf &
                                      "      </parameter>" & vbCrLf &
                                      "      <parameter>" & vbCrLf &
                                      "        <parameterName value=""@file"" />" & vbCrLf &
                                      "        <dbType value=""String"" />" & vbCrLf &
                                      "        <size value=""255"" />" & vbCrLf &
                                      "        <layout type=""log4net.Layout.PatternLayout"">" & vbCrLf &
                                      "          <conversionPattern value=""%file"" />" & vbCrLf &
                                      "        </layout>" & vbCrLf &
                                      "      </parameter>" & vbCrLf &
                                      "      <parameter>" & vbCrLf &
                                      "        <parameterName value=""@method"" />" & vbCrLf &
                                      "        <dbType value=""String"" />" & vbCrLf &
                                      "        <size value=""50"" />" & vbCrLf &
                                      "        <layout type=""log4net.Layout.PatternLayout"">" & vbCrLf &
                                      "          <conversionPattern value=""%M"" />" & vbCrLf &
                                      "        </layout>" & vbCrLf &
                                      "      </parameter>" & vbCrLf &
                                      "      <parameter>" & vbCrLf &
                                      "        <parameterName value=""@line"" />" & vbCrLf &
                                      "        <dbType value=""Int32"" />" & vbCrLf &
                                      "        <layout type=""log4net.Layout.PatternLayout"">" & vbCrLf &
                                      "          <conversionPattern value=""%line"" />" & vbCrLf &
                                      "        </layout>" & vbCrLf &
                                      "      </parameter>" & vbCrLf &
                                      "      <parameter>" & vbCrLf &
                                      "        <parameterName value=""@areaID"" />" & vbCrLf &
                                      "        <dbType value=""Int32"" />" & vbCrLf &
                                      "        <layout type=""log4net.Layout.PatternLayout"">" & vbCrLf &
                                      "          <conversionPattern value=""%property{AreaID}"" />" & vbCrLf &
                                      "        </layout>" & vbCrLf &
                                      "      </parameter>" & vbCrLf &
                                      "      <parameter>" & vbCrLf &
                                      "        <parameterName value=""@exception"" />" & vbCrLf &
                                      "        <dbType value=""String"" />" & vbCrLf &
                                      "        <size value=""2000"" />" & vbCrLf &
                                      "        <layout type=""log4net.Layout.ExceptionLayout"" />" & vbCrLf &
                                      "      </parameter>" & vbCrLf &
                                      "      <parameter>" & vbCrLf &
                                      "        <parameterName value=""@message"" />" & vbCrLf &
                                      "        <dbType value=""String"" />" & vbCrLf &
                                      "        <size value=""4000"" />" & vbCrLf &
                                      "        <layout type=""log4net.Layout.PatternLayout"">" & vbCrLf &
                                      "          <conversionPattern value=""%message"" />" & vbCrLf &
                                      "        </layout>" & vbCrLf &
                                      "      </parameter>" & vbCrLf &
                                      "    </appender>" & vbCrLf &
                                      vbCrLf &
                                      "    <logger name = ""SQL"">" & vbCrLf &
                                      "      <level value=""ALL""/>" & vbCrLf &
                                      "      <appender-ref ref=""AdoNetAppender""/>" & vbCrLf &
                                      "    </logger>" & vbCrLf &
                                      vbCrLf &
                                      "    <logger name = ""Error"">" & vbCrLf &
                                      "      <level value=""ERROR""/>" & vbCrLf &
                                      "      <appender-ref ref=""ExceptionAppender""/>" & vbCrLf &
                                      "      <appender-ref ref=""AdoNetAppender""/>" & vbCrLf &
                                      "    </logger>" & vbCrLf &
                                      vbCrLf &
                                      "    <root>" & vbCrLf &
                                      "      <level value=""ALL""/>" & vbCrLf &
                                      "      <appender-ref ref=""DefaultAppender""/>" & vbCrLf &
                                      "    </root>" & vbCrLf &
                                      "  </log4net>" & vbCrLf &
                                      "</configuration>"

        ' 將配置內容寫入文件
        File.WriteAllText(FilePath, configContent)
    End Sub

    Function GetLocalIPAddress() As String
        Dim host As IPHostEntry = Dns.GetHostEntry(Dns.GetHostName())
        For Each ip As IPAddress In host.AddressList
            If ip.AddressFamily = Sockets.AddressFamily.InterNetwork Then
                Return ip.ToString()
            End If
        Next
        Throw New Exception("No network adapters with an IPv4 address in the system!")
    End Function

    Sub UnhandledExceptionTrapper(sender As Object, e As UnhandledExceptionEventArgs)
        log_Error.Error("未處理的例外：", CType(e.ExceptionObject, Exception))
        Environment.Exit(1)
    End Sub


    Public Class SetProperty
        Public Shared Property AreaID As String
            Get
                Return GlobalContext.Properties("AreaID").ToString
            End Get
            Set(value As String)
                If value Is Nothing Then
                    GlobalContext.Properties("AreaID") = Nothing
                Else
                    GlobalContext.Properties("AreaID") = value
                End If
            End Set
        End Property
    End Class
End Class
