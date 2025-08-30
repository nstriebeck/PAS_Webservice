Imports System.Configuration
Imports System.IO

Module ServiceConfig
    ' Konfigurationswerte
    Public ReadOnly Property ConnectionString As String
        Get
            Dim connStr As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("PAS_Database")
            If connStr IsNot Nothing Then
                Return connStr.ConnectionString
            Else
                Return "Server=SILINSQL\PatientenAufruf;Database=PAS_Database;User Id=sa;Password=PatientenAufruf4711;"
            End If
        End Get
    End Property

    Public ReadOnly Property WebServerPort As Integer
        Get
            Dim port As String = ConfigurationManager.AppSettings("WebServerPort")
            If String.IsNullOrEmpty(port) Then
                Return 9090
            Else
                Return Integer.Parse(port)
            End If
        End Get
    End Property

    Public ReadOnly Property EnableLogging As Boolean
        Get
            Dim setting As String = ConfigurationManager.AppSettings("EnableLogging")
            If String.IsNullOrEmpty(setting) Then
                Return True
            Else
                Return Boolean.Parse(setting)
            End If
        End Get
    End Property

    Public ReadOnly Property LogPath As String
        Get
            Dim path As String = ConfigurationManager.AppSettings("LogPath")
            If String.IsNullOrEmpty(path) Then
                path = "C:\PAS\Logs"
            End If

            If Not Directory.Exists(path) Then
                Directory.CreateDirectory(path)
            End If
            Return path
        End Get
    End Property

    ' Logging-Funktionen
    Public Sub LogInfo(message As String)
        If EnableLogging Then
            WriteLog("INFO", message)
        End If
    End Sub

    Public Sub LogError(message As String)
        If EnableLogging Then
            WriteLog("ERROR", message)
        End If
    End Sub

    Public Sub LogWarning(message As String)
        If EnableLogging Then
            WriteLog("WARN", message)
        End If
    End Sub

    Private Sub WriteLog(level As String, message As String)
        Try
            Dim logFile As String = Path.Combine(LogPath, $"PAS_WebService_{DateTime.Now:yyyy-MM-dd}.log")
            Dim logEntry As String = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{level}] {message}"

            File.AppendAllText(logFile, logEntry & Environment.NewLine)

            ' Auch in Konsole ausgeben
            Console.WriteLine(logEntry)
        Catch
            ' Logging-Fehler ignorieren
        End Try
    End Sub

    ' Service-Status
    Public Sub DisplayStartupInfo()
        Console.WriteLine("=====================================")
        Console.WriteLine("PAS Web-Service v1.0")
        Console.WriteLine("=====================================")
        Console.WriteLine($"Server Port: {WebServerPort}")
        Console.WriteLine($"Logging: {If(EnableLogging, "Aktiviert", "Deaktiviert")}")
        Console.WriteLine($"Log-Pfad: {LogPath}")
        Console.WriteLine($"Datenbank: {GetDatabaseName()}")
        Console.WriteLine("=====================================")
        Console.WriteLine()
    End Sub

    Private Function GetDatabaseName() As String
        Try
            Dim builder As New SqlClient.SqlConnectionStringBuilder(ConnectionString)
            Return $"{builder.DataSource}\{builder.InitialCatalog}"
        Catch
            Return "Unbekannt"
        End Try
    End Function
End Module