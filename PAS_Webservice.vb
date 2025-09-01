Imports System.ComponentModel
Imports System.Configuration.Install
Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports System.ServiceProcess
Imports System.Text
Imports System.Threading
Imports Newtonsoft.Json
Imports System.Web


Public Class PAS_WebService
    Inherits ServiceBase

    Private _httpListener As HttpListener
    Private _listenerThread As Thread
    Private _isRunning As Boolean = False

    Public Sub New()
        Me.ServiceName = "PAS-Service"
        Me.CanStop = True
        Me.CanPauseAndContinue = False
        Me.AutoLog = True
    End Sub

    Protected Overrides Sub OnStart(args() As String)
        _isRunning = True
        _listenerThread = New Thread(AddressOf RunServer)
        _listenerThread.Start()
        EventLog.WriteEntry("PAS-Service", "Service gestartet", EventLogEntryType.Information)
    End Sub

    Private Sub RunServer()
        Try
            _httpListener = New HttpListener()
            _httpListener.Prefixes.Add("http://+:9090/")
            _httpListener.Start()

            EventLog.WriteEntry("PAS-Service", "HTTP Listener gestartet auf Port 9090", EventLogEntryType.Information)

            While _isRunning AndAlso _httpListener.IsListening
                Try
                    Dim context = _httpListener.GetContext()
                    ThreadPool.QueueUserWorkItem(Sub() ProcessRequest(context))
                Catch ex As HttpListenerException
                    If _isRunning Then
                        EventLog.WriteEntry("PAS-Service", "Listener Error: " & ex.Message, EventLogEntryType.Warning)
                    End If
                Catch ex As Exception
                    If _isRunning Then
                        EventLog.WriteEntry("PAS-Service", "Error: " & ex.Message, EventLogEntryType.Error)
                    End If
                End Try
            End While

        Catch ex As Exception
            EventLog.WriteEntry("PAS-Service", "RunServer Error: " & ex.ToString(), EventLogEntryType.Error)
        End Try
    End Sub

    Private Sub ProcessRequest(context As HttpListenerContext)
        Try
            Dim request = context.Request
            Dim response = context.Response
            Dim html As String = ""

            Select Case request.HttpMethod
                Case "GET"
                    Select Case request.Url.AbsolutePath
                        Case "/", "/waiting"
                            html = GetWaitingRoomHTML()
                        Case "/api/patients", "/api/warteschlange"
                            Dim datum As String = Nothing
                            If request.Url.Query.Contains("datum=") Then
                                datum = request.QueryString("datum")
                            End If
                            html = GetPatientsJSON(datum)
                            response.ContentType = "application/json"
                        Case "/api/wartezimmer-anonym"
                            ' Anonymisierte Daten für öffentliche Anzeige
                            html = GetAnonymePatientenJSON()
                            response.ContentType = "application/json"
                        Case Else
                            If request.Url.AbsolutePath.StartsWith("/patient/") Then
                                Dim patNr = request.Url.AbsolutePath.Replace("/patient/", "")
                                html = GetPatientHTML(patNr)
                            Else
                                response.StatusCode = 404
                                html = "<html><body><h1>404 - Seite nicht gefunden</h1></body></html>"
                            End If
                    End Select

                Case "POST"
                    Select Case request.Url.AbsolutePath
                        Case "/api/neuerpatient"
                            Try
                                ' POST-Daten lesen
                                Dim reader As New StreamReader(request.InputStream)
                                Dim postData = reader.ReadToEnd()
                                Dim formData = ParseFormData(postData)

                                Using conn As New SqlConnection(ConnectionString)
                                    conn.Open()
                                    EventLog.WriteEntry("PAS-Service", $"Neuer Patient empfangen: {formData("patientenID")}", EventLogEntryType.Information)

                                    Dim query = "INSERT INTO dbo.Warteschlange (PatNr, Name, Status, Bereich, Ankunft, Prioritaet, Bemerkung, Wartezeit) 
VALUES (@patnr, @name, @status, @bereich, @ankunft, @prio, @bemerkung, 0)"

                                    Using cmd As New SqlCommand(query, conn)
                                        ' Nutze explizite Typ-Deklaration:
                                        cmd.Parameters.Add("@patnr", SqlDbType.NVarChar, 20).Value = formData("patientenID")
                                        cmd.Parameters.Add("@name", SqlDbType.NVarChar, 100).Value = formData("name")
                                        cmd.Parameters.Add("@status", SqlDbType.NVarChar, 20).Value = If(formData.ContainsKey("status"), formData("status"), "Wartend")
                                        cmd.Parameters.Add("@bereich", SqlDbType.NVarChar, 50).Value = If(formData.ContainsKey("zimmer"), formData("zimmer"), "Wartezimmer")
                                        cmd.Parameters.Add("@bemerkung", SqlDbType.NVarChar, 500).Value = If(formData.ContainsKey("bemerkung"), formData("bemerkung"), "")

                                        ' Ankunftszeit aus dem Form-Parameter holen
                                        Dim ankunftszeitStr = If(formData.ContainsKey("ankunftszeit"), formData("ankunftszeit"), "")
                                        Dim ankunftszeit As DateTime

                                        If Not String.IsNullOrEmpty(ankunftszeitStr) AndAlso DateTime.TryParse(ankunftszeitStr, ankunftszeit) Then
                                            cmd.Parameters.AddWithValue("@ankunft", ankunftszeit)
                                        Else
                                            cmd.Parameters.AddWithValue("@ankunft", DateTime.Now)
                                        End If

                                        cmd.Parameters.AddWithValue("@prio", If(formData.ContainsKey("prioritaet"), Integer.Parse(formData("prioritaet")), 0))


                                        cmd.ExecuteNonQuery()
                                    End Using
                                End Using

                                html = "{""success"": true, ""message"": ""Patient hinzugefügt""}"
                                response.ContentType = "application/json"
                                response.StatusCode = 200

                            Catch ex As Exception
                                html = "{""success"": false, ""error"": """ & ex.Message.Replace("""", "'") & """}"
                                response.ContentType = "application/json"
                                response.StatusCode = 500
                                EventLog.WriteEntry("PAS-Service", "Fehler beim Hinzufügen: " & ex.ToString(), EventLogEntryType.Error)
                            End Try

                        Case "/api/statusupdate"
                            Try
                                ' POST-Daten lesen
                                Dim reader As New StreamReader(request.InputStream)
                                Dim postData = reader.ReadToEnd()
                                Dim formData = ParseFormData(postData)

                                Using conn As New SqlConnection(ConnectionString)
                                    conn.Open()

                                    Dim query As String = ""
                                    Dim status = formData("status")

                                    Select Case status
                                        Case "Aufgerufen"
                                            query = "UPDATE dbo.Warteschlange SET Status = @status, Aufgerufen = @zeit WHERE PatNr = @patnr"
                                        Case "InBehandlung"
                                            query = "UPDATE dbo.Warteschlange SET Status = @status, Behandlungsbeginn = @zeit WHERE PatNr = @patnr"
                                        Case "Fertig"
                                            query = "UPDATE dbo.Warteschlange SET Status = @status, Behandlungsende = @zeit WHERE PatNr = @patnr"
                                        Case Else
                                            query = "UPDATE dbo.Warteschlange SET Status = @status WHERE PatNr = @patnr"
                                    End Select

                                    Using cmd As New SqlCommand(query, conn)
                                        cmd.Parameters.AddWithValue("@patnr", formData("patientenID"))
                                        cmd.Parameters.AddWithValue("@status", status)

                                        If formData.ContainsKey("zeitpunkt") Then
                                            Dim zeitpunkt As DateTime
                                            If DateTime.TryParse(formData("zeitpunkt"), zeitpunkt) Then
                                                cmd.Parameters.AddWithValue("@zeit", zeitpunkt)
                                            Else
                                                cmd.Parameters.AddWithValue("@zeit", DateTime.Now)
                                            End If
                                        End If

                                        cmd.ExecuteNonQuery()
                                    End Using
                                End Using

                                html = "{""success"": true}"
                                response.ContentType = "application/json"
                                response.StatusCode = 200

                            Catch ex As Exception
                                html = "{""success"": false, ""error"": """ & ex.Message.Replace("""", "'") & """}"
                                response.ContentType = "application/json"
                                response.StatusCode = 500
                                EventLog.WriteEntry("PAS-Service", "StatusUpdate Fehler: " & ex.ToString(), EventLogEntryType.Error)
                            End Try

                        Case "/api/zimmerwechsel"
                            Try
                                Dim reader As New StreamReader(request.InputStream)
                                Dim postData = reader.ReadToEnd()
                                Dim formData = ParseFormData(postData)

                                Using conn As New SqlConnection(ConnectionString)
                                    conn.Open()

                                    Dim query = "UPDATE dbo.Warteschlange SET Bereich = @zimmer WHERE PatNr = @patnr"

                                    Using cmd As New SqlCommand(query, conn)
                                        cmd.Parameters.AddWithValue("@patnr", formData("patientenID"))
                                        cmd.Parameters.AddWithValue("@zimmer", formData("zimmer"))
                                        cmd.ExecuteNonQuery()
                                    End Using
                                End Using

                                html = "{""success"": true}"
                                response.ContentType = "application/json"
                                response.StatusCode = 200

                            Catch ex As Exception
                                html = "{""success"": false, ""error"": """ & ex.Message.Replace("""", "'") & """}"
                                response.ContentType = "application/json"
                                response.StatusCode = 500
                            End Try

                        Case "/api/checkin"
                            Try
                                Dim reader As New StreamReader(request.InputStream)
                                Dim postData = reader.ReadToEnd()
                                Dim formData = ParseFormData(postData)

                                Using conn As New SqlConnection(ConnectionString)
                                    conn.Open()

                                    ' Wartezeit auf 0 und Ankunft auf jetzt setzen
                                    Dim query = "UPDATE dbo.Warteschlange 
                        SET Status = 'Wartend',
                            Ankunft = @ankunft,
                            Wartezeit = 0
                        WHERE PatNr = @patnr"

                                    Using cmd As New SqlCommand(query, conn)
                                        cmd.Parameters.AddWithValue("@patnr", formData("patientenID"))
                                        cmd.Parameters.AddWithValue("@ankunft", DateTime.Parse(formData("ankunftszeit")))

                                        Dim rowsAffected = cmd.ExecuteNonQuery()

                                        If rowsAffected > 0 Then
                                            EventLog.WriteEntry("PAS-Service",
                                                $"Check-In: Patient {formData("patientenID")} - Wartezeit auf 0 gesetzt",
                                                EventLogEntryType.Information)
                                            html = "{""success"": true, ""message"": ""Check-In erfolgreich""}"
                                        Else
                                            html = "{""success"": false, ""message"": ""Patient nicht gefunden""}"
                                            response.StatusCode = 404
                                        End If
                                    End Using
                                End Using

                                response.ContentType = "application/json"

                            Catch ex As Exception
                                html = "{""success"": false, ""error"": """ & ex.Message.Replace("""", "'") & """}"
                                response.ContentType = "application/json"
                                response.StatusCode = 500
                                EventLog.WriteEntry("PAS-Service", "Check-In Fehler: " & ex.ToString(), EventLogEntryType.Error)
                            End Try

                        Case "/api/aufgerufene"
                            Try
                                Dim aufgerufene As New List(Of Object)

                                Using conn As New SqlConnection(ConnectionString)
                                    conn.Open()

                                    ' Alle aufgerufenen und in Behandlung befindlichen Patienten
                                    Dim query = "SELECT PatNr, Name, Bereich, Status, Aufgerufen, Prioritaet,
                               CASE WHEN Prioritaet = 2 THEN 'Notfall' 
                                    WHEN Prioritaet = 1 THEN 'Dringend'
                                    ELSE 'Normal' END as PrioritaetText
                        FROM dbo.Warteschlange 
                        WHERE Status IN ('Aufgerufen', 'InBehandlung')
                          AND CAST(Ankunft AS DATE) = CAST(GETDATE() AS DATE)
                          AND (Aufgerufen IS NULL OR Aufgerufen > DATEADD(minute, -30, GETDATE()))
                        ORDER BY Aufgerufen DESC"

                                    Using cmd As New SqlCommand(query, conn)
                                        Using reader = cmd.ExecuteReader()
                                            While reader.Read()
                                                aufgerufene.Add(New With {
                                                    .PatientenID = reader("PatNr").ToString(),
                                                    .Name = reader("Name").ToString(),
                                                    .Zimmer = reader("Bereich").ToString(),
                                                    .Status = reader("Status").ToString(),
                                                    .Aufrufzeit = If(IsDBNull(reader("Aufgerufen")),
                                                                   DateTime.Now,
                                                                   Convert.ToDateTime(reader("Aufgerufen"))),
                                                    .Prioritaet = reader("PrioritaetText").ToString()
                                                })
                                            End While
                                        End Using
                                    End Using
                                End Using

                                html = JsonConvert.SerializeObject(aufgerufene)
                                response.ContentType = "application/json"

                            Catch ex As Exception
                                html = "{""error"": """ & ex.Message.Replace("""", "'") & """}"
                                response.ContentType = "application/json"
                                response.StatusCode = 500
                                EventLog.WriteEntry("PAS-Service", "Fehler bei /api/aufgerufene: " & ex.ToString(), EventLogEntryType.Error)
                            End Try

                        Case "/api/updatepatient"
                            Try
                                Dim reader As New StreamReader(request.InputStream)
                                Dim postData = reader.ReadToEnd()
                                Dim formData = ParseFormData(postData)

                                Using conn As New SqlConnection(ConnectionString)
                                    conn.Open()

                                    Dim query = "UPDATE dbo.Warteschlange SET " &
                                               "Name = @name, " &
                                               "Prioritaet = @prio, " &
                                               "Bereich = @zimmer, " &
                                               "Bemerkung = @bemerkung " &
                                               "WHERE PatNr = @patnr"

                                    Using cmd As New SqlCommand(query, conn)
                                        cmd.Parameters.AddWithValue("@patnr", formData("patientenID"))
                                        cmd.Parameters.AddWithValue("@name", If(formData.ContainsKey("name"), formData("name"), ""))
                                        cmd.Parameters.AddWithValue("@prio", If(formData.ContainsKey("prioritaet"), Integer.Parse(formData("prioritaet")), 0))
                                        cmd.Parameters.AddWithValue("@zimmer", If(formData.ContainsKey("zimmer"), formData("zimmer"), ""))
                                        cmd.Parameters.AddWithValue("@bemerkung", If(formData.ContainsKey("bemerkung"), formData("bemerkung"), ""))

                                        cmd.ExecuteNonQuery()
                                    End Using
                                End Using

                                html = "{""success"": true, ""message"": ""Patient aktualisiert""}"
                                response.ContentType = "application/json"
                                response.StatusCode = 200

                            Catch ex As Exception
                                html = "{""success"": false, ""error"": """ & ex.Message.Replace("""", "'") & """}"
                                response.ContentType = "application/json"
                                response.StatusCode = 500
                                EventLog.WriteEntry("PAS-Service", "UpdatePatient Fehler: " & ex.ToString(), EventLogEntryType.Error)
                            End Try

                        Case Else
                            response.StatusCode = 404
                            html = "{""error"": ""Route nicht gefunden""}"
                            response.ContentType = "application/json"
                    End Select

                Case Else
                    response.StatusCode = 405
                    html = "<html><body><h1>405 - Methode nicht erlaubt</h1></body></html>"
            End Select

            Dim buffer = Encoding.UTF8.GetBytes(html)
            response.ContentLength64 = buffer.Length
            If response.ContentType = "" Then response.ContentType = "text/html; charset=utf-8"
            response.OutputStream.Write(buffer, 0, buffer.Length)
            response.OutputStream.Close()

        Catch ex As Exception
            EventLog.WriteEntry("PAS-Service", "ProcessRequest Error: " & ex.ToString(), EventLogEntryType.Error)
        End Try
    End Sub

    ' Hilfsfunktion zum Parsen der POST-Daten
    Private Function ParseFormData(data As String) As Dictionary(Of String, String)
        Dim result As New Dictionary(Of String, String)
        If String.IsNullOrEmpty(data) Then Return result

        ' StreamReader sollte UTF-8 verwenden
        Dim pairs = data.Split("&"c)
        For Each pair In pairs
            Dim parts = pair.Split("="c)
            If parts.Length = 2 Then
                ' Beide Varianten testen:
                Dim key = HttpUtility.UrlDecode(parts(0), Encoding.UTF8)
                Dim value = HttpUtility.UrlDecode(parts(1), Encoding.UTF8)

                EventLog.WriteEntry("PAS-Service", $"Parsed: {key}={value}", EventLogEntryType.Information)
                result(key) = value
            End If
        Next
        Return result
    End Function


    Private Function GetWaitingRoomHTML() As String
        Try
            Dim html As New StringBuilder()
            html.AppendLine("<!DOCTYPE html>")
            html.AppendLine("<html><head>")
            html.AppendLine("<title>Wartezimmer-Anzeige</title>")
            html.AppendLine("<meta charset='utf-8'>")
            html.AppendLine("<meta http-equiv='refresh' content='5'>") ' Schnellere Updates
            html.AppendLine("<style>")
            html.AppendLine("body { font-family: 'Segoe UI', Arial; margin: 0; padding: 20px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); min-height: 100vh; }")
            html.AppendLine(".container { max-width: 1400px; margin: 0 auto; }")
            html.AppendLine(".header { text-align: center; color: white; margin-bottom: 30px; }")
            html.AppendLine(".header h1 { font-size: 3em; margin: 0; text-shadow: 2px 2px 4px rgba(0,0,0,0.3); }")
            html.AppendLine(".aufrufe-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(350px, 1fr)); gap: 20px; margin-bottom: 40px; }")
            html.AppendLine(".aufruf-card { background: white; padding: 25px; border-radius: 15px; box-shadow: 0 10px 30px rgba(0,0,0,0.2); position: relative; overflow: hidden; }")
            html.AppendLine(".aufruf-card.notfall { animation: notfallPulse 1.5s infinite; border: 3px solid #ff0000; }")
            html.AppendLine("@keyframes notfallPulse { 0%, 100% { background: white; } 50% { background: #ffe0e0; } }")
            html.AppendLine(".patient-nummer { font-size: 2.5em; font-weight: bold; color: #333; text-align: center; margin: 15px 0; }")
            html.AppendLine(".zimmer-info { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 15px; border-radius: 10px; text-align: center; font-size: 1.5em; }")
            html.AppendLine(".bereich-label { position: absolute; top: 10px; right: 10px; padding: 5px 15px; border-radius: 20px; font-size: 0.9em; font-weight: bold; color: white; }")
            html.AppendLine(".bereich-labor { background: #9f7aea; }")
            html.AppendLine(".bereich-ekg { background: #4299e1; }")
            html.AppendLine(".bereich-roentgen { background: #ed8936; }")
            html.AppendLine(".bereich-behandlung { background: #48bb78; }")
            html.AppendLine(".wartende-section { background: white; padding: 30px; border-radius: 15px; margin-top: 30px; }")
            html.AppendLine(".wartend-item { display: flex; justify-content: space-between; padding: 15px; margin: 10px 0; background: #f8f9fa; border-radius: 8px; border-left: 4px solid #667eea; }")
            html.AppendLine(".keine-aufrufe { text-align: center; color: white; font-size: 1.5em; padding: 50px; }")
            html.AppendLine(".clock { position: fixed; top: 20px; right: 20px; background: rgba(0,0,0,0.3); color: white; padding: 10px 20px; border-radius: 10px; font-size: 1.3em; }")
            html.AppendLine("</style>")
            html.AppendLine("</head><body>")
            html.AppendLine("<div class='clock' id='clock'></div>")
            html.AppendLine("<div class='container'>")
            html.AppendLine("<div class='header'><h1>Patientenaufruf</h1></div>")

            Using conn As New SqlConnection(ConnectionString)
                conn.Open()

                ' Alle aufgerufenen Patienten (können mehrere sein!)
                html.AppendLine("<div class='aufrufe-grid'>")

                Dim cmdAufgerufen As New SqlCommand(
                "SELECT PatNr, Bereich, Prioritaet, Aufgerufen 
                 FROM dbo.Warteschlange 
                 WHERE Status IN ('Aufgerufen', 'InBehandlung')
                 AND CAST(Ankunft AS DATE) = CAST(GETDATE() AS DATE)
                 ORDER BY Aufgerufen DESC", conn)

                Dim hatAufrufe = False
                Using reader = cmdAufgerufen.ExecuteReader()
                    While reader.Read()
                        hatAufrufe = True
                        Dim patNr = reader("PatNr").ToString()
                        Dim zimmer = reader("Bereich").ToString()
                        Dim prioritaet = Convert.ToInt32(If(IsDBNull(reader("Prioritaet")), 0, reader("Prioritaet")))
                        Dim notfallClass = If(prioritaet = 2, "notfall", "")
                        Dim bereichClass = GetBereichClass(zimmer)

                        html.AppendLine($"<div class='aufruf-card {notfallClass}'>")
                        If bereichClass <> "" Then
                            html.AppendLine($"<div class='bereich-label {bereichClass}'>{GetBereichName(zimmer)}</div>")
                        End If
                        html.AppendLine($"<div class='patient-nummer'>Patient {patNr}</div>")
                        html.AppendLine($"<div class='zimmer-info'>→ {zimmer}</div>")
                        html.AppendLine("</div>")
                    End While
                End Using

                If Not hatAufrufe Then
                    html.AppendLine("<div class='keine-aufrufe'>Keine aktuellen Aufrufe</div>")
                End If

                html.AppendLine("</div>") ' Ende aufrufe-grid

                ' Wartende Patienten in kompakter Liste
                html.AppendLine("<div class='wartende-section'>")
                html.AppendLine("<h2 style='color: #333; margin-bottom: 20px;'>Wartende Patienten</h2>")

                Dim cmdWartend As New SqlCommand(
                "SELECT PatNr, Ankunft, Prioritaet 
                 FROM dbo.Warteschlange 
                 WHERE Status = 'Wartend' 
                 AND CAST(Ankunft AS DATE) = CAST(GETDATE() AS DATE)
                 ORDER BY Prioritaet DESC, Ankunft", conn)

                Using reader = cmdWartend.ExecuteReader()
                    Dim position = 1
                    While reader.Read()
                        Dim patNr = reader("PatNr").ToString()
                        Dim wartezeit = BerechneWartezeit(CDate(reader("Ankunft")))
                        Dim prioritaet = Convert.ToInt32(If(IsDBNull(reader("Prioritaet")), 0, reader("Prioritaet")))
                        Dim prioText = If(prioritaet = 2, " 🚨", If(prioritaet = 1, " ⚠️", ""))

                        html.AppendLine("<div class='wartend-item'>")
                        html.AppendLine($"<div><strong>#{position}</strong> Patient {patNr}{prioText}</div>")
                        html.AppendLine($"<div>Wartezeit: {wartezeit}</div>")
                        html.AppendLine("</div>")
                        position += 1
                    End While
                End Using

                html.AppendLine("</div>") ' Ende wartende-section
            End Using

            html.AppendLine("</div>") ' Ende container

            ' JavaScript für Uhr
            html.AppendLine("<script>")
            html.AppendLine("function updateClock() {")
            html.AppendLine("  const now = new Date();")
            html.AppendLine("  document.getElementById('clock').textContent = ")
            html.AppendLine("    now.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit', second: '2-digit' });")
            html.AppendLine("}")
            html.AppendLine("updateClock();")
            html.AppendLine("setInterval(updateClock, 1000);")
            html.AppendLine("</script>")

            html.AppendLine("</body></html>")
            Return html.ToString()

        Catch ex As Exception
            EventLog.WriteEntry("PAS-Service", "GetWaitingRoomHTML Error: " & ex.Message)
            Return "<html><body><h1>Fehler beim Laden der Wartezimmer-Anzeige</h1></body></html>"
        End Try
    End Function



    Private Function GetPatientHTML(patNr As String) As String
        Try
            Dim patientData As Object = Nothing
            Dim position As Integer = 0

            ' Patientendaten und Position ermitteln
            Using conn As New SqlConnection(ConnectionString)
                conn.Open()

                ' Patient suchen
                Dim queryPatient = "SELECT PatNr, Name, Status, Bereich, Ankunft, Prioritaet, Bemerkung 
                               FROM dbo.Warteschlange 
                               WHERE PatNr = @patnr 
                               AND CAST(Ankunft AS DATE) = CAST(GETDATE() AS DATE)"

                Using cmd As New SqlCommand(queryPatient, conn)
                    cmd.Parameters.AddWithValue("@patnr", patNr)
                    Using reader = cmd.ExecuteReader()
                        If reader.Read() Then
                            patientData = New With {
                            .PatNr = reader("PatNr").ToString(),
                            .Name = reader("Name").ToString(),
                            .Status = reader("Status").ToString(),
                            .Bereich = reader("Bereich").ToString(),
                            .Ankunft = Convert.ToDateTime(reader("Ankunft")),
                            .Prioritaet = Convert.ToInt32(If(IsDBNull(reader("Prioritaet")), 1, reader("Prioritaet"))),
                            .Bemerkung = reader("Bemerkung").ToString()
                        }
                        End If
                    End Using
                End Using

                ' Position in Warteschlange ermitteln (nur wenn Status = Wartend)
                If patientData IsNot Nothing AndAlso patientData.Status = "Wartend" Then
                    Dim queryPosition = "SELECT COUNT(*) + 1 as Position 
                                    FROM dbo.Warteschlange 
                                    WHERE Status = 'Wartend' 
                                    AND CAST(Ankunft AS DATE) = CAST(GETDATE() AS DATE)
                                    AND (Prioritaet > @prio 
                                         OR (Prioritaet = @prio AND Ankunft < @ankunft))"

                    Using cmd As New SqlCommand(queryPosition, conn)
                        cmd.Parameters.AddWithValue("@prio", patientData.Prioritaet)
                        cmd.Parameters.AddWithValue("@ankunft", patientData.Ankunft)
                        position = Convert.ToInt32(cmd.ExecuteScalar())
                    End Using
                End If
            End Using

            ' HTML generieren
            If patientData Is Nothing Then
                Return "<html><body style='font-family: Arial; text-align: center; padding: 50px;'><h1>Patient nicht gefunden</h1><p>Bitte prüfen Sie Ihre Nummer oder wenden Sie sich an die Anmeldung.</p></body></html>"
            End If

            ' Wartezeit berechnen
            Dim ankunftZeit As DateTime = Convert.ToDateTime(patientData.Ankunft)
            Dim wartezeit = Math.Floor((DateTime.Now - ankunftZeit).TotalMinutes)
            Dim geschaetzteWartezeit = Math.Max(0, (position - 1) * 10)

            Dim html As New StringBuilder()
            html.AppendLine("<!DOCTYPE html>")
            html.AppendLine("<html>")
            html.AppendLine("<head>")
            html.AppendLine("    <meta charset='utf-8'>")
            html.AppendLine("    <meta name='viewport' content='width=device-width, initial-scale=1'>")
            html.AppendLine("    <meta http-equiv='refresh' content='10'>") ' Öfter aktualisieren
            html.AppendLine("    <title>Ihre Warteposition</title>")
            html.AppendLine("    <style>")
            html.AppendLine("        * { margin: 0; padding: 0; box-sizing: border-box; }")
            html.AppendLine("        body {")
            html.AppendLine("            font-family: 'Segoe UI', Arial, sans-serif;")
            html.AppendLine("            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);")
            html.AppendLine("            min-height: 100vh;")
            html.AppendLine("            display: flex;")
            html.AppendLine("            align-items: center;")
            html.AppendLine("            justify-content: center;")
            html.AppendLine("            padding: 20px;")
            html.AppendLine("        }")
            html.AppendLine("        .card {")
            html.AppendLine("            background: white;")
            html.AppendLine("            padding: 40px;")
            html.AppendLine("            border-radius: 20px;")
            html.AppendLine("            box-shadow: 0 20px 60px rgba(0,0,0,0.3);")
            html.AppendLine("            max-width: 500px;")
            html.AppendLine("            width: 100%;")
            html.AppendLine("        }")
            html.AppendLine("        h2 { color: #333; margin-bottom: 30px; text-align: center; }")
            html.AppendLine("        .patient-number { ")
            html.AppendLine("            background: #f0f0f0;")
            html.AppendLine("            padding: 15px;")
            html.AppendLine("            border-radius: 10px;")
            html.AppendLine("            text-align: center;")
            html.AppendLine("            font-size: 24px;")
            html.AppendLine("            font-weight: bold;")
            html.AppendLine("            color: #764ba2;")
            html.AppendLine("            margin-bottom: 30px;")
            html.AppendLine("        }")
            html.AppendLine("        .position {")
            html.AppendLine("            font-size: 120px;")
            html.AppendLine("            font-weight: bold;")
            html.AppendLine("            color: #667eea;")
            html.AppendLine("            text-align: center;")
            html.AppendLine("            line-height: 1;")
            html.AppendLine("            margin: 30px 0;")
            html.AppendLine("        }")
            html.AppendLine("        .info {")
            html.AppendLine("            text-align: center;")
            html.AppendLine("            font-size: 18px;")
            html.AppendLine("            color: #666;")
            html.AppendLine("            margin: 15px 0;")
            html.AppendLine("        }")
            html.AppendLine("        .status-wartend { background: #fff3cd; border: 2px solid #ffc107; }")
            html.AppendLine("        .status-aufgerufen {")
            html.AppendLine("            background: linear-gradient(135deg, #4CAF50, #45a049);")
            html.AppendLine("            color: white;")
            html.AppendLine("            padding: 30px;")
            html.AppendLine("            border-radius: 15px;")
            html.AppendLine("            text-align: center;")
            html.AppendLine("            animation: pulse 2s infinite;")
            html.AppendLine("        }")
            html.AppendLine("        .status-behandlung {")
            html.AppendLine("            background: #d4edda;")
            html.AppendLine("            border: 2px solid #28a745;")
            html.AppendLine("            color: #155724;")
            html.AppendLine("            padding: 20px;")
            html.AppendLine("            border-radius: 10px;")
            html.AppendLine("            text-align: center;")
            html.AppendLine("        }")
            html.AppendLine("        @keyframes pulse {")
            html.AppendLine("            0% { transform: scale(1); }")
            html.AppendLine("            50% { transform: scale(1.05); }")
            html.AppendLine("            100% { transform: scale(1); }")
            html.AppendLine("        }")
            html.AppendLine("        .priority-notfall {")
            html.AppendLine("            background: #ff4444;")
            html.AppendLine("            color: white;")
            html.AppendLine("            padding: 5px 15px;")
            html.AppendLine("            border-radius: 20px;")
            html.AppendLine("            display: inline-block;")
            html.AppendLine("            font-weight: bold;")
            html.AppendLine("            margin-bottom: 20px;")
            html.AppendLine("        }")
            html.AppendLine("    </style>")
            html.AppendLine("</head>")
            html.AppendLine("<body>")
            html.AppendLine("    <div class='card'>")
            html.AppendLine($"        <h2>Patienteninformation</h2>")
            html.AppendLine($"        <div class='patient-number'>Ihre Nummer: {patientData.PatNr}</div>")

            ' Priorität anzeigen wenn Notfall/Dringend
            If patientData.Prioritaet = 2 Then
                html.AppendLine("        <div style='text-align: center;'><span class='priority-notfall'>NOTFALL - Bevorzugte Behandlung</span></div>")
            End If

            ' Je nach Status unterschiedliche Anzeige
            Select Case patientData.Status
                Case "Wartend"
                    html.AppendLine($"        <div class='position'>{position}</div>")
                    html.AppendLine("        <div class='info'>Ihre Position in der Warteschlange</div>")
                    html.AppendLine($"        <div class='info'>⏱️ Bisherige Wartezeit: {wartezeit} Minuten</div>")
                    html.AppendLine($"        <div class='info'>📊 Geschätzte Restwartezeit: ca. {geschaetzteWartezeit} Minuten</div>")

                Case "Aufgerufen"
                    html.AppendLine("        <div class='status-aufgerufen'>")
                    html.AppendLine("            <h3 style='font-size: 36px; margin-bottom: 15px;'>Sie werden aufgerufen!</h3>")
                    html.AppendLine($"            <p style='font-size: 24px;'>Bitte begeben Sie sich zu:</p>")
                    html.AppendLine($"            <p style='font-size: 32px; font-weight: bold;'>{patientData.Bereich}</p>")
                    html.AppendLine("        </div>")

                Case "InBehandlung"
                    html.AppendLine("        <div class='status-behandlung'>")
                    html.AppendLine("            <h3>Sie werden gerade behandelt</h3>")
                    html.AppendLine($"            <p>Behandlungsort: {patientData.Bereich}</p>")
                    html.AppendLine("        </div>")

                Case "Fertig"
                    html.AppendLine("        <div class='info' style='font-size: 24px; color: #28a745;'>")
                    html.AppendLine("            ✅ Ihre Behandlung ist abgeschlossen")
                    html.AppendLine("            <p style='font-size: 16px; margin-top: 10px;'>Vielen Dank für Ihren Besuch!</p>")
                    html.AppendLine("        </div>")
            End Select

            html.AppendLine("        <div class='info' style='margin-top: 30px; font-size: 14px; color: #999;'>")
            html.AppendLine("            Diese Seite aktualisiert sich automatisch")
            html.AppendLine("        </div>")
            html.AppendLine("    </div>")
            html.AppendLine("</body>")
            html.AppendLine("</html>")

            Return html.ToString()

        Catch ex As Exception
            EventLog.WriteEntry("PAS-Service", "GetPatientHTML Error: " & ex.ToString(), EventLogEntryType.Error)
            Return "<html><body style='font-family: Arial; text-align: center; padding: 50px;'><h1>Fehler beim Laden der Daten</h1><p>Bitte versuchen Sie es später erneut.</p></body></html>"
        End Try
    End Function

    Private Function GetAnonymePatientenJSON() As String
        Try
            Using conn As New SqlConnection(ConnectionString)
                conn.Open()

                ' Nur wartende/aufgerufene Patienten
                Dim query = "SELECT PatNr, Bereich, Ankunft, Status 
                        FROM dbo.Warteschlange 
                        WHERE Status IN ('Wartend', 'Aufgerufen')
                        AND CAST(Ankunft AS DATE) = CAST(GETDATE() AS DATE)
                        ORDER BY Ankunft"

                Using cmd As New SqlCommand(query, conn)
                    Using reader = cmd.ExecuteReader()
                        Dim patienten As New List(Of AnonymerPatient)
                        Dim position = 1

                        While reader.Read()
                            patienten.Add(New AnonymerPatient With {
                            .Position = position,
                            .PatientenNummer = reader("PatNr").ToString(),
                            .Zimmer = AnonymisiereZimmer(reader("Bereich").ToString()),
                            .Wartezeit = BerechneWartezeit(CDate(reader("Ankunft")))
                        })
                            position += 1
                        End While

                        Return JsonConvert.SerializeObject(patienten)
                    End Using
                End Using
            End Using
        Catch ex As Exception
            EventLog.WriteEntry("PAS-Service", $"Fehler in GetAnonymePatientenJSON: {ex.Message}")
            Return "[]"
        End Try
    End Function



    ' Neue Hilfsfunktion für Zimmer-Anonymisierung
    Private Function AnonymisiereZimmer(zimmer As String) As String
        ' Nur Zimmernummer ohne medizinische Details
        If zimmer.Contains("-") Then
            Return zimmer.Split("-"c)(0).Trim()
        End If

        ' Spezialräume generisch machen
        Select Case zimmer.ToLower()
            Case "labor"
                Return "Zimmer 1"
            Case "ekg"
                Return "Zimmer 2"
            Case "ultraschall"
                Return "Zimmer 3"
            Case "röntgen"
                Return "Zimmer 4"
            Case "behandlung", "behandlungszimmer"
                Return "Behandlungszimmer"
            Case Else
                Return zimmer
        End Select
    End Function

    Private Function BerechneWartezeit(ankunft As DateTime) As String
        Dim minuten = CInt(Math.Floor((DateTime.Now - ankunft).TotalMinutes))
        Return $"{minuten} Min."
    End Function



    Private Function GetPatientsJSON(Optional datum As String = Nothing) As String
        Try
            Dim patients As New List(Of Object)

            Using conn As New SqlConnection(ConnectionString)
                conn.Open()

                Dim query As String

                If String.IsNullOrEmpty(datum) Then
                    ' Ohne Datum-Parameter: nur heute
                    query = "SELECT PatNr, Name, Status, Bereich, Ankunft, Wartezeit, 
                    Prioritaet, Bemerkung FROM dbo.Warteschlange 
                    WHERE CAST(Ankunft AS DATE) = CAST(GETDATE() AS DATE)
                    ORDER BY CASE WHEN Status = 'Fertig' THEN 0 ELSE 1 END,
                    Prioritaet DESC, Ankunft"
                Else

                    ' Mit Datum-Parameter: spezifisches Datum
                    query = "SELECT PatNr, Name, Status, Bereich, Ankunft, Wartezeit, 
                    Prioritaet, Bemerkung FROM dbo.Warteschlange 
                    WHERE CAST(Ankunft AS DATE) = @datum
                    ORDER BY CASE WHEN Status = 'Fertig' THEN 0 ELSE 1 END,
                    Prioritaet DESC, Ankunft"
                End If

                Using cmd As New SqlCommand(query, conn)
                    If Not String.IsNullOrEmpty(datum) Then
                        EventLog.WriteEntry("PAS-Service", $"Datum-Parameter erhalten: {datum}", EventLogEntryType.Information)

                        Dim parsedDate = DateTime.Parse(datum)
                        EventLog.WriteEntry("PAS-Service", $"Datum geparst als: {parsedDate:yyyy-MM-dd}", EventLogEntryType.Information)

                        cmd.Parameters.AddWithValue("@datum", parsedDate.Date)
                    End If

                    Using reader = cmd.ExecuteReader()
                        Dim count = 0
                        While reader.Read()
                            count += 1
                            Dim nameRaw = reader("Name").ToString()
                            EventLog.WriteEntry("PAS-Service", $"Name aus DB: {nameRaw} | Bytes: {String.Join(",", Encoding.Default.GetBytes(nameRaw))}", EventLogEntryType.Information)

                            patients.Add(New With {
                            .PatientenID = reader("PatNr").ToString(),
                            .Name = reader("Name").ToString(),
                            .Status = reader("Status").ToString(),
                            .Zimmer = reader("Bereich").ToString(),
                            .Ankunftszeit = Convert.ToDateTime(reader("Ankunft")),
                            .Wartezeit = Convert.ToInt32(If(IsDBNull(reader("Wartezeit")), 0, reader("Wartezeit"))),
                            .Prioritaet = Convert.ToInt32(If(IsDBNull(reader("Prioritaet")), 1, reader("Prioritaet"))),
                            .Bemerkung = reader("Bemerkung").ToString()
                        })
                        End While
                        EventLog.WriteEntry("PAS-Service", $"Query ergab {count} Einträge für Datum {datum}", EventLogEntryType.Information)
                    End Using
                End Using
            End Using

            Return JsonConvert.SerializeObject(patients)
        Catch ex As Exception
            EventLog.WriteEntry("PAS-Service", $"GetPatientsJSON Fehler: {ex.Message}", EventLogEntryType.Error)
            Return "[]"
        End Try
    End Function
    Protected Overrides Sub OnStop()
        Try
            _isRunning = False

            If _httpListener IsNot Nothing AndAlso _httpListener.IsListening Then
                _httpListener.Stop()
                _httpListener.Close()
            End If

            If _listenerThread IsNot Nothing AndAlso _listenerThread.IsAlive Then
                _listenerThread.Join(5000)
            End If

            EventLog.WriteEntry("PAS-Service", "Service gestoppt", EventLogEntryType.Information)
        Catch ex As Exception
            EventLog.WriteEntry("PAS-Service", "OnStop Error: " & ex.ToString(), EventLogEntryType.Error)
        End Try
    End Sub

    ' Main Methode für Debug und Service-Modus
    Public Shared Sub Main(args As String())
        If args.Length > 0 AndAlso args(0).ToLower() = "debug" Then
            ' Debug-Modus
            Console.WriteLine("PAS-Service im Debug-Modus")
            Dim service As New PAS_WebService()
            service.OnStart(Nothing)
            Console.WriteLine("Service läuft auf http://localhost:9090")
            Console.WriteLine("Drücke eine Taste zum Beenden...")
            Console.ReadKey()
            service.OnStop()
        Else
            ' Service-Modus
            Dim ServicesToRun() As ServiceBase
            ServicesToRun = New ServiceBase() {New PAS_WebService()}
            ServiceBase.Run(ServicesToRun)
        End If
    End Sub

    ' Hilfsfunktionen für Bereichs-Styling
    Private Function GetBereichClass(zimmer As String) As String
        Dim zimmerLower = zimmer.ToLower()
        If zimmerLower.Contains("labor") Then Return "bereich-labor"
        If zimmerLower.Contains("ekg") Then Return "bereich-ekg"
        If zimmerLower.Contains("röntgen") OrElse zimmerLower.Contains("roentgen") Then Return "bereich-roentgen"
        If zimmerLower.Contains("behandlung") Then Return "bereich-behandlung"
        Return ""
    End Function

    Private Function GetBereichName(zimmer As String) As String
        Dim zimmerLower = zimmer.ToLower()
        If zimmerLower.Contains("labor") Then Return "LABOR"
        If zimmerLower.Contains("ekg") Then Return "EKG"
        If zimmerLower.Contains("röntgen") OrElse zimmerLower.Contains("roentgen") Then Return "RÖNTGEN"
        If zimmerLower.Contains("behandlung") Then Return "BEHANDLUNG"
        If zimmerLower.Contains("zimmer") Then Return "ZIMMER"
        Return "AUFRUF"
    End Function

    ' Neue Klasse für anonyme Patienten
    Private Class AnonymerPatient
        Public Property Position As Integer
        Public Property PatientenNummer As String
        Public Property Zimmer As String
        Public Property Wartezeit As String
    End Class
End Class

