Imports System.Data.SqlClient

Public Class PASDatabase
    Private Shared ConnectionString As String = "Server=SILINSQL\PatientenAufruf;Database=PAS_Database;User Id=sa;Password=PatientenAufruf4711;"

    Public Shared Function GetWaitingPatients() As List(Of WaitingPatient)
        Dim patients As New List(Of WaitingPatient)

        Using conn As New SqlConnection(ConnectionString)
            conn.Open()
            Using cmd As New SqlCommand("
                SELECT ID, PatNr, Name, Termin, Ankunft, Wartezeit, 
                       Status, Bereich, Behandlungsgrund, Notfall, IstPatient
                FROM Warteschlange 
                WHERE Status = 'wartend' 
                  AND CAST(ErstelltAm AS DATE) = CAST(GETDATE() AS DATE)
                ORDER BY Notfall DESC, Prioritaet DESC, Ankunft", conn)

                Using reader = cmd.ExecuteReader()
                    Dim position As Integer = 1
                    While reader.Read()
                        patients.Add(New WaitingPatient With {
                            .ID = reader.GetInt32(0),
                            .PatNr = If(reader.IsDBNull(1), "", reader.GetString(1)),
                            .Name = reader.GetString(2),
                            .Termin = If(reader.IsDBNull(3), "ohne", reader.GetString(3)),
                            .Ankunft = If(reader.IsDBNull(4), Nothing, reader.GetDateTime(4)),
                            .Wartezeit = If(reader.IsDBNull(5), 0, reader.GetInt32(5)),
                            .Status = reader.GetString(6),
                            .Bereich = reader.GetString(7),
                            .Behandlungsgrund = If(reader.IsDBNull(8), "", reader.GetString(8)),
                            .Notfall = reader.GetBoolean(9),
                            .IstPatient = reader.GetBoolean(10),
                            .Position = position
                        })
                        position += 1
                    End While
                End Using
            End Using
        End Using

        Return patients
    End Function

    Public Shared Function GetPatientInfo(patNr As String) As WaitingPatient
        Using conn As New SqlConnection(ConnectionString)
            conn.Open()
            Using cmd As New SqlCommand("
            SELECT ID, PatNr, Name, 
                   CASE WHEN Status = 'in Behandlung' THEN 0 ELSE Wartezeit END as Wartezeit,
                   Status, Bereich,
                   (SELECT COUNT(*) FROM Warteschlange w2 
                    WHERE w2.Status = 'wartend' 
                      AND w2.Ankunft < w1.Ankunft
                      AND CAST(w2.ErstelltAm AS DATE) = CAST(GETDATE() AS DATE)) + 1 as Position
            FROM Warteschlange w1
            WHERE PatNr = @patNr 
              AND CAST(ErstelltAm AS DATE) = CAST(GETDATE() AS DATE)", conn)

                cmd.Parameters.AddWithValue("@patNr", patNr)

                Using reader = cmd.ExecuteReader()
                    If reader.Read() Then
                        Return New WaitingPatient With {
                        .ID = reader.GetInt32(0),
                        .PatNr = reader.GetString(1),
                        .Name = reader.GetString(2),
                        .Wartezeit = reader.GetInt32(3),
                        .Status = reader.GetString(4),
                        .Bereich = reader.GetString(5),
                        .Position = If(reader.GetString(4) = "wartend", reader.GetInt32(6), 0)
                    }
                    End If
                End Using
            End Using
        End Using
        Return Nothing
    End Function

    ' Patient aufrufen
    Public Shared Sub CallPatient(patientID As Integer, bereich As String)
        Using conn As New SqlConnection(ConnectionString)
            conn.Open()
            Using cmd As New SqlCommand("
                UPDATE Warteschlange 
                SET Status = 'in Behandlung',
                    Behandlungsbeginn = GETDATE(),
                    Bereich = @bereich
                WHERE ID = @id", conn)

                cmd.Parameters.AddWithValue("@id", patientID)
                cmd.Parameters.AddWithValue("@bereich", bereich)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub
End Class

Public Class WaitingPatient
    Public Property ID As Integer
    Public Property PatNr As String
    Public Property Name As String
    Public Property Termin As String
    Public Property Ankunft As DateTime?
    Public Property Wartezeit As Integer
    Public Property Status As String
    Public Property Bereich As String
    Public Property Behandlungsgrund As String
    Public Property Notfall As Boolean
    Public Property IstPatient As Boolean
    Public Property Position As Integer
End Class
