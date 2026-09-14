Imports System
Imports Microsoft.Data.SqlClient

Namespace BTA_OSG
    Public Class RfidCardService
        Private ReadOnly _connFactory As Object
        Private ReadOnly _auditService As Object

        Public Sub New(connFactory As Object, auditService As Object)
            _connFactory = connFactory
            _auditService = auditService
        End Sub

        Public Function IssueCard(userId As Integer, cardPublicId As String, cardLabel As String, issuedBy As Integer) As Object
            Dim card = Nothing ' New RfidCard
            cardPublicId = cardPublicId.Trim().ToUpper()
            
            Using conn As SqlConnection = _connFactory.CreateConnection()
                Using cmd As New SqlCommand("INSERT INTO tbl_RfidCards (UserID, CardPublicID, CardLabel, CreatedByUserID, CreatedAtUTC, IsActive) VALUES (@u, @c, @l, @ib, SYSUTCDATETIME(), 1); SELECT SCOPE_IDENTITY();", conn)
                    cmd.Parameters.AddWithValue("@u", userId)
                    cmd.Parameters.AddWithValue("@c", cardPublicId)
                    cmd.Parameters.AddWithValue("@l", cardLabel)
                    cmd.Parameters.AddWithValue("@ib", issuedBy)
                    cmd.Parameters.AddWithValue("@d", DateTime.Now)
                    Dim newId = Convert.ToInt32(cmd.ExecuteScalar())
                    card = newId
                End Using
            End Using

            _auditService.LogEvent("RFID_ISSUED", "RfidCard", card, cardPublicId, Nothing, Nothing, True, Nothing)
            Return card
        End Function

        Public Sub RevokeCard(cardId As Integer, revokedBy As Integer, reason As String)
            Using conn As SqlConnection = _connFactory.CreateConnection()
                Using cmd As New SqlCommand("UPDATE tbl_RfidCards SET IsActive = 0, RevokedByUserID = @rb, RevocationReason = @rr, RevokedAtUTC = SYSUTCDATETIME() WHERE RfidCardID = @id", conn)
                    cmd.Parameters.AddWithValue("@rb", revokedBy)
                    cmd.Parameters.AddWithValue("@rr", reason)
                    cmd.Parameters.AddWithValue("@d", DateTime.Now)
                    cmd.Parameters.AddWithValue("@id", cardId)
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            _auditService.LogEvent("RFID_REVOKED", "RfidCard", cardId, Nothing, Nothing, Nothing, True, reason)
        End Sub
    End Class
End Namespace
