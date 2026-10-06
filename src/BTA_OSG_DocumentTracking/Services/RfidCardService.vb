Option Explicit On
Option Strict On

Imports System
Imports Microsoft.Data.SqlClient

Namespace BTA_OSG
    Public Class RfidCardService
        Private ReadOnly _connFactory As IDbConnectionFactory
        Private ReadOnly _auditService As AuditService

        Public Sub New(connFactory As IDbConnectionFactory, auditService As AuditService)
            _connFactory = connFactory
            _auditService = auditService
        End Sub

        Public Function IssueCard(userId As Integer, cardPublicId As String, cardLabel As String, issuedBy As Integer, Optional transaction As Microsoft.Data.SqlClient.SqlTransaction = Nothing) As Integer
            If String.IsNullOrWhiteSpace(cardPublicId) Then Throw New ArgumentException("CardPublicID cannot be empty.")
            Dim cleanId As String = EmbeddedDB.SanitizeCardUid(cardPublicId)
            Dim newId As Integer = 0

            Dim conn As SqlConnection = If(transaction IsNot Nothing, transaction.Connection, _connFactory.CreateConnection())
            Try
                Using cmd As New SqlCommand("INSERT INTO tbl_RfidCards (UserID, CardPublicID, CardLabel, CreatedByUserID, CreatedAtUTC, IsActive) VALUES (@u, @c, @l, @ib, SYSUTCDATETIME(), 1); SELECT SCOPE_IDENTITY();", conn, transaction)
                    cmd.Parameters.AddWithValue("@u", userId)
                    cmd.Parameters.AddWithValue("@c", cleanId)
                    cmd.Parameters.AddWithValue("@l", If(cardLabel IsNot Nothing, CType(cardLabel, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@ib", issuedBy)
                    newId = Convert.ToInt32(cmd.ExecuteScalar())
                End Using
            Finally
                If transaction Is Nothing Then conn.Dispose()
            End Try

            If _auditService IsNot Nothing Then
                ' Masked like every other card-credential log site (AGENTS.md non-negotiable 5):
                ' the raw card number lives in tbl_RfidCards and nowhere else.
                _auditService.LogEvent("RFID_ISSUED", "RfidCard", newId.ToString(), If(cleanId.Length > 4, "****" & Right(cleanId, 4), cleanId), Nothing, Nothing, True, Nothing, transaction)
            End If
            Return newId
        End Function

        Public Sub RevokeCard(cardId As Integer, revokedBy As Integer, reason As String)
            Using conn As SqlConnection = _connFactory.CreateConnection()
                Using cmd As New SqlCommand("UPDATE tbl_RfidCards SET IsActive = 0, RevokedByUserID = @rb, RevocationReason = @rr, RevokedAtUTC = SYSUTCDATETIME() WHERE RfidCardID = @id", conn)
                    cmd.Parameters.AddWithValue("@rb", revokedBy)
                    cmd.Parameters.AddWithValue("@rr", If(reason IsNot Nothing, CType(reason, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@id", cardId)
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            If _auditService IsNot Nothing Then
                _auditService.LogEvent("RFID_REVOKED", "RfidCard", cardId.ToString(), Nothing, Nothing, Nothing, True, reason)
            End If
        End Sub

        ''' <summary>
        ''' Lost-badge recovery: deactivates every active card of the account and returns how
        ''' many were revoked. The Admin grid carries no card id, so revocation keys on the
        ''' user; the Users mirror drops revoked cards, so the badge stops authenticating on
        ''' every desk at the next pull.
        ''' </summary>
        Public Function RevokeActiveCardsForUser(userId As Integer, revokedBy As Integer, reason As String) As Integer
            Dim revoked As Integer
            Using conn As SqlConnection = _connFactory.CreateConnection()
                Using cmd As New SqlCommand("UPDATE tbl_RfidCards SET IsActive = 0, RevokedByUserID = @rb, RevocationReason = @rr, RevokedAtUTC = SYSUTCDATETIME() WHERE UserID = @u AND IsActive = 1; SELECT @@ROWCOUNT;", conn)
                    cmd.Parameters.AddWithValue("@rb", revokedBy)
                    cmd.Parameters.AddWithValue("@rr", If(reason IsNot Nothing, CType(reason, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@u", userId)
                    revoked = Convert.ToInt32(cmd.ExecuteScalar())
                End Using
            End Using

            If _auditService IsNot Nothing Then
                _auditService.LogEvent("RFID_REVOKED", "RfidCard", userId.ToString(), Nothing, Nothing,
                                       "{""RevokedCards"":" & revoked.ToString() & "}", True, reason)
            End If
            Return revoked
        End Function
    End Class
End Namespace
