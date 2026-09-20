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

        Public Function IssueCard(userId As Integer, cardPublicId As String, cardLabel As String, issuedBy As Integer) As Integer
            If String.IsNullOrWhiteSpace(cardPublicId) Then Throw New ArgumentException("CardPublicID cannot be empty.")
            Dim cleanId As String = cardPublicId.Trim().ToUpperInvariant()
            Dim newId As Integer = 0

            Using conn As SqlConnection = _connFactory.CreateConnection()
                Using cmd As New SqlCommand("INSERT INTO tbl_RfidCards (UserID, CardPublicID, CardLabel, CreatedByUserID, CreatedAtUTC, IsActive) VALUES (@u, @c, @l, @ib, SYSUTCDATETIME(), 1); SELECT SCOPE_IDENTITY();", conn)
                    cmd.Parameters.AddWithValue("@u", userId)
                    cmd.Parameters.AddWithValue("@c", cleanId)
                    cmd.Parameters.AddWithValue("@l", If(cardLabel IsNot Nothing, CType(cardLabel, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@ib", issuedBy)
                    newId = Convert.ToInt32(cmd.ExecuteScalar())
                End Using
            End Using

            If _auditService IsNot Nothing Then
                _auditService.LogEvent("RFID_ISSUED", "RfidCard", newId.ToString(), cleanId, Nothing, Nothing, True, Nothing)
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
    End Class
End Namespace
