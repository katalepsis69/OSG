Imports System.Collections.Generic
Imports System.Data
Imports Microsoft.Data.SqlClient

Namespace BTA_OSG
    Public Class AuditRepository
        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        Public Function Insert(entry As AuditEntry) As Long
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "INSERT INTO tbl_AuditTrail (EventAtUTC, UserID, UsernameSnapshot, FullNameSnapshot, RoleSnapshot, ActionType, EntityType, EntityID, DocumentCode, MachineName, Success, FailureReason) " &
                           "OUTPUT INSERTED.AuditID " &
                           "VALUES (@EventAtUTC, @UserID, @UsernameSnapshot, @FullNameSnapshot, @RoleSnapshot, @ActionType, @EntityType, @EntityID, @DocumentCode, @MachineName, @Success, @FailureReason)"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@EventAtUTC", entry.EventAtUTC)
                    cmd.Parameters.AddWithValue("@UserID", If(CType(entry.UserID, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@UsernameSnapshot", If(entry.UsernameSnapshot IsNot Nothing, CType(entry.UsernameSnapshot, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@FullNameSnapshot", If(entry.FullNameSnapshot IsNot Nothing, CType(entry.FullNameSnapshot, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@RoleSnapshot", If(entry.RoleSnapshot IsNot Nothing, CType(entry.RoleSnapshot, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@ActionType", entry.ActionType)
                    cmd.Parameters.AddWithValue("@EntityType", If(entry.EntityType IsNot Nothing, CType(entry.EntityType, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@EntityID", If(entry.EntityID IsNot Nothing, CType(entry.EntityID, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@DocumentCode", If(entry.DocumentCode IsNot Nothing, CType(entry.DocumentCode, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@MachineName", If(entry.MachineName IsNot Nothing, CType(entry.MachineName, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Success", entry.Success)
                    cmd.Parameters.AddWithValue("@FailureReason", If(entry.FailureReason IsNot Nothing, CType(entry.FailureReason, Object), DBNull.Value))
                    Return Convert.ToInt64(cmd.ExecuteScalar())
                End Using
            End Using
        End Function

        Public Function GetByFilter(userIdFilter As Integer?, actionType As String, entityType As String, dateFrom As DateTime?, dateTo As DateTime?, pageSize As Integer, pageNumber As Integer) As List(Of AuditEntry)
            Dim list As New List(Of AuditEntry)()
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT AuditID, EventAtUTC, UserID, UsernameSnapshot, FullNameSnapshot, RoleSnapshot, ActionType, EntityType, EntityID, DocumentCode, MachineName, Success, FailureReason FROM tbl_AuditTrail WHERE 1=1 "
                pageSize = Math.Max(1, Math.Min(100, pageSize))
                pageNumber = Math.Max(1, pageNumber)
                If userIdFilter.HasValue Then sql &= " AND UserID = @userId "
                If Not String.IsNullOrEmpty(actionType) Then sql &= " AND ActionType = @actionType "
                If Not String.IsNullOrEmpty(entityType) Then sql &= " AND EntityType = @entityType "
                If dateFrom.HasValue Then sql &= " AND EventAtUTC >= @dateFrom "
                If dateTo.HasValue Then sql &= " AND EventAtUTC <= @dateTo "
                
                sql &= " ORDER BY AuditID DESC OFFSET @offset ROWS FETCH NEXT @pageSize ROWS ONLY"

                Using cmd = New SqlCommand(sql, conn)
                    If userIdFilter.HasValue Then cmd.Parameters.AddWithValue("@userId", userIdFilter.Value)
                    If Not String.IsNullOrEmpty(actionType) Then cmd.Parameters.AddWithValue("@actionType", actionType)
                    If Not String.IsNullOrEmpty(entityType) Then cmd.Parameters.AddWithValue("@entityType", entityType)
                    If dateFrom.HasValue Then cmd.Parameters.AddWithValue("@dateFrom", dateFrom.Value)
                    If dateTo.HasValue Then cmd.Parameters.AddWithValue("@dateTo", dateTo.Value)
                    cmd.Parameters.AddWithValue("@offset", (pageNumber - 1) * pageSize)
                    cmd.Parameters.AddWithValue("@pageSize", pageSize)

                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            list.Add(New AuditEntry With {
                                .AuditID = Convert.ToInt64(reader("AuditID")),
                                .EventAtUTC = Convert.ToDateTime(reader("EventAtUTC")),
                                .UserID = If(IsDBNull(reader("UserID")), CType(Nothing, Integer?), Convert.ToInt32(reader("UserID"))),
                                .UsernameSnapshot = If(IsDBNull(reader("UsernameSnapshot")), Nothing, Convert.ToString(reader("UsernameSnapshot"))),
                                .FullNameSnapshot = If(IsDBNull(reader("FullNameSnapshot")), Nothing, Convert.ToString(reader("FullNameSnapshot"))),
                                .RoleSnapshot = If(IsDBNull(reader("RoleSnapshot")), Nothing, Convert.ToString(reader("RoleSnapshot"))),
                                .ActionType = Convert.ToString(reader("ActionType")),
                                .EntityType = If(IsDBNull(reader("EntityType")), Nothing, Convert.ToString(reader("EntityType"))),
                                .EntityID = If(IsDBNull(reader("EntityID")), Nothing, Convert.ToString(reader("EntityID"))),
                                .DocumentCode = If(IsDBNull(reader("DocumentCode")), Nothing, Convert.ToString(reader("DocumentCode"))),
                                .MachineName = If(IsDBNull(reader("MachineName")), Nothing, Convert.ToString(reader("MachineName"))),
                                .Success = Convert.ToBoolean(reader("Success")),
                                .FailureReason = If(IsDBNull(reader("FailureReason")), Nothing, Convert.ToString(reader("FailureReason")))
                            })
                        End While
                    End Using
                End Using
            End Using
            Return list
        End Function

        Public Function InsertFailedAttempt(attempt As RfidFailedAttempt) As Long
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "INSERT INTO tbl_RfidFailedAttempts (CardPublicID, UserID, AttemptAtUTC, MachineName, Reason, Success) " &
                          "OUTPUT INSERTED.FailedAttemptID " &
                          "VALUES (@CardPublicID, @UserID, @AttemptAtUTC, @MachineName, @Reason, @Success)"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@CardPublicID", attempt.CardPublicID)
                    cmd.Parameters.AddWithValue("@UserID", If(CType(attempt.UserID, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@AttemptAtUTC", attempt.AttemptAtUTC)
                    cmd.Parameters.AddWithValue("@MachineName", If(attempt.MachineName IsNot Nothing, CType(attempt.MachineName, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Reason", attempt.Reason)
                    cmd.Parameters.AddWithValue("@Success", attempt.Success)
                    Return Convert.ToInt64(cmd.ExecuteScalar())
                End Using
            End Using
        End Function
    End Class
End Namespace
