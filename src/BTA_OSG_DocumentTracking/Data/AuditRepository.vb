Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports System.Data
Imports Microsoft.Data.SqlClient

Namespace BTA_OSG
    Public Class AuditRepository
        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        ' The columns the chain hashes, in the order AuditChain.ComputeRowHash expects.
        Private Const ChainColumns As String =
            "AuditID, EventAtUTC, SessionID, UserID, UsernameSnapshot, FullNameSnapshot, RoleSnapshot, " &
            "ActionType, EntityType, EntityID, DocumentCode, OldValuesJson, NewValuesJson, MachineName, " &
            "ClientInfo, Success, FailureReason, CorrelationId, CardPublicIDMasked, ApplicationVersion, " &
            "PrevHash, RowHash"

        ' With a transaction supplied the audit row joins it, so an audit entry for a
        ' rolled-back write never lands; otherwise the method owns its connection.
        Public Function Insert(entry As AuditEntry, Optional transaction As SqlTransaction = Nothing) As Long
            If transaction IsNot Nothing Then
                Return AppendChained(transaction, entry)
            End If

            Using conn = _connectionFactory.CreateConnection()
                Using owned = conn.BeginTransaction()
                    Dim auditId = AppendChained(owned, entry)
                    owned.Commit()
                    Return auditId
                End Using
            End Using
        End Function

        ' Seals after the insert because the identity is hashed content, and inside the caller's
        ' transaction so a rolled-back write leaves no sealed row.
        Private Function AppendChained(tran As SqlTransaction, entry As AuditEntry) As Long
            ' The predecessor is locked before the insert: insert first, read the tail after, and
            ' two seats can each seal against the same entry and fork the chain.
            Dim pred = ReadChainPredecessor(tran, Long.MaxValue)

            ' One column list feeds the OUTPUT list and the SELECT paths, so the hashed fields
            ' cannot drift apart from the fields the verifier reads back.
            Dim sql = "INSERT INTO tbl_AuditTrail (" &
                      "EventAtUTC, SessionID, UserID, UsernameSnapshot, FullNameSnapshot, RoleSnapshot, " &
                      "ActionType, EntityType, EntityID, DocumentCode, OldValuesJson, NewValuesJson, " &
                      "MachineName, ClientInfo, Success, FailureReason, CorrelationId, CardPublicIDMasked, ApplicationVersion" &
                      ") OUTPUT INSERTED." & ChainColumns.Replace(", ", ", INSERTED.") & " VALUES (" &
                      "@EventAtUTC, @SessionID, @UserID, @UsernameSnapshot, @FullNameSnapshot, @RoleSnapshot, " &
                      "@ActionType, @EntityType, @EntityID, @DocumentCode, @OldValuesJson, @NewValuesJson, " &
                      "@MachineName, @ClientInfo, @Success, @FailureReason, @CorrelationId, @CardPublicIDMasked, @ApplicationVersion)"

            Using cmd = New SqlCommand(sql, tran.Connection, tran)
                cmd.Parameters.AddWithValue("@EventAtUTC", entry.EventAtUTC)
                cmd.Parameters.AddWithValue("@SessionID", If(entry.SessionID.HasValue, CType(entry.SessionID.Value, Object), DBNull.Value))
                cmd.Parameters.AddWithValue("@UserID", If(entry.UserID.HasValue, CType(entry.UserID.Value, Object), DBNull.Value))
                cmd.Parameters.AddWithValue("@UsernameSnapshot", If(entry.UsernameSnapshot IsNot Nothing, CType(entry.UsernameSnapshot, Object), DBNull.Value))
                cmd.Parameters.AddWithValue("@FullNameSnapshot", If(entry.FullNameSnapshot IsNot Nothing, CType(entry.FullNameSnapshot, Object), DBNull.Value))
                cmd.Parameters.AddWithValue("@RoleSnapshot", If(entry.RoleSnapshot IsNot Nothing, CType(entry.RoleSnapshot, Object), DBNull.Value))
                cmd.Parameters.AddWithValue("@ActionType", entry.ActionType)
                cmd.Parameters.AddWithValue("@EntityType", If(entry.EntityType IsNot Nothing, CType(entry.EntityType, Object), DBNull.Value))
                cmd.Parameters.AddWithValue("@EntityID", If(entry.EntityID IsNot Nothing, CType(entry.EntityID, Object), DBNull.Value))
                cmd.Parameters.AddWithValue("@DocumentCode", If(entry.DocumentCode IsNot Nothing, CType(entry.DocumentCode, Object), DBNull.Value))
                cmd.Parameters.AddWithValue("@OldValuesJson", If(entry.OldValuesJson IsNot Nothing, CType(entry.OldValuesJson, Object), DBNull.Value))
                cmd.Parameters.AddWithValue("@NewValuesJson", If(entry.NewValuesJson IsNot Nothing, CType(entry.NewValuesJson, Object), DBNull.Value))
                cmd.Parameters.AddWithValue("@MachineName", If(entry.MachineName IsNot Nothing, CType(entry.MachineName, Object), DBNull.Value))
                cmd.Parameters.AddWithValue("@ClientInfo", If(entry.ClientInfo IsNot Nothing, CType(entry.ClientInfo, Object), DBNull.Value))
                cmd.Parameters.AddWithValue("@Success", entry.Success)
                cmd.Parameters.AddWithValue("@FailureReason", If(entry.FailureReason IsNot Nothing, CType(entry.FailureReason, Object), DBNull.Value))
                cmd.Parameters.AddWithValue("@CorrelationId", If(entry.CorrelationId.HasValue, CType(entry.CorrelationId.Value, Object), DBNull.Value))
                cmd.Parameters.AddWithValue("@CardPublicIDMasked", If(entry.CardPublicIDMasked IsNot Nothing, CType(entry.CardPublicIDMasked, Object), DBNull.Value))
                cmd.Parameters.AddWithValue("@ApplicationVersion", If(entry.ApplicationVersion IsNot Nothing, CType(entry.ApplicationVersion, Object), DBNull.Value))

                ' OUTPUT INSERTED is what the server stored, not what was sent: hashing the sent
                ' values would fail on every row, because DATETIME rounds the timestamp.
                Using reader = cmd.ExecuteReader()
                    If Not reader.Read() Then Throw New InvalidOperationException("Audit insert returned no row.")
                    Dim stored = MapChainRow(reader)
                    reader.Close()

                    If Not pred.HasRow OrElse pred.RowHash IsNot Nothing Then
                        WriteHashes(tran, stored.AuditID, pred.RowHash, AuditChain.ComputeRowHash(stored, pred.RowHash))
                    End If
                    Return stored.AuditID
                End Using
            End Using
        End Function

        ''' Newest entry below the ceiling, under an update plus range lock: concurrent seals block
        ''' here, then seal against what the first one wrote. Older entry is always locked first,
        ''' so a wait cannot form a cycle.
        Private Shared Function ReadChainPredecessor(tran As SqlTransaction, beforeAuditId As Long) As (HasRow As Boolean, RowHash As Byte())
            Using cmd = New SqlCommand("SELECT TOP 1 RowHash FROM tbl_AuditTrail WITH (UPDLOCK, HOLDLOCK) WHERE AuditID < @before ORDER BY AuditID DESC", tran.Connection, tran)
                cmd.Parameters.AddWithValue("@before", beforeAuditId)
                Using reader = cmd.ExecuteReader()
                    If Not reader.Read() Then Return (False, Nothing)
                    Return (True, If(IsDBNull(reader("RowHash")), Nothing, CType(reader("RowHash"), Byte())))
                End Using
            End Using
        End Function

        Private Shared Sub WriteHashes(tran As SqlTransaction, auditId As Long, prevHash As Byte(), rowHash As Byte())
            Using cmd = New SqlCommand("UPDATE tbl_AuditTrail SET PrevHash = @prevHash, RowHash = @rowHash WHERE AuditID = @auditId", tran.Connection, tran)
                ' Typed, not AddWithValue: a DBNull parameter defaults to NVarChar, and SQL Server
                ' will not convert that implicitly into the VARBINARY columns the chain writes.
                cmd.Parameters.Add("@prevHash", SqlDbType.VarBinary, 32).Value =
                    If(prevHash IsNot Nothing, CType(prevHash, Object), DBNull.Value)
                cmd.Parameters.Add("@rowHash", SqlDbType.VarBinary, 32).Value = rowHash
                cmd.Parameters.Add("@auditId", SqlDbType.BigInt).Value = auditId
                cmd.ExecuteNonQuery()
            End Using
        End Sub

        Public Function VerifyChain() As AuditChainResult
            Using conn = _connectionFactory.CreateConnection()
                Return AuditChain.Verify(ReadChainRows(conn))
            End Using
        End Function

        ''' <summary>
        ''' Seals every entry migration 014 left unsealed, oldest first. Only rows with a NULL
        ''' RowHash are written: re-hashing a sealed entry would seal in an alteration.
        ''' </summary>
        Public Function SealUnchainedRows() As Integer
            ' Exit before taking the update and range locks when there is nothing to seal:
            ' the locking scan below blocks appends for its duration, and this pass runs on
            ' every connected launch forever.
            Using probeConn = _connectionFactory.CreateConnection()
                Using probe As New SqlCommand("SELECT TOP 1 1 FROM tbl_AuditTrail WHERE RowHash IS NULL", probeConn)
                    If probe.ExecuteScalar() Is Nothing Then Return 0
                End Using
            End Using

            Using conn = _connectionFactory.CreateConnection()
                Using tran = conn.BeginTransaction()
                    ' The scan carries the locks: update plus range over the unsealed rows blocks
                    ' appends until the pass commits.
                    Dim batch As New List(Of AuditEntry)()
                    Using cmd = New SqlCommand("SELECT " & ChainColumns & " FROM tbl_AuditTrail WITH (UPDLOCK, HOLDLOCK) WHERE RowHash IS NULL ORDER BY AuditID ASC", conn, tran)
                        Using reader = cmd.ExecuteReader()
                            While reader.Read()
                                batch.Add(MapChainRow(reader))
                            End While
                        End Using
                    End Using

                    Dim sealedCount As Integer = 0
                    For Each row In batch
                        Dim pred = ReadChainPredecessor(tran, row.AuditID)
                        ' An unsealed predecessor gets sealed further down this pass, so wait for
                        ' the next run instead of linking around the gap.
                        If pred.HasRow AndAlso pred.RowHash Is Nothing Then Continue For
                        WriteHashes(tran, row.AuditID, pred.RowHash, AuditChain.ComputeRowHash(row, pred.RowHash))
                        sealedCount += 1
                    Next

                    tran.Commit()
                    Return sealedCount
                End Using
            End Using
        End Function

        Private Shared Function ReadChainRows(conn As SqlConnection) As List(Of AuditEntry)
            Dim list As New List(Of AuditEntry)()
            Using cmd = New SqlCommand("SELECT " & ChainColumns & " FROM tbl_AuditTrail ORDER BY AuditID ASC", conn)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        list.Add(MapChainRow(reader))
                    End While
                End Using
            End Using
            Return list
        End Function

        Private Shared Function MapChainRow(reader As System.Data.IDataReader) As AuditEntry
            Return New AuditEntry With {
                .AuditID = Convert.ToInt64(reader("AuditID")),
                .EventAtUTC = Convert.ToDateTime(reader("EventAtUTC")),
                .SessionID = If(IsDBNull(reader("SessionID")), CType(Nothing, Guid?), CType(reader("SessionID"), Guid)),
                .UserID = If(IsDBNull(reader("UserID")), CType(Nothing, Integer?), Convert.ToInt32(reader("UserID"))),
                .UsernameSnapshot = If(IsDBNull(reader("UsernameSnapshot")), Nothing, Convert.ToString(reader("UsernameSnapshot"))),
                .FullNameSnapshot = If(IsDBNull(reader("FullNameSnapshot")), Nothing, Convert.ToString(reader("FullNameSnapshot"))),
                .RoleSnapshot = If(IsDBNull(reader("RoleSnapshot")), Nothing, Convert.ToString(reader("RoleSnapshot"))),
                .ActionType = Convert.ToString(reader("ActionType")),
                .EntityType = If(IsDBNull(reader("EntityType")), Nothing, Convert.ToString(reader("EntityType"))),
                .EntityID = If(IsDBNull(reader("EntityID")), Nothing, Convert.ToString(reader("EntityID"))),
                .DocumentCode = If(IsDBNull(reader("DocumentCode")), Nothing, Convert.ToString(reader("DocumentCode"))),
                .OldValuesJson = If(IsDBNull(reader("OldValuesJson")), Nothing, Convert.ToString(reader("OldValuesJson"))),
                .NewValuesJson = If(IsDBNull(reader("NewValuesJson")), Nothing, Convert.ToString(reader("NewValuesJson"))),
                .MachineName = If(IsDBNull(reader("MachineName")), Nothing, Convert.ToString(reader("MachineName"))),
                .ClientInfo = If(IsDBNull(reader("ClientInfo")), Nothing, Convert.ToString(reader("ClientInfo"))),
                .Success = Convert.ToBoolean(reader("Success")),
                .FailureReason = If(IsDBNull(reader("FailureReason")), Nothing, Convert.ToString(reader("FailureReason"))),
                .CorrelationId = If(IsDBNull(reader("CorrelationId")), CType(Nothing, Guid?), CType(reader("CorrelationId"), Guid)),
                .CardPublicIDMasked = If(IsDBNull(reader("CardPublicIDMasked")), Nothing, Convert.ToString(reader("CardPublicIDMasked"))),
                .ApplicationVersion = If(IsDBNull(reader("ApplicationVersion")), Nothing, Convert.ToString(reader("ApplicationVersion"))),
                .PrevHash = If(IsDBNull(reader("PrevHash")), Nothing, CType(reader("PrevHash"), Byte())),
                .RowHash = If(IsDBNull(reader("RowHash")), Nothing, CType(reader("RowHash"), Byte()))
            }
        End Function

        Public Function GetByFilter(userIdFilter As Integer?, actionType As String, entityType As String, dateFrom As DateTime?, dateTo As DateTime?, pageSize As Integer, pageNumber As Integer) As List(Of AuditEntry)
            Dim list As New List(Of AuditEntry)()
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT AuditID, EventAtUTC, SessionID, UserID, UsernameSnapshot, FullNameSnapshot, RoleSnapshot, " &
                          "ActionType, EntityType, EntityID, DocumentCode, OldValuesJson, NewValuesJson, MachineName, " &
                          "ClientInfo, Success, FailureReason, CorrelationId, CardPublicIDMasked, ApplicationVersion FROM tbl_AuditTrail WHERE 1=1 "
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
                                .SessionID = If(IsDBNull(reader("SessionID")), CType(Nothing, Guid?), CType(reader("SessionID"), Guid)),
                                .UserID = If(IsDBNull(reader("UserID")), CType(Nothing, Integer?), Convert.ToInt32(reader("UserID"))),
                                .UsernameSnapshot = If(IsDBNull(reader("UsernameSnapshot")), Nothing, Convert.ToString(reader("UsernameSnapshot"))),
                                .FullNameSnapshot = If(IsDBNull(reader("FullNameSnapshot")), Nothing, Convert.ToString(reader("FullNameSnapshot"))),
                                .RoleSnapshot = If(IsDBNull(reader("RoleSnapshot")), Nothing, Convert.ToString(reader("RoleSnapshot"))),
                                .ActionType = Convert.ToString(reader("ActionType")),
                                .EntityType = If(IsDBNull(reader("EntityType")), Nothing, Convert.ToString(reader("EntityType"))),
                                .EntityID = If(IsDBNull(reader("EntityID")), Nothing, Convert.ToString(reader("EntityID"))),
                                .DocumentCode = If(IsDBNull(reader("DocumentCode")), Nothing, Convert.ToString(reader("DocumentCode"))),
                                .OldValuesJson = If(IsDBNull(reader("OldValuesJson")), Nothing, Convert.ToString(reader("OldValuesJson"))),
                                .NewValuesJson = If(IsDBNull(reader("NewValuesJson")), Nothing, Convert.ToString(reader("NewValuesJson"))),
                                .MachineName = If(IsDBNull(reader("MachineName")), Nothing, Convert.ToString(reader("MachineName"))),
                                .ClientInfo = If(IsDBNull(reader("ClientInfo")), Nothing, Convert.ToString(reader("ClientInfo"))),
                                .Success = Convert.ToBoolean(reader("Success")),
                                .FailureReason = If(IsDBNull(reader("FailureReason")), Nothing, Convert.ToString(reader("FailureReason"))),
                                .CorrelationId = If(IsDBNull(reader("CorrelationId")), CType(Nothing, Guid?), CType(reader("CorrelationId"), Guid)),
                                .CardPublicIDMasked = If(IsDBNull(reader("CardPublicIDMasked")), Nothing, Convert.ToString(reader("CardPublicIDMasked"))),
                                .ApplicationVersion = If(IsDBNull(reader("ApplicationVersion")), Nothing, Convert.ToString(reader("ApplicationVersion")))
                            })
                        End While
                    End Using
                End Using
            End Using
            Return list
        End Function
    End Class
End Namespace
