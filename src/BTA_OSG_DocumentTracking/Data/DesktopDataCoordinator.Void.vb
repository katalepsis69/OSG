Option Explicit On
Option Strict On

Imports System
Imports System.Data

Namespace BTA_OSG
    Partial Public Class DesktopDataCoordinator

        Private Class VoidTarget
            Public Property DocCode As String
            Public Property Title As String
            Public Property StatusCode As String
            Public Property IsLocalOnly As Boolean
        End Class

        ''' <summary>
        ''' Voiding removes a filed record from the active registry, so it is narrower than
        ''' management: only the two records-owner roles get it. Both the display names the
        ''' mirror carries and the seeded role codes are matched.
        ''' </summary>
        Public Shared Function IsVoidAuthorized(role As String) As Boolean
            Dim r = If(role, "").Trim()
            Return r.Equals("Secretary-General", StringComparison.OrdinalIgnoreCase) OrElse
                   r.Equals("System Administrator", StringComparison.OrdinalIgnoreCase) OrElse
                   r.Equals("SG", StringComparison.OrdinalIgnoreCase) OrElse
                   r.Equals("SYSADMIN", StringComparison.OrdinalIgnoreCase)
        End Function

        ''' <summary>
        ''' Removes a mis-registered document. A synced document is soft-deleted on the server
        ''' with its audit row in one transaction. A document that has never reached SQL is
        ''' removed from the local cache instead. A synced document cannot be voided while the
        ''' seat is offline: the next pull would mirror it straight back.
        ''' </summary>
        Public Function VoidDocument(docId As Integer, reason As String, staffName As String, staffUserId As Integer, staffRole As String) As Boolean
            If docId <= 0 OrElse String.IsNullOrWhiteSpace(reason) Then Return False
            If Not IsVoidAuthorized(staffRole) Then Return False
            Dim target = CachedVoidTarget(docId)
            If target Is Nothing Then Return False
            If DocumentStatus.IsTerminalStatus(target.StatusCode) Then Return False

            If target.IsLocalOnly Then
                Return RemoveLocalOnlyDocument(docId, target, reason, staffName, staffUserId)
            End If
            If Not IsDatabaseConnected Then Return False
            If AppStartup.DocumentRepo Is Nothing OrElse AppStartup.ConnectionFactory Is Nothing Then Return False

            Try
                Using conn = AppStartup.ConnectionFactory.CreateConnection()
                    Using tx = conn.BeginTransaction()
                        Try
                            ' The local mirror can be one poll old. Lock and re-read the server
                            ' status inside the transaction so a release made on another seat
                            ' cannot slip a terminal document through the void.
                            Dim serverStatus = ReadServerStatusCode(conn, tx, docId)
                            If serverStatus Is Nothing OrElse DocumentStatus.IsTerminalStatus(serverStatus) Then
                                tx.Rollback()
                                Return False
                            End If
                            Dim affected = AppStartup.DocumentRepo.SoftDelete(docId, staffUserId, reason, tx)
                            If affected = 0 Then
                                tx.Rollback()
                                Return False
                            End If
                            If AppStartup.AuditService IsNot Nothing Then
                                AppStartup.AuditService.LogEvent("DOCUMENT_VOIDED", "Document", docId.ToString(), target.DocCode,
                                                                 Nothing, BuildVoidSnapshot(target, reason), True, Nothing, tx)
                            End If
                            tx.Commit()
                        Catch
                            tx.Rollback()
                            Throw
                        End Try
                    End Using
                End Using
                ' The pull drops the row once SQL omits it, but an empty Documents pull and the
                ' local ID 1-4 seed guard can both leave it on screen. Remove it explicitly.
                RemoveCachedDocumentTree(docId)
                PullAfterWrite("Documents", "AuditTrail")
                Return True
            Catch ex As Exception
                _sqlAvailable = False
                Program.ReportOperatorWarning("Void failed: " & ex.Message)
                Return False
            End Try
        End Function

        Private Shared Function CachedVoidTarget(docId As Integer) As VoidTarget
            SyncLock EmbeddedDB.SyncRoot
                Dim table = EmbeddedDB.DataSet.Tables("Documents")
                If table Is Nothing Then Return Nothing
                Dim row = table.Rows.Find(docId)
                If row Is Nothing Then Return Nothing
                Dim localOnly As Boolean = False
                If table.Columns.Contains("IsNewOfflineRecord") AndAlso Not IsDBNull(row("IsNewOfflineRecord")) Then
                    localOnly = CBool(row("IsNewOfflineRecord"))
                End If
                Return New VoidTarget With {
                    .DocCode = row("DocCode").ToString(),
                    .Title = row("Title").ToString(),
                    .StatusCode = row("CurrentStatus").ToString(),
                    .IsLocalOnly = localOnly
                }
            End SyncLock
        End Function

        Private Shared Sub RemoveCachedDocumentTree(docId As Integer)
            SyncLock EmbeddedDB.SyncRoot
                For Each tableName In {"Documents", "Directives", "RoutingLogs", "Movements"}
                    Dim table = EmbeddedDB.DataSet.Tables(tableName)
                    If table Is Nothing OrElse Not table.Columns.Contains("DocumentID") Then Continue For
                    For Each row As DataRow In table.Select("DocumentID = " & docId.ToString())
                        row.Delete()
                    Next
                Next
            End SyncLock
            ' Persist the removal before the pull or audit, so a crash cannot resurrect a row
            ' the operator already voided.
            EmbeddedDB.Save()
        End Sub

        Private Shared Function ReadServerStatusCode(conn As Microsoft.Data.SqlClient.SqlConnection, tx As Microsoft.Data.SqlClient.SqlTransaction, docId As Integer) As String
            Dim sql = "SELECT st.StatusCode FROM dbo.tbl_Documents d WITH (UPDLOCK, HOLDLOCK) " &
                      "LEFT JOIN dbo.tbl_DocumentStatuses st ON st.StatusID = d.StatusID " &
                      "WHERE d.DocumentID = @id AND d.IsDeleted = 0"
            Using cmd As New Microsoft.Data.SqlClient.SqlCommand(sql, conn)
                cmd.Transaction = tx
                cmd.Parameters.AddWithValue("@id", docId)
                Dim value = cmd.ExecuteScalar()
                If value Is Nothing OrElse IsDBNull(value) Then Return Nothing
                Return value.ToString()
            End Using
        End Function

        Private Shared Function RemoveLocalOnlyDocument(docId As Integer, target As VoidTarget, reason As String, staffName As String, staffUserId As Integer) As Boolean
            RemoveCachedDocumentTree(docId)
            ' The row never reached SQL, so its void is local-only; the audit entry stays on
            ' this seat and must not replay as a document operation.
            EmbeddedDB.LogAudit(staffName, "Voided local document " & target.DocCode & ": " & reason,
                                isOffline:=False, userId:=staffUserId, actionType:="DOCUMENT_VOIDED")
            EmbeddedDB.Save()
            Return True
        End Function

        Private Shared Function BuildVoidSnapshot(target As VoidTarget, reason As String) As String
            Return "{""DocCode"":" & System.Text.Json.JsonSerializer.Serialize(If(target.DocCode, "")) &
                   ",""Title"":" & System.Text.Json.JsonSerializer.Serialize(If(target.Title, "")) &
                   ",""Status"":" & System.Text.Json.JsonSerializer.Serialize(If(target.StatusCode, "")) &
                   ",""VoidReason"":" & System.Text.Json.JsonSerializer.Serialize(If(reason, "")) & "}"
        End Function

    End Class
End Namespace
