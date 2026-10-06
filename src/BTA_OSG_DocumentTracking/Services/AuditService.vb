Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic

Namespace BTA_OSG
    Public Class AuditService
        Private ReadOnly _auditRepo As AuditRepository

        Public Sub New(auditRepo As AuditRepository)
            _auditRepo = auditRepo
        End Sub

        ''' <summary>
        ''' Minimal JSON string escaping for the audit snapshot fields. The trail stores
        ''' small hand-built before/after bags, so the escaper stays smaller than a
        ''' serializer dependency.
        ''' </summary>
        Public Shared Function JsonText(value As String) As String
            If value Is Nothing Then Return ""
            Dim sb As New System.Text.StringBuilder(value.Length + 8)
            For Each ch In value
                Select Case ch
                    Case """"c : sb.Append("\""")
                    Case "\"c : sb.Append("\\")
                    Case Else
                        If ch < " "c Then
                            sb.Append(String.Format("\u{0:x4}", AscW(ch)))
                        Else
                            sb.Append(ch)
                        End If
                End Select
            Next
            Return sb.ToString()
        End Function

        Public Sub LogEvent(actionType As String, entityType As String, entityId As String, documentCode As String, oldValues As String, newValues As String, success As Boolean, failureReason As String, Optional transaction As Microsoft.Data.SqlClient.SqlTransaction = Nothing)
            Dim entry As New AuditEntry With {
                .EventAtUTC = DateTime.UtcNow,
                .ActionType = actionType,
                .EntityType = entityType,
                .EntityID = entityId,
                .DocumentCode = documentCode,
                .OldValuesJson = oldValues,
                .NewValuesJson = newValues,
                .Success = success,
                .FailureReason = failureReason,
                .MachineName = Environment.MachineName
            }

            If BTA_OSG.SessionManager.CurrentSession IsNot Nothing Then
                Dim session = BTA_OSG.SessionManager.CurrentSession
                entry.SessionID = session.SessionID
                entry.UserID = session.User.UserID
                entry.UsernameSnapshot = session.User.Username
                entry.FullNameSnapshot = session.User.FullName
                If session.Roles IsNot Nothing AndAlso session.Roles.Count > 0 Then
                    entry.RoleSnapshot = session.Roles(0).RoleName
                End If
            End If

            _auditRepo.Insert(entry, transaction)
        End Sub

        Public Sub LogLoginEvent(actionType As String, userId As Integer?, username As String, fullName As String, role As String, cardPublicIdMasked As String, success As Boolean, failureReason As String)
            Dim entry As New AuditEntry With {
                .EventAtUTC = DateTime.UtcNow,
                .ActionType = actionType,
                .UserID = userId,
                .UsernameSnapshot = username,
                .FullNameSnapshot = fullName,
                .RoleSnapshot = role,
                .CardPublicIDMasked = cardPublicIdMasked,
                .Success = success,
                .FailureReason = failureReason,
                .MachineName = Environment.MachineName
            }

            _auditRepo.Insert(entry)
        End Sub

        Public Function GetAuditLog(actionType As String, dateFrom As DateTime?, dateTo As DateTime?, pageSize As Integer, pageNumber As Integer) As List(Of AuditEntry)
            Return _auditRepo.GetByFilter(Nothing, actionType, Nothing, dateFrom, dateTo, pageSize, pageNumber)
        End Function
    End Class
End Namespace
