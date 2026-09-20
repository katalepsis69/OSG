Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic

Namespace BTA_OSG
    Public Class AuthenticationService
        Private ReadOnly _userRepo As UserRepository
        Private ReadOnly _auditRepo As AuditRepository
        Private ReadOnly _rfidSettings As RfidSettings

        Public Sub New(userRepo As UserRepository, auditRepo As AuditRepository, rfidSettings As RfidSettings)
            _userRepo = userRepo
            _auditRepo = auditRepo
            _rfidSettings = rfidSettings
        End Sub

        Public Function AuthenticateByCard(cardPublicId As String) As SessionContext
            If String.IsNullOrWhiteSpace(cardPublicId) Then Return Nothing
            Dim cleanCardId As String = cardPublicId.Trim().ToUpperInvariant()

            Dim user As User = _userRepo.GetByCardPublicID(cleanCardId)
            If user Is Nothing Then
                LogAudit("LOGIN_FAILURE", Nothing, Nothing, cleanCardId, False, "Card not recognized")
                Return Nothing
            End If

            If Not user.IsActive OrElse user.IsLocked Then
                LogAudit("LOGIN_FAILURE", CType(user.UserID, Integer?), user.Username, cleanCardId, False, "User account inactive or locked")
                Return Nothing
            End If

            If user.FailedTapCount >= _rfidSettings.LockoutThreshold Then
                _userRepo.LockUser(user.UserID)
                LogAudit("LOGIN_FAILURE", CType(user.UserID, Integer?), user.Username, cleanCardId, False, "Account locked out due to consecutive failed attempts")
                Return Nothing
            End If

            ' Reset failed tap counter upon successful authentication
            _userRepo.ResetFailedTaps(user.UserID)

            Dim roles As List(Of Role) = _userRepo.GetUserRoles(user.UserID)
            Dim perms As HashSet(Of String) = _userRepo.GetUserPermissions(user.UserID)

            Dim session As New SessionContext With {
                .SessionID = Guid.NewGuid(),
                .User = user,
                .Roles = roles,
                .Permissions = perms,
                .LoginAtUTC = DateTime.UtcNow,
                .LastActivityUTC = DateTime.UtcNow
            }

            LogAudit("LOGIN_SUCCESS", CType(user.UserID, Integer?), user.Username, cleanCardId, True, Nothing)
            Return session
        End Function

        Private Sub LogAudit(actionType As String, userId As Integer?, username As String, cardId As String, success As Boolean, reason As String)
            If _auditRepo Is Nothing Then Return
            Dim maskedCard As String = If(cardId.Length > 4, "****" & cardId.Substring(cardId.Length - 4), cardId)
            Dim entry As New AuditEntry With {
                .EventAtUTC = DateTime.UtcNow,
                .ActionType = actionType,
                .UserID = userId,
                .UsernameSnapshot = username,
                .CardPublicIDMasked = maskedCard,
                .Success = success,
                .FailureReason = reason,
                .MachineName = Environment.MachineName
            }
            Try
                _auditRepo.Insert(entry)
            Catch
            End Try
        End Sub
    End Class
End Namespace
