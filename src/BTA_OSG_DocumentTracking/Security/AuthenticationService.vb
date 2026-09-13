Imports System
Imports System.Collections.Generic

Namespace BTA_OSG
    Public Class AuthenticationService
        Private ReadOnly _userRepo As Object ' UserRepository
        Private ReadOnly _auditRepo As Object ' AuditRepository
        Private ReadOnly _rfidSettings As Object ' RfidSettings

        Public Sub New(userRepo As Object, auditRepo As Object, rfidSettings As Object)
            _userRepo = userRepo
            _auditRepo = auditRepo
            _rfidSettings = rfidSettings
        End Sub

        Public Function AuthenticateByCard(cardPublicId As String) As Object ' SessionContext
            If String.IsNullOrWhiteSpace(cardPublicId) Then Return Nothing
            cardPublicId = cardPublicId.Trim().ToUpper()

            Dim user As Object = _userRepo.GetUserByCard(cardPublicId)
            If user Is Nothing Then
                LogAudit("LOGIN_FAILURE", Nothing, False, "Card not found")
                Return Nothing
            End If

            If Not user.IsActive OrElse user.IsLocked Then
                LogAudit("LOGIN_FAILURE", user.UserID, False, "User inactive or locked")
                Return Nothing
            End If

            If user.FailedTaps >= _rfidSettings.LockoutThreshold Then
                LogAudit("LOGIN_FAILURE", user.UserID, False, "Account locked out")
                Return Nothing
            End If

            ' Success
            _userRepo.ResetFailedTaps(user.UserID)
            
            Dim session As Object = _userRepo.CreateSessionContext(user.UserID) ' Assume this exists
            session.LoginTime = DateTime.Now
            session.LastActivityTime = DateTime.Now

            LogAudit("LOGIN_SUCCESS", user.UserID, True, Nothing)
            Return session
        End Function

        Private Sub LogAudit(action As String, userId As Integer?, success As Boolean, reason As String)
            ' Implementation for _auditRepo.Insert()
        End Sub
    End Class
End Namespace
