Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Drawing
Imports System.Windows.Forms

Namespace BTA_OSG
    ' Session and authentication: the RFID login dialog, mirror-based badge
    ' authentication, the inactivity timeout, and the session-row rebinding that keeps
    ' UserID pointed at the SQL identity after a sync. Split from FormMain.vb (same
    ' partial class, zero behavior change) so the shell file stays under the 1k-line
    ' review threshold.
    Partial Public Class FormMain

        ''' <summary>
        ''' The configured inactivity timeout is enforced here: the clock timer checks real
        ''' user input (UiActivityMonitor) every second and ends the session when the desk
        ''' has been idle past the limit. Until this existed the TimeoutMinutes setting had
        ''' no caller and sessions lived until manual logout.
        ''' </summary>
        Private Sub EnforceSessionTimeout()
            If CurrentUser Is Nothing OrElse Me.IsDisposed Then Return
            Dim timeoutMinutes = AppSettings.Instance.SessionSettings.TimeoutMinutes
            If timeoutMinutes <= 0 Then Return
            If UiActivityMonitor.Instance.IdleSeconds < timeoutMinutes * 60 Then Return

            Dim who As String = "Unknown Staff"
            Try
                who = CurrentUser("FullName").ToString()
            Catch
            End Try
            EmbeddedDB.LogAudit(who, String.Format("Session ended automatically after {0} minutes of inactivity.", timeoutMinutes), userId:=0, actionType:="SESSION_TIMEOUT")
            AuthenticateUser("")
            lblStatusMessage.Text = String.Format("Session timed out after {0} minutes of inactivity. Tap your RFID card to sign in again.", timeoutMinutes)
        End Sub
        ''' <summary>
        ''' The session row is captured at login, while the mirror moves Users rows from
        ''' cache-local ids onto SQL identities. Re-point it at the mirrored row for the same
        ''' badge so UserID stays the SQL identity every write path puts into a foreign key.
        ''' When the mirrored table no longer carries the badge (disabled account) the held
        ''' row is detached and every later read throws RowNotInTableException, so the
        ''' session drops to the logged-out state that every CurrentUser call site checks
        ''' for. The drop is now announced on the status banner: offline enrolments replay
        ''' to SQL before the pull, so a vanished badge means the account was disabled.
        ''' </summary>
        Private Sub RebindCurrentUser()
            If CurrentUser Is Nothing Then Return
            If Not CurrentUser.Table.Columns.Contains("RFID_UID") Then Return
            Dim uid = CurrentUser("RFID_UID").ToString()
            If String.IsNullOrWhiteSpace(uid) Then Return
            Dim match = EmbeddedDB.DataSet.Tables("Users").Select(String.Format("RFID_UID = '{0}' AND IsActive = True", uid.Replace("'", "''")))
            If match.Length > 0 Then
                CurrentUser = match(0)
            Else
                System.Diagnostics.Trace.WriteLine("Session badge " & uid & " is not in the mirrored user list; ending the session view.")
                CurrentUser = Nothing
                lblUserBadge.Text = "[ RFID Logged Out: Access Restricted ]"
                lblUserBadge.ForeColor = CivicCalmTheme.ColorDanger
                lblStatusMessage.Text = "Session ended: your badge is no longer active in the user registry."
            End If
        End Sub
        Private Sub ShowRFIDLoginDialog()
            Using dlg As New FormLogin()
                If dlg.ShowDialog(Me) = DialogResult.OK AndAlso Not String.IsNullOrEmpty(dlg.ScannedUID) Then
                    AuthenticateUser(dlg.ScannedUID, dlg.PickedOnScreen)
                End If
            End Using
        End Sub

        Public Sub AuthenticateUser(uid As String, Optional simulated As Boolean = False)
            If String.IsNullOrEmpty(uid) Then
                If CurrentUser IsNot Nothing Then
                    Dim who As String = ""
                    Try
                        who = CurrentUser("FullName").ToString()
                    Catch
                    End Try
                    EmbeddedDB.LogAudit(who, "User logged out.", actionType:="LOGOUT")
                End If
                CurrentUser = Nothing
                lblUserBadge.Text = "[ RFID Logged Out: Access Restricted ]"
                lblUserBadge.ForeColor = CivicCalmTheme.ColorDanger
                lblStatusMessage.Text = "User logged out."
                If cmbDashSection IsNot Nothing Then
                    cmbDashSection.Enabled = True
                    cmbDashSection.SelectedIndex = 0
                End If
            Else
                Dim maskedUid As String = EmbeddedDB.MaskUid(uid)
                Dim user = EmbeddedDB.AuthenticateRFID(uid)
                If user IsNot Nothing Then
                    ' The mirror carries IsLocked: a locked account must not unlock itself by
                    ' tapping, and the success path below clears the SQL-side counter, which
                    ' would reset the lock along with it. Refuse here instead.
                    If user.Table.Columns.Contains("IsLocked") AndAlso Not IsDBNull(user("IsLocked")) AndAlso Convert.ToBoolean(user("IsLocked")) Then
                        EmbeddedDB.LogAudit(user("FullName").ToString(), "RFID Badge Tap Blocked: account is locked [Card: " & maskedUid & "]", actionType:="LOGIN_LOCKED")
                        ReportRefusal("Access Denied: this account is locked after repeated failed badge reads. Ask a System Administrator to unlock it.")
                        lblStatusMessage.ForeColor = CivicCalmTheme.ColorDanger
                        Return
                    End If
                    CurrentUser = user
                    Dim role = user("Role").ToString()
                    Dim office = If(user.Table.Columns.Contains("Office") AndAlso Not IsDBNull(user("Office")), user("Office").ToString(), "")
                    Dim effectiveDesk = If(Not String.IsNullOrWhiteSpace(office), office, role)
                    Dim isGlobal As Boolean = (role = "Secretary-General" OrElse role = "System Administrator" OrElse role = "OSG Chief")
                    Dim displayRole = If(Not String.IsNullOrWhiteSpace(office) AndAlso Not office.Equals(role, StringComparison.OrdinalIgnoreCase), String.Format("{0} - {1}", role.ToUpperInvariant(), office.ToUpperInvariant()), role.ToUpperInvariant())
                    lblUserBadge.Text = String.Format("AUTHENTICATED: {0} [{1}] : {2}", user("FullName").ToString().ToUpperInvariant(), displayRole, If(isGlobal, "GLOBAL ACCESS", "SECTION DESK VIEW"))
                    lblUserBadge.ForeColor = CivicCalmTheme.ColorPrimary
                    lblStatusMessage.Text = "Authenticated: " & user("FullName").ToString()
                    lblStatusMessage.ForeColor = CivicCalmTheme.ColorInk
                    EmbeddedDB.LogAudit(user("FullName").ToString(), If(simulated,
                        "On-screen badge selection authenticated [Card: " & maskedUid & "] (no reader tap)",
                        "RFID Badge Tap Authenticated [Card: " & maskedUid & "]"), actionType:="LOGIN_SUCCESS")

                    ' Successful tap clears the SQL-side failed tap counter so the per-user
                    ' lockout (tbl_Users.FailedTapCount) tracks live state, not stale counts.
                    If AppStartup.UserRepo IsNot Nothing AndAlso Program.IsDatabaseConnected Then
                        Try
                            Dim userId = If(user.Table.Columns.Contains("UserID") AndAlso Not IsDBNull(user("UserID")), Convert.ToInt32(user("UserID")), 0)
                            If userId > 0 Then AppStartup.UserRepo.ResetFailedTaps(userId)
                        Catch
                        End Try
                    End If

                    If cmbDashSection IsNot Nothing Then
                        If isGlobal Then
                            cmbDashSection.Enabled = True
                        Else
                            If effectiveDesk.EndsWith("Section", StringComparison.OrdinalIgnoreCase) OrElse effectiveDesk.Equals("Secretariat", StringComparison.OrdinalIgnoreCase) Then
                                Dim idx = cmbDashSection.Items.IndexOf(effectiveDesk)
                                If idx >= 0 Then
                                    cmbDashSection.SelectedIndex = idx
                                End If
                            End If
                            cmbDashSection.Enabled = False
                        End If
                    End If
                Else
                    ' A cache miss with a live SQL host means one of two things: a truly
                    ' unknown card, or a badge the server knows but this seat's Users mirror
                    ' predates. GetByCardPublicID only matches active badges, so a hit here is
                    ' the stale-mirror case: the owner must not be punished for the seat's
                    ' lag. Refresh the mirror and invite a second tap instead of counting a
                    ' failed tap against the account.
                    Dim staleMirrorResolved As Boolean = False
                    If AppStartup.UserRepo IsNot Nothing AndAlso Program.IsDatabaseConnected Then
                        Try
                            Dim knownUser = AppStartup.UserRepo.GetByCardPublicID(uid.Trim().ToUpperInvariant())
                            If knownUser IsNot Nothing Then
                                If Program.Coordinator IsNot Nothing Then Program.Coordinator.RefreshFromServer("Users")
                                EmbeddedDB.LogAudit("SYSTEM", "Badge recognized on the server; this seat's staff mirror was stale and has been refreshed [Card: " & maskedUid & "]", actionType:="MIRROR_REFRESH")
                                lblStatusMessage.Text = "This badge is valid on the server; the seat's staff list was stale and has been refreshed. Tap the badge again."
                                lblStatusMessage.ForeColor = CivicCalmTheme.ColorWarning
                                staleMirrorResolved = True
                            End If
                        Catch
                        End Try
                    End If
                    If Not staleMirrorResolved Then
                        If EmbeddedDB.IsTerminalLockedOut() Then
                            lblStatusMessage.Text = "Security Lockout: Terminal locked for 5 minutes due to 5 consecutive failed card reads."
                            MessageBox.Show("Terminal has been temporarily locked out due to 5 consecutive failed RFID smart card badge reads. Please notify the System Administrator or wait 5 minutes.", "Security Lockout", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Else
                            ReportRefusal("Access Denied: Unrecognized RFID card [" & maskedUid & "]")
                            MessageBox.Show("Unrecognized RFID Smart Card Badge UID: " & maskedUid, "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    End If
                End If
            End If
            PopulateStaffDropdowns()
            RefreshActiveTabGrid()
        End Sub

        ''' <summary>
        ''' The signed-in officer's user id for attribution: the session row's UserID, or
        ''' the shipped sentinel 1 when the row predates the column (and when no session
        ''' is live, which every caller guards separately before reading other fields).
        ''' </summary>
        Friend Function CurrentStaffUserId() As Integer
            Dim userId As Integer = 1
            If CurrentUser IsNot Nothing AndAlso CurrentUser.Table.Columns.Contains("UserID") AndAlso Not IsDBNull(CurrentUser("UserID")) Then
                userId = Convert.ToInt32(CurrentUser("UserID"))
            End If
            Return userId
        End Function
    End Class
End Namespace
