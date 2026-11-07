Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.IO

Namespace BTA_OSG
    ' Offline authentication: the RFID badge lookup over the mirrored Users table, the
    ' card-credential hashing/masking/sanitizing vocabulary, and the persisted terminal
    ' lockout (five consecutive failed taps close the terminal for five minutes, and
    ' closing the app does not bypass it). Split from EmbeddedDB.vb (same partial class,
    ' zero behavior change) so the store file stays under the 1k-line review threshold.
    Partial Public Class EmbeddedDB

        Private Shared _terminalFailedTaps As Integer = 0
        Private Shared _terminalLockoutUntilUTC As DateTime = DateTime.MinValue
        Private Shared ReadOnly _lockoutFileLock As New Object()

        ' Closing and reopening the app must not be a free bypass of the five-tap rule, so
        ' the lockout expiry is persisted. LocalAppData survives a portable install being
        ' replaced during an update, which the exe folder does not.
        Friend Shared LockoutFilePath As String = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "BTA_OSG_DocumentTracking", "terminal-lockout.txt")

        Public Shared Function IsTerminalLockedOut() As Boolean
            Dim storedUntil = ReadStoredLockoutUntilUtc()
            Dim effectiveUntil = If(storedUntil > _terminalLockoutUntilUTC, storedUntil, _terminalLockoutUntilUTC)
            Return DateTime.UtcNow < effectiveUntil
        End Function

        Public Shared Sub ResetTerminalLockout()
            System.Threading.Interlocked.Exchange(_terminalFailedTaps, 0)
            _terminalLockoutUntilUTC = DateTime.MinValue
            ClearStoredLockout()
        End Sub

        Private Shared Function ReadStoredLockoutUntilUtc() As DateTime
            Try
                SyncLock _lockoutFileLock
                    If File.Exists(LockoutFilePath) Then
                        Dim raw = File.ReadAllText(LockoutFilePath).Trim()
                        Dim parsed As DateTime
                        If DateTime.TryParse(raw, Globalization.CultureInfo.InvariantCulture,
                                             Globalization.DateTimeStyles.RoundtripKind, parsed) Then
                            Return parsed.ToUniversalTime()
                        End If
                    End If
                End SyncLock
            Catch ex As Exception
                System.Diagnostics.Trace.TraceWarning("Terminal lockout state unreadable: " & ex.Message)
            End Try
            Return DateTime.MinValue
        End Function

        Private Shared Sub StoreLockoutUntilUtc(untilUtc As DateTime)
            Try
                SyncLock _lockoutFileLock
                    Directory.CreateDirectory(Path.GetDirectoryName(LockoutFilePath))
                    File.WriteAllText(LockoutFilePath, untilUtc.ToString("o", Globalization.CultureInfo.InvariantCulture))
                End SyncLock
            Catch ex As Exception
                ' An unwritable state file weakens the restart gate but must not block login.
                System.Diagnostics.Trace.TraceWarning("Terminal lockout could not be persisted: " & ex.Message)
            End Try
        End Sub

        Private Shared Sub ClearStoredLockout()
            Try
                SyncLock _lockoutFileLock
                    If File.Exists(LockoutFilePath) Then File.Delete(LockoutFilePath)
                End SyncLock
            Catch ex As Exception
                System.Diagnostics.Trace.TraceWarning("Terminal lockout file could not be cleared: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' The card credential is stored and matched only as a SHA-256 hex digest of the
        ''' uppercase ASCII card id. The Users mirror projection computes the same digest
        ''' with SQL Server's HASHBYTES('SHA2_256', CAST(... AS varchar)), so the two sides
        ''' hash the same bytes. Legacy caches that still carry the raw UID in RFID_UID
        ''' keep authenticating through the raw fallback branches.
        ''' </summary>
        Friend Shared Function Sha256Hex(value As String) As String
            Using sha = System.Security.Cryptography.SHA256.Create()
                Dim bytes = sha.ComputeHash(System.Text.Encoding.ASCII.GetBytes(If(value, "")))
                Dim sb As New System.Text.StringBuilder(bytes.Length * 2)
                For Each b In bytes
                    sb.Append(b.ToString("X2"))
                Next
                Return sb.ToString()
            End Using
        End Function

        Friend Shared Function MaskUid(uid As String) As String
            Return If(If(uid, "").Length > 4, "****" & uid.Substring(uid.Length - 4), uid)
        End Function

        ' Readers pad card IDs with trailing control characters (STX/ETX). Every entry point
        ' that accepts a typed or pasted UID funnels through this one sanitizer so what is
        ' stored always matches what the login lookup cleans.
        Public Shared Function SanitizeCardUid(value As String) As String
            Dim sb As New System.Text.StringBuilder(If(value, "").Length)
            For Each ch In If(value, "")
                If Not Char.IsControl(ch) Then sb.Append(ch)
            Next
            Return sb.ToString().Trim().ToUpperInvariant()
        End Function

        ' The offline terminal lockout reads the same configured threshold the connected
        ' path enforces at FormMain, so a tuned setting is not silently ignored offline.
        Private Shared ReadOnly Property TerminalLockoutThreshold As Integer
            Get
                Dim threshold As Integer = 5
                Try
                    threshold = AppSettings.Instance.RfidSettings.LockoutThreshold
                Catch
                End Try
                Return If(threshold > 0, threshold, 5)
            End Get
        End Property

        Public Shared Function AuthenticateRFID(uid As String) As DataRow
            EnsureInitialized()
            If IsTerminalLockedOut() Then
                Return Nothing
            End If

            If String.IsNullOrWhiteSpace(uid) Then Return Nothing
            Dim cleanUid = uid.Trim().ToUpperInvariant()
            ' The RFID listener thread and the UI can tap concurrently while the snapshot
            ' poller mutates the same tables, so the lookup runs under the store lock.
            SyncLock _syncLock
                Dim rows = DataSet.Tables("Users").Select(String.Format(
                    "(RFID_UID_HASH = '{0}' OR RFID_UID = '{1}' OR RFID_UID_RAW = '{1}') AND IsActive = True",
                    Sha256Hex(cleanUid).Replace("'", "''"), cleanUid.Replace("'", "''")))
                If rows.Length > 0 Then
                    System.Threading.Interlocked.Exchange(_terminalFailedTaps, 0)
                    ClearStoredLockout()
                    Return rows(0)
                End If
            End SyncLock
            If System.Threading.Interlocked.Increment(_terminalFailedTaps) >= TerminalLockoutThreshold Then
                _terminalLockoutUntilUTC = DateTime.UtcNow.AddMinutes(5)
                StoreLockoutUntilUtc(_terminalLockoutUntilUTC)
            End If
            Return Nothing
        End Function
    End Class
End Namespace
