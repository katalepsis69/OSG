Option Explicit On
Option Strict On

Imports System.Windows.Forms

Namespace BTA_OSG
    Public Module AppGlobalExceptionHandler
        Public Sub Setup()
            AddHandler Application.ThreadException, Sub(sender, e) HandleException(e.Exception)
            AddHandler AppDomain.CurrentDomain.UnhandledException, Sub(sender, e)
                Dim ex = TryCast(e.ExceptionObject, Exception)
                If ex IsNot Nothing Then HandleException(ex)
            End Sub
            AddHandler TaskScheduler.UnobservedTaskException, Sub(sender, e)
                ' Observed nowhere else, so log it and mark it handled: the default policy
                ' already ignores these, but a silent Task.Run bug should leave a trace.
                ' Unobserved task exceptions must never trigger modal user dialogs.
                e.SetObserved()
                If e.Exception IsNot Nothing Then
                    Dim baseEx = e.Exception.GetBaseException()
                    If baseEx IsNot Nothing Then
                        HandleException(baseEx, isUnobservedTask:=True)
                    End If
                End If
            End Sub
        End Sub

        ' A WinForms timer keeps firing while a modal error dialog pumps messages, so a
        ' handler that throws every tick would stack dialogs faster than they can be
        ' dismissed. Repeat faults within the debounce window go to the log only.
        Private _lastDialogShownUtc As DateTime = DateTime.MinValue

        Private ReadOnly LogPath As String = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BTA_OSG", "error.log")

        ' Two threads faulting at once would otherwise race File.AppendAllText and lose both
        ' stack traces to an IOException, so appends are serialized and the file is rotated
        ' before it outgrows disk sanity.
        Private ReadOnly LogLock As New Object()

        Private Function IsAbortedOrCanceledException(ex As Exception) As Boolean
            If ex Is Nothing Then Return False
            If TypeOf ex Is OperationCanceledException OrElse TypeOf ex Is Threading.ThreadAbortException Then
                Return True
            End If
            Dim baseEx = ex.GetBaseException()
            If baseEx IsNot Nothing AndAlso baseEx IsNot ex Then
                If TypeOf baseEx Is OperationCanceledException OrElse TypeOf baseEx Is Threading.ThreadAbortException Then
                    Return True
                End If
            End If
            Dim cur As Exception = ex
            While cur IsNot Nothing
                If TypeOf cur Is System.Net.Sockets.SocketException Then
                    Dim sockEx = DirectCast(cur, System.Net.Sockets.SocketException)
                    If sockEx.NativeErrorCode = 995 OrElse sockEx.SocketErrorCode = System.Net.Sockets.SocketError.OperationAborted Then
                        Return True
                    End If
                End If
                If Not String.IsNullOrEmpty(cur.Message) AndAlso
                   cur.Message.Contains("aborted because of either a thread exit or an application request", StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
                cur = cur.InnerException
            End While
            Return False
        End Function

        Private Sub HandleException(ex As Exception, Optional isUnobservedTask As Boolean = False)
            If ex Is Nothing Then Return
            If Environment.HasShutdownStarted Then Return

            ' Aborted or canceled operations (e.g. socket teardown on app exit) should not trigger popups
            If IsAbortedOrCanceledException(ex) Then
                System.Diagnostics.Trace.TraceInformation("Ignored aborted/canceled operation: {0}: {1}", ex.GetType().FullName, ex.Message)
                Return
            End If

            Try
                If AppStartup.AuditService IsNot Nothing AndAlso SessionManager.CurrentSession IsNot Nothing Then
                    Try
                        AppStartup.AuditService.LogEvent("APP_ERROR", "System", Nothing, Nothing, Nothing, ex.GetType().Name, False, ex.GetType().Name)
                    Catch
                    End Try
                End If
                System.Diagnostics.Trace.TraceError("{0}: {1}{2}{3}", ex.GetType().FullName, ex.Message, Environment.NewLine, ex.StackTrace)
                Try
                    SyncLock LogLock
                        System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(LogPath))
                        Dim logFile As New System.IO.FileInfo(LogPath)
                        ' Rotate by rename instead of delete so the tail of the previous log
                        ' stays diagnosable; the .old file is overwritten each rotation. A
                        ' locked .old file must not drop the current crash tail, so the
                        ' append falls through and the log simply exceeds the size cap.
                        If logFile.Exists AndAlso logFile.Length > 1024 * 1024 Then
                            Try
                                Dim oldPath = LogPath & ".old"
                                If System.IO.File.Exists(oldPath) Then System.IO.File.Delete(oldPath)
                                logFile.MoveTo(oldPath)
                            Catch
                            End Try
                        End If
                        System.IO.File.AppendAllText(LogPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {ex.GetType().FullName}: {ex.Message}{Environment.NewLine}{ex.StackTrace}{Environment.NewLine}{Environment.NewLine}")
                    End SyncLock
                Catch
                End Try

                ' Background task exceptions should be logged to trace and file, but not interrupt
                ' the operator with modal message boxes.
                If isUnobservedTask Then Return

                If DateTime.UtcNow.Subtract(_lastDialogShownUtc).TotalSeconds < 10 Then Return
                _lastDialogShownUtc = DateTime.UtcNow
                MessageBox.Show("An unexpected error occurred: " & ex.Message & Environment.NewLine & Environment.NewLine &
                                "Technical details were saved to " & LogPath,
                                "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Catch
                If Not isUnobservedTask Then
                    MessageBox.Show("A critical error occurred. Contact support.", "Critical Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End Try
        End Sub
    End Module
End Namespace
