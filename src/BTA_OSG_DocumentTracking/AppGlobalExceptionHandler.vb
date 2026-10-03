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
                e.SetObserved()
                If e.Exception IsNot Nothing Then HandleException(e.Exception.GetBaseException())
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

        Private Sub HandleException(ex As Exception)
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
                If DateTime.UtcNow.Subtract(_lastDialogShownUtc).TotalSeconds < 10 Then Return
                _lastDialogShownUtc = DateTime.UtcNow
                MessageBox.Show("An unexpected error occurred: " & ex.Message & Environment.NewLine & Environment.NewLine &
                                "Technical details were saved to " & LogPath,
                                "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Catch
                MessageBox.Show("A critical error occurred. Contact support.", "Critical Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
    End Module
End Namespace
