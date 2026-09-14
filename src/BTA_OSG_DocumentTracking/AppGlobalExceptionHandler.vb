Imports System.Windows.Forms

Namespace BTA_OSG
    Public Module AppGlobalExceptionHandler
        Public Sub Setup()
            AddHandler Application.ThreadException, Sub(sender, e) HandleException(e.Exception)
            AddHandler AppDomain.CurrentDomain.UnhandledException, Sub(sender, e)
                Dim ex = TryCast(e.ExceptionObject, Exception)
                If ex IsNot Nothing Then HandleException(ex)
            End Sub
        End Sub

        Private Sub HandleException(ex As Exception)
            Try
                System.Diagnostics.Trace.TraceError($"Unhandled: {ex}")
                If AppStartup.AuditService IsNot Nothing AndAlso SessionManager.CurrentSession IsNot Nothing Then
                    Try
                        AppStartup.AuditService.LogEvent("APP_ERROR", "System", Nothing, Nothing, Nothing, ex.GetType().Name, False, ex.GetType().Name)
                    Catch
                    End Try
                End If
                MessageBox.Show("An unexpected error occurred. Contact support if this persists.", "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Catch
                MessageBox.Show("A critical error occurred. Contact support.", "Critical Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
    End Module
End Namespace
