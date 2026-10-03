Option Explicit On
Option Strict On

Namespace BTA_OSG
    Public Class DatabaseSettings
        Public Property ConnectionString As String
        Public Property UseSqlServer As Boolean = True

        ' Seconds between sync polls. The floor lives here so every reader (poll timer,
        ' stale-seat heuristic) shares one definition.
        Public Property SyncIntervalSeconds As Integer = 10

        Public Function EffectiveSyncIntervalSeconds() As Integer
            Return Math.Max(5, SyncIntervalSeconds)
        End Function
    End Class
End Namespace
