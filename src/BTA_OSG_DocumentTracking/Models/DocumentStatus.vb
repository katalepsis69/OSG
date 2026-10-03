Option Explicit On
Option Strict On

Namespace BTA_OSG
    Public Class DocumentStatus
        Public Property StatusID As Integer
        Public Property StatusCode As String
        Public Property StatusName As String
        Public Property SortOrder As Integer
        Public Property IsActive As Boolean

        ' The seeded status codes an operator can route a document into (db/scripts 007/008).
        ' Single source of truth: the route dialog offers exactly this list, and the offline
        ' routing path moves CurrentStatus for exactly this set. Deliberately a constant, not
        ' a live lookup, so the connected and offline paths can never disagree.
        Public Shared ReadOnly RouteTargetStatusCodes As String() = {
            "ROUTED", "FOR_REVIEW", "FOR_REVISION", "APPROVED", "RELEASED", "FILED",
            "IN_PROGRESS", "COMPLETED", "ARCHIVED"
        }

        Public Shared Function IsValidStatus(code As String) As Boolean
            If String.IsNullOrWhiteSpace(code) Then Return False
            Dim c = code.Trim().ToUpperInvariant()
            Return c = "RECEIVED" OrElse c = "FOR_REVIEW" OrElse c = "FOR_REVISION" OrElse c = "APPROVED" OrElse c = "RELEASED" OrElse c = "FILED"
        End Function
        ' Names match tbl_DocumentStatuses.StatusName; callers keep the raw code for filters and audit.
        Public Shared Function DisplayName(statusCode As String) As String
            If String.IsNullOrWhiteSpace(statusCode) Then Return ""

            Select Case statusCode.Trim().ToUpperInvariant()
                Case "RECEIVED", "LOGGED"
                    Return "Received"
                Case "FOR_REVIEW", "PENDING_REVIEW"
                    Return "For Review"
                Case "FOR_REVISION", "REVISION_REQUESTED"
                    Return "For Revision"
                Case "APPROVED"
                    Return "Approved"
                Case "RELEASED"
                    Return "Released"
                Case "FILED", "ARCHIVED"
                    Return "Filed"
                Case "PENDING"
                    Return "Pending"
                Case "IN_TRANSIT", "IN TRANSIT"
                    Return "In Transit"
                Case Else
                    ' Legacy and imported rows can carry alias codes, so unknown text shows as stored.
                    Return statusCode.Trim()
            End Select
        End Function


    End Class
End Namespace
