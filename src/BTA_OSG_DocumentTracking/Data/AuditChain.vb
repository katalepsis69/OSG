Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Security.Cryptography
Imports System.Text

Namespace BTA_OSG
    ''' SHA-256 hash chain for tbl_AuditTrail: each entry seals its own stored columns plus the
    ''' hash of the entry before it, so altering, deleting, or reordering history breaks the link.
    ''' The canonical form is defined here once for the append, seal, and verify paths.
    ''' Known limit: deleting only the newest entry leaves no trace, which no in-table scheme can
    ''' close; that needs the chain head witnessed outside the database.
    Public Class AuditChain

        ''' Bumped if the canonical form changes, so old entries stay verifiable under their own version.
        Public Const ChainVersion As Integer = 1

        Private Shared ReadOnly Utf8 As New UTF8Encoding(False)

        ''' row must carry the values as the server stored them: DATETIME rounds a .NET timestamp
        ''' on write, so hashing the intended values would produce a hash nothing can reproduce.
        Public Shared Function ComputeRowHash(row As AuditEntry, prevHash As Byte()) As Byte()
            If row Is Nothing Then Throw New ArgumentNullException(NameOf(row))
            Dim bytes As New List(Of Byte)(1024)
            AppendText(bytes, row.AuditID.ToString("D", CultureInfo.InvariantCulture))
            AppendText(bytes, row.EventAtUTC.ToString("yyyy-MM-dd HH:mm:ss.fffffff", CultureInfo.InvariantCulture))
            AppendText(bytes, NullableGuidText(row.SessionID))
            AppendText(bytes, NullableIntText(row.UserID))
            AppendText(bytes, row.UsernameSnapshot)
            AppendText(bytes, row.FullNameSnapshot)
            AppendText(bytes, row.RoleSnapshot)
            AppendText(bytes, row.ActionType)
            AppendText(bytes, row.EntityType)
            AppendText(bytes, row.EntityID)
            AppendText(bytes, row.DocumentCode)
            AppendText(bytes, row.OldValuesJson)
            AppendText(bytes, row.NewValuesJson)
            AppendText(bytes, row.MachineName)
            AppendText(bytes, row.ClientInfo)
            AppendText(bytes, If(row.Success, "1", "0"))
            AppendText(bytes, row.FailureReason)
            AppendText(bytes, NullableGuidText(row.CorrelationId))
            AppendText(bytes, row.CardPublicIDMasked)
            AppendText(bytes, row.ApplicationVersion)
            AppendRaw(bytes, prevHash)
            AppendText(bytes, "v" & ChainVersion.ToString("D", CultureInfo.InvariantCulture))
            Return SHA256.HashData(bytes.ToArray())
        End Function

        ''' Unsealed entries are skipped, not rejected: they are the pre-014 prefix and any tail the
        ''' seal has not reached. RowsUnsealed tells the two apart from a stripped hash.
        Public Shared Function Verify(rowsInAuditIdOrder As IEnumerable(Of AuditEntry)) As AuditChainResult
            Dim result As New AuditChainResult With {.IsValid = True, .Reason = "Chain verified."}
            Dim prevHash As Byte() = Nothing

            For Each row In rowsInAuditIdOrder
                If row.RowHash Is Nothing Then
                    result.RowsUnsealed += 1
                    Continue For
                End If

                If Not BytesEqual(row.PrevHash, prevHash) Then
                    Dim reason = If(prevHash Is Nothing,
                                    "Chain does not start where it should: the first sealed entry carries a predecessor hash.",
                                    "Chain link broken: an entry was deleted, reordered, or stripped of its hash.")
                    Return Fail(result, row.AuditID, reason)
                End If

                Dim recomputed = ComputeRowHash(row, prevHash)
                If Not BytesEqual(recomputed, row.RowHash) Then
                    Return Fail(result, row.AuditID, "Entry content altered: the stored columns no longer match the hash sealed over them.")
                End If

                prevHash = row.RowHash
                result.RowsChecked += 1
            Next

            Return result
        End Function

        Private Shared Function Fail(result As AuditChainResult, auditId As Long, reason As String) As AuditChainResult
            result.IsValid = False
            result.FirstBrokenAuditID = auditId
            result.Reason = reason
            Return result
        End Function

        ' Length prefixes stop one value borrowing characters from its neighbour, so "AB" then "C"
        ' cannot hash like "A" then "BC". -1 marks NULL, distinct from an empty string.
        Private Shared Sub AppendText(bytes As List(Of Byte), value As String)
            If value Is Nothing Then
                AddLength(bytes, -1)
                Return
            End If
            Dim payload = Utf8.GetBytes(value)
            AddLength(bytes, payload.Length)
            bytes.AddRange(payload)
        End Sub

        Private Shared Sub AppendRaw(bytes As List(Of Byte), payload As Byte())
            If payload Is Nothing Then
                AddLength(bytes, -1)
                Return
            End If
            AddLength(bytes, payload.Length)
            bytes.AddRange(payload)
        End Sub

        Private Shared Sub AddLength(bytes As List(Of Byte), length As Integer)
            bytes.Add(CByte((length >> 24) And &HFF))
            bytes.Add(CByte((length >> 16) And &HFF))
            bytes.Add(CByte((length >> 8) And &HFF))
            bytes.Add(CByte(length And &HFF))
        End Sub

        Private Shared Function NullableGuidText(value As Guid?) As String
            Return If(value.HasValue, value.Value.ToString("N", CultureInfo.InvariantCulture).ToUpperInvariant(), Nothing)
        End Function

        Private Shared Function NullableIntText(value As Integer?) As String
            Return If(value.HasValue, value.Value.ToString("D", CultureInfo.InvariantCulture), Nothing)
        End Function

        Public Shared Function BytesEqual(a As Byte(), b As Byte()) As Boolean
            If a Is Nothing OrElse b Is Nothing Then Return a Is b
            If a.Length <> b.Length Then Return False
            For i = 0 To a.Length - 1
                If a(i) <> b(i) Then Return False
            Next
            Return True
        End Function
    End Class

    Public Class AuditChainResult
        Public Property IsValid As Boolean
        Public Property RowsChecked As Integer
        Public Property RowsUnsealed As Integer
        Public Property FirstBrokenAuditID As Long
        Public Property Reason As String
    End Class
End Namespace
