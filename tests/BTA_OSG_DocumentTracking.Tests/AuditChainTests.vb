Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

Namespace BTA_OSG.Tests
    <TestClass>
    Public Class AuditChainTests
        Private Shared Function Sealed(count As Integer) As List(Of AuditEntry)
            Dim rows As New List(Of AuditEntry)()
            Dim prev As Byte() = Nothing
            For i = 1 To count
                Dim row As New AuditEntry With {
                    .AuditID = i,
                    .EventAtUTC = New DateTime(2026, 10, 2, 8, 0, 0, DateTimeKind.Utc).AddMinutes(i),
                    .ActionType = "ACTION_" & i.ToString(),
                    .DocumentCode = "RES-2026-00" & i.ToString(),
                    .NewValuesJson = "{""v"":1}",
                    .Success = True,
                    .MachineName = "OSG-SEAT-01"
                }
                row.PrevHash = prev
                row.RowHash = AuditChain.ComputeRowHash(row, prev)
                prev = row.RowHash
                rows.Add(row)
            Next
            Return rows
        End Function

        <TestMethod>
        Public Sub AuditChain_CleanChainVerifies()
            Dim rows = Sealed(4)
            Dim result = AuditChain.Verify(rows)
            Assert.IsTrue(result.IsValid, result.Reason)
            Assert.AreEqual(4, result.RowsChecked)
            Assert.AreEqual(0, result.RowsUnsealed)
        End Sub

        <TestMethod>
        Public Sub AuditChain_AlteredEntry_FailsAtThatEntry()
            Dim rows = Sealed(5)
            rows(2).NewValuesJson = "{""v"":99}"
            Dim result = AuditChain.Verify(rows)
            Assert.IsFalse(result.IsValid, "an altered entry verified as clean")
            Assert.AreEqual(3, result.FirstBrokenAuditID)
            Assert.IsTrue(result.Reason.Contains("content altered"), result.Reason)
        End Sub

        <TestMethod>
        Public Sub AuditChain_DeletedMiddleEntry_BreaksTheLink()
            ' The successor keeps pointing at the removed entry's hash, and no surviving entry
            ' produces it, so the deletion surfaces without anyone having rewritten anything.
            Dim rows = Sealed(5)
            rows.RemoveAt(2)
            Dim result = AuditChain.Verify(rows)
            Assert.IsFalse(result.IsValid, "a deleted entry verified as clean")
            Assert.AreEqual(4, result.FirstBrokenAuditID)
            Assert.IsTrue(result.Reason.Contains("link broken"), result.Reason)
        End Sub

        <TestMethod>
        Public Sub AuditChain_StrippedHashOnMiddleEntry_BreaksTheLink()
            ' Dropping a hash to hide an edit leaves the next entry sealed against a value that
            ' is no longer recomputable.
            Dim rows = Sealed(5)
            rows(2).RowHash = Nothing
            Dim result = AuditChain.Verify(rows)
            Assert.IsFalse(result.IsValid, "a stripped hash verified as clean")
            Assert.AreEqual(4, result.FirstBrokenAuditID)
        End Sub

        <TestMethod>
        Public Sub AuditChain_FieldBoundary_ShiftedCharacters_ChangeTheHash()
            ' Without the length prefix these two rows would hash alike, which is how a value
            ' could be moved between columns without leaving a trace.
            Dim one As New AuditEntry With {.AuditID = 1, .ActionType = "X", .UsernameSnapshot = "AB", .FullNameSnapshot = "C", .Success = True}
            Dim other As New AuditEntry With {.AuditID = 1, .ActionType = "X", .UsernameSnapshot = "A", .FullNameSnapshot = "BC", .Success = True}
            Assert.IsFalse(AuditChain.BytesEqual(AuditChain.ComputeRowHash(one, Nothing), AuditChain.ComputeRowHash(other, Nothing)),
                           "two different field splits produced the same hash")
        End Sub

        <TestMethod>
        Public Sub AuditChain_NullAndEmptyString_AreDistinct()
            Dim filled As New AuditEntry With {.AuditID = 1, .ActionType = "X", .EntityType = "Doc", .Success = True}
            Dim empty As New AuditEntry With {.AuditID = 1, .ActionType = "X", .EntityType = "", .Success = True}
            Dim missing As New AuditEntry With {.AuditID = 1, .ActionType = "X", .EntityType = Nothing, .Success = True}
            Assert.IsFalse(AuditChain.BytesEqual(AuditChain.ComputeRowHash(empty, Nothing), AuditChain.ComputeRowHash(missing, Nothing)),
                           "an empty value hashed like a NULL")
            Assert.IsFalse(AuditChain.BytesEqual(AuditChain.ComputeRowHash(filled, Nothing), AuditChain.ComputeRowHash(empty, Nothing)),
                           "a value hashed like its absence")
        End Sub

        <TestMethod>
        Public Sub AuditChain_TimestampIsCoveredByTheHash()
            ' The write path hashes what the server stored, so a rounding difference at the
            ' seventh decimal has to be a different hash or the seal could not be reproduced.
            Dim first As New AuditEntry With {.AuditID = 1, .ActionType = "X", .Success = True, .EventAtUTC = New DateTime(2026, 10, 2, 8, 0, 0, DateTimeKind.Utc)}
            Dim second As New AuditEntry With {.AuditID = 1, .ActionType = "X", .Success = True, .EventAtUTC = New DateTime(2026, 10, 2, 8, 0, 0, 1, DateTimeKind.Utc)}
            Assert.IsFalse(AuditChain.BytesEqual(AuditChain.ComputeRowHash(first, Nothing), AuditChain.ComputeRowHash(second, Nothing)),
                           "a changed timestamp kept the same hash")
        End Sub

        <TestMethod>
        Public Sub AuditChain_UnsealedPrefixIsReportedNotRejected()
            ' Entries written before migration 014 carry no hash, and an append that lands while
            ' its predecessor is still unsealed stays unsealed too. The first entry the seal pass
            ' leaves sealed therefore opens the chain with no predecessor hash, which the
            ' verifier must report as unfinished rather than as tampering.
            Dim rows = Sealed(3)
            rows(0).RowHash = Nothing
            rows(0).PrevHash = Nothing
            rows(1).RowHash = Nothing
            rows(1).PrevHash = Nothing
            rows(2).PrevHash = Nothing
            rows(2).RowHash = AuditChain.ComputeRowHash(rows(2), Nothing)

            Dim result = AuditChain.Verify(rows)
            Assert.IsTrue(result.IsValid, result.Reason)
            Assert.AreEqual(1, result.RowsChecked)
            Assert.AreEqual(2, result.RowsUnsealed)
        End Sub

        <TestMethod>
        Public Sub AuditChain_EachEntryIsSealedAgainstItsOwnPredecessor()
            Dim rows = Sealed(3)
            Assert.IsNull(rows(0).PrevHash)
            Assert.IsTrue(AuditChain.BytesEqual(rows(1).PrevHash, rows(0).RowHash))
            Assert.IsTrue(AuditChain.BytesEqual(rows(2).PrevHash, rows(1).RowHash))
            Assert.IsFalse(AuditChain.BytesEqual(rows(0).RowHash, rows(1).RowHash))
        End Sub
    End Class
End Namespace
