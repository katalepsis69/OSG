Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

Namespace BTA_OSG.Tests
    <TestClass>
    Public Class RbacTests
        <TestMethod>
        Public Sub SessionContext_SG_HasDirectiveIssuePermission()
            Dim user As New User With {.UserID = 1, .FullName = "Prof. Ali B. Pangalian", .IsActive = True}
            Dim roles As New List(Of Role) From {New Role With {.RoleCode = "SG", .RoleName = "Secretary-General"}}
            Dim perms As New HashSet(Of String) From {"DOCUMENT_VIEW_ALL", "DIRECTIVE_ISSUE", "AUDIT_VIEW"}
            Dim session As New SessionContext With {.User = user, .Roles = roles, .Permissions = perms}

            Assert.IsTrue(session.HasPermission("DIRECTIVE_ISSUE"))
            Assert.IsTrue(session.HasPermission("DOCUMENT_VIEW_ALL"))
            Assert.IsFalse(session.HasPermission("USER_MANAGE"))
        End Sub

        <TestMethod>
        Public Sub SessionContext_Staff_LacksDirectiveIssuePermission()
            Dim user As New User With {.UserID = 5, .FullName = "Hassim A. Ibrahim", .IsActive = True}
            Dim roles As New List(Of Role) From {New Role With {.RoleCode = "ADMIN_STAFF", .RoleName = "Administrative Staff"}}
            Dim perms As New HashSet(Of String) From {"DOCUMENT_VIEW_ASSIGNED", "DOCUMENT_CREATE", "STORAGE_UPDATE", "PDF_OPEN"}
            Dim session As New SessionContext With {.User = user, .Roles = roles, .Permissions = perms}

            Assert.IsFalse(session.HasPermission("DIRECTIVE_ISSUE"))
            Assert.IsFalse(session.HasPermission("DOCUMENT_VIEW_ALL"))
            Assert.IsTrue(session.HasPermission("DOCUMENT_CREATE"))
        End Sub

        <TestMethod>
        Public Sub EmbeddedDB_AdminStaff_WithAssignedSection_IsolatesDocuments()
            EmbeddedDB.Initialize()
            Dim cardUid = "TEST_STAFF_CARD_01"
            Dim staffName = "BTA Admin Staff Test"
            Dim deskSection = "Records Section"

            EmbeddedDB.AddUser(cardUid, staffName, "Administrative Staff", deskSection)
            Dim userRow = EmbeddedDB.AuthenticateRFID(cardUid)

            Assert.IsNotNull(userRow)
            Assert.AreEqual("Administrative Staff", userRow("Role").ToString())
            Assert.AreEqual(deskSection, userRow("Office").ToString())

            ' Check visible documents uses assigned desk section
            Dim visibleDocs = EmbeddedDB.GetVisibleDocuments(userRow)
            Assert.IsNotNull(visibleDocs)
            For Each r As System.Data.DataRowView In visibleDocs
                Dim sec = r("AssignedSection").ToString()
                Dim assigned = r("AssignedStaff").ToString()
                Assert.IsTrue(sec = deskSection OrElse assigned = staffName OrElse String.IsNullOrEmpty(assigned))
            Next
        End Sub

        <TestMethod>
        Public Sub EmbeddedDB_AdminStaff_Privileges_ArePersistedAndReadable()
            EmbeddedDB.Initialize()
            Dim cardUid = "TEST_STAFF_PERMS_01"
            Dim staffName = "Restricted Staff Test"
            Dim deskSection = "Travel Section"

            ' Register with CanRoute = False, CanMove = True, CanSoftCopy = False
            EmbeddedDB.AddUser(cardUid, staffName, "Administrative Staff", deskSection, False, True, False)
            Dim userRow = EmbeddedDB.AuthenticateRFID(cardUid)

            Assert.IsNotNull(userRow)
            Assert.IsFalse(Convert.ToBoolean(userRow("CanRoute")))
            Assert.IsTrue(Convert.ToBoolean(userRow("CanMove")))
            Assert.IsFalse(Convert.ToBoolean(userRow("CanSoftCopy")))
        End Sub

        <TestMethod>
        Public Sub EmbeddedDB_RfidAuthentication_EnforcesLockoutAfterFiveFailedAttempts()
            EmbeddedDB.Initialize()
            EmbeddedDB.ResetTerminalLockout()
            Dim fakeCard = "BAD_CARD_9999"
            ' The store starts empty, so the valid badge this test needs is one it enrols.
            Dim validCard = "TESTSG0002"
            EmbeddedDB.AddUser(validCard, "Test Lockout Officer", "Secretary-General", "Office of the Secretary-General")

            ' 4 consecutive failed taps do not lock out
            For i As Integer = 1 To 4
                Dim user = EmbeddedDB.AuthenticateRFID(fakeCard)
                Assert.IsNull(user)
                Assert.IsFalse(EmbeddedDB.IsTerminalLockedOut())
            Next

            ' 5th failed tap triggers terminal lockout
            Dim fifthTry = EmbeddedDB.AuthenticateRFID(fakeCard)
            Assert.IsNull(fifthTry)
            Assert.IsTrue(EmbeddedDB.IsTerminalLockedOut())

            ' Even a valid card is rejected while terminal is locked out
            Dim validUser = EmbeddedDB.AuthenticateRFID(validCard)
            Assert.IsNull(validUser)

            ' Reset terminal lockout
            EmbeddedDB.ResetTerminalLockout()
            Assert.IsFalse(EmbeddedDB.IsTerminalLockedOut())
            validUser = EmbeddedDB.AuthenticateRFID(validCard)
            Assert.IsNotNull(validUser)
        End Sub
    End Class
End Namespace
