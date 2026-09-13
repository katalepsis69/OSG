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
    End Class
End Namespace
