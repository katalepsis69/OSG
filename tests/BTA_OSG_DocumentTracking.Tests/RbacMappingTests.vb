Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

Namespace BTA_OSG.Tests
    ''' <summary>
    ''' Exercises the production permission resolution (SessionManager.HasPermission with
    ''' its alias map) rather than hand-built permission sets: a regression that breaks the
    ''' DOC_DIRECTIVE to DIRECTIVE_ISSUE alias, or grants a staff view-only set the
    ''' directive lever, must fail here.
    ''' </summary>
    <TestClass>
    Public Class RbacMappingTests

        <TestCleanup>
        Public Sub Cleanup()
            SessionManager.Logout()
        End Sub

        <TestMethod>
        Public Sub AliasMap_DocDirectiveGrantsDirectiveIssue()
            SessionManager.CurrentSession = New SessionContext With {
                .Permissions = New HashSet(Of String) From {RbacPolicy.DOC_DIRECTIVE}
            }
            Assert.IsTrue(SessionManager.HasPermission(RbacPolicy.DIRECTIVE_ISSUE),
                          "the alias map must resolve DOC_DIRECTIVE to DIRECTIVE_ISSUE; this is the fact that scopes directives to the SG role")
        End Sub

        <TestMethod>
        Public Sub AliasMap_RouteGrantResolvesRoutingCreate()
            SessionManager.CurrentSession = New SessionContext With {
                .Permissions = New HashSet(Of String) From {RbacPolicy.DOC_ROUTE}
            }
            Assert.IsTrue(SessionManager.HasPermission(RbacPolicy.ROUTING_CREATE))
            Assert.IsFalse(SessionManager.HasPermission(RbacPolicy.DIRECTIVE_ISSUE),
                           "a routing-only grant must not reach the directive lever")
        End Sub

        <TestMethod>
        Public Sub ViewOnlySet_CannotIssueDirectives()
            SessionManager.CurrentSession = New SessionContext With {
                .Permissions = New HashSet(Of String) From {RbacPolicy.DOC_VIEW}
            }
            Assert.IsFalse(SessionManager.HasPermission(RbacPolicy.DIRECTIVE_ISSUE))
            Assert.IsTrue(SessionManager.HasPermission(RbacPolicy.DOC_VIEW))
        End Sub

        <TestMethod>
        Public Sub NoSession_DeniesEverything()
            SessionManager.CurrentSession = Nothing
            Assert.IsFalse(SessionManager.HasPermission(RbacPolicy.DIRECTIVE_ISSUE))
            Assert.IsFalse(SessionManager.HasPermission(RbacPolicy.DOC_VIEW))
        End Sub
    End Class
End Namespace
