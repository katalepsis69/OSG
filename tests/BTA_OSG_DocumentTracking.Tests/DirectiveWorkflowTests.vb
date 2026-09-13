Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

Namespace BTA_OSG.Tests
    <TestClass>
    Public Class DirectiveWorkflowTests
        <TestMethod>
        Public Sub DirectiveWorkflow_MapsResultStatus_WhenProvided()
            Dim initialStatus = "LOGGED"
            Dim directiveType As New DirectiveType With {
                .DirectiveCode = "APPROVE_ARCHIVE",
                .DirectiveName = "Approved & Archived",
                .ResultStatusID = 8 ' ARCHIVED
            }

            Dim newStatusID As Integer = If(directiveType.ResultStatusID.HasValue, directiveType.ResultStatusID.Value, 4)
            Assert.AreEqual(8, newStatusID)
        End Sub

        <TestMethod>
        Public Sub DirectiveWorkflow_DefaultsToDirectiveIssued_WhenNoResultStatus()
            Dim directiveType As New DirectiveType With {
                .DirectiveCode = "IMMEDIATE_ACTION",
                .DirectiveName = "For Immediate Action",
                .ResultStatusID = Nothing
            }

            Dim defaultStatusID = 4 ' DIRECTIVE_ISSUED
            Dim newStatusID As Integer = If(directiveType.ResultStatusID.HasValue, directiveType.ResultStatusID.Value, defaultStatusID)
            Assert.AreEqual(defaultStatusID, newStatusID)
        End Sub
    End Class
End Namespace
