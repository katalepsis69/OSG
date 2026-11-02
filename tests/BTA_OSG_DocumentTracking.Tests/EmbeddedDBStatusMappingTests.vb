Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

Namespace BTA_OSG.Tests
    ' The route dialog offers ROUTED/IN_PROGRESS/COMPLETED as routing targets; these codes
    ' must map onto the step-roadmap's virtual status space instead of regressing to
    ' RECEIVED through the fallback, which froze the pipeline at step 2.
    <TestClass>
    Public Class EmbeddedDBStatusMappingTests
        Private Shared Function MapStatus(statusCode As String) As Integer
            Dim dt As New DataTable()
            dt.Columns.Add("CurrentStatus", GetType(String))
            Dim r = dt.NewRow()
            r("CurrentStatus") = statusCode
            Return EmbeddedDB.MapRowToDocument(r).StatusID
        End Function

        <TestMethod>
        Public Sub MapRowToDocument_Completed_MapsTerminal()
            Assert.AreEqual(8, MapStatus("COMPLETED"))
        End Sub

        <TestMethod>
        Public Sub MapRowToDocument_RoutedAndInProgress_StayAtDeskStage()
            Assert.AreEqual(1, MapStatus("ROUTED"))
            Assert.AreEqual(1, MapStatus("IN_PROGRESS"))
        End Sub

        <TestMethod>
        Public Sub MapRowToDocument_RegisteredAtUTC_ComesFromRow()
            Dim dt As New DataTable()
            dt.Columns.Add("CurrentStatus", GetType(String))
            dt.Columns.Add("DateReceived", GetType(String))
            Dim r = dt.NewRow()
            r("CurrentStatus") = "RECEIVED"
            r("DateReceived") = "2026-09-15 08:30:00"
            Dim doc = EmbeddedDB.MapRowToDocument(r)
            Assert.AreEqual(New DateTime(2026, 9, 15, 8, 30, 0), doc.RegisteredAtUTC)
        End Sub
    End Class
End Namespace
