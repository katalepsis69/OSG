Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

Namespace BTA_OSG.Tests
    <TestClass>
    Public Class RoutingStepServiceTests
        <TestMethod>
        Public Sub GetStepsForDocument_Finance_ContainsCashierStage()
            Dim doc As New Document With {
                .DocumentID = 10,
                .DocCode = "FIN-2026-001",
                .Title = "Procurement Voucher for OSG Equipment",
                .DocumentTypeID = 3,
                .StatusID = 1,
                .RegisteredAtUTC = DateTime.UtcNow
            }

            Dim steps = RoutingStepService.GetStepsForDocument(doc)

            Assert.AreEqual(5, steps.Count)
            Assert.AreEqual("Records Section", steps(0).ResponsibleOffice)
            Assert.AreEqual("Finance Section", steps(1).ResponsibleOffice)
            Assert.AreEqual("Office of the Secretary-General", steps(2).ResponsibleOffice)
            Assert.AreEqual("Cashier / Disbursing Unit", steps(3).ResponsibleOffice)
            Assert.AreEqual("Records Section", steps(4).ResponsibleOffice)
        End Sub

        <TestMethod>
        Public Sub GetStepsForDocument_Legislative_ContainsReviewAndReferral()
            Dim doc As New Document With {
                .DocumentID = 11,
                .DocCode = "LEG-2026-001",
                .Title = "Proposed Parliamentary Bill No. 42",
                .DocumentTypeID = 2,
                .StatusID = 2,
                .RegisteredAtUTC = DateTime.UtcNow
            }

            Dim steps = RoutingStepService.GetStepsForDocument(doc)

            Assert.AreEqual(5, steps.Count)
            Assert.AreEqual("Legislative Section", steps(1).ResponsibleOffice)
            Assert.AreEqual("Office of the Secretary-General", steps(2).ResponsibleOffice)
            Assert.AreEqual("Legislative Section", steps(3).ResponsibleOffice)
        End Sub

        <TestMethod>
        Public Sub GetStepsForDocument_Travel_ContainsAdvanceAndLiquidation()
            Dim doc As New Document With {
                .DocumentID = 12,
                .DocCode = "TO-2026-001",
                .Title = "Official Travel Order to Cotabato City",
                .DocumentTypeID = 4,
                .StatusID = 1,
                .RegisteredAtUTC = DateTime.UtcNow
            }

            Dim steps = RoutingStepService.GetStepsForDocument(doc)

            Assert.AreEqual(5, steps.Count)
            Assert.AreEqual("Travel Section", steps(1).ResponsibleOffice)
            Assert.AreEqual("Cashier / Travel Desk", steps(3).ResponsibleOffice)
        End Sub

        <TestMethod>
        Public Sub RoutingStepService_Progression_Approved_CashierIsCurrent()
            Dim doc As New Document With {
                .DocumentID = 13,
                .DocCode = "FIN-2026-005",
                .Title = "Approved Financial Assistance",
                .DocumentTypeID = 3,
                .StatusID = 6, ' APPROVED
                .RegisteredAtUTC = DateTime.UtcNow
            }

            Dim steps = RoutingStepService.GetStepsForDocument(doc)
            Dim currentStation = RoutingStepService.GetCurrentStation(steps)
            Dim nextStation = RoutingStepService.GetNextStation(steps)

            Assert.IsNotNull(currentStation)
            Assert.AreEqual(4, currentStation.StepNumber)
            Assert.AreEqual("Cashier / Disbursing Unit", currentStation.ResponsibleOffice)
            Assert.AreEqual("CURRENT", currentStation.StepStatus)

            Assert.IsNotNull(nextStation)
            Assert.AreEqual(5, nextStation.StepNumber)
            Assert.AreEqual("Records Section", nextStation.ResponsibleOffice)
        End Sub

        <TestMethod>
        Public Sub RoutingStepService_Progression_Filed_100PercentComplete()
            Dim doc As New Document With {
                .DocumentID = 14,
                .DocCode = "COMM-2026-009",
                .Title = "Archived Communications Record",
                .DocumentTypeID = 1,
                .StatusID = 8, ' FILED
                .RegisteredAtUTC = DateTime.UtcNow
            }

            Dim steps = RoutingStepService.GetStepsForDocument(doc)
            Dim progress = RoutingStepService.GetProgressPercentage(steps)

            Assert.AreEqual(100, progress)
            For Each stp In steps
                Assert.AreEqual("COMPLETED", stp.StepStatus)
            Next
        End Sub

        ' Registration auto-routes a new document to its section desk with status FOR_REVIEW;
        ' the roadmap must show the DESK as the current station, not skip ahead to the SG.
        <TestMethod>
        Public Sub RoutingStepService_FreshRegisteredDocument_DeskIsCurrent()
            Dim doc As New Document With {
                .DocumentID = 15,
                .DocCode = "COMM-2026-010",
                .Title = "Freshly Registered Communication",
                .DocumentTypeID = 1,
                .StatusID = 2, ' FOR_REVIEW, as registration sets it
                .AssignedSection = "Secretariat",
                .RegisteredAtUTC = DateTime.UtcNow
            }

            Dim steps = RoutingStepService.GetStepsForDocument(doc)
            Dim currentStation = RoutingStepService.GetCurrentStation(steps)

            Assert.AreEqual(2, currentStation.StepNumber, "a freshly registered document sits at its desk, not the SG")
            Assert.AreEqual("COMPLETED", steps(0).StepStatus)
            Assert.AreEqual("CURRENT", steps(1).StepStatus)
            Assert.AreEqual("UPCOMING", steps(2).StepStatus)
        End Sub

        <TestMethod>
        Public Sub RoutingStepService_ForReviewAtSGDesk_SgStepIsCurrent()
            Dim doc As New Document With {
                .DocumentID = 16,
                .DocCode = "COMM-2026-011",
                .Title = "Resubmitted Communication at SG",
                .DocumentTypeID = 1,
                .StatusID = 2,
                .AssignedSection = "Secretary-General",
                .RegisteredAtUTC = DateTime.UtcNow
            }

            Dim currentStation = RoutingStepService.GetCurrentStation(RoutingStepService.GetStepsForDocument(doc))

            Assert.AreEqual(3, currentStation.StepNumber, "a document parked at the SG's desk shows the executive step")
        End Sub
    End Class
End Namespace
