Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic

Namespace BTA_OSG
    Public Class RoutingStepService
        Public Shared Function GetStepsForDocument(doc As Document, Optional routingLogs As List(Of RoutingLog) = Nothing) As List(Of RoutingStep)
            Dim steps As List(Of RoutingStep)
            If doc Is Nothing Then Return New List(Of RoutingStep)()

            Dim docType As String = GetDocTypeIdentifier(doc)

            Select Case docType
                Case "FIN"
                    steps = CreateFinancePipeline()
                Case "TRAVEL"
                    steps = CreateTravelPipeline()
                Case "LEG"
                    steps = CreateLegislativePipeline()
                Case Else
                    steps = CreateCommunicationPipeline()
            End Select

            AssignStepStatuses(doc, steps, routingLogs)
            Return steps
        End Function

        Private Shared Function CreateFinancePipeline() As List(Of RoutingStep)
            Return New List(Of RoutingStep) From {
                New RoutingStep With {
                    .StepNumber = 1,
                    .StageCode = "INTAKE",
                    .StageName = "Intake & Document Verification",
                    .ResponsibleOffice = "Records Section",
                    .ActionRequired = "Verify digital attachments, assign control code, and log electronic intake"
                },
                New RoutingStep With {
                    .StepNumber = 2,
                    .StageCode = "OBLIGATION",
                    .StageName = "Obligation & Pre-Audit Processing",
                    .ResponsibleOffice = "Finance Section",
                    .ActionRequired = "Review budget allocation and process obligation request"
                },
                New RoutingStep With {
                    .StepNumber = 3,
                    .StageCode = "APPROVAL",
                    .StageName = "Executive Review & Sign-Off",
                    .ResponsibleOffice = "Office of the Secretary-General",
                    .ActionRequired = "Review and approve disbursement directive"
                },
                New RoutingStep With {
                    .StepNumber = 4,
                    .StageCode = "CASHIER",
                    .StageName = "Cashiering & Disbursement",
                    .ResponsibleOffice = "Cashier / Disbursing Unit",
                    .ActionRequired = "Disburse check or payment advice to recipient"
                },
                New RoutingStep With {
                    .StepNumber = 5,
                    .StageCode = "RELEASE",
                    .StageName = "Official Release & Archiving",
                    .ResponsibleOffice = "Records Section",
                    .ActionRequired = "Issue official release and file voucher in archive"
                }
            }
        End Function

        Private Shared Function CreateTravelPipeline() As List(Of RoutingStep)
            Return New List(Of RoutingStep) From {
                New RoutingStep With {
                    .StepNumber = 1,
                    .StageCode = "INTAKE",
                    .StageName = "Travel Request Intake",
                    .ResponsibleOffice = "Records Section",
                    .ActionRequired = "Receive travel request and supporting itinerary"
                },
                New RoutingStep With {
                    .StepNumber = 2,
                    .StageCode = "VERIFICATION",
                    .StageName = "Travel Desk & Budget Verification",
                    .ResponsibleOffice = "Travel Section",
                    .ActionRequired = "Verify per diem allowances and travel fund availability"
                },
                New RoutingStep With {
                    .StepNumber = 3,
                    .StageCode = "APPROVAL",
                    .StageName = "Executive Authorization",
                    .ResponsibleOffice = "Office of the Secretary-General",
                    .ActionRequired = "Review and sign official travel order"
                },
                New RoutingStep With {
                    .StepNumber = 4,
                    .StageCode = "ADVANCE",
                    .StageName = "Cash Advance & Ticket Issuance",
                    .ResponsibleOffice = "Cashier / Travel Desk",
                    .ActionRequired = "Issue travel advance and confirmed itinerary"
                },
                New RoutingStep With {
                    .StepNumber = 5,
                    .StageCode = "RELEASE",
                    .StageName = "Post-Travel Liquidation & Filing",
                    .ResponsibleOffice = "Travel Section",
                    .ActionRequired = "File certificate of appearance and travel liquidation"
                }
            }
        End Function

        Private Shared Function CreateLegislativePipeline() As List(Of RoutingStep)
            Return New List(Of RoutingStep) From {
                New RoutingStep With {
                    .StepNumber = 1,
                    .StageCode = "INTAKE",
                    .StageName = "Intake & Legislative Logging",
                    .ResponsibleOffice = "Records Section",
                    .ActionRequired = "Log legislative measure and generate digital tracking identifier"
                },
                New RoutingStep With {
                    .StepNumber = 2,
                    .StageCode = "REVIEW",
                    .StageName = "Legislative Technical Review",
                    .ResponsibleOffice = "Legislative Section",
                    .ActionRequired = "Substantive legal analysis and committee preparation"
                },
                New RoutingStep With {
                    .StepNumber = 3,
                    .StageCode = "APPROVAL",
                    .StageName = "Executive Review & Sign-Off",
                    .ResponsibleOffice = "Office of the Secretary-General",
                    .ActionRequired = "Executive review and official endorsement"
                },
                New RoutingStep With {
                    .StepNumber = 4,
                    .StageCode = "REFERRAL",
                    .StageName = "Plenary & Committee Referral",
                    .ResponsibleOffice = "Legislative Section",
                    .ActionRequired = "Transmit to committee or calendar for plenary session"
                },
                New RoutingStep With {
                    .StepNumber = 5,
                    .StageCode = "RELEASE",
                    .StageName = "Gazetting & Archival",
                    .ResponsibleOffice = "Records Section",
                    .ActionRequired = "Release official copy and archive legislative record"
                }
            }
        End Function

        Private Shared Function CreateCommunicationPipeline() As List(Of RoutingStep)
            Return New List(Of RoutingStep) From {
                New RoutingStep With {
                    .StepNumber = 1,
                    .StageCode = "INTAKE",
                    .StageName = "Intake & Classification",
                    .ResponsibleOffice = "Records Section",
                    .ActionRequired = "Register electronic transmittal and generate digital tracking barcode"
                },
                New RoutingStep With {
                    .StepNumber = 2,
                    .StageCode = "PROCESSING",
                    .StageName = "Desk Processing & Endorsement",
                    .ResponsibleOffice = "Secretariat",
                    .ActionRequired = "Review communication and draft executive endorsement"
                },
                New RoutingStep With {
                    .StepNumber = 3,
                    .StageCode = "APPROVAL",
                    .StageName = "Executive Action & Directive",
                    .ResponsibleOffice = "Office of the Secretary-General",
                    .ActionRequired = "Review communication and issue executive directive"
                },
                New RoutingStep With {
                    .StepNumber = 4,
                    .StageCode = "DISPATCH",
                    .StageName = "Transmittal & Dispatch",
                    .ResponsibleOffice = "Records Section",
                    .ActionRequired = "Transmit official communication to destination office"
                },
                New RoutingStep With {
                    .StepNumber = 5,
                    .StageCode = "ARCHIVE",
                    .StageName = "Central Archival",
                    .ResponsibleOffice = "Records Section",
                    .ActionRequired = "File official electronic record in OSG digital records archive"
                }
            }
        End Function

        Private Shared Function GetDocTypeIdentifier(doc As Document) As String
            If Not String.IsNullOrWhiteSpace(doc.DocCode) Then
                Dim prefix = doc.DocCode.Split("-"c)(0).ToUpperInvariant()
                If prefix = "FIN" Then Return "FIN"
                If prefix = "TO" OrElse prefix = "TRAVEL" Then Return "TRAVEL"
                If prefix = "LEG" OrElse prefix = "BLL" OrElse prefix = "RES" Then Return "LEG"
                If prefix = "COMM" OrElse prefix = "REG" Then Return "REG_COMM"
            End If

            If doc.DocumentTypeID = 3 Then Return "FIN"
            If doc.DocumentTypeID = 4 Then Return "TRAVEL"
            If doc.DocumentTypeID = 2 Then Return "LEG"
            Return "REG_COMM"
        End Function

        Private Shared Sub AssignStepStatuses(doc As Document, steps As List(Of RoutingStep), routingLogs As List(Of RoutingLog))
            Dim currentStepNumber As Integer = 1

            ' Derive current active station by status code or ID
            Select Case doc.StatusID
                Case 1 ' RECEIVED / LOGGED
                    currentStepNumber = 2
                Case 2 ' FOR_REVIEW
                    currentStepNumber = 3
                Case 3 ' FOR_REVISION
                    currentStepNumber = 2
                Case 6 ' APPROVED
                    currentStepNumber = 4
                Case 7 ' RELEASED
                    currentStepNumber = 5
                Case 8 ' FILED / ARCHIVED
                    currentStepNumber = 6 ' Beyond last step, all done
                Case Else
                    ' Check LastActionTaken or AssignedSection heuristic
                    If doc.LastActionTaken IsNot Nothing AndAlso doc.LastActionTaken.IndexOf("Approved", StringComparison.OrdinalIgnoreCase) >= 0 Then
                        currentStepNumber = 4
                    ElseIf doc.AssignedSection = "Secretary-General" Then
                        currentStepNumber = 3
                    ElseIf doc.AssignedSection = "Records Section" AndAlso doc.LastActionTaken IsNot Nothing AndAlso doc.LastActionTaken.IndexOf("Released", StringComparison.OrdinalIgnoreCase) >= 0 Then
                        currentStepNumber = 5
                    Else
                        currentStepNumber = 2
                    End If
            End Select

            For Each stp In steps
                If stp.StepNumber < currentStepNumber Then
                    stp.StepStatus = "COMPLETED"
                    stp.CompletedAtUTC = doc.RegisteredAtUTC.AddHours(stp.StepNumber * 2)
                    stp.CompletedBy = stp.ResponsibleOffice
                ElseIf stp.StepNumber = currentStepNumber Then
                    stp.StepStatus = "CURRENT"
                    If doc.StatusID = 3 Then
                        stp.Remarks = "Revision Punchlist: " & doc.RevisionPunchlist
                    End If
                Else
                    stp.StepStatus = "UPCOMING"
                End If
            Next

            ' Correlate actual timestamps from routing logs where available
            If routingLogs IsNot Nothing AndAlso routingLogs.Count > 0 Then
                Dim logIndex As Integer = 0
                For Each stp In steps
                    If stp.StepStatus = "COMPLETED" AndAlso logIndex < routingLogs.Count Then
                        stp.CompletedAtUTC = routingLogs(logIndex).RoutedAtUTC
                        stp.Remarks = routingLogs(logIndex).RoutingRemarks
                        logIndex += 1
                    End If
                Next
            End If
        End Sub

        Public Shared Function GetCurrentStation(steps As List(Of RoutingStep)) As RoutingStep
            If steps Is Nothing Then Return Nothing
            For Each stp In steps
                If stp.StepStatus = "CURRENT" Then Return stp
            Next
            If steps.Count > 0 AndAlso steps(steps.Count - 1).StepStatus = "COMPLETED" Then
                Return steps(steps.Count - 1)
            End If
            Return If(steps.Count > 0, steps(0), Nothing)
        End Function

        Public Shared Function GetNextStation(steps As List(Of RoutingStep)) As RoutingStep
            If steps Is Nothing Then Return Nothing
            Dim foundCurrent As Boolean = False
            For Each stp In steps
                If foundCurrent AndAlso stp.StepStatus = "UPCOMING" Then
                    Return stp
                End If
                If stp.StepStatus = "CURRENT" Then
                    foundCurrent = True
                End If
            Next
            Return Nothing
        End Function

        Public Shared Function GetProgressPercentage(steps As List(Of RoutingStep)) As Integer
            If steps Is Nothing OrElse steps.Count = 0 Then Return 0
            Dim completed As Integer = 0
            For Each stp In steps
                If stp.StepStatus = "COMPLETED" Then completed += 1
            Next
            Return CInt(Math.Round((completed / CDbl(steps.Count)) * 100.0))
        End Function
    End Class
End Namespace
