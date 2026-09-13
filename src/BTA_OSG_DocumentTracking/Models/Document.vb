Imports System

Namespace BTA_OSG
    Public Class Document
        Public Property DocumentID As Integer
        Public Property DocCode As String
        Public Property Title As String
        Public Property DocumentTypeID As Integer
        Public Property OriginOffice As String
        Public Property DestinationOffice As String
        Public Property StatusID As Integer
        Public Property RegisteredByUserID As Integer
        Public Property RegisteredAtUTC As DateTime
        Public Property ReceivedDate As Date?
        Public Property CurrentStorageLocationID As Integer?
        Public Property GoogleDriveUrl As String
        Public Property Remarks As String
        Public Property IsDeleted As Boolean
        Public Property DeletedByUserID As Integer?
        Public Property DeletedAtUTC As DateTime?
        Public Property DeletionReason As String
    End Class
End Namespace
