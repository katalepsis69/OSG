Namespace BTA_OSG
    Public Class DocumentCodingService
        Private ReadOnly _seqRepo As SequenceRepository
        Private ReadOnly _refRepo As ReferenceDataRepository

        Public Sub New(seqRepo As SequenceRepository, refRepo As ReferenceDataRepository)
            _seqRepo = seqRepo
            _refRepo = refRepo
        End Sub

        Public Function GenerateCode(typeCode As String) As String
            Dim docType = _refRepo.GetDocumentTypeByCode(typeCode)
            If docType Is Nothing Then Throw New ArgumentException("Unknown document type code: " & typeCode)
            Dim prefix = docType.Prefix
            Dim year = CShort(DateTime.Now.Year)
            Return _seqRepo.GetNextDocCode(typeCode, prefix, year)
        End Function
    End Class
End Namespace
