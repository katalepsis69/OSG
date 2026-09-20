Imports System.Collections.Generic
Imports System.Data
Imports Microsoft.Data.SqlClient

Namespace BTA_OSG
    Public Class DocumentRepository
        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        Private Const DOC_COLS As String = "DocumentID, DocCode, Title, DocumentTypeID, OriginOffice, DestinationOffice, StatusID, ReceivedDate, CurrentStorageLocationID, Remarks, IsDeleted"
        Public Function GetById(docId As Integer) As Document
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT " & DOC_COLS & " FROM tbl_Documents WHERE DocumentID = @id"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@id", docId)
                    Using reader = cmd.ExecuteReader()
                        If reader.Read() Then
                            Return MapDocument(reader)
                        End If
                    End Using
                End Using
            End Using
            Return Nothing
        End Function

        Public Function GetByDocCode(code As String) As Document
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT " & DOC_COLS & " FROM tbl_Documents WHERE DocCode = @code"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@code", code)
                    Using reader = cmd.ExecuteReader()
                        If reader.Read() Then
                            Return MapDocument(reader)
                        End If
                    End Using
                End Using
            End Using
            Return Nothing
        End Function

        Public Function GetAll(includeDeleted As Boolean) As List(Of Document)
            Dim list As New List(Of Document)()
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT " & DOC_COLS & " FROM tbl_Documents"
                If Not includeDeleted Then
                    sql &= " WHERE IsDeleted = 0"
                End If
                Using cmd = New SqlCommand(sql, conn)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            list.Add(MapDocument(reader))
                        End While
                    End Using
                End Using
            End Using
            Return list
        End Function

        Public Function GetByFilter(titleLike As String, typeId As Integer?, statusId As Integer?, originLike As String, destLike As String, storageId As Integer?, dateFrom As Date?, dateTo As Date?, pageSize As Integer, pageNumber As Integer) As List(Of Document)
            pageSize = Math.Max(1, Math.Min(100, pageSize))
            pageNumber = Math.Max(1, pageNumber)
            Dim list As New List(Of Document)()
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT " & DOC_COLS & " FROM tbl_Documents WHERE IsDeleted = 0 "
                If Not String.IsNullOrEmpty(titleLike) Then sql &= " AND Title LIKE @title "
                If typeId.HasValue Then sql &= " AND DocumentTypeID = @typeId "
                If statusId.HasValue Then sql &= " AND StatusID = @statusId "
                If Not String.IsNullOrEmpty(originLike) Then sql &= " AND OriginOffice LIKE @origin "
                If Not String.IsNullOrEmpty(destLike) Then sql &= " AND DestinationOffice LIKE @dest "
                If storageId.HasValue Then sql &= " AND CurrentStorageLocationID = @storageId "
                If dateFrom.HasValue Then sql &= " AND ReceivedDate >= @dateFrom "
                If dateTo.HasValue Then sql &= " AND ReceivedDate <= @dateTo "
                
                sql &= " ORDER BY DocumentID OFFSET @offset ROWS FETCH NEXT @pageSize ROWS ONLY"

                Using cmd = New SqlCommand(sql, conn)
                    If Not String.IsNullOrEmpty(titleLike) Then cmd.Parameters.AddWithValue("@title", "%" & titleLike & "%")
                    If typeId.HasValue Then cmd.Parameters.AddWithValue("@typeId", typeId.Value)
                    If statusId.HasValue Then cmd.Parameters.AddWithValue("@statusId", statusId.Value)
                    If Not String.IsNullOrEmpty(originLike) Then cmd.Parameters.AddWithValue("@origin", "%" & originLike & "%")
                    If Not String.IsNullOrEmpty(destLike) Then cmd.Parameters.AddWithValue("@dest", "%" & destLike & "%")
                    If storageId.HasValue Then cmd.Parameters.AddWithValue("@storageId", storageId.Value)
                    If dateFrom.HasValue Then cmd.Parameters.AddWithValue("@dateFrom", dateFrom.Value)
                    If dateTo.HasValue Then cmd.Parameters.AddWithValue("@dateTo", dateTo.Value)
                    cmd.Parameters.AddWithValue("@offset", (pageNumber - 1) * pageSize)
                    cmd.Parameters.AddWithValue("@pageSize", pageSize)

                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            list.Add(MapDocument(reader))
                        End While
                    End Using
                End Using
            End Using
            Return list
        End Function

        Public Function GetVisibleToUser(userId As Integer, roles As List(Of Role), office As String, hasViewAll As Boolean, titleLike As String, pageSize As Integer, pageNumber As Integer) As List(Of Document)
            ' Simplified logic for visibility, normally this would check assignments or permissions
            Return GetByFilter(titleLike, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, pageSize, pageNumber)
        End Function

        Public Function Insert(doc As Document) As Integer
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "INSERT INTO tbl_Documents (DocCode, Title, DocumentTypeID, StatusID, OriginOffice, DestinationOffice, CurrentStorageLocationID, ReceivedDate, Remarks, CreatedByUserID) " &
                           "OUTPUT INSERTED.DocumentID " &
                           "VALUES (@DocCode, @Title, @DocumentTypeID, @StatusID, @OriginOffice, @DestinationOffice, @CurrentStorageLocationID, @ReceivedDate, @Remarks, @CreatedByUserID)"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@DocCode", doc.DocCode)
                    cmd.Parameters.AddWithValue("@Title", doc.Title)
                    cmd.Parameters.AddWithValue("@DocumentTypeID", doc.DocumentTypeID)
                    cmd.Parameters.AddWithValue("@StatusID", doc.StatusID)
                    cmd.Parameters.AddWithValue("@OriginOffice", If(doc.OriginOffice IsNot Nothing, CType(doc.OriginOffice, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@DestinationOffice", If(doc.DestinationOffice IsNot Nothing, CType(doc.DestinationOffice, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@CurrentStorageLocationID", If(CType(doc.CurrentStorageLocationID, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@ReceivedDate", If(CType(doc.ReceivedDate, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Remarks", If(doc.Remarks IsNot Nothing, CType(doc.Remarks, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@CreatedByUserID", doc.RegisteredByUserID)
                    Return Convert.ToInt32(cmd.ExecuteScalar())
                End Using
            End Using
        End Function

        Public Sub Update(doc As Document, modifiedBy As Integer?)
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "UPDATE tbl_Documents SET Title = @Title, DocumentTypeID = @DocumentTypeID, StatusID = @StatusID, " &
                           "OriginOffice = @OriginOffice, DestinationOffice = @DestinationOffice, CurrentStorageLocationID = @CurrentStorageLocationID, " &
                           "ReceivedDate = @ReceivedDate, Remarks = @Remarks, ModifiedByUserID=@ModifiedByUserID, ModifiedAtUTC=SYSUTCDATETIME() " &
                           "WHERE DocumentID = @id"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@Title", doc.Title)
                    cmd.Parameters.AddWithValue("@DocumentTypeID", doc.DocumentTypeID)
                    cmd.Parameters.AddWithValue("@StatusID", doc.StatusID)
                    cmd.Parameters.AddWithValue("@OriginOffice", If(doc.OriginOffice IsNot Nothing, CType(doc.OriginOffice, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@DestinationOffice", If(doc.DestinationOffice IsNot Nothing, CType(doc.DestinationOffice, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@CurrentStorageLocationID", If(CType(doc.CurrentStorageLocationID, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@ReceivedDate", If(CType(doc.ReceivedDate, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Remarks", If(doc.Remarks IsNot Nothing, CType(doc.Remarks, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@ModifiedByUserID", If(modifiedBy.HasValue, CType(modifiedBy.Value, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@id", doc.DocumentID)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Sub UpdateStatus(docId As Integer, statusId As Integer)
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "UPDATE tbl_Documents SET StatusID = @statusId, ModifiedAtUTC = SYSUTCDATETIME() WHERE DocumentID = @id"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@statusId", statusId)
                    cmd.Parameters.AddWithValue("@id", docId)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Sub UpdateStorageLocation(docId As Integer, storageLocationId As Integer)
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "UPDATE tbl_Documents SET CurrentStorageLocationID = @storageId, ModifiedAtUTC = SYSUTCDATETIME() WHERE DocumentID = @id"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@storageId", storageLocationId)
                    cmd.Parameters.AddWithValue("@id", docId)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Sub SoftDelete(docId As Integer, deletedBy As Integer, reason As String)
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "UPDATE tbl_Documents SET IsDeleted = 1, DeletedByUserID = @deletedBy, DeletedAtUTC = SYSUTCDATETIME(), DeletionReason = @reason WHERE DocumentID = @id"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@deletedBy", deletedBy)
                    cmd.Parameters.AddWithValue("@reason", If(reason IsNot Nothing, CType(reason, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@id", docId)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Function GetAssignments(docId As Integer) As List(Of DocumentAssignment)
            Dim list As New List(Of DocumentAssignment)()
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT AssignmentID, DocumentID, AssignedUserID, AssignedAtUTC, Remarks FROM tbl_DocumentAssignments WHERE DocumentID = @id"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@id", docId)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            list.Add(New DocumentAssignment With {
                                .AssignmentID = Convert.ToInt32(reader("AssignmentID")),
                                .DocumentID = Convert.ToInt32(reader("DocumentID")),
                                .AssignedUserID = If(IsDBNull(reader("AssignedUserID")), CType(Nothing, Integer?), Convert.ToInt32(reader("AssignedUserID"))),
                                .AssignedAtUTC = Convert.ToDateTime(reader("AssignedAtUTC")),
                                .Remarks = If(IsDBNull(reader("Remarks")), Nothing, Convert.ToString(reader("Remarks")))
                            })
                        End While
                    End Using
                End Using
            End Using
            Return list
        End Function

        Public Sub AddAssignment(assignment As DocumentAssignment)
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "INSERT INTO tbl_DocumentAssignments (DocumentID, AssignedUserID, AssignedByUserID, AssignedAtUTC, Remarks) " &
                           "VALUES (@DocumentID, @AssignedUserID, @AssignedByUserID, @AssignedAtUTC, @Remarks)"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@DocumentID", assignment.DocumentID)
                    cmd.Parameters.AddWithValue("@AssignedUserID", assignment.AssignedUserID)
                    cmd.Parameters.AddWithValue("@AssignedByUserID", assignment.AssignedByUserID)
                    cmd.Parameters.AddWithValue("@AssignedAtUTC", assignment.AssignedAtUTC)
                    cmd.Parameters.AddWithValue("@Remarks", If(assignment.Remarks IsNot Nothing, CType(assignment.Remarks, Object), DBNull.Value))
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Function SearchDocuments(hasViewAll As Boolean, userId As Integer, titleLike As String, typeId As Integer?, statusId As Integer?, originLike As String, destLike As String, storageId As Integer?, dateFrom As DateTime?, dateTo As DateTime?, pageSize As Integer, pageNumber As Integer) As List(Of Document)
            Return GetByFilter(titleLike, typeId, statusId, originLike, destLike, storageId, dateFrom, dateTo, pageSize, pageNumber)
        End Function

        Private Function MapDocument(reader As IDataReader) As Document
            Return New Document With {
                .DocumentID = Convert.ToInt32(reader("DocumentID")),
                .DocCode = Convert.ToString(reader("DocCode")),
                .Title = Convert.ToString(reader("Title")),
                .DocumentTypeID = Convert.ToInt32(reader("DocumentTypeID")),
                .StatusID = Convert.ToInt32(reader("StatusID")),
                .OriginOffice = If(IsDBNull(reader("OriginOffice")), Nothing, Convert.ToString(reader("OriginOffice"))),
                .DestinationOffice = If(IsDBNull(reader("DestinationOffice")), Nothing, Convert.ToString(reader("DestinationOffice"))),
                .CurrentStorageLocationID = If(IsDBNull(reader("CurrentStorageLocationID")), CType(Nothing, Integer?), Convert.ToInt32(reader("CurrentStorageLocationID"))),
                .ReceivedDate = If(IsDBNull(reader("ReceivedDate")), CType(Nothing, Date?), Convert.ToDateTime(reader("ReceivedDate"))),
                .Remarks = If(IsDBNull(reader("Remarks")), Nothing, Convert.ToString(reader("Remarks"))),
                .IsDeleted = If(IsDBNull(reader("IsDeleted")), False, Convert.ToBoolean(reader("IsDeleted")))
            }
        End Function
    End Class
End Namespace
