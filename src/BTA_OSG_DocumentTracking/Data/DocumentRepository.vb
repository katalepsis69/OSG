Imports System.Collections.Generic
Imports System.Data
Imports Microsoft.Data.SqlClient

Namespace BTA_OSG
    Public Class DocumentRepository
        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        Public Function GetById(docId As Integer) As Document
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT * FROM tbl_Documents WHERE DocumentID = @id"
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
                Dim sql = "SELECT * FROM tbl_Documents WHERE DocumentCode = @code"
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
                Dim sql = "SELECT * FROM tbl_Documents"
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
            Dim list As New List(Of Document)()
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT * FROM tbl_Documents WHERE IsDeleted = 0 "
                If Not String.IsNullOrEmpty(titleLike) Then sql &= " AND Title LIKE @title "
                If typeId.HasValue Then sql &= " AND DocumentTypeID = @typeId "
                If statusId.HasValue Then sql &= " AND StatusID = @statusId "
                If Not String.IsNullOrEmpty(originLike) Then sql &= " AND OriginatingOffice LIKE @origin "
                If Not String.IsNullOrEmpty(destLike) Then sql &= " AND DestinationOffice LIKE @dest "
                If storageId.HasValue Then sql &= " AND CurrentStorageID = @storageId "
                If dateFrom.HasValue Then sql &= " AND DocumentDate >= @dateFrom "
                If dateTo.HasValue Then sql &= " AND DocumentDate <= @dateTo "
                
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
                Dim sql = "INSERT INTO tbl_Documents (DocumentCode, Title, DocumentTypeID, StatusID, OriginatingOffice, DestinationOffice, CurrentStorageID, DocumentDate, Description) " &
                          "OUTPUT INSERTED.DocumentID " &
                          "VALUES (@DocumentCode, @Title, @DocumentTypeID, @StatusID, @OriginatingOffice, @DestinationOffice, @CurrentStorageID, @DocumentDate, @Description)"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@DocumentCode", doc.DocCode)
                    cmd.Parameters.AddWithValue("@Title", doc.Title)
                    cmd.Parameters.AddWithValue("@DocumentTypeID", doc.DocumentTypeID)
                    cmd.Parameters.AddWithValue("@StatusID", doc.StatusID)
                    cmd.Parameters.AddWithValue("@OriginatingOffice", If(doc.OriginOffice, DBNull.Value))
                    cmd.Parameters.AddWithValue("@DestinationOffice", If(doc.DestinationOffice, DBNull.Value))
                    cmd.Parameters.AddWithValue("@CurrentStorageID", If(CType(doc.CurrentStorageLocationID, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@DocumentDate", doc.ReceivedDate)
                    cmd.Parameters.AddWithValue("@Description", If(doc.Remarks, DBNull.Value))
                    Return Convert.ToInt32(cmd.ExecuteScalar())
                End Using
            End Using
        End Function

        Public Sub Update(doc As Document, modifiedBy As Integer?)
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "UPDATE tbl_Documents SET Title = @Title, DocumentTypeID = @DocumentTypeID, StatusID = @StatusID, " &
                          "OriginatingOffice = @OriginatingOffice, DestinationOffice = @DestinationOffice, CurrentStorageID = @CurrentStorageID, " &
                          "DocumentDate = @DocumentDate, Description = @Description " &
                          "WHERE DocumentID = @id"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@Title", doc.Title)
                    cmd.Parameters.AddWithValue("@DocumentTypeID", doc.DocumentTypeID)
                    cmd.Parameters.AddWithValue("@StatusID", doc.StatusID)
                    cmd.Parameters.AddWithValue("@OriginatingOffice", If(doc.OriginOffice, DBNull.Value))
                    cmd.Parameters.AddWithValue("@DestinationOffice", If(doc.DestinationOffice, DBNull.Value))
                    cmd.Parameters.AddWithValue("@CurrentStorageID", If(CType(doc.CurrentStorageLocationID, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@DocumentDate", doc.ReceivedDate)
                    cmd.Parameters.AddWithValue("@Description", If(doc.Remarks, DBNull.Value))
                    cmd.Parameters.AddWithValue("@id", doc.DocumentID)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Sub SoftDelete(docId As Integer, deletedBy As Integer, reason As String)
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "UPDATE tbl_Documents SET IsDeleted = 1, DeletedBy = @deletedBy, DeletedDate = GETDATE(), DeletionReason = @reason WHERE DocumentID = @id"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@deletedBy", deletedBy)
                    cmd.Parameters.AddWithValue("@reason", If(reason, DBNull.Value))
                    cmd.Parameters.AddWithValue("@id", docId)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Function GetAssignments(docId As Integer) As List(Of DocumentAssignment)
            Dim list As New List(Of DocumentAssignment)()
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT * FROM tbl_DocumentAssignments WHERE DocumentID = @id"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@id", docId)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            list.Add(New DocumentAssignment With {
                                .AssignmentID = Convert.ToInt32(reader("AssignmentID")),
                                .DocumentID = Convert.ToInt32(reader("DocumentID")),
                                .AssignedUserID = Convert.ToInt32(reader("AssignedToUserID")),
                                .AssignedAtUTC = Convert.ToDateTime(reader("AssignedDate")),
                                .Remarks = If(IsDBNull(reader("Notes")), Nothing, Convert.ToString(reader("Notes")))
                            })
                        End While
                    End Using
                End Using
            End Using
            Return list
        End Function

        Public Sub AddAssignment(assignment As DocumentAssignment)
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "INSERT INTO tbl_DocumentAssignments (DocumentID, AssignedToUserID, AssignedDate, Notes) " &
                          "VALUES (@DocumentID, @AssignedToUserID, @AssignedDate, @Notes)"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@DocumentID", assignment.DocumentID)
                    cmd.Parameters.AddWithValue("@AssignedToUserID", assignment.AssignedUserID)
                    cmd.Parameters.AddWithValue("@AssignedDate", assignment.AssignedAtUTC)
                    cmd.Parameters.AddWithValue("@Notes", If(assignment.Remarks, DBNull.Value))
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Private Function MapDocument(reader As IDataReader) As Document
            Return New Document With {
                .DocumentID = Convert.ToInt32(reader("DocumentID")),
                .DocCode = Convert.ToString(reader("DocumentCode")),
                .Title = Convert.ToString(reader("Title")),
                .DocumentTypeID = Convert.ToInt32(reader("DocumentTypeID")),
                .StatusID = Convert.ToInt32(reader("StatusID")),
                .OriginOffice = If(IsDBNull(reader("OriginatingOffice")), Nothing, Convert.ToString(reader("OriginatingOffice"))),
                .DestinationOffice = If(IsDBNull(reader("DestinationOffice")), Nothing, Convert.ToString(reader("DestinationOffice"))),
                .CurrentStorageLocationID = If(IsDBNull(reader("CurrentStorageID")), CType(Nothing, Integer?), Convert.ToInt32(reader("CurrentStorageID"))),
                .ReceivedDate = If(IsDBNull(reader("DocumentDate")), CType(Nothing, Date?), Convert.ToDateTime(reader("DocumentDate"))),
                .Remarks = If(IsDBNull(reader("Description")), Nothing, Convert.ToString(reader("Description"))),
                .IsDeleted = If(IsDBNull(reader("IsDeleted")), False, Convert.ToBoolean(reader("IsDeleted")))
            }
        End Function
    End Class
End Namespace
