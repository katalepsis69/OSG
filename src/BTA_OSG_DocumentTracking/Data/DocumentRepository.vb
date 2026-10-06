Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports System.Data
Imports Microsoft.Data.SqlClient

Namespace BTA_OSG
    Public Class DocumentRepository
        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        Public ReadOnly Property ConnectionFactory As IDbConnectionFactory
            Get
                Return _connectionFactory
            End Get
        End Property

        Private Const DOC_COLS As String = "DocumentID, DocCode, Title, DocumentTypeID, OriginOffice, DestinationOffice, StatusID, ReceivedDate, CurrentStorageLocationID, Remarks, IsDeleted, FlowDirection, AssignedSection, TargetDeadlineUTC, RevisionPunchlist, LastActionTaken, ExternalControlNumber, GoogleDriveUrl, RequesterGender"
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

        ' When a transaction is supplied its connection is already open and the command
        ' joins it; otherwise the method owns a short-lived connection, as before.
        Public Function Insert(doc As Document, Optional transaction As SqlTransaction = Nothing) As Integer
            Dim sql = "INSERT INTO tbl_Documents (DocCode, Title, DocumentTypeID, StatusID, OriginOffice, DestinationOffice, CurrentStorageLocationID, ReceivedDate, Remarks, CreatedByUserID, FlowDirection, AssignedSection, TargetDeadlineUTC, RevisionPunchlist, LastActionTaken, ExternalControlNumber, GoogleDriveUrl, RequesterGender) " &
                       "OUTPUT INSERTED.DocumentID " &
                       "VALUES (@DocCode, @Title, @DocumentTypeID, @StatusID, @OriginOffice, @DestinationOffice, @CurrentStorageLocationID, @ReceivedDate, @Remarks, @CreatedByUserID, @FlowDirection, @AssignedSection, @TargetDeadlineUTC, @RevisionPunchlist, @LastActionTaken, @ExternalControlNumber, @GoogleDriveUrl, @RequesterGender)"
            Dim conn As SqlConnection = If(transaction IsNot Nothing, transaction.Connection, _connectionFactory.CreateConnection())
            Try
                Using cmd = New SqlCommand(sql, conn, transaction)
                    cmd.Parameters.AddWithValue("@DocCode", doc.DocCode)
                    cmd.Parameters.AddWithValue("@Title", doc.Title)
                    cmd.Parameters.AddWithValue("@DocumentTypeID", doc.DocumentTypeID)
                    cmd.Parameters.AddWithValue("@StatusID", doc.StatusID)
                    cmd.Parameters.AddWithValue("@OriginOffice", If(doc.OriginOffice IsNot Nothing, CType(doc.OriginOffice, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@DestinationOffice", If(doc.DestinationOffice IsNot Nothing, CType(doc.DestinationOffice, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@CurrentStorageLocationID", If(doc.CurrentStorageLocationID.HasValue, CType(doc.CurrentStorageLocationID.Value, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@ReceivedDate", If(doc.ReceivedDate.HasValue, CType(doc.ReceivedDate.Value, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Remarks", If(doc.Remarks IsNot Nothing, CType(doc.Remarks, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@CreatedByUserID", doc.RegisteredByUserID)
                    cmd.Parameters.AddWithValue("@FlowDirection", If(doc.FlowDirection IsNot Nothing, CType(doc.FlowDirection, Object), "INCOMING"))
                    cmd.Parameters.AddWithValue("@AssignedSection", If(doc.AssignedSection IsNot Nothing, CType(doc.AssignedSection, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@TargetDeadlineUTC", If(doc.TargetDeadlineUTC.HasValue, CType(doc.TargetDeadlineUTC.Value, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@RevisionPunchlist", If(doc.RevisionPunchlist IsNot Nothing, CType(doc.RevisionPunchlist, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@LastActionTaken", If(doc.LastActionTaken IsNot Nothing, CType(doc.LastActionTaken, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@ExternalControlNumber", ExternalControlNumberParam(doc.ExternalControlNumber))
                    cmd.Parameters.AddWithValue("@GoogleDriveUrl", EmptyToDbNull(doc.GoogleDriveUrl))
                    cmd.Parameters.AddWithValue("@RequesterGender", EmptyToDbNull(doc.RequesterGender))
                    Return Convert.ToInt32(cmd.ExecuteScalar())
                End Using
            Finally
                If transaction Is Nothing Then conn.Dispose()
            End Try
        End Function

        Public Sub Update(doc As Document, modifiedBy As Integer?)
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "UPDATE tbl_Documents SET Title = @Title, DocumentTypeID = @DocumentTypeID, StatusID = @StatusID, " &
                           "OriginOffice = @OriginOffice, DestinationOffice = @DestinationOffice, CurrentStorageLocationID = @CurrentStorageLocationID, " &
                           "ReceivedDate = @ReceivedDate, Remarks = @Remarks, FlowDirection = @FlowDirection, AssignedSection = @AssignedSection, " &
                           "TargetDeadlineUTC = @TargetDeadlineUTC, RevisionPunchlist = @RevisionPunchlist, LastActionTaken = @LastActionTaken, " &
                           "ExternalControlNumber = @ExternalControlNumber, GoogleDriveUrl = @GoogleDriveUrl, RequesterGender = @RequesterGender, " &
                           "ModifiedByUserID=@ModifiedByUserID, ModifiedAtUTC=SYSUTCDATETIME() " &
                           "WHERE DocumentID = @id"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@Title", doc.Title)
                    cmd.Parameters.AddWithValue("@DocumentTypeID", doc.DocumentTypeID)
                    cmd.Parameters.AddWithValue("@StatusID", doc.StatusID)
                    cmd.Parameters.AddWithValue("@OriginOffice", If(doc.OriginOffice IsNot Nothing, CType(doc.OriginOffice, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@DestinationOffice", If(doc.DestinationOffice IsNot Nothing, CType(doc.DestinationOffice, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@CurrentStorageLocationID", If(doc.CurrentStorageLocationID.HasValue, CType(doc.CurrentStorageLocationID.Value, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@ReceivedDate", If(doc.ReceivedDate.HasValue, CType(doc.ReceivedDate.Value, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Remarks", If(doc.Remarks IsNot Nothing, CType(doc.Remarks, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@FlowDirection", If(doc.FlowDirection IsNot Nothing, CType(doc.FlowDirection, Object), "INCOMING"))
                    cmd.Parameters.AddWithValue("@AssignedSection", If(doc.AssignedSection IsNot Nothing, CType(doc.AssignedSection, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@TargetDeadlineUTC", If(doc.TargetDeadlineUTC.HasValue, CType(doc.TargetDeadlineUTC.Value, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@RevisionPunchlist", If(doc.RevisionPunchlist IsNot Nothing, CType(doc.RevisionPunchlist, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@LastActionTaken", If(doc.LastActionTaken IsNot Nothing, CType(doc.LastActionTaken, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@ExternalControlNumber", ExternalControlNumberParam(doc.ExternalControlNumber))
                    cmd.Parameters.AddWithValue("@GoogleDriveUrl", EmptyToDbNull(doc.GoogleDriveUrl))
                    cmd.Parameters.AddWithValue("@RequesterGender", EmptyToDbNull(doc.RequesterGender))
                    cmd.Parameters.AddWithValue("@ModifiedByUserID", If(modifiedBy.HasValue, CType(modifiedBy.Value, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@id", doc.DocumentID)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Sub UpdateStatus(docId As Integer, statusId As Integer, Optional transaction As SqlTransaction = Nothing)
            Dim sql = "UPDATE tbl_Documents SET StatusID = @statusId, ModifiedAtUTC = SYSUTCDATETIME() WHERE DocumentID = @id"
            If transaction IsNot Nothing Then
                Using cmd = New SqlCommand(sql, transaction.Connection, transaction)
                    cmd.Parameters.AddWithValue("@statusId", statusId)
                    cmd.Parameters.AddWithValue("@id", docId)
                    cmd.ExecuteNonQuery()
                End Using
            Else
                Using conn = _connectionFactory.CreateConnection()
                    Using cmd = New SqlCommand(sql, conn)
                        cmd.Parameters.AddWithValue("@statusId", statusId)
                        cmd.Parameters.AddWithValue("@id", docId)
                        cmd.ExecuteNonQuery()
                    End Using
                End Using
            End If
        End Sub

        Public Sub UpdateWorkflowState(docId As Integer, statusId As Integer, assignedSection As String, punchlist As String, lastAction As String, modifiedBy As Integer?, Optional transaction As SqlTransaction = Nothing, Optional destinationOffice As String = Nothing, Optional originOffice As String = Nothing)
            Dim sql = "UPDATE tbl_Documents SET StatusID = @statusId, " &
                       "AssignedSection = COALESCE(@assignedSection, AssignedSection), " &
                       "RevisionPunchlist = COALESCE(@punchlist, RevisionPunchlist), " &
                       "LastActionTaken = COALESCE(@lastAction, LastActionTaken), " &
                       "DestinationOffice = COALESCE(@destinationOffice, DestinationOffice), " &
                       "OriginOffice = COALESCE(@originOffice, OriginOffice), " &
                       "ModifiedByUserID = @modifiedBy, ModifiedAtUTC = SYSUTCDATETIME() " &
                       "WHERE DocumentID = @id"
            If transaction IsNot Nothing Then
                Using cmd = New SqlCommand(sql, transaction.Connection, transaction)
                    cmd.Parameters.AddWithValue("@statusId", statusId)
                    cmd.Parameters.AddWithValue("@assignedSection", If(assignedSection IsNot Nothing, CType(assignedSection, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@punchlist", If(punchlist IsNot Nothing, CType(punchlist, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@lastAction", If(lastAction IsNot Nothing, CType(lastAction, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@destinationOffice", If(destinationOffice IsNot Nothing, CType(destinationOffice, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@originOffice", If(originOffice IsNot Nothing, CType(originOffice, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@modifiedBy", If(modifiedBy.HasValue, CType(modifiedBy.Value, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@id", docId)
                    cmd.ExecuteNonQuery()
                End Using
            Else
                Using conn = _connectionFactory.CreateConnection()
                    Using cmd = New SqlCommand(sql, conn)
                        cmd.Parameters.AddWithValue("@statusId", statusId)
                        cmd.Parameters.AddWithValue("@assignedSection", If(assignedSection IsNot Nothing, CType(assignedSection, Object), DBNull.Value))
                        cmd.Parameters.AddWithValue("@punchlist", If(punchlist IsNot Nothing, CType(punchlist, Object), DBNull.Value))
                        cmd.Parameters.AddWithValue("@lastAction", If(lastAction IsNot Nothing, CType(lastAction, Object), DBNull.Value))
                        cmd.Parameters.AddWithValue("@destinationOffice", If(destinationOffice IsNot Nothing, CType(destinationOffice, Object), DBNull.Value))
                        cmd.Parameters.AddWithValue("@originOffice", If(originOffice IsNot Nothing, CType(originOffice, Object), DBNull.Value))
                        cmd.Parameters.AddWithValue("@modifiedBy", If(modifiedBy.HasValue, CType(modifiedBy.Value, Object), DBNull.Value))
                        cmd.Parameters.AddWithValue("@id", docId)
                        cmd.ExecuteNonQuery()
                    End Using
                End Using
            End If
        End Sub

        ''' <summary>
        ''' From-status guard for the connected workflow actions: the update lands only when
        ''' the server row still carries the status the operator's form was showing, so a
        ''' document another seat already approved, released, or archived cannot be pushed
        ''' through the same transition again. Returns the rows affected; 0 means the state
        ''' moved and the caller must surface that instead of writing a duplicate.
        ''' </summary>
        Public Function UpdateWorkflowStateFromStatus(docId As Integer, expectedFromStatusId As Integer, statusId As Integer, assignedSection As String, punchlist As String, lastAction As String, modifiedBy As Integer?, Optional transaction As SqlTransaction = Nothing, Optional destinationOffice As String = Nothing, Optional originOffice As String = Nothing) As Integer
            Dim sql = "UPDATE tbl_Documents SET StatusID = @statusId, " &
                       "AssignedSection = COALESCE(@assignedSection, AssignedSection), " &
                       "RevisionPunchlist = COALESCE(@punchlist, RevisionPunchlist), " &
                       "LastActionTaken = COALESCE(@lastAction, LastActionTaken), " &
                       "DestinationOffice = COALESCE(@destinationOffice, DestinationOffice), " &
                       "OriginOffice = COALESCE(@originOffice, OriginOffice), " &
                       "ModifiedByUserID = @modifiedBy, ModifiedAtUTC = SYSUTCDATETIME() " &
                       "WHERE DocumentID = @id AND StatusID = @expectedFrom"
            Dim Execute = Sub(cmd As SqlCommand)
                              cmd.Parameters.AddWithValue("@statusId", statusId)
                              cmd.Parameters.AddWithValue("@assignedSection", If(assignedSection IsNot Nothing, CType(assignedSection, Object), DBNull.Value))
                              cmd.Parameters.AddWithValue("@punchlist", If(punchlist IsNot Nothing, CType(punchlist, Object), DBNull.Value))
                              cmd.Parameters.AddWithValue("@lastAction", If(lastAction IsNot Nothing, CType(lastAction, Object), DBNull.Value))
                              cmd.Parameters.AddWithValue("@destinationOffice", If(destinationOffice IsNot Nothing, CType(destinationOffice, Object), DBNull.Value))
                              cmd.Parameters.AddWithValue("@originOffice", If(originOffice IsNot Nothing, CType(originOffice, Object), DBNull.Value))
                              cmd.Parameters.AddWithValue("@modifiedBy", If(modifiedBy.HasValue, CType(modifiedBy.Value, Object), DBNull.Value))
                              cmd.Parameters.AddWithValue("@id", docId)
                              cmd.Parameters.AddWithValue("@expectedFrom", expectedFromStatusId)
                          End Sub
            If transaction IsNot Nothing Then
                Using cmd = New SqlCommand(sql, transaction.Connection, transaction)
                    Execute(cmd)
                    Return cmd.ExecuteNonQuery()
                End Using
            Else
                Using conn = _connectionFactory.CreateConnection()
                    Using cmd = New SqlCommand(sql, conn)
                        Execute(cmd)
                        Return cmd.ExecuteNonQuery()
                    End Using
                End Using
            End If
        End Function

        ''' <summary>
        ''' Optimistic-concurrency variant for the offline replay: the update lands only when
        ''' the server row still carries the rowversion this seat mirrored before going
        ''' offline. Returns the rows affected, so 0 means another seat already moved the
        ''' document past that version and the caller must not stomp the newer truth.
        ''' </summary>
        Public Function UpdateWorkflowStateChecked(docId As Integer, expectedRowVersion As Byte(), statusId As Integer, assignedSection As String, punchlist As String, lastAction As String, modifiedBy As Integer?, Optional destinationOffice As String = Nothing, Optional originOffice As String = Nothing) As Integer
            Dim sql = "UPDATE tbl_Documents SET StatusID = @statusId, " &
                       "AssignedSection = COALESCE(@assignedSection, AssignedSection), " &
                       "RevisionPunchlist = COALESCE(@punchlist, RevisionPunchlist), " &
                       "LastActionTaken = COALESCE(@lastAction, LastActionTaken), " &
                       "DestinationOffice = COALESCE(@destinationOffice, DestinationOffice), " &
                       "OriginOffice = COALESCE(@originOffice, OriginOffice), " &
                       "ModifiedByUserID = @modifiedBy, ModifiedAtUTC = SYSUTCDATETIME() " &
                       "WHERE DocumentID = @id AND RowVersion = @rv"
            Using conn = _connectionFactory.CreateConnection()
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@statusId", statusId)
                    cmd.Parameters.AddWithValue("@assignedSection", If(assignedSection IsNot Nothing, CType(assignedSection, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@punchlist", If(punchlist IsNot Nothing, CType(punchlist, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@lastAction", If(lastAction IsNot Nothing, CType(lastAction, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@destinationOffice", If(destinationOffice IsNot Nothing, CType(destinationOffice, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@originOffice", If(originOffice IsNot Nothing, CType(originOffice, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@modifiedBy", If(modifiedBy.HasValue, CType(modifiedBy.Value, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@id", docId)
                    cmd.Parameters.AddWithValue("@rv", expectedRowVersion)
                    Return cmd.ExecuteNonQuery()
                End Using
            End Using
        End Function

        Public Sub UpdateStorageLocation(docId As Integer, storageLocationId As Integer, Optional transaction As SqlTransaction = Nothing)
            Dim sql = "UPDATE tbl_Documents SET CurrentStorageLocationID = @storageId, ModifiedAtUTC = SYSUTCDATETIME() WHERE DocumentID = @id"
            If transaction IsNot Nothing Then
                Using cmd = New SqlCommand(sql, transaction.Connection, transaction)
                    cmd.Parameters.AddWithValue("@storageId", storageLocationId)
                    cmd.Parameters.AddWithValue("@id", docId)
                    cmd.ExecuteNonQuery()
                End Using
            Else
                Using conn = _connectionFactory.CreateConnection()
                    Using cmd = New SqlCommand(sql, conn)
                        cmd.Parameters.AddWithValue("@storageId", storageLocationId)
                        cmd.Parameters.AddWithValue("@id", docId)
                        cmd.ExecuteNonQuery()
                    End Using
                End Using
            End If
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

        Public Sub AddAssignment(assignment As DocumentAssignment, Optional transaction As SqlTransaction = Nothing)
            Dim sql = "INSERT INTO tbl_DocumentAssignments (DocumentID, AssignedUserID, AssignedByUserID, AssignedAtUTC, Remarks) " &
                       "VALUES (@DocumentID, @AssignedUserID, @AssignedByUserID, @AssignedAtUTC, @Remarks)"
            Dim conn As SqlConnection = If(transaction IsNot Nothing, transaction.Connection, _connectionFactory.CreateConnection())
            Try
                Using cmd = New SqlCommand(sql, conn, transaction)
                    cmd.Parameters.AddWithValue("@DocumentID", assignment.DocumentID)
                    cmd.Parameters.AddWithValue("@AssignedUserID", assignment.AssignedUserID)
                    cmd.Parameters.AddWithValue("@AssignedByUserID", assignment.AssignedByUserID)
                    cmd.Parameters.AddWithValue("@AssignedAtUTC", assignment.AssignedAtUTC)
                    cmd.Parameters.AddWithValue("@Remarks", If(assignment.Remarks IsNot Nothing, CType(assignment.Remarks, Object), DBNull.Value))
                    cmd.ExecuteNonQuery()
                End Using
            Finally
                If transaction Is Nothing Then conn.Dispose()
            End Try
        End Sub

        Public Function GetBySection(sectionName As String, ongoingOnly As Boolean, pageSize As Integer, pageNumber As Integer) As List(Of Document)
            pageSize = Math.Max(1, Math.Min(100, pageSize))
            pageNumber = Math.Max(1, pageNumber)
            Dim list As New List(Of Document)()
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT " & DOC_COLS & " FROM tbl_Documents WHERE IsDeleted = 0 "
                If Not String.IsNullOrEmpty(sectionName) AndAlso sectionName <> "All Sections" Then
                    sql &= " AND AssignedSection = @sec "
                End If
                If ongoingOnly Then
                    sql &= " AND StatusID NOT IN (SELECT StatusID FROM tbl_DocumentStatuses WHERE StatusCode IN ('FILED', 'RELEASED')) "
                End If
                sql &= " ORDER BY DocumentID DESC OFFSET @offset ROWS FETCH NEXT @pageSize ROWS ONLY"

                Using cmd = New SqlCommand(sql, conn)
                    If Not String.IsNullOrEmpty(sectionName) AndAlso sectionName <> "All Sections" Then
                        cmd.Parameters.AddWithValue("@sec", sectionName)
                    End If
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
                .IsDeleted = If(IsDBNull(reader("IsDeleted")), False, Convert.ToBoolean(reader("IsDeleted"))),
                .FlowDirection = If(IsDBNull(reader("FlowDirection")), "INCOMING", Convert.ToString(reader("FlowDirection"))),
                .AssignedSection = If(IsDBNull(reader("AssignedSection")), Nothing, Convert.ToString(reader("AssignedSection"))),
                .TargetDeadlineUTC = If(IsDBNull(reader("TargetDeadlineUTC")), CType(Nothing, DateTime?), Convert.ToDateTime(reader("TargetDeadlineUTC"))),
                .RevisionPunchlist = If(IsDBNull(reader("RevisionPunchlist")), Nothing, Convert.ToString(reader("RevisionPunchlist"))),
                .LastActionTaken = If(IsDBNull(reader("LastActionTaken")), Nothing, Convert.ToString(reader("LastActionTaken"))),
                .ExternalControlNumber = If(IsDBNull(reader("ExternalControlNumber")), Nothing, Convert.ToString(reader("ExternalControlNumber"))),
                .GoogleDriveUrl = If(IsDBNull(reader("GoogleDriveUrl")), Nothing, Convert.ToString(reader("GoogleDriveUrl"))),
                .RequesterGender = If(IsDBNull(reader("RequesterGender")), Nothing, Convert.ToString(reader("RequesterGender")))
            }
        End Function

        ' The filtered unique index on ExternalControlNumber excludes only NULL, so an
        ' absent control number must be stored as NULL: an empty string would collide
        ' across every non-portal document in the registry.
        Private Shared Function ExternalControlNumberParam(value As String) As Object
            If String.IsNullOrWhiteSpace(value) Then Return DBNull.Value
            Return CType(value, Object)
        End Function

        ' GoogleDriveUrl and RequesterGender have no unique index, but NULL still reads
        ' better than '' for rows that never carried the field.
        Private Shared Function EmptyToDbNull(value As String) As Object
            If String.IsNullOrWhiteSpace(value) Then Return DBNull.Value
            Return CType(value, Object)
        End Function

        Public Function GetByExternalControlNumber(externalCn As String) As Document
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT " & DOC_COLS & " FROM tbl_Documents WHERE ExternalControlNumber = @ecn AND IsDeleted = 0"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@ecn", externalCn)
                    Using reader = cmd.ExecuteReader()
                        If reader.Read() Then
                            Return MapDocument(reader)
                        End If
                    End Using
                End Using
            End Using
            Return Nothing
        End Function
    End Class
End Namespace
