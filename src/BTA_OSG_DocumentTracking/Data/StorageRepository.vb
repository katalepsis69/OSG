Imports System.Collections.Generic
Imports System.Data
Imports Microsoft.Data.SqlClient

Namespace BTA_OSG
    Public Class StorageRepository
        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        Public Function GetAll() As List(Of StorageLocation)
            Dim list As New List(Of StorageLocation)()
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT * FROM tbl_StorageLocations"
                Using cmd = New SqlCommand(sql, conn)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            list.Add(MapStorageLocation(reader))
                        End While
                    End Using
                End Using
            End Using
            Return list
        End Function

        Public Function GetById(id As Integer) As StorageLocation
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT * FROM tbl_StorageLocations WHERE StorageID = @id"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@id", id)
                    Using reader = cmd.ExecuteReader()
                        If reader.Read() Then
                            Return MapStorageLocation(reader)
                        End If
                    End Using
                End Using
            End Using
            Return Nothing
        End Function

        Public Function Insert(loc As StorageLocation) As Integer
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "INSERT INTO tbl_StorageLocations (CabinetID, ShelfNo, BoxCode, Description, IsActive) " &
                          "OUTPUT INSERTED.StorageLocationID " &
                          "VALUES (@CabinetID, @ShelfNo, @BoxCode, @Description, @IsActive)"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@CabinetID", If(loc.CabinetID, DBNull.Value))
                    cmd.Parameters.AddWithValue("@ShelfNo", If(loc.ShelfNo, DBNull.Value))
                    cmd.Parameters.AddWithValue("@BoxCode", If(loc.BoxCode, DBNull.Value))
                    cmd.Parameters.AddWithValue("@Description", If(loc.Description, DBNull.Value))
                    cmd.Parameters.AddWithValue("@IsActive", loc.IsActive)
                    Return Convert.ToInt32(cmd.ExecuteScalar())
                End Using
            End Using
        End Function

        Public Function GetMovements(docId As Integer) As List(Of DocumentMovement)
            Dim list As New List(Of DocumentMovement)()
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT * FROM tbl_DocumentMovements WHERE DocumentID = @id ORDER BY MovedAtUTC DESC"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@id", docId)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            list.Add(New DocumentMovement With {
                                .MovementID = Convert.ToInt32(reader("MovementID")),
                                .DocumentID = Convert.ToInt32(reader("DocumentID")),
                                .StorageLocationID = Convert.ToInt32(reader("StorageLocationID")),
                                .MovedAtUTC = Convert.ToDateTime(reader("MovedAtUTC")),
                                .MovedByUserID = Convert.ToInt32(reader("MovedByUserID")),
                                .MovementReason = If(IsDBNull(reader("MovementReason")), Nothing, Convert.ToString(reader("MovementReason")))
                            })
                        End While
                    End Using
                End Using
            End Using
            Return list
        End Function

        Public Function InsertMovement(movement As DocumentMovement) As Integer
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "INSERT INTO tbl_DocumentMovements (DocumentID, StorageLocationID, MovedAtUTC, MovedByUserID, MovementReason) " &
                          "OUTPUT INSERTED.MovementID " &
                          "VALUES (@DocumentID, @StorageLocationID, @MovedAtUTC, @MovedByUserID, @MovementReason)"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@DocumentID", movement.DocumentID)
                    cmd.Parameters.AddWithValue("@StorageLocationID", movement.StorageLocationID)
                    cmd.Parameters.AddWithValue("@MovedAtUTC", movement.MovedAtUTC)
                    cmd.Parameters.AddWithValue("@MovedByUserID", movement.MovedByUserID)
                    cmd.Parameters.AddWithValue("@MovementReason", If(movement.MovementReason, DBNull.Value))
                    Return Convert.ToInt32(cmd.ExecuteScalar())
                End Using
            End Using
        End Function

        Private Function MapStorageLocation(reader As IDataReader) As StorageLocation
            Return New StorageLocation With {
                .StorageLocationID = Convert.ToInt32(reader("StorageLocationID")),
                .CabinetID = If(IsDBNull(reader("CabinetID")), Nothing, Convert.ToString(reader("CabinetID"))),
                .ShelfNo = If(IsDBNull(reader("ShelfNo")), Nothing, Convert.ToString(reader("ShelfNo"))),
                .BoxCode = If(IsDBNull(reader("BoxCode")), Nothing, Convert.ToString(reader("BoxCode"))),
                .Description = If(IsDBNull(reader("Description")), Nothing, Convert.ToString(reader("Description"))),
                .IsActive = Convert.ToBoolean(reader("IsActive"))
            }
        End Function
    End Class
End Namespace
