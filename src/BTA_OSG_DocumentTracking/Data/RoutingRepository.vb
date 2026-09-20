Imports System.Collections.Generic
Imports System.Data
Imports Microsoft.Data.SqlClient

Namespace BTA_OSG
    Public Class RoutingRepository
        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        Public Function GetByDocumentId(docId As Integer) As List(Of RoutingLog)
            Dim list As New List(Of RoutingLog)()
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT * FROM tbl_RoutingLogs WHERE DocumentID = @id ORDER BY RoutedAtUTC DESC"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@id", docId)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            list.Add(New RoutingLog With {
                                .RoutingLogID = Convert.ToInt32(reader("RoutingLogID")),
                                .DocumentID = Convert.ToInt32(reader("DocumentID")),
                                .FromStatusID = If(IsDBNull(reader("FromStatusID")), CType(Nothing, Integer?), Convert.ToInt32(reader("FromStatusID"))),
                                .ToStatusID = Convert.ToInt32(reader("ToStatusID")),
                                .FromOffice = If(IsDBNull(reader("FromOffice")), Nothing, Convert.ToString(reader("FromOffice"))),
                                .ToOffice = If(IsDBNull(reader("ToOffice")), Nothing, Convert.ToString(reader("ToOffice"))),
                                .RoutedByUserID = Convert.ToInt32(reader("RoutedByUserID")),
                                .RoutedAtUTC = Convert.ToDateTime(reader("RoutedAtUTC")),
                                .RoutingRemarks = If(IsDBNull(reader("RoutingRemarks")), Nothing, Convert.ToString(reader("RoutingRemarks")))
                            })
                        End While
                    End Using
                End Using
            End Using
            Return list
        End Function

        Public Function Insert(log As RoutingLog) As Integer
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "INSERT INTO tbl_RoutingLogs (DocumentID, FromStatusID, ToStatusID, FromOffice, ToOffice, RoutedByUserID, RoutedAtUTC, RoutingRemarks) " &
                          "OUTPUT INSERTED.RoutingLogID " &
                          "VALUES (@DocumentID, @FromStatusID, @ToStatusID, @FromOffice, @ToOffice, @RoutedByUserID, @RoutedAtUTC, @RoutingRemarks)"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@DocumentID", log.DocumentID)
                    cmd.Parameters.AddWithValue("@FromStatusID", If(log.FromStatusID.HasValue, CType(log.FromStatusID.Value, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@ToStatusID", log.ToStatusID)
                    cmd.Parameters.AddWithValue("@FromOffice", If(log.FromOffice IsNot Nothing, CType(log.FromOffice, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@ToOffice", If(log.ToOffice IsNot Nothing, CType(log.ToOffice, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@RoutedByUserID", log.RoutedByUserID)
                    cmd.Parameters.AddWithValue("@RoutedAtUTC", log.RoutedAtUTC)
                    cmd.Parameters.AddWithValue("@RoutingRemarks", If(log.RoutingRemarks IsNot Nothing, CType(log.RoutingRemarks, Object), DBNull.Value))
                    Return Convert.ToInt32(cmd.ExecuteScalar())
                End Using
            End Using
        End Function
    End Class
End Namespace
