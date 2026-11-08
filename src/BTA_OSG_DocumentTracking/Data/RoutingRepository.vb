Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports System.Data
Imports Microsoft.Data.SqlClient

Namespace BTA_OSG
    Public Class RoutingRepository
        Private ReadOnly _connectionFactory As SqlConnectionFactory

        Public Sub New(connectionFactory As SqlConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        Public Function Insert(log As RoutingLog, Optional transaction As SqlTransaction = Nothing) As Integer
            Dim sql = "INSERT INTO tbl_RoutingLogs (DocumentID, FromStatusID, ToStatusID, FromOffice, ToOffice, RoutedByUserID, RoutedAtUTC, RoutingRemarks) " &
                      "OUTPUT INSERTED.RoutingLogID " &
                      "VALUES (@DocumentID, @FromStatusID, @ToStatusID, @FromOffice, @ToOffice, @RoutedByUserID, @RoutedAtUTC, @RoutingRemarks)"
            If transaction IsNot Nothing Then
                Using cmd = New SqlCommand(sql, transaction.Connection, transaction)
                    cmd.Parameters.AddWithValue("@DocumentID", log.DocumentID)
                    cmd.Parameters.AddWithValue("@FromStatusID", If(log.FromStatusID.HasValue, CType(log.FromStatusID.Value, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@ToStatusID", If(log.ToStatusID.HasValue, CType(log.ToStatusID.Value, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@FromOffice", If(log.FromOffice IsNot Nothing, CType(log.FromOffice, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@ToOffice", If(log.ToOffice IsNot Nothing, CType(log.ToOffice, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@RoutedByUserID", log.RoutedByUserID)
                    cmd.Parameters.AddWithValue("@RoutedAtUTC", log.RoutedAtUTC)
                    cmd.Parameters.AddWithValue("@RoutingRemarks", If(log.RoutingRemarks IsNot Nothing, CType(log.RoutingRemarks, Object), DBNull.Value))
                    Return Convert.ToInt32(cmd.ExecuteScalar())
                End Using
            Else
                Using conn = _connectionFactory.CreateConnection()
                    Using cmd = New SqlCommand(sql, conn)
                        cmd.Parameters.AddWithValue("@DocumentID", log.DocumentID)
                        cmd.Parameters.AddWithValue("@FromStatusID", If(log.FromStatusID.HasValue, CType(log.FromStatusID.Value, Object), DBNull.Value))
                        cmd.Parameters.AddWithValue("@ToStatusID", If(log.ToStatusID.HasValue, CType(log.ToStatusID.Value, Object), DBNull.Value))
                        cmd.Parameters.AddWithValue("@FromOffice", If(log.FromOffice IsNot Nothing, CType(log.FromOffice, Object), DBNull.Value))
                        cmd.Parameters.AddWithValue("@ToOffice", If(log.ToOffice IsNot Nothing, CType(log.ToOffice, Object), DBNull.Value))
                        cmd.Parameters.AddWithValue("@RoutedByUserID", log.RoutedByUserID)
                        cmd.Parameters.AddWithValue("@RoutedAtUTC", log.RoutedAtUTC)
                        cmd.Parameters.AddWithValue("@RoutingRemarks", If(log.RoutingRemarks IsNot Nothing, CType(log.RoutingRemarks, Object), DBNull.Value))
                        Return Convert.ToInt32(cmd.ExecuteScalar())
                    End Using
                End Using
            End If
        End Function
    End Class
End Namespace
