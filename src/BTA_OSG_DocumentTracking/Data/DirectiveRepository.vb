Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports System.Data
Imports Microsoft.Data.SqlClient

Namespace BTA_OSG
    Public Class DirectiveRepository
        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        Public Function GetByDocumentId(docId As Integer) As List(Of ActionDirective)
            Dim list As New List(Of ActionDirective)()
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT * FROM tbl_ActionDirectives WHERE DocumentID = @id"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@id", docId)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            list.Add(New ActionDirective With {
                                .DirectiveID = Convert.ToInt32(reader("DirectiveID")),
                                .DocumentID = Convert.ToInt32(reader("DocumentID")),
                                .DirectiveTypeID = Convert.ToInt32(reader("DirectiveTypeID")),
                                .DirectiveText = If(IsDBNull(reader("DirectiveText")), Nothing, Convert.ToString(reader("DirectiveText"))),
                                .IssuedByUserID = Convert.ToInt32(reader("IssuedByUserID")),
                                .IssuedAtUTC = Convert.ToDateTime(reader("IssuedAtUTC")),
                                .IsActive = Convert.ToBoolean(reader("IsActive")),
                                .SupersededByDirectiveID = If(IsDBNull(reader("SupersededByDirectiveID")), CType(Nothing, Integer?), Convert.ToInt32(reader("SupersededByDirectiveID"))),
                                .Remarks = If(IsDBNull(reader("Remarks")), Nothing, Convert.ToString(reader("Remarks")))
                            })
                        End While
                    End Using
                End Using
            End Using
            Return list
        End Function

        Public Function Insert(directive As ActionDirective, Optional transaction As SqlTransaction = Nothing) As Integer
            Dim sql = "INSERT INTO tbl_ActionDirectives (DocumentID, DirectiveTypeID, DirectiveText, IssuedByUserID, IssuedAtUTC, IsActive, Remarks) " &
                      "OUTPUT INSERTED.DirectiveID " &
                      "VALUES (@DocumentID, @DirectiveTypeID, @DirectiveText, @IssuedByUserID, @IssuedAtUTC, @IsActive, @Remarks)"
            If transaction IsNot Nothing Then
                Using cmd = New SqlCommand(sql, transaction.Connection, transaction)
                    FillInsertParameters(cmd, directive)
                    Return Convert.ToInt32(cmd.ExecuteScalar())
                End Using
            Else
                Using conn = _connectionFactory.CreateConnection()
                    Using cmd = New SqlCommand(sql, conn)
                        FillInsertParameters(cmd, directive)
                        Return Convert.ToInt32(cmd.ExecuteScalar())
                    End Using
                End Using
            End If
        End Function

        Private Shared Sub FillInsertParameters(cmd As SqlCommand, directive As ActionDirective)
            cmd.Parameters.AddWithValue("@DocumentID", directive.DocumentID)
            cmd.Parameters.AddWithValue("@DirectiveTypeID", directive.DirectiveTypeID)
            cmd.Parameters.AddWithValue("@DirectiveText", If(directive.DirectiveText IsNot Nothing, CType(directive.DirectiveText, Object), DBNull.Value))
            cmd.Parameters.AddWithValue("@IssuedByUserID", directive.IssuedByUserID)
            cmd.Parameters.AddWithValue("@IssuedAtUTC", directive.IssuedAtUTC)
            cmd.Parameters.AddWithValue("@IsActive", directive.IsActive)
            cmd.Parameters.AddWithValue("@Remarks", If(directive.Remarks IsNot Nothing, CType(directive.Remarks, Object), DBNull.Value))
        End Sub

        Public Sub Supersede(directiveId As Integer, newDirectiveId As Integer)
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "UPDATE tbl_ActionDirectives SET IsActive = 0, SupersededByDirectiveID = @newId WHERE DirectiveID = @id"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@newId", newDirectiveId)
                    cmd.Parameters.AddWithValue("@id", directiveId)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub
    End Class
End Namespace
