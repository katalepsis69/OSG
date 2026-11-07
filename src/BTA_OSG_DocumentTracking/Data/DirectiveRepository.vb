Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports System.Data
Imports Microsoft.Data.SqlClient

Namespace BTA_OSG
    Public Class DirectiveRepository
        Private ReadOnly _connectionFactory As SqlConnectionFactory

        Public Sub New(connectionFactory As SqlConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

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
    End Class
End Namespace
