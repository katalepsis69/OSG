Imports System.Collections.Generic
Imports System.Data
Imports Microsoft.Data.SqlClient

Namespace BTA_OSG
    Public Class ReferenceDataRepository
        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        Public Function GetDocumentTypes() As List(Of DocumentType)
            Dim list As New List(Of DocumentType)()
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT * FROM tbl_DocumentTypes WHERE IsActive = 1"
                Using cmd = New SqlCommand(sql, conn)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            list.Add(New DocumentType With {
                                .DocumentTypeID = Convert.ToInt32(reader("DocumentTypeID")),
                                .TypeCode = Convert.ToString(reader("TypeCode")),
                                .TypeName = Convert.ToString(reader("TypeName")),
                                .Prefix = If(IsDBNull(reader("Prefix")), Nothing, Convert.ToString(reader("Prefix"))),
                                .IsActive = Convert.ToBoolean(reader("IsActive")),
                                .SortOrder = If(IsDBNull(reader("SortOrder")), 0, Convert.ToInt32(reader("SortOrder")))
                            })
                        End While
                    End Using
                End Using
            End Using
            Return list
        End Function

        Public Function GetDocumentStatuses() As List(Of DocumentStatus)
            Dim list As New List(Of DocumentStatus)()
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT * FROM tbl_DocumentStatuses WHERE IsActive = 1"
                Using cmd = New SqlCommand(sql, conn)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            list.Add(New DocumentStatus With {
                                .StatusID = Convert.ToInt32(reader("StatusID")),
                                .StatusCode = Convert.ToString(reader("StatusCode")),
                                .StatusName = Convert.ToString(reader("StatusName")),
                                .IsActive = Convert.ToBoolean(reader("IsActive"))
                            })
                        End While
                    End Using
                End Using
            End Using
            Return list
        End Function

        Public Function GetDirectiveTypes() As List(Of DirectiveType)
            Dim list As New List(Of DirectiveType)()
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT * FROM tbl_DirectiveTypes WHERE IsActive = 1"
                Using cmd = New SqlCommand(sql, conn)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            list.Add(New DirectiveType With {
                                .DirectiveTypeID = Convert.ToInt32(reader("DirectiveTypeID")),
                                .DirectiveCode = Convert.ToString(reader("DirectiveCode")),
                                .DirectiveName = Convert.ToString(reader("DirectiveName")),
                                .ResultStatusID = If(IsDBNull(reader("ResultStatusID")), CType(Nothing, Integer?), Convert.ToInt32(reader("ResultStatusID"))),
                                .IsActive = Convert.ToBoolean(reader("IsActive"))
                            })
                        End While
                    End Using
                End Using
            End Using
            Return list
        End Function

        Public Function GetRoles() As List(Of Role)
            Dim list As New List(Of Role)()
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT * FROM tbl_Roles WHERE IsActive = 1"
                Using cmd = New SqlCommand(sql, conn)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            list.Add(New Role With {
                                .RoleID = Convert.ToInt32(reader("RoleID")),
                                .RoleName = Convert.ToString(reader("RoleName")),
                                .Description = If(IsDBNull(reader("Description")), Nothing, Convert.ToString(reader("Description"))),
                                .IsActive = Convert.ToBoolean(reader("IsActive"))
                            })
                        End While
                    End Using
                End Using
            End Using
            Return list
        End Function

        Public Function GetStatusByCode(code As String) As DocumentStatus
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT * FROM tbl_DocumentStatuses WHERE StatusCode = @code"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@code", code)
                    Using reader = cmd.ExecuteReader()
                        If reader.Read() Then
                            Return New DocumentStatus With {
                                .StatusID = Convert.ToInt32(reader("StatusID")),
                                .StatusCode = Convert.ToString(reader("StatusCode")),
                                .StatusName = Convert.ToString(reader("StatusName")),
                                .IsActive = Convert.ToBoolean(reader("IsActive"))
                            }
                        End If
                    End Using
                End Using
            End Using
            Return Nothing
        End Function

        Public Function GetDocumentTypeByCode(code As String) As DocumentType
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT * FROM tbl_DocumentTypes WHERE TypeCode = @code"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@code", code)
                    Using reader = cmd.ExecuteReader()
                        If reader.Read() Then
                            Return New DocumentType With {
                                .DocumentTypeID = Convert.ToInt32(reader("DocumentTypeID")),
                                .TypeCode = Convert.ToString(reader("TypeCode")),
                                .TypeName = Convert.ToString(reader("TypeName")),
                                .Prefix = If(IsDBNull(reader("Prefix")), Nothing, Convert.ToString(reader("Prefix"))),
                                .IsActive = Convert.ToBoolean(reader("IsActive")),
                                .SortOrder = If(IsDBNull(reader("SortOrder")), 0, Convert.ToInt32(reader("SortOrder")))
                            }
                        End If
                    End Using
                End Using
            End Using
            Return Nothing
        End Function
    End Class
End Namespace
