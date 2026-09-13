Imports System.Collections.Generic
Imports System.Data
Imports Microsoft.Data.SqlClient

Namespace BTA_OSG
    Public Class UserRepository
        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        Public Function GetByCardPublicID(cardId As String) As User
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT u.* FROM tbl_Users u JOIN tbl_RfidCards c ON u.UserID = c.UserID WHERE c.CardPublicID = @id AND c.IsActive = 1 AND u.IsActive = 1"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@id", cardId.Trim().ToUpper())
                    Using reader = cmd.ExecuteReader()
                        If reader.Read() Then
                            Return MapUser(reader)
                        End If
                    End Using
                End Using
            End Using
            Return Nothing
        End Function

        Public Function GetById(userId As Integer) As User
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT * FROM tbl_Users WHERE UserID = @id"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@id", userId)
                    Using reader = cmd.ExecuteReader()
                        If reader.Read() Then
                            Return MapUser(reader)
                        End If
                    End Using
                End Using
            End Using
            Return Nothing
        End Function

        Public Function GetAll() As List(Of User)
            Dim list As New List(Of User)()
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT * FROM tbl_Users"
                Using cmd = New SqlCommand(sql, conn)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            list.Add(MapUser(reader))
                        End While
                    End Using
                End Using
            End Using
            Return list
        End Function

        Public Function Insert(user As User, createdBy As Integer?) As Integer
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "INSERT INTO tbl_Users (Username, DisplayName, Email, OfficeName, IsActive, CreatedBy, CreatedDate) " &
                          "OUTPUT INSERTED.UserID " &
                          "VALUES (@Username, @DisplayName, @Email, @OfficeName, @IsActive, @CreatedBy, GETDATE())"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@Username", user.Username)
                    cmd.Parameters.AddWithValue("@DisplayName", If(user.FullName, DBNull.Value))
                    cmd.Parameters.AddWithValue("@Email", If(user.Email, DBNull.Value))
                    cmd.Parameters.AddWithValue("@OfficeName", If(user.Office, DBNull.Value))
                    cmd.Parameters.AddWithValue("@IsActive", user.IsActive)
                    cmd.Parameters.AddWithValue("@CreatedBy", If(CType(createdBy, Object), DBNull.Value))
                    Return Convert.ToInt32(cmd.ExecuteScalar())
                End Using
            End Using
        End Function

        Public Sub Update(user As User, modifiedBy As Integer?)
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "UPDATE tbl_Users SET Username = @Username, DisplayName = @DisplayName, Email = @Email, " &
                          "OfficeName = @OfficeName, IsActive = @IsActive, ModifiedBy = @ModifiedBy, ModifiedDate = GETDATE() " &
                          "WHERE UserID = @id"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@Username", user.Username)
                    cmd.Parameters.AddWithValue("@DisplayName", If(user.FullName, DBNull.Value))
                    cmd.Parameters.AddWithValue("@Email", If(user.Email, DBNull.Value))
                    cmd.Parameters.AddWithValue("@OfficeName", If(user.Office, DBNull.Value))
                    cmd.Parameters.AddWithValue("@IsActive", user.IsActive)
                    cmd.Parameters.AddWithValue("@ModifiedBy", If(CType(modifiedBy, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@id", user.UserID)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Function GetUserRoles(userId As Integer) As List(Of Role)
            Dim list As New List(Of Role)()
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT r.* FROM tbl_Roles r JOIN tbl_UserRoles ur ON r.RoleID = ur.RoleID " &
                          "WHERE ur.UserID = @id AND ur.IsActive = 1 AND r.IsActive = 1"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@id", userId)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim role As New Role With {
                                .RoleID = Convert.ToInt32(reader("RoleID")),
                                .RoleName = Convert.ToString(reader("RoleName")),
                                .Description = If(IsDBNull(reader("Description")), Nothing, Convert.ToString(reader("Description"))),
                                .IsActive = Convert.ToBoolean(reader("IsActive"))
                            }
                            list.Add(role)
                        End While
                    End Using
                End Using
            End Using
            Return list
        End Function

        Public Function GetUserPermissions(userId As Integer) As HashSet(Of String)
            Dim permissions As New HashSet(Of String)()
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT p.PermissionCode FROM tbl_Permissions p " &
                          "JOIN tbl_RolePermissions rp ON p.PermissionID=rp.PermissionID " &
                          "JOIN tbl_UserRoles ur ON rp.RoleID=ur.RoleID " &
                          "WHERE ur.UserID=@id AND ur.IsActive=1 AND rp.IsActive=1 AND p.IsActive=1"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@id", userId)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            permissions.Add(Convert.ToString(reader("PermissionCode")))
                        End While
                    End Using
                End Using
            End Using
            Return permissions
        End Function

        Public Sub IncrementFailedTaps(userId As Integer)
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "UPDATE tbl_Users SET FailedTapAttempts = ISNULL(FailedTapAttempts, 0) + 1 WHERE UserID = @id"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@id", userId)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Sub LockUser(userId As Integer)
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "UPDATE tbl_Users SET IsLocked = 1 WHERE UserID = @id"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@id", userId)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Sub ResetFailedTaps(userId As Integer)
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "UPDATE tbl_Users SET FailedTapAttempts = 0, IsLocked = 0 WHERE UserID = @id"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@id", userId)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Private Function MapUser(reader As IDataReader) As User
            Return New User With {
                .UserID = Convert.ToInt32(reader("UserID")),
                .Username = Convert.ToString(reader("Username")),
                .FullName = If(IsDBNull(reader("DisplayName")), Nothing, Convert.ToString(reader("DisplayName"))),
                .Email = If(IsDBNull(reader("Email")), Nothing, Convert.ToString(reader("Email"))),
                .Office = If(IsDBNull(reader("OfficeName")), Nothing, Convert.ToString(reader("OfficeName"))),
                .FailedTapCount = If(IsDBNull(reader("FailedTapAttempts")), 0, Convert.ToInt32(reader("FailedTapAttempts"))),
                .IsLocked = If(IsDBNull(reader("IsLocked")), False, Convert.ToBoolean(reader("IsLocked"))),
                .IsActive = If(IsDBNull(reader("IsActive")), True, Convert.ToBoolean(reader("IsActive")))
            }
        End Function
    End Class
End Namespace
