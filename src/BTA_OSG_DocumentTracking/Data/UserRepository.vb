Option Explicit On
Option Strict On

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
                Dim sql = "SELECT u.UserID, u.Username, u.FullName, u.Office, u.Email, u.IsActive, u.CanRoute, u.CanMove, u.CanSoftCopy, u.IsLocked, u.FailedTapCount, u.LastFailedTapUTC, u.CreatedByUserID, u.CreatedAtUTC, u.ModifiedByUserID, u.ModifiedAtUTC, u.RowVersion FROM tbl_Users u JOIN tbl_RfidCards c ON u.UserID = c.UserID WHERE c.CardPublicID = @id AND c.IsActive = 1 AND u.IsActive = 1"
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
                Dim sql = "SELECT UserID, Username, FullName, Office, Email, IsActive, CanRoute, CanMove, CanSoftCopy, IsLocked, FailedTapCount, LastFailedTapUTC, CreatedByUserID, CreatedAtUTC, ModifiedByUserID, ModifiedAtUTC, RowVersion FROM tbl_Users WHERE UserID = @id"
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
                Dim sql = "SELECT UserID, Username, FullName, Office, Email, IsActive, CanRoute, CanMove, CanSoftCopy, IsLocked, FailedTapCount FROM tbl_Users"
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

        Public Function Insert(user As User, createdBy As Integer?, Optional transaction As Microsoft.Data.SqlClient.SqlTransaction = Nothing) As Integer
            Dim conn As Microsoft.Data.SqlClient.SqlConnection = If(transaction IsNot Nothing, transaction.Connection, _connectionFactory.CreateConnection())
            Try
                Dim sql = "INSERT INTO tbl_Users (Username, FullName, Email, Office, IsActive, CanRoute, CanMove, CanSoftCopy, CreatedByUserID, CreatedAtUTC) " &
                           "OUTPUT INSERTED.UserID " &
                           "VALUES (@Username, @FullName, @Email, @Office, @IsActive, @CanRoute, @CanMove, @CanSoftCopy, @CreatedByUserID, SYSUTCDATETIME())"
                Using cmd = New SqlCommand(sql, conn, transaction)
                    cmd.Parameters.AddWithValue("@Username", user.Username)
                    cmd.Parameters.AddWithValue("@FullName", If(user.FullName IsNot Nothing, CType(user.FullName, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Email", If(user.Email IsNot Nothing, CType(user.Email, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Office", If(user.Office IsNot Nothing, CType(user.Office, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@IsActive", user.IsActive)
                    cmd.Parameters.AddWithValue("@CanRoute", user.CanRoute)
                    cmd.Parameters.AddWithValue("@CanMove", user.CanMove)
                    cmd.Parameters.AddWithValue("@CanSoftCopy", user.CanSoftCopy)
                    cmd.Parameters.AddWithValue("@CreatedByUserID", If(createdBy.HasValue, CType(createdBy.Value, Object), DBNull.Value))
                    Return Convert.ToInt32(cmd.ExecuteScalar())
                End Using
            Finally
                If transaction Is Nothing Then conn.Dispose()
            End Try
        End Function

        Public Sub Update(user As User, modifiedBy As Integer?, Optional transaction As Microsoft.Data.SqlClient.SqlTransaction = Nothing)
            Dim conn As Microsoft.Data.SqlClient.SqlConnection = If(transaction IsNot Nothing, transaction.Connection, _connectionFactory.CreateConnection())
            Try
                Dim sql = "UPDATE tbl_Users SET Username = @Username, FullName = @FullName, Email = @Email, " &
                           "Office = @Office, IsActive = @IsActive, CanRoute = @CanRoute, CanMove = @CanMove, " &
                           "CanSoftCopy = @CanSoftCopy, ModifiedByUserID = @ModifiedByUserID, ModifiedAtUTC = SYSUTCDATETIME() " &
                           "WHERE UserID = @id"
                Using cmd = New SqlCommand(sql, conn, transaction)
                    cmd.Parameters.AddWithValue("@Username", user.Username)
                    cmd.Parameters.AddWithValue("@FullName", If(user.FullName IsNot Nothing, CType(user.FullName, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Email", If(user.Email IsNot Nothing, CType(user.Email, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Office", If(user.Office IsNot Nothing, CType(user.Office, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@IsActive", user.IsActive)
                    cmd.Parameters.AddWithValue("@CanRoute", user.CanRoute)
                    cmd.Parameters.AddWithValue("@CanMove", user.CanMove)
                    cmd.Parameters.AddWithValue("@CanSoftCopy", user.CanSoftCopy)
                    cmd.Parameters.AddWithValue("@ModifiedByUserID", If(modifiedBy.HasValue, CType(modifiedBy.Value, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@id", user.UserID)
                    cmd.ExecuteNonQuery()
                End Using
            Finally
                If transaction Is Nothing Then conn.Dispose()
            End Try
        End Sub

        ''' <summary>
        ''' Links a user to a role by code (SG, SYSADMIN, OSG_CHIEF, RECORDS, SECRETARIAT,
        ''' LEGISLATIVE, FINANCE, TRAVEL, ADMIN_STAFF, seeded by db/scripts/007 and 008).
        ''' Idempotent, so re-enrolling a badge does not stack roles. Returns False when the
        ''' role code is unknown.
        ''' </summary>
        Public Function AssignRole(userId As Integer, roleCode As String, assignedBy As Integer?, Optional transaction As Microsoft.Data.SqlClient.SqlTransaction = Nothing) As Boolean
            Dim conn As Microsoft.Data.SqlClient.SqlConnection = If(transaction IsNot Nothing, transaction.Connection, _connectionFactory.CreateConnection())
            Try
                Dim sql = "INSERT INTO tbl_UserRoles (UserID, RoleID, AssignedByUserID, AssignedAtUTC, IsActive) " &
                          "SELECT @uid, RoleID, @by, SYSUTCDATETIME(), 1 FROM tbl_Roles WHERE RoleCode = @code " &
                          "AND NOT EXISTS (SELECT 1 FROM tbl_UserRoles ur JOIN tbl_Roles r ON r.RoleID = ur.RoleID " &
                          "WHERE ur.UserID = @uid AND r.RoleCode = @code)"
                Using cmd = New SqlCommand(sql, conn, transaction)
                    cmd.Parameters.AddWithValue("@uid", userId)
                    cmd.Parameters.AddWithValue("@by", If(assignedBy.HasValue, CType(assignedBy.Value, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@code", roleCode)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            Finally
                If transaction Is Nothing Then conn.Dispose()
            End Try
        End Function

        ''' <summary>
        ''' True when at least one active user carries the role. The first-run claim gate needs
        ''' exactly this question, and counting users instead would answer wrong on a server
        ''' whose staff were enrolled without roles.
        ''' </summary>
        Public Function HasAnyUserWithRole(roleCode As String) As Boolean
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT COUNT(1) FROM tbl_Users u " &
                          "JOIN tbl_UserRoles ur ON ur.UserID = u.UserID " &
                          "JOIN tbl_Roles r ON r.RoleID = ur.RoleID " &
                          "WHERE r.RoleCode = @code AND ur.IsActive = 1 AND u.IsActive = 1"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@code", roleCode)
                    Return Convert.ToInt32(cmd.ExecuteScalar()) > 0
                End Using
            End Using
        End Function

        Public Function GetUserRoles(userId As Integer) As List(Of Role)
            Dim list As New List(Of Role)()
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "SELECT r.RoleID, r.RoleCode, r.RoleName, r.Description, r.IsActive FROM tbl_Roles r JOIN tbl_UserRoles ur ON r.RoleID = ur.RoleID " &
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
                Dim sql = "UPDATE tbl_Users SET FailedTapCount = ISNULL(FailedTapCount, 0) + 1, LastFailedTapUTC = SYSUTCDATETIME() WHERE UserID = @id"
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
                Dim sql = "UPDATE tbl_Users SET FailedTapCount = 0, IsLocked = 0 WHERE UserID = @id"
                Using cmd = New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@id", userId)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Sub Disable(userId As Integer)
            Using conn = _connectionFactory.CreateConnection()
                Dim sql = "UPDATE tbl_Users SET IsActive = 0 WHERE UserID = @id"
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
                .FullName = If(IsDBNull(reader("FullName")), Nothing, Convert.ToString(reader("FullName"))),
                .Email = If(IsDBNull(reader("Email")), Nothing, Convert.ToString(reader("Email"))),
                .Office = If(IsDBNull(reader("Office")), Nothing, Convert.ToString(reader("Office"))),
                .CanRoute = If(IsDBNull(reader("CanRoute")), True, Convert.ToBoolean(reader("CanRoute"))),
                .CanMove = If(IsDBNull(reader("CanMove")), True, Convert.ToBoolean(reader("CanMove"))),
                .CanSoftCopy = If(IsDBNull(reader("CanSoftCopy")), True, Convert.ToBoolean(reader("CanSoftCopy"))),
                .FailedTapCount = If(IsDBNull(reader("FailedTapCount")), 0, Convert.ToInt32(reader("FailedTapCount"))),
                .IsLocked = If(IsDBNull(reader("IsLocked")), False, Convert.ToBoolean(reader("IsLocked"))),
                .IsActive = If(IsDBNull(reader("IsActive")), True, Convert.ToBoolean(reader("IsActive")))
            }
        End Function
    End Class
End Namespace
