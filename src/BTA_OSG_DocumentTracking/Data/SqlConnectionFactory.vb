Imports Microsoft.Data.SqlClient

Namespace BTA_OSG
    Public Class SqlConnectionFactory
        Implements IDbConnectionFactory

        Private ReadOnly _connectionString As String

        Public Sub New(connectionString As String)
            _connectionString = connectionString
        End Sub

        Public Function CreateConnection() As SqlConnection Implements IDbConnectionFactory.CreateConnection
            Dim conn As New SqlConnection(_connectionString)
            conn.Open()
            Return conn
        End Function
    End Class
End Namespace
