Option Explicit On
Option Strict On

Imports Microsoft.Data.SqlClient

Namespace BTA_OSG
    Public Interface IDbConnectionFactory
        Function CreateConnection() As SqlConnection
    End Interface
End Namespace
