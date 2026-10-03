Option Explicit On
Option Strict On

Imports System.Data
Imports Microsoft.Data.SqlClient

Namespace BTA_OSG
    Public Class SequenceRepository
        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        ''' <summary>
        ''' Mints the next DocCode under UPDLOCK so two workstations cannot take the same
        ''' number. When the caller supplies a transaction the reservation joins it and the
        ''' lock holds until the caller commits; without one, the reserve is its own
        ''' committed unit exactly as before.
        ''' </summary>
        Public Function GetNextDocCode(typeCode As String, prefix As String, year As Short, Optional transaction As SqlTransaction = Nothing) As String
            Dim lastNumber As Integer

            If transaction IsNot Nothing Then
                lastNumber = ReserveNextSequenceNumber(typeCode, year, transaction)
            Else
                Using conn = _connectionFactory.CreateConnection()
                    Using trans = conn.BeginTransaction()
                        Try
                            lastNumber = ReserveNextSequenceNumber(typeCode, year, trans)
                            trans.Commit()
                        Catch
                            trans.Rollback()
                            Throw
                        End Try
                    End Using
                End Using
            End If

            ' Build code: PREFIX-yyyy-### (3 digit padded, expand if >999)
            Dim paddedNumber As String = lastNumber.ToString("D3")
            Return $"{prefix}-{year}-{paddedNumber}"
        End Function

        ''' <summary>
        ''' Moves the shared counter past a code that arrived from outside it: an offline seat
        ''' mints its own number, and unless the counter learns about it the counter later
        ''' re-mints that same code and every registration after it fails the unique key.
        ''' Codes that do not read as PREFIX-yyyy-number are ignored.
        ''' </summary>
        Public Sub RaiseSequenceTo(docCode As String, typeCode As String, Optional transaction As SqlTransaction = Nothing)
            Dim parts = docCode.Split("-"c)
            If parts.Length < 3 Then Return
            Dim year As Short
            Dim number As Integer
            If Not Short.TryParse(parts(1), year) OrElse Not Integer.TryParse(parts(2), number) Then Return

            Dim conn As SqlConnection = If(transaction IsNot Nothing, transaction.Connection, _connectionFactory.CreateConnection())
            Try
                Using cmd = New SqlCommand(
                    "UPDATE tbl_DocumentSequences WITH (UPDLOCK, HOLDLOCK) " &
                    "SET LastNumber = CASE WHEN LastNumber < @n THEN @n ELSE LastNumber END " &
                    "WHERE DocumentTypeCode = @tc AND SequenceYear = @yr; " &
                    "IF @@ROWCOUNT = 0 INSERT INTO tbl_DocumentSequences (DocumentTypeCode, SequenceYear, LastNumber) VALUES (@tc, @yr, @n)", conn, transaction)
                    cmd.Parameters.AddWithValue("@n", number)
                    cmd.Parameters.AddWithValue("@tc", typeCode)
                    cmd.Parameters.AddWithValue("@yr", year)
                    cmd.ExecuteNonQuery()
                End Using
            Finally
                If transaction Is Nothing Then conn.Dispose()
            End Try
        End Sub

        ' No commit here: the transaction owner decides whether the reservation sticks.
        Private Shared Function ReserveNextSequenceNumber(typeCode As String, year As Short, trans As SqlTransaction) As Integer
            Dim conn = trans.Connection
            Dim lastNumber As Integer = 0

            ' SELECT LastNumber FROM tbl_DocumentSequences WITH (UPDLOCK, HOLDLOCK) WHERE DocumentTypeCode=@tc AND SequenceYear=@yr
            Dim selectSql = "SELECT LastNumber FROM tbl_DocumentSequences WITH (UPDLOCK, HOLDLOCK) WHERE DocumentTypeCode=@tc AND SequenceYear=@yr"
            Using cmdSelect = New SqlCommand(selectSql, conn, trans)
                cmdSelect.Parameters.AddWithValue("@tc", typeCode)
                cmdSelect.Parameters.AddWithValue("@yr", year)
                Dim result = cmdSelect.ExecuteScalar()

                If result Is Nothing OrElse IsDBNull(result) Then
                    ' Not found: INSERT with LastNumber=1
                    lastNumber = 1
                    Dim insertSql = "INSERT INTO tbl_DocumentSequences (DocumentTypeCode, SequenceYear, LastNumber) VALUES (@tc, @yr, @last)"
                    Using cmdInsert = New SqlCommand(insertSql, conn, trans)
                        cmdInsert.Parameters.AddWithValue("@tc", typeCode)
                        cmdInsert.Parameters.AddWithValue("@yr", year)
                        cmdInsert.Parameters.AddWithValue("@last", lastNumber)
                        cmdInsert.ExecuteNonQuery()
                    End Using
                Else
                    ' Else: UPDATE SET LastNumber=LastNumber+1
                    lastNumber = Convert.ToInt32(result) + 1
                    Dim updateSql = "UPDATE tbl_DocumentSequences SET LastNumber = @last WHERE DocumentTypeCode=@tc AND SequenceYear=@yr"
                    Using cmdUpdate = New SqlCommand(updateSql, conn, trans)
                        cmdUpdate.Parameters.AddWithValue("@last", lastNumber)
                        cmdUpdate.Parameters.AddWithValue("@tc", typeCode)
                        cmdUpdate.Parameters.AddWithValue("@yr", year)
                        cmdUpdate.ExecuteNonQuery()
                    End Using
                End If
            End Using

            Return lastNumber
        End Function
    End Class
End Namespace
