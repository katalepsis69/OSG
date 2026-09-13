Imports System.Data
Imports Microsoft.Data.SqlClient

Namespace BTA_OSG
    Public Class SequenceRepository
        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        Public Function GetNextDocCode(typeCode As String, prefix As String, year As Short) As String
            Dim lastNumber As Integer = 0
            
            Using conn = _connectionFactory.CreateConnection()
                Using trans = conn.BeginTransaction()
                    Try
                        ' b) SELECT LastNumber FROM tbl_DocumentSequences WITH (UPDLOCK, HOLDLOCK) WHERE DocumentTypeCode=@tc AND SequenceYear=@yr
                        Dim selectSql = "SELECT LastNumber FROM tbl_DocumentSequences WITH (UPDLOCK, HOLDLOCK) WHERE DocumentTypeCode=@tc AND SequenceYear=@yr"
                        Using cmdSelect = New SqlCommand(selectSql, conn, trans)
                            cmdSelect.Parameters.AddWithValue("@tc", typeCode)
                            cmdSelect.Parameters.AddWithValue("@yr", year)
                            Dim result = cmdSelect.ExecuteScalar()
                            
                            If result Is Nothing OrElse IsDBNull(result) Then
                                ' c) If not found: INSERT with LastNumber=1
                                lastNumber = 1
                                Dim insertSql = "INSERT INTO tbl_DocumentSequences (DocumentTypeCode, SequenceYear, LastNumber) VALUES (@tc, @yr, @last)"
                                Using cmdInsert = New SqlCommand(insertSql, conn, trans)
                                    cmdInsert.Parameters.AddWithValue("@tc", typeCode)
                                    cmdInsert.Parameters.AddWithValue("@yr", year)
                                    cmdInsert.Parameters.AddWithValue("@last", lastNumber)
                                    cmdInsert.ExecuteNonQuery()
                                End Using
                            Else
                                ' d) Else: UPDATE SET LastNumber=LastNumber+1
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
                        
                        ' f) COMMIT
                        trans.Commit()
                    Catch
                        trans.Rollback()
                        Throw
                    End Try
                End Using
            End Using
            
            ' e) Build code: PREFIX-yyyy-### (3 digit padded, expand if >999)
            Dim paddedNumber As String = lastNumber.ToString("D3")
            Return $"{prefix}-{year}-{paddedNumber}"
        End Function
    End Class
End Namespace
