Option Explicit On
Option Strict On

Namespace BTA_OSG
    Public Class User
        Public Property UserID As Integer
        Public Property Username As String
        Public Property FullName As String
        Public Property Office As String
        Public Property Email As String
        Public Property IsActive As Boolean
        Public Property IsLocked As Boolean
        Public Property FailedTapCount As Integer
        Public Property LastFailedTapUTC As DateTime?
        Public Property CreatedByUserID As Integer?
        Public Property CreatedAtUTC As DateTime
        Public Property ModifiedByUserID As Integer?
        Public Property ModifiedAtUTC As DateTime?
        ' Desktop operational privileges. The Admin screen owns these three checkboxes and
        ' the Users mirror carries them back, so an unchecked box is not silently reset.
        Public Property CanRoute As Boolean = True
        Public Property CanMove As Boolean = True
        Public Property CanSoftCopy As Boolean = True
        Public Property RowVersion As Byte()
    End Class
End Namespace
