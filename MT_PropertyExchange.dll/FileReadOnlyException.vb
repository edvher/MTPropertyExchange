<ComVisible(False)> _
Public Class FileReadOnlyException
    Inherits Exception

    Public Sub New()
    End Sub 'New

    Public Sub New(ByVal message As String)
        MyBase.New(message)
    End Sub 'New

    Public Sub New(ByVal message As String, ByVal inner As Exception)
        MyBase.New(message, inner)
    End Sub 'New
End Class

