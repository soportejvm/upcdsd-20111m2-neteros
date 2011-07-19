Public Class Usuarios
    Private pcCodUsu As String
    Private pcPass As String

    Public Property Usuario As Integer
        Get
            Usuario = pcCodUsu
        End Get
        Set(ByVal value As Integer)
            Return
        End Set
    End Property

    Public Property Password As String
        Get
            Password = pcPass
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property
End Class
