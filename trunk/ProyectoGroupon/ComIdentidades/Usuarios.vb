Public Class Usuarios
    Private pcEmail As String
    Private pcNombres As String
    Private pcApellidos As String
    Private pcCodUsu As String
    Private pcPass As String

    Public Property Email As String
        Get
            Email = pcEmail
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property

    Public Property Nombres As String
        Get
            Nombres = pcNombres
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property

    Public Property Apellidos As String
        Get
            Apellidos = pcApellidos
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property

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
