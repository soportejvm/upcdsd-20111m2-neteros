Public Class Usuarios
    Private pcEmail As String
    Private pcNombres As String
    Private pcApellidos As String
    Private pcCodUsu As String
    Private pcPass As String

    Public Property Email As String
        Get
            Return pcEmail
        End Get
        Set(ByVal value As String)
            pcEmail = value
        End Set
    End Property

    Public Property Nombres As String
        Get
            Return pcNombres
        End Get
        Set(ByVal value As String)
            pcNombres = value
        End Set
    End Property

    Public Property Apellidos As String
        Get
            Return pcApellidos
        End Get
        Set(ByVal value As String)
            pcApellidos = value
        End Set
    End Property

    Public Property Usuario As String
        Get
            Return pcCodUsu
        End Get
        Set(ByVal value As String)
            pcCodUsu = value
        End Set
    End Property

    Public Property Password As String
        Get
            Return pcPass
        End Get
        Set(ByVal value As String)
            pcPass = value
        End Set
    End Property
End Class
