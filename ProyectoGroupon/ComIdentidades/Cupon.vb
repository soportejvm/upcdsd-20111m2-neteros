Public Class Cupon
    Private pcTitulo As String
    Private pcDescripcion As String
    Private nCosto As Double
    Private nDias As Integer

    Private nDesc As Integer
    Private oFoto As Byte()
    Private pcRestricciones As String

    Public Property Titulo As String
        Get
            Return pcTitulo
        End Get
        Set(ByVal value As String)
            pcTitulo = value
        End Set
    End Property

    Public Property Descripcion As String
        Get
            Return pcDescripcion
        End Get
        Set(ByVal value As String)
            pcDescripcion = value
        End Set
    End Property

    Public Property Costo As Double
        Get
            Return nCosto
        End Get
        Set(ByVal value As Double)
            nCosto = value
        End Set
    End Property

    Public Property Dias As Integer
        Get
            Return nDias
        End Get
        Set(ByVal value As Integer)
            nDias = value
        End Set
    End Property


    Public Property Descuento As Double
        Get
            Return nDesc
        End Get
        Set(ByVal value As Double)
            nDesc = value
        End Set
    End Property

    Public Property Foto As Byte()
        Get
            Return oFoto
        End Get
        Set(ByVal value As Byte())
            oFoto = value
        End Set
    End Property

    Public Property Restricciones As String
        Get
            Return pcRestricciones
        End Get
        Set(ByVal value As String)
            pcRestricciones = value
        End Set
    End Property

End Class
