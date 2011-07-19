Public Class Cupon
    Private pcDescripcion As String
    Private nCosto As Double
    Private nDias As Integer


    Public Property Descripcion As String
        Get
            Descripcion = pcDescripcion
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property

    Public Property Costo As Double
        Get
            Costo = nCosto
        End Get
        Set(ByVal value As Double)
            Return
        End Set
    End Property

    Public Property Dias As Integer
        Get
            Dias = nDias
        End Get
        Set(ByVal value As Integer)
            Return
        End Set
    End Property
End Class
