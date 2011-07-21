Public Class Pagos
    Private nIdCupon As Integer
    Private dFecPago As Date
    Private nMonto As Double
    Private nMedioPago As Integer
    Private nNrotarjeta As String
    Private cCCV As String

    Public Property Cupon As Integer
        Get
            Return nIdCupon
        End Get
        Set(ByVal value As Integer)
            nIdCupon = value
        End Set
    End Property
    Public Property FechaPago As Date
        Get
            Return dFecPago
        End Get
        Set(ByVal value As Date)
            dFecPago = value
        End Set
    End Property
    Public Property Monto As Double
        Get
            Return nMonto
        End Get
        Set(ByVal value As Double)
            nMonto = value
        End Set
    End Property
    Public Property MedioPago As Integer
        Get
            Return nMedioPago
        End Get
        Set(ByVal value As Integer)
            nMedioPago = value
        End Set
    End Property

    Public Property NroTarjeta As String
        Get
            Return nNrotarjeta
        End Get
        Set(ByVal value As String)
            nNrotarjeta = value
        End Set
    End Property

    Public Property CCV As String
        Get
            Return cCCV
        End Get
        Set(ByVal value As String)
            cCCV = value
        End Set
    End Property
End Class
