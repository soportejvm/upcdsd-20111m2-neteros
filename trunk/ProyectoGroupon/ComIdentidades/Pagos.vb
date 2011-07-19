Public Class Pagos
    Private nIdCupon As Integer
    Private dFecPago As Date
    Private nMonto As Double
    Private nMedioPago As Integer

    Public Property Cupon As Integer
        Get
            Cupon = nIdCupon
        End Get
        Set(ByVal value As Integer)
            Return
        End Set
    End Property
    Public Property FechaPago As Date
        Get
            FechaPago = dFecPago
        End Get
        Set(ByVal value As Date)
            Return
        End Set
    End Property
    Public Property Monto As Double
        Get
            Monto = nMonto
        End Get
        Set(ByVal value As Double)
            Return
        End Set
    End Property
    Public Property MedioPago As Integer
        Get
            MedioPago = nMedioPago
        End Get
        Set(ByVal value As Integer)
            Return
        End Set
    End Property
End Class
