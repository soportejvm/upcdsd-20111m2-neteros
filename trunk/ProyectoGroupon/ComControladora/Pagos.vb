Public Class Pagos : Inherits MarshalByRefObject

    Public Function RegistrarPagoCupon(ByVal oPago As ComIdentidades.Pagos) As Integer
        Dim D As ComMetodos.PagosImpl
        D = New ComMetodos.PagosImpl
        RegistrarPagoCupon = D.RegistrarPagoCupon(oPago)
        D = Nothing
    End Function

End Class
