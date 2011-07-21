Public Class Cupon : Inherits MarshalByRefObject

    Public Sub RegistrarCupon(ByVal oCupon As ComIdentidades.Cupon, ByVal oFoto() As Byte)
        Dim D As ComMetodos.CuponImpl
        D = New ComMetodos.CuponImpl
        D.RegistrarCupon(oCupon, oFoto)
        D = Nothing
    End Sub

    Public Function DevuelveOfertas() As ComTipos.clsDataSet
        Dim C As ComMetodos.CuponImpl
        C = New ComMetodos.CuponImpl
        DevuelveOfertas = C.DevuelveOfertas()
        C = Nothing
    End Function

End Class
