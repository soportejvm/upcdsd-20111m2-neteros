Public Class Usuarios : Inherits MarshalByRefObject

    Public Function RegistrarUsuario(ByVal oUsuario As ComIdentidades.Usuarios) As Integer
        Dim D As ComMetodos.UsuarioImpl
        D = New ComMetodos.UsuarioImpl
        RegistrarUsuario = D.RegistrarUsuario(oUsuario)
        D = Nothing
    End Function
End Class
