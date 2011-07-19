Imports System.Data
Imports System.Data.SqlClient

Public Class UsuarioImpl

    Public Function RegistrarUsuario(ByVal oUsuario As ComIdentidades.Usuarios) As Integer
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("RegistrarUsuarios", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cMail", SqlDbType.VarChar, 50).Value = oUsuario.Email
        oComm.Parameters.Add("@cNombres", SqlDbType.VarChar, 50).Value = oUsuario.Nombres
        oComm.Parameters.Add("@cApellidos", SqlDbType.VarChar, 50).Value = oUsuario.Apellidos
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 20).Value = oUsuario.Usuario
        oComm.Parameters.Add("@cPass", SqlDbType.VarChar, 8).Value = oUsuario.Password
        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output


        oComm.ExecuteNonQuery()

        RegistrarUsuario = oComm.Parameters("@nResp").Value

        oComm = Nothing
        objCon = Nothing
    End Function

    Public Function ValidarUsuario(ByVal cEmail As String) As ComTipos.clsDataSet
        Dim D As ComTipos.clsDataSet
        D = New ComTipos.clsDataSet
        D.CargaTabla(, , "ValidarUsuarios '" & cEmail & "'")
        ValidarUsuario = D
        D = Nothing
    End Function
End Class
