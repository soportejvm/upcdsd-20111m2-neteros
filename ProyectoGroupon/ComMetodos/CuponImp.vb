Public Class CuponImp
    Public Function RegistrarCupon(ByVal oCupon As ComIdentidades.Cupon) As Integer
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("RegistrarCupon", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cTitulo", SqlDbType.VarChar, 20).Value = oCupon.Usuario
        oComm.Parameters.Add("@cDescripcion", SqlDbType.VarChar, 8).Value = oCupon.Password
        oComm.Parameters.Add("@nCosto", SqlDbType.VarChar, 8).Value = oCupon.Password
        oComm.Parameters.Add("@nDias", SqlDbType.VarChar, 8).Value = oCupon.Password
        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output


        oComm.ExecuteNonQuery()

        RegistrarCupon = oComm.Parameters("@nResp").Value

        oComm = Nothing
        objCon = Nothing
    End Function
End Class
