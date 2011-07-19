Public Class NombrePagosImpl

    Public Function RegistrarPagoCupon(ByVal oPagos As ComIdentidades.Pagos) As Integer
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("RegistrarPagoCupon", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nIdCupon", SqlDbType.Int).Value = oPagos.IdCupon
        oComm.Parameters.Add("@dFecPago", SqlDbType.DateTime).Value = oPagos.FechaPago
        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = oPagos.Monto
        oComm.Parameters.Add("@nMedioPago", SqlDbType.Int).Value = oPagos.MedioPago
        oComm.Parameters.Add("@nMonto", SqlDbType.Int).Direction = ParameterDirection.Output


        oComm.ExecuteNonQuery()

        RegistrarPagoCupon = oComm.Parameters("@nResp").Value

        oComm = Nothing
        objCon = Nothing
    End Function


End Class
