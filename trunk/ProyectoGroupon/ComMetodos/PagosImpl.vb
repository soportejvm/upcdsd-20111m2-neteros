Imports System.Data
Imports System.Data.SqlClient
Public Class PagosImpl

    Public Function RegistrarPagoCupon(ByVal oPagos As ComIdentidades.Pagos) As Integer
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("RegistrarPagoCupon", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nIdCupon", SqlDbType.Int).Value = oPagos.Cupon
        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = oPagos.Monto
        oComm.Parameters.Add("@nMedioPago", SqlDbType.Int).Value = oPagos.MedioPago
        oComm.Parameters.Add("@nNroTarjeta", SqlDbType.VarChar, 30).Value = oPagos.NroTarjeta
        oComm.Parameters.Add("@cCCV", SqlDbType.VarChar, 30).Value = oPagos.CCV
        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output


        oComm.ExecuteNonQuery()

        RegistrarPagoCupon = oComm.Parameters("@nResp").Value

        oComm = Nothing
        objCon = Nothing
    End Function

End Class
