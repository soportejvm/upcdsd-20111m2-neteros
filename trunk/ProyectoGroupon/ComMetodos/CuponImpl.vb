Imports System.Data
Imports System.Data.SqlClient

Public Class CuponImpl

    Public Sub RegistrarCupon(ByVal oCupon As ComIdentidades.Cupon, ByVal oFoto() As Byte)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("RegistraCupon", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cTitulo", SqlDbType.VarChar, 250).Value = oCupon.Titulo
        oComm.Parameters.Add("@cDescpricion", SqlDbType.VarChar, 250).Value = oCupon.Descripcion
        oComm.Parameters.Add("@nCosto", SqlDbType.Money).Value = oCupon.Costo
        oComm.Parameters.Add("@nDias", SqlDbType.Int).Value = oCupon.Dias
        oComm.Parameters.Add("@nDesc", SqlDbType.Money).Value = oCupon.Descuento
        oComm.Parameters.Add("@oFotoOferta", SqlDbType.Image).Value = IIf(oFoto Is Nothing, DBNull.Value, oFoto)

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing

    End Sub

    Public Function DevuelveOfertas() As ComTipos.clsDataSet

        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataSet
        D.CargaTabla(, , "DevueleOfertas ")
        DevuelveOfertas = D
        D = Nothing
    End Function


    
End Class
