
Imports System.Data
Imports System.Data.SqlClient

Public Class clsConecta

    Private oConex As SqlConnection
    Private oTransac As SqlTransaction

    Public Function ConexionActual() As SqlConnection
        Return oConex
    End Function

    Public Function AbrirTransaccion() As SqlTransaction
        AbrirTransaccion = oConex.BeginTransaction()
    End Function

    Public Sub AbrirConexion()
        Dim C As SqlCommand

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing
        Exit Sub
    End Sub

    Public Function DameConexion() As SqlConnection
        Dim C As SqlCommand

        'Produccion 
        'oConex = New SqlConnection("Initial Catalog=BDVENT;Data Source=.;User=sa;password=1234")
        oConex = New SqlConnection("Data Source=KEPLER-PC;Initial Catalog=bdgroupon;Integrated Security=True")

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        Return oConex
    End Function

    Public Sub CierraConexion()
        oConex.Close()
        oConex = Nothing

    End Sub

    Sub EjecutaTransaccion(ByVal oCommand As SqlCommand)
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            RevierteTransaccion()
        End Try
    End Sub

    Sub IniciaTransaccion()
        oTransac = oConex.BeginTransaction("SampleTransaction")
    End Sub

    Sub FinalizaTransaccion()
        oTransac.Commit()
        oTransac = Nothing
        oConex.Close()
    End Sub

    Sub RevierteTransaccion()
        oTransac.Rollback()
        oTransac = Nothing
        oConex.Close()
    End Sub

    Public Function TransaccionActual() As SqlTransaction
        Return oTransac
    End Function

End Class

