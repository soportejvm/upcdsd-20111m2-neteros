' NOTA: puede usar el comando "Cambiar nombre" del menú contextual para cambiar el nombre de clase "Service1" en el código, en svc y en el archivo de configuración.
Imports System.IO
Imports System.Xml.Linq
Imports System.Linq
Imports System.Xml
<Serializable()>
Public Class ServiceGroupon
    Implements IServiceGroupon
    Private Function DevuelveXDocument(ByVal sDatos As String) As XDocument
        Dim S As TextReader = New StringReader(sDatos)
        Dim doc As XDocument = XDocument.Load(S)
        Return doc
    End Function

    Public Function RegistraCliente(ByVal oCliente As String) As Integer Implements IServiceGroupon.RegistraCliente
        Dim C As ComControladora.Usuarios

        Dim oUsuario As ComIdentidades.Usuarios

        C = New ComControladora.Usuarios
        oUsuario = New ComIdentidades.Usuarios

        Dim Cliente = From element In DevuelveXDocument(oCliente).Descendants("DATA") Select cCodUsu = element.Descendants("cCodUsu").Value, cPass = element.Descendants("cPass").Value

        For Each item In Cliente
            oUsuario.Usuario = item.cCodUsu
            oUsuario.Password = item.cPass
        Next



        RegistraCliente = C.RegistrarUsuario(oUsuario)

        Return RegistraCliente
    End Function

    Public Function ValidarUsuario(ByVal pcMail As String) As String Implements IServiceGroupon.ValidarUsuario
        Dim C As ComControladora.Usuarios
        Dim D As New ComTipos.clsDataset

        C = New ComControladora.Usuarios

        D = New ComTipos.clsDataSet
        D = C.ValidarUsuario(pcMail)


        Return D.DevuelveXMLLinQ
    End Function

End Class
