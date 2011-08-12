' NOTA: puede usar el comando "Cambiar nombre" del menú contextual para cambiar el nombre de clase "Service1" en el código, en svc y en el archivo de configuración.
Imports System.IO
Imports System.Xml.Linq
Imports System.Linq
Imports System.Xml
Imports System.Net.Mail
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


        Dim objUsuario As ComIdentidades.Usuarios

        C = New ComControladora.Usuarios



        '<GENESYS><DATA><pcEmail>carlospua@hotmail.com</pcEmail><pcNombres>Carlos Jose</pcNombres><pcApellidos>Pua Soplapuco</pcApellidos><pcCodUsu>carlospua@hotmail.com</pcCodUsu><pcPass>42206636</pcPass></DATA></GENESYS>

        Dim Cliente = From element In DevuelveXDocument(oCliente).Descendants("DATA") Select pcEmail = element.Descendants("pcEmail").Value, pcNombres = element.Descendants("pcNombres").Value, _
        pcApellidos = element.Descendants("pcApellidos").Value, pcCodUsu = element.Descendants("pcCodUsu").Value, pcPass = element.Descendants("pcPass").Value

        For Each item In Cliente
            objUsuario = New ComIdentidades.Usuarios
            objUsuario.Nombres = item.pcNombres
            objUsuario.Apellidos = item.pcApellidos
            objUsuario.Email = item.pcEmail
            objUsuario.Usuario = item.pcCodUsu
            objUsuario.Password = item.pcPass            
        Next

        RegistraCliente = C.RegistrarUsuario(objUsuario)

        Call EnviaCorreo(objUsuario.Email, "Registro Groupon", "Bienvenido a Groupon aqui encontrar muchas ofertas en distintos rubros, gracias por registrarse")

        objUsuario = Nothing
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

    Function RegistrarCupon(ByVal oCupon As String,
                         ByVal oFotoSumin() As Byte) As Integer Implements IServiceGroupon.RegistrarCupon

        Dim C As ComControladora.Cupon
        Dim objCupon As ComIdentidades.Cupon

        C = New ComControladora.Cupon

        Dim Cupones = From element In DevuelveXDocument(oCupon).Descendants("DATA") Select pcTitulo = element.Descendants("pcTitulo").Value, pcDescripcion = element.Descendants("pcDescripcion").Value, _
        nCosto = element.Descendants("nCosto").Value, nDias = element.Descendants("nDias").Value, nDesc = element.Descendants("nDesc").Value, cRestricciones = element.Descendants("cRestricciones").Value


        For Each item In Cupones
            objCupon = New ComIdentidades.Cupon
            objCupon.Titulo = item.pcTitulo
            objCupon.Descripcion = item.pcDescripcion
            objCupon.Dias = item.nDias
            objCupon.Costo = item.nCosto
            objCupon.Descuento = item.nDesc
            objCupon.Restricciones = item.cRestricciones
        Next


        C.RegistrarCupon(objCupon, oFotoSumin)


        C = Nothing
        Return 1
    End Function

    Public Function DevuelveOfertas() As String Implements IServiceGroupon.DevuelveOfertas
        Dim C As ComControladora.Cupon
        Dim D As ComTipos.clsDataSet

        D = New ComTipos.clsDataSet
        C = New ComControladora.Cupon


        D = C.DevuelveOfertas()

        Return D.DevuelveXMLLinQ

        C = Nothing
        D = Nothing
    End Function

    Public Function RegistrarPagoCupon(ByVal oPagos As String, ByVal pcMail As String) As Integer Implements IServiceGroupon.RegistrarPagoCupon
        Dim C As ComControladora.Pagos
        Dim D As ComIdentidades.Pagos


        Dim Pagos = From element In DevuelveXDocument(oPagos).Descendants("DATA") Select nIdCupon = element.Descendants("nIdCupon").Value, nMonto = element.Descendants("nMonto").Value, _
        nMedioPago = element.Descendants("nMedioPago").Value, nNrotarjeta = element.Descendants("nNrotarjeta").Value, cCCV = element.Descendants("cCCV").Value

        For Each item In Pagos
            D = New ComIdentidades.Pagos
            D.Cupon = item.nIdCupon
            D.MedioPago = item.nMedioPago
            D.Monto = item.nMonto
            D.NroTarjeta = item.nNrotarjeta
            D.CCV = item.cCCV
        Next

        C = New ComControladora.Pagos
        C.RegistrarPagoCupon(D)

        Call EnviaCorreo(pcMail, "Compra en Groupon", "Gracias por comprar nuestras ofertas")

        C = Nothing
        D = Nothing

        Return 1
    End Function

    Public Sub EnviaCorreo(ByVal pcCorreo As String, ByVal pcTitulo As String, ByVal pcCuerpo As String)
        Dim correo As New System.Net.Mail.MailMessage()
        correo.From = New System.Net.Mail.MailAddress("estopa1712@hotmail.com")
        correo.To.Add(pcCorreo)

        Dim s As String = pcTitulo

        correo.Subject = s

        correo.Body = pcCuerpo
        correo.IsBodyHtml = False
        correo.Priority = System.Net.Mail.MailPriority.Normal

        Dim smtp As New System.Net.Mail.SmtpClient("smtp.live.com", 25)
        smtp.Credentials = New System.Net.NetworkCredential("estopa1712@hotmail.com", "42206636")
        smtp.EnableSsl = True

        Try
            smtp.Send(correo)
            ' Label1.Text = "Mensaje enviado satisfactoriamente"
        Catch ex As Exception
            'Label1.Text = "ERROR: " & ex.Message
        End Try
    End Sub
End Class
