Public Class Pagos
    Public nIdCupon As Integer
    Public nMonto As Double
    Public nMedioPago As Integer
    Public nNrotarjeta As String
    Public cCCV As String

    Public Function DevuelveXMLPago() As String
        Dim oXml As String = ""

        oXml = "<DATA>" + "<nIdCupon>" + CStr(nIdCupon) + "</nIdCupon>" + _
                        "<nMonto>" + CStr(nMonto) + "</nMonto>" + _
                        "<nMedioPago>" + CStr(nMedioPago) + "</nMedioPago>" + _
                        "<nNrotarjeta>" + CStr(nNrotarjeta) + "</nNrotarjeta>" + _
                        "<cCCV>" + CStr(cCCV) + "</cCCV>" + _
                "</DATA>"

        Return oXml

    End Function
End Class
