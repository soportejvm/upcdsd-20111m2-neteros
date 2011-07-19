Public Class Cliente
    Public pcEmail As String
    Public pcNombres As String
    Public pcApellidos As String
    Public pcCodUsu As String
    Public pcPass As String

    Public Function DevuelveXMLCliente() As String
        Dim oXml As String = ""

        oXml = "<DATA>" + "<pcEmail>" + CStr(pcEmail) + "</pcEmail>" + _
                        "<pcNombres>" + CStr(pcNombres) + "</pcNombres>" + _
                        "<pcApellidos>" + CStr(pcApellidos) + "</pcApellidos>" + _
                        "<pcCodUsu>" + CStr(pcCodUsu) + "</pcCodUsu>" + _
                        "<pcPass>" + CStr(pcPass) + "</pcPass>" + _
                "</DATA>"

        Return oXml

    End Function
End Class
