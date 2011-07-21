Partial Public Class Login
    Inherits ChildWindow
    Dim cEmail As String
    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub HyperlinkButton1_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles HyperlinkButton1.Click
        Dim webService As ServiceReference1.ServiceGrouponClient = New ServiceReference1.ServiceGrouponClient
        AddHandler webService.ValidarUsuarioCompleted, AddressOf ValidarUsuario
        webService.ValidarUsuarioAsync(Me.txtEmail.Text)
        webService = Nothing
    End Sub

    Public Sub IniciaPantalla(ByRef pcEmail As String)
        Me.Show()        
    End Sub

    Sub ValidarUsuario(ByVal sender As Object, ByVal e As ServiceReference1.ValidarUsuarioCompletedEventArgs)

        Dim Datos = From element In DevuelveXDocument(e.Result).Descendants("DATA") Select pcEmail = element.Descendants("cMail").Value, pcNombres = element.Descendants("cNombres").Value, _
        pcApellidos = element.Descendants("cApellidos").Value

        For Each item In Datos
            If item.pcEmail = "" Then
                MessageBox.Show("No se encuentra Registrado en nuestra Pagina", "Regístrese", MessageBoxButton.OK)
            End If
            oLogin = item.pcEmail
        Next
        Me.DialogResult = True
    End Sub
End Class
