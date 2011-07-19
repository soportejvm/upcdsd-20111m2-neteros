Partial Public Class RegistroCliente
    Inherits ChildWindow
    Dim oClientes As List(Of Cliente)
    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub OKButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles OKButton.Click
        Dim sCadenaClientes As String = ""

        Call LlenaDatosCliente()

        For Each item In oClientes
            sCadenaClientes = sCadenaClientes + item.DevuelveXMLCliente()
        Next

        Dim webService As ServiceReference1.ServiceGrouponClient = New ServiceReference1.ServiceGrouponClient
        AddHandler webService.RegistraClienteCompleted, AddressOf RegistraCliente
        webService.RegistraClienteAsync(GeneraCabeceraXML(sCadenaClientes))
        webService = Nothing

        Me.DialogResult = True
    End Sub

    Sub RegistraCliente(ByVal sender As Object, ByVal e As ServiceReference1.RegistraClienteCompletedEventArgs)
        If e.Result = 0 Then
            MessageBox.Show("Se Registro el Cliente Correctamente, se le envio un mail a su correo para la confirmacion", "Registro de Cliente", MessageBoxButton.OK)
        End If

    End Sub

    Public Sub LlenaDatosCliente()
        Dim Clientes As Cliente

        oClientes = New List(Of Cliente)

        Clientes = New Cliente
        With Clientes
            .pcEmail = Trim(Me.txtMail.Text)
            .pcNombres = Trim(Me.txtNombres.Text)
            .pcApellidos = Trim(Me.txtApellidos.Text)
            .pcCodUsu = Trim(Mid(Me.txtNombres.Text, 1, 1)) & Trim(Me.txtApellidos.Text)
            .pcPass = Trim(Me.txtContraseña.PasswordChar)
        End With

        oClientes.Add(Clientes)
        Clientes = Nothing
    End Sub
End Class
