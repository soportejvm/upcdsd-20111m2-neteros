
Partial Public Class RegistroCliente
    Inherits ChildWindow
    Dim oClientes As List(Of Cliente)
    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub OKButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles OKButton.Click
        Dim sCadenaClientes As String = ""

        If Trim(Me.txtNombres.Text) = "" Then
            txtMensaje.Text = "Ingrese Nombres"
            txtNombres.Focus()
            Exit Sub
        End If

        If Trim(Me.txtApellidos.Text) = "" Then
            txtMensaje.Text = "Ingrese Apellidos"
            txtApellidos.Focus()
            Exit Sub
        End If


        If Trim(Me.txtMail.Text) = "" Then
            txtMensaje.Text = "Ingrese E-Mail"
            txtMail.Focus()
            Exit Sub
        End If


        If Trim(Me.txtContraseña.Password) = "" Then
            txtMensaje.Text = "Ingrese Contraseña"
            txtContraseña.Focus()
            Exit Sub
        End If

        If Trim(Me.txtRepContraseña.Password) = "" Then
            txtMensaje.Text = "Repita la Contraseña"
            txtRepContraseña.Focus()
            Exit Sub
        End If

        txtMensaje.Text = ""


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
            .pcCodUsu = Trim(Me.txtMail.Text)
            .pcPass = Trim(Me.txtContraseña.Password)
        End With

        oClientes.Add(Clientes)
        Clientes = Nothing
    End Sub
End Class
