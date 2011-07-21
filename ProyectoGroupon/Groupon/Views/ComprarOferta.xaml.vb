Partial Public Class ComprarOferta
    Inherits ChildWindow
    Dim oPagos As List(Of Pagos)
    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub OKButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles OKButton.Click
        Dim sCadenaPagos As String = ""
        If optTarjeta.IsChecked Then
            If txtTarj1.Text = "" Then
                txtMensaje.Text = "Ingrese Nro de Tarjeta Correcto"
                txtTarj1.Focus()
                Exit Sub
            End If

            If txtTarj2.Text = "" Then
                txtMensaje.Text = "Ingrese Nro de Tarjeta Correcto"
                txtTarj2.Focus()
                Exit Sub
            End If

            If txtTarj3.Text = "" Then
                txtMensaje.Text = "Ingrese Nro de Tarjeta Correcto"
                txtTarj3.Focus()
                Exit Sub
            End If

            If txtTarj4.Text = "" Then
                txtMensaje.Text = "Ingrese Nro de Tarjeta Correcto"
                txtTarj4.Focus()
                Exit Sub
            End If

            If txtCCV.Text = "" Then
                txtMensaje.Text = "Ingrese Valor de Validación de la Tarjeta de Crédito"
                txtCCV.Focus()
                Exit Sub
            End If
        End If

        Call LlenaDatosPagos()
        For Each item In oPagos
            sCadenaPagos = sCadenaPagos + item.DevuelveXMLPago()
        Next

        Dim webService As ServiceReference1.ServiceGrouponClient = New ServiceReference1.ServiceGrouponClient
        AddHandler webService.RegistrarPagoCuponCompleted, AddressOf RegistraPagoCupon
        webService.RegistrarPagoCuponAsync(GeneraCabeceraXML(sCadenaPagos))
        webService = Nothing

        Me.DialogResult = True

    End Sub

    Sub RegistraPagoCupon(ByVal sender As Object, ByVal e As ServiceReference1.RegistrarPagoCuponCompletedEventArgs)
        If e.Result = 1 Then
            MessageBox.Show("Se Realizo la Compra, se le envio una Constancia a su mail ", "Compra de Ofertas", MessageBoxButton.OK)
        End If

    End Sub

    Public Sub IniciaPantalla(ByVal pnCodigo As Integer, ByVal pcTitulo As String, ByVal pnMonto As Double)
        Me.txtCodigo.Text = CInt(pnCodigo)
        Me.txtTitulo.Text = pcTitulo
        Me.txtMonto.Text = CDbl(pnMonto)

        Me.Show()
    End Sub

    Public Sub LlenaDatosPagos()
        Dim Pagos As Pagos

        oPagos = New List(Of Pagos)

        Pagos = New Pagos
        With Pagos
            .nIdCupon = Trim(Me.txtCodigo.Text)
            If optTarjeta.IsChecked Then
                .nMedioPago = 1
            Else
                .nMedioPago = 2
            End If
            .nMonto = CDbl(Me.txtMonto.Text)
            .nNrotarjeta = Trim(Me.txtTarj1.Text) & Trim(Me.txtTarj2.Text) & Trim(Me.txtTarj3.Text) & Trim(Me.txtTarj4.Text)
            .cCCV = Trim(Me.txtCCV.Text)
        End With

        oPagos.Add(Pagos)
        Pagos = Nothing
    End Sub

    Private Sub optTarjeta_Checked(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles optTarjeta.Checked
        If Not optTarjeta.IsChecked Then
            txtTarj1.IsEnabled = False
            txtTarj2.IsEnabled = False
            txtTarj3.IsEnabled = False
            txtTarj4.IsEnabled = False
        Else
            txtTarj1.IsEnabled = True
            txtTarj2.IsEnabled = True
            txtTarj3.IsEnabled = True
            txtTarj4.IsEnabled = True
        End If
    End Sub
End Class
