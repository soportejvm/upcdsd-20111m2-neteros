Partial Public Class MainPage
    Inherits UserControl

    Public Sub New()
        InitializeComponent()
    End Sub


    Private Sub UserControl_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        lblUsuario.Content = ""
        Dim webService As ServiceReference1.ServiceGrouponClient = New ServiceReference1.ServiceGrouponClient
        AddHandler webService.DevuelveOfertasCompleted, AddressOf DevuelveOfertas
        webService.DevuelveOfertasAsync()
        webService = Nothing

    End Sub

    Sub DevuelveOfertas(ByVal sender As Object, ByVal e As ServiceReference1.DevuelveOfertasCompletedEventArgs)

        DGOfertas.ItemsSource = (From element In DevuelveXDocument(e.Result).Descendants("DATA") Select New Cupon With {.Codigo = element.Descendants("id").Value, .Titulo = element.Descendants("cTitulo").Value, .Descripcion = element.Descendants("cDescripcion").Value, _
                                                                                            .Costo = element.Descendants("nCosto").Value, .Descuento = element.Descendants("nDesc").Value})
        '.Foto = Bytes2Image(element.Descendants("oFotoOferta").Value)})


    End Sub

    Private Sub HyperlinkButton1_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles HyperlinkButton1.Click
        Dim frm As RegistroCliente
        frm = New RegistroCliente
        frm.Show()
        frm = Nothing
    End Sub

    Private Sub HyperlinkButton3_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles HyperlinkButton3.Click
        Dim frm As RegistroCupon
        frm = New RegistroCupon
        frm.Show()
        frm = Nothing
    End Sub

    Private Sub DGOfertas_SelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles DGOfertas.SelectionChanged
        Dim nItem As Integer = DGOfertas.SelectedIndex
        Dim pcmail As String = ""

        If nItem >= 0 Then
            Dim i As Integer = 0
            Dim oLista As IEnumerable(Of Cupon)

            oLista = DGOfertas.ItemsSource

            For Each item In oLista
                If i = nItem Then
                    If lblUsuario.Content = "" Then
                        MessageBox.Show("Para acceder a la oferta debe loguearse", "Compra de Ofertas", MessageBoxButton.OK)
                        Dim frmLog As Login
                        frmLog = New Login
                        frmLog.IniciaPantalla(pcmail)
                        AddHandler frmLog.Closed, AddressOf logeo
                        frmLog = Nothing
                        Exit Sub
                    End If
                    Dim frm As ComprarOferta
                    frm = New ComprarOferta

                    frm.IniciaPantalla(item.Codigo, item.Titulo, item.Costo)
                    AddHandler frm.Closed, AddressOf Comprar
                    frm = Nothing
                End If
                i = i + 1
            Next
            'Else
            '    MessageBox.Show("Debe seleccionar un elemento de la lista", "Mensaje", MessageBoxButton.OK)
        End If
    End Sub

    Sub Comprar(ByVal sender As Object, ByVal e As EventArgs)

    End Sub

    Sub logeo(ByVal sender As Object, ByVal e As EventArgs)
        lblUsuario.Content = oLogin
    End Sub

    Private Sub HyperlinkButton2_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles HyperlinkButton2.Click
        Dim frm As Login
        frm = New Login
        frm.Show()
        lblUsuario.Content = oLogin
        frm = Nothing
    End Sub
End Class