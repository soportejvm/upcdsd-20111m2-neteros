
Imports System.Windows.Media.Imaging
Imports System.IO
Imports System.Text

Partial Public Class RegistroCupon
    Inherits ChildWindow
    Dim oImgOferta As Byte() = Nothing
    Dim oCupon As List(Of Cupon)
    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub OKButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles OKButton.Click
        Dim sCadenaCupones As String = ""

        If txtTitulo.Text = "" Then
            MessageBox.Show("Ingrese Titulo del Cupón", "Mensaje", MessageBoxButton.OK)
            Exit Sub
        End If

        If txtDescripcion.Text = "" Then
            MessageBox.Show("Ingrese Descripcion del Cupón", "Mensaje", MessageBoxButton.OK)
            Exit Sub
        End If

        If txtDias.Text = 0 Then
            MessageBox.Show("Ingrese Dias de Duracion del Cupon", "Mensaje", MessageBoxButton.OK)
            Exit Sub
        End If

        If txtPrecio.Text = 0 Then
            MessageBox.Show("Ingrese Precio del Cupón", "Mensaje", MessageBoxButton.OK)
            Exit Sub
        End If

        If txtDscto.Text = 0 Then
            MessageBox.Show("Ingrese Descuento del Cupón", "Mensaje", MessageBoxButton.OK)
            Exit Sub
        End If

        If txtRestricciones.Text = "" Then
            MessageBox.Show("Ingrese Restricciones del Cupón", "Mensaje", MessageBoxButton.OK)
            Exit Sub
        End If

        If ValidaLongitudFotos() = False Then Exit Sub

        Call LlenaDatosCupon()


        For Each item In oCupon
            sCadenaCupones = sCadenaCupones + item.DevuelveXMLCupon
        Next

        Dim webService As ServiceReference1.ServiceGrouponClient = New ServiceReference1.ServiceGrouponClient
        AddHandler webService.RegistrarCuponCompleted, AddressOf RegistraCupon
        webService.RegistrarCuponAsync(GeneraCabeceraXML(sCadenaCupones), oImgOferta)
        webService = Nothing

    End Sub

    Sub RegistraCupon(ByVal sender As Object, ByVal e As ServiceReference1.RegistrarCuponCompletedEventArgs)
        MessageBox.Show("Se Registro el Cupón correctamente", "Mensaje", MessageBoxButton.OK)
        Me.DialogResult = True

    End Sub

    Public Sub LlenaDatosCupon()
        Dim objCupon As Cupon

        oCupon = New List(Of Cupon)

        objCupon = New Cupon
        With objCupon
            .Titulo = Trim(Me.txtTitulo.Text)
            .Descripcion = Trim(Me.txtDescripcion.Text)
            .Costo = CDbl(Me.txtPrecio.Text)
            .Descuento = CDbl(Me.txtDscto.Text)
            .Dias = CInt(Me.txtDias.Text)
            .Restricciones = Trim(Me.txtRestricciones.Text)
        End With

        oCupon.Add(objCupon)
        objCupon = Nothing
    End Sub

    Private Sub CancelButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles CancelButton.Click
        Me.DialogResult = False
    End Sub

    Private Sub cmdCargaD_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cmdCargaD.Click
        Dim oFileDia As OpenFileDialog = New OpenFileDialog
        ' Set filter options and filter index.
        oFileDia.Filter = "Jpg Files (*.jpg)|*.jpg"
        oFileDia.FilterIndex = 1
        oFileDia.Multiselect = True

        Dim UserClickedOK As Boolean = oFileDia.ShowDialog

        ' Process input if the user clicked OK.
        If (UserClickedOK = True) Then
            Me.txtRutaD.Content = oFileDia.File.ToString

            Dim imageSource As New BitmapImage()
            Dim stream As Stream = oFileDia.File.OpenRead()
            Dim binaryReader As New BinaryReader(stream)
            oImgOferta = binaryReader.ReadBytes(CInt(stream.Length))
            stream.Position = 0
            imageSource.SetSource(stream)
            Me.imgOferta.Source = imageSource
        End If
    End Sub

    Function ValidaLongitudFotos() As Boolean
        ValidaLongitudFotos = True

        If Not (oImgOferta Is Nothing) Then
            If oImgOferta.Length > 500 * 1024 Then
                MessageBox.Show("El tamaño de la foto del domicilio 1 es muy grande" & vbCrLf & "Tamaño máximo 500 Kb", "Validación", MessageBoxButton.OK)
                ValidaLongitudFotos = False
                Exit Function
            End If
        End If

    End Function
End Class
