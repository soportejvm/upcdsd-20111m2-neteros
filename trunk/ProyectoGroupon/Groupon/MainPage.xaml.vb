Partial Public Class MainPage
    Inherits UserControl

    Public Sub New()
        InitializeComponent()
    End Sub


    Private Sub UserControl_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded

    End Sub

    Private Sub HyperlinkButton1_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles HyperlinkButton1.Click
        Dim frm As RegistroCliente
        frm = New RegistroCliente
        frm.Show()
        frm = Nothing
    End Sub
End Class