Partial Public Class MainPage
    Inherits UserControl

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub TreeViewItem1_Selected(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles TreeViewItem1.Selected
        Dim frm As RegistroCliente
        frm = New RegistroCliente
        frm.Show()
        frm = Nothing
    End Sub
End Class