Imports System.IO
Imports System.Xml.Linq
Imports System.Linq
Imports Microsoft.VisualBasic.CompilerServices
Imports System.Windows.Media.Imaging

Module ModControles

    Public Function DevuelveXDocument(ByVal sDatos As String) As XDocument        
        'Dim S As Stream = New MemoryStream(Len(sDatos))
        'S.Write(System.Text.Encoding.UTF8.GetBytes(sDatos), 0, Len(sDatos))
        'S.Flush()
        'S.Position = 0
        Dim S As TextReader = New StringReader(sDatos)
        Dim doc As XDocument = XDocument.Load(S)
        Return doc
    End Function

    Public Sub CargaCombo(ByVal Combo As ComboBox, ByVal sDato As String, _
                            Optional ByVal pnValor As String = "NVALOR", _
                            Optional ByVal pnNombre As String = "CNOMCOD")

        'Dim S As Stream = New MemoryStream(Len(sDato))
        'S.Write(System.Text.Encoding.Unicode.GetBytes(sDato), 0, Len(sDato))
        'S.Flush()
        'S.Position = 0

        Dim S As TextReader = New StringReader(sDato)
        Dim doc As XDocument = XDocument.Load(S)
        Dim datos = From element In doc.Descendants("DATA") Select nValor = element.Descendants(pnValor).Value, cNomCod = element.Descendants(pnNombre).Value

        Dim dc As New Dictionary(Of String, String)

        For Each item In datos
            dc.Add(item.nValor, item.cNomCod)
        Next
        Combo.ItemsSource = dc
        Combo.DisplayMemberPath = "Value"
        Combo.SelectedValuePath = "Key"
    End Sub

    Public Sub CargaComboDpto(ByVal Combo As ComboBox, ByVal sDato As String, _
                        Optional ByVal pnValor As String = "CCODDPTO", _
                        Optional ByVal pnNombre As String = "CDPTO")

        'Dim S As Stream = New MemoryStream(Len(sDato))
        'S.Write(System.Text.Encoding.UTF8.GetBytes(sDato), 0, Len(sDato))
        'S.Flush()
        'S.Position = 0
        Dim S As TextReader = New StringReader(sDato)
        Dim doc As XDocument = XDocument.Load(S)
        Dim datos = From element In doc.Descendants("DATA") Select nValor = element.Descendants(pnValor).Value, cNomCod = element.Descendants(pnNombre).Value

        Dim dc As New Dictionary(Of String, String)

        For Each item In datos
            dc.Add(item.nValor, item.cNomCod)
        Next
        Combo.ItemsSource = dc
        Combo.DisplayMemberPath = "Value"
        Combo.SelectedValuePath = "Key"

    End Sub

    Public Sub CargaComboProv(ByVal Combo As ComboBox, ByVal sDato As String, _
                        Optional ByVal pnValor As String = "CCODPROV", _
                        Optional ByVal pnNombre As String = "CPROV")

        'Dim S As Stream = New MemoryStream(Len(sDato))
        'S.Write(System.Text.Encoding.UTF8.GetBytes(sDato), 0, Len(sDato))
        'S.Flush()
        'S.Position = 0
        Dim S As TextReader = New StringReader(sDato)
        Dim doc As XDocument = XDocument.Load(S)
        Dim datos = From element In doc.Descendants("DATA") Select nValor = element.Descendants(pnValor).Value, cNomCod = element.Descendants(pnNombre).Value

        Dim dc As New Dictionary(Of String, String)

        For Each item In datos
            dc.Add(item.nValor, item.cNomCod)
        Next
        Combo.ItemsSource = dc
        Combo.DisplayMemberPath = "Value"
        Combo.SelectedValuePath = "Key"
    End Sub

    Public Function fValidaCaracteres(ByVal psNombre As String) As String
        Dim lsNomMod As String
        Dim letra As String = ""

        If psNombre = "" Then
            Return ""
            Exit Function
        End If

        lsNomMod = psNombre
        'If InStr(1, lsNomMod, "'", vbTextCompare) <> 0 Then
        '    lsNomMod = Replace(lsNomMod, "'", "''", , , vbTextCompare)
        'End If

        For i = 1 To Len(psNombre)
            letra = Mid(psNombre, i, 1)
            'If letra = "í" Then Stop
            If InStr("áéíóúÁÉÍÓÚabcdefghijklmnopqrstuvwxyz0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ ", letra) <= 0 Then
                Select Case Mid(lsNomMod, i, 1)
                    Case Is = "á"
                        lsNomMod = Replace(lsNomMod, letra, "a", , , vbTextCompare)
                    Case Is = "é"
                        lsNomMod = Replace(lsNomMod, letra, "e", , , vbTextCompare)
                    Case Is = "í"
                        lsNomMod = Replace(lsNomMod, letra, "i", , , vbTextCompare)
                    Case Is = "ó"
                        lsNomMod = Replace(lsNomMod, letra, "o", , , vbTextCompare)
                    Case Is = "ú"
                        lsNomMod = Replace(lsNomMod, letra, "u", , , vbTextCompare)
                    Case Is = "Á"
                        lsNomMod = Replace(lsNomMod, letra, "A", , , vbTextCompare)
                    Case Is = "É"
                        lsNomMod = Replace(lsNomMod, letra, "E", , , vbTextCompare)
                    Case Is = "Í"
                        lsNomMod = Replace(lsNomMod, letra, "I", , , vbTextCompare)
                    Case Is = "Ó"
                        lsNomMod = Replace(lsNomMod, letra, "O", , , vbTextCompare)
                    Case Is = "Ú"
                        lsNomMod = Replace(lsNomMod, letra, "U", , , vbTextCompare)
                    Case Else
                        lsNomMod = Replace(lsNomMod, letra, "", , , vbTextCompare)
                End Select

            End If
        Next


        fValidaCaracteres = lsNomMod
    End Function

    Public Function fReemplazaCaracterEspecial(ByVal psNombre As String) As String
        Dim lsNomMod As String

        If psNombre = "" Then
            Return ""
            Exit Function
        End If

        lsNomMod = psNombre
        If InStr(1, lsNomMod, "'", vbTextCompare) <> 0 Then
            lsNomMod = Replace(lsNomMod, "'", "''", , , vbTextCompare)
        End If
        'ARCV 17-12-2007
        'lsNomMod = Replace(lsNomMod, "Ñ", "&#209;", , , vbTextCompare)
        'lsNomMod = Replace(lsNomMod, "ñ", "#", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "–", "", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "-", "", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "|", "", , , vbTextCompare)
        'lsNomMod = Replace(lsNomMod, ".", "", , , vbTextCompare) '07-10-2010
        'al final
        'lsNomMod = Replace(lsNomMod, "   ", " ", , , vbTextCompare)
        'lsNomMod = Replace(lsNomMod, "  ", " ", , , vbTextCompare)
        'lsNomMod = Trim(lsNomMod)

        'Ultimas observaciones
        'lsNomMod = Replace(lsNomMod, "Ñ", "&#209;", , , vbTextCompare)
        'lsNomMod = Replace(lsNomMod, "ñ", "&#241;", , , vbTextCompare)
        'lsNomMod = Replace(lsNomMod, "_", "-", , , vbTextCompare)
        'lsNomMod = Replace(lsNomMod, "Á", "A", , , vbTextCompare)
        'lsNomMod = Replace(lsNomMod, "É", "E", , , vbTextCompare)
        'lsNomMod = Replace(lsNomMod, "Í", "I", , , vbTextCompare)
        'lsNomMod = Replace(lsNomMod, "Ó", "O", , , vbTextCompare)
        'lsNomMod = Replace(lsNomMod, "Ú", "U", , , vbTextCompare)
        'lsNomMod = Replace(lsNomMod, "á", "a", , , vbTextCompare)
        'lsNomMod = Replace(lsNomMod, "é", "e", , , vbTextCompare)
        'lsNomMod = Replace(lsNomMod, "í", "i", , , vbTextCompare)
        'lsNomMod = Replace(lsNomMod, "ó", "o", , , vbTextCompare)
        'lsNomMod = Replace(lsNomMod, "ú", "u", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, ",", "", , , vbTextCompare)

        lsNomMod = Replace(lsNomMod, "&", "Y", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, ";", " ", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, ",", " ", , , vbTextCompare)
        'lsNomMod = Replace(lsNomMod, ".", "", , , vbTextCompare)
        '--------

        If lsNomMod Is Nothing Then lsNomMod = ""

        fReemplazaCaracterEspecial = lsNomMod
    End Function
    
    Public Function fReemplazaCaracterEspecial_soloÑ(ByVal psNombre As String) As String
        Dim lsNomMod As String

        If psNombre = "" Then
            Return ""
            Exit Function
        End If

        lsNomMod = psNombre

        lsNomMod = Replace(lsNomMod, "Ñ", "&#209;", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "ñ", "&#241;", , , vbTextCompare)
        ''lsNomMod = Replace(lsNomMod, "&", "&amp;", , , vbTextCompare)
        ''lsNomMod = Replace(lsNomMod, "<", "Menor", , , vbTextCompare)
        ''lsNomMod = Replace(lsNomMod, ">", "Mayor", , , vbTextCompare)
        ''lsNomMod = Replace(lsNomMod, """", "&#39;", , , vbTextCompare)
        ''lsNomMod = Replace(lsNomMod, "'", "&#34;", , , vbTextCompare)
        ''lsNomMod = Replace(lsNomMod, "&", "&amp;", , , vbTextCompare)

        ''lsNomMod = Replace(lsNomMod, "&", "", , , vbTextCompare)
        'lsNomMod = Replace(lsNomMod, ":", "", , , vbTextCompare)
        'lsNomMod = Replace(lsNomMod, "%", "Porc", , , vbTextCompare)
        'lsNomMod = Replace(lsNomMod, """", "&quot;", , , vbTextCompare)

        'lsNomMod = Replace(lsNomMod, ".", "", , , vbTextCompare)

        '--------

        If lsNomMod Is Nothing Then lsNomMod = ""

        fReemplazaCaracterEspecial_soloÑ = lsNomMod


    End Function
    Public Function NumRegistroDataGrid(ByVal oDataGrid As DataGrid, _
                                        ByVal oLista As Object) As Integer

        Dim i As Integer = 0
        oLista = oDataGrid.ItemsSource

        For Each item In oLista
            i = i + 1
        Next
        Return i

    End Function
    'Public Function IsNumeric(ByVal oValor As Object) As Boolean
    '    Dim returnValue As Type
    '    Dim sTipoDato As String
    '    returnValue = oValor.GetType

    '    sTipoDato = returnValue.

    '    'If sTipoDato = "System.Double" Or sTipoDato = "System.Integer" Then
    '    IsNumeric = True
    '    'Else
    '    'IsNumeric = False
    '    'End If
    'End Function
    'Declaration



    Public Function IsNumeric(ByVal oValor As Object) As Boolean
        'Dim Expression As Object
        Dim returnValue As Boolean

        returnValue = Versioned.IsNumeric(oValor)

        Return returnValue
    End Function

    'Public Function DevuelveSumaColumnaDataGrid(ByRef DGControl As DataGrid, _
    '                                            ByVal sNomColumna As String) As Double
    '    DGControl.ItemsSource
    '    DevuelveSumaColumnaDataGrid = 0
    'End Function

    Public Sub ValidaNumeros(ByRef e As Object, _
                             Optional ByVal bDecimales As Boolean = False)
        If Not IsNumeric(e) Then
            e = IIf(bDecimales, "0.00", "0")
        End If
    End Sub

    Public Sub EnfocaControlTexto(ByRef oTexto As TextBox)
        oTexto.Focus()
        oTexto.SelectionStart = 0
        oTexto.SelectionLength = Len(oTexto.Text)
    End Sub

    Public Function DevuelveTextoCombo(ByVal oCombo As ComboBox) As String
        Dim nItem As Integer = oCombo.SelectedIndex
        Dim i As Integer = 0
        DevuelveTextoCombo = ""
        If nItem >= 0 Then
            Dim dc As New Dictionary(Of String, String)

            dc = oCombo.ItemsSource

            For Each item In dc
                If i = nItem Then
                    DevuelveTextoCombo = item.Value
                End If
                i = i + 1
            Next
        End If
    End Function

    Public Function GeneraEtiquetasXML(ByVal sValor As String) As String
        Return "<DATA>" + sValor + "</DATA>"
    End Function

    Public Function GeneraCabeceraXML(ByVal sValor As String) As String
        Return "<GENESYS>" + sValor + "</GENESYS>"
    End Function

    Public Sub HabilitaIngresoMat(ByVal pbHab As Boolean, ByVal Grid As DataGrid, ByVal btn1 As Button, ByVal btn2 As Button, ByVal btn3 As Button)
        If pbHab Then
            Grid.Height = 88
        Else
            Grid.Height = 204
        End If        

        btn1.IsEnabled = Not pbHab
        btn2.IsEnabled = Not pbHab
        btn3.IsEnabled = Not pbHab
    End Sub

    Public Sub HabilitaIngresoIns(ByVal pbHab As Boolean, ByVal Grid As DataGrid, ByVal btn1 As Button, ByVal btn2 As Button, ByVal btn3 As Button)
        If pbHab Then
            Grid.Height = 89
        Else
            Grid.Height = 204
        End If

        btn1.IsEnabled = Not pbHab
        btn2.IsEnabled = Not pbHab
        btn3.IsEnabled = Not pbHab
    End Sub

    Public Function Bytes2Image(ByVal bytes() As Byte) As BitmapImage
        If bytes Is Nothing Then Return Nothing

        Dim imageSource As New BitmapImage()
        Dim stream As New MemoryStream(bytes)
        Dim binaryReader As New BinaryReader(stream)
        stream.Position = 0
        imageSource.SetSource(stream)
        Return imageSource
    End Function

    Public Function RellenaCaracter(ByVal sCadena As String, _
                        Optional ByVal sCaracter As String = " ", _
                        Optional ByVal nNum As Integer = -1, _
                        Optional ByVal nAlinear As HorizontalAlignment = HorizontalAlignment.Left) As String
        Dim sTemp As String = ""
        Dim i As Integer

        If nNum = -1 Then nNum = sCadena.Length

        If sCadena.Length > nNum Then
            Return Mid(sCadena, 1, nNum)    'Retornar solo los nNum caracteres
        End If
        For i = 1 To nNum - sCadena.Length
            sTemp = sTemp & sCaracter
        Next

        If nAlinear = HorizontalAlignment.Left Then
            sTemp = sCadena & sTemp
        ElseIf nAlinear = HorizontalAlignment.Right Then
            sTemp = sTemp & sCadena
        Else 'Centro
            If (nNum - sCadena.Length) Mod 2 = 0 Then
                sTemp = Space((nNum - sCadena.Length) / 2) & sCadena & Space((nNum - sCadena.Length) / 2)
            Else
                'sTemp = Space(CInt((nNum - sCadena.Length) / 2)) & sCadena & Space(CInt((nNum - sCadena.Length) / 2) - 1)
                sTemp = Space(CInt((nNum - sCadena.Length) / 2)) & sCadena & Space(IIf(CInt((nNum - sCadena.Length) / 2) - 1 < 0, 0, CInt((nNum - sCadena.Length) / 2) - 1))
            End If
        End If

        Return sTemp
    End Function

    Public Function SaltoLinea(Optional ByVal nNumSaltos As Integer = 1) As String
        Dim sTemp As String = ""
        Dim i As Integer

        For i = 1 To nNumSaltos
            sTemp = sTemp & ChrW(10)
        Next
        Return sTemp
    End Function

    'Public Sub LLenaTreeView(ByVal indicePadre As Integer, ByVal nodePadre As TreeViewItem, ByVal D As List(Of ClsTreeView), ByRef TreeView1 As TreeView)

    '    Dim var = From p In D
    '             Where p.nNivel Like indicePadre.ToString
    '             Select p

    '    For Each item In var
    '        Dim nuevoNodo As New TreeViewItem
    '        nuevoNodo.Header = UCase(item.cDescrip)
    '        nuevoNodo.Tag = item.nCodOpe

    '        If nodePadre Is Nothing Then
    '            TreeView1.Items.Add(nuevoNodo)
    '        Else
    '            nodePadre.Items.Add(nuevoNodo)
    '        End If
    '        LLenaTreeView(Int32.Parse(item.nCodOpe), nuevoNodo, D, TreeView1)
    '    Next
    'End Sub
End Module
