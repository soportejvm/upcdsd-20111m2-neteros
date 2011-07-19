
Imports System.Data.SqlClient

<Serializable()> Public Class clsDataSet
    Public Columna() As String
    Public Dato(,) As String
    Public nNumFilas As Integer
    Public nNumColumn As Integer
    Private nCursor As Integer
    Public cNomTab As String

    Public Sub Seek(ByVal nNumFila As Integer)
        nCursor = nNumFila
    End Sub

    Public Sub EliminaRegistro(ByVal pnRegistro As Integer)
        Dim DatoTemp As Object
        Dim i As Integer
        Dim j As Integer
        Dim nPunt As Integer

        DatoTemp = New Object
        DatoTemp = Dato
        ReDim Dato(0, 0)
        ReDim Dato(nNumFilas - 1, nNumColumn)
        nPunt = 1
        For i = 1 To nNumFilas
            If i <> pnRegistro Then
                For j = 1 To nNumColumn
                    Dato(nPunt, j) = DatoTemp(i, j)
                Next
                nPunt = nPunt + 1
            End If
        Next
        nNumFilas = nNumFilas - 1


    End Sub

    Public Function AddRegistro() As Integer
        Dim DatoTemp As Object
        Dim i As Integer
        Dim j As Integer

        DatoTemp = New Object
        DatoTemp = Dato
        ReDim Dato(0, 0)
        ReDim Dato(nNumFilas + 1, nNumColumn)

        For i = 1 To nNumFilas
            For j = 1 To nNumColumn
                Dato(i, j) = DatoTemp(i, j)
            Next
        Next
        nNumFilas = nNumFilas + 1

        For j = 1 To nNumColumn
            Dato(nNumFilas, j) = ""
        Next

        Seek(nNumFilas) 'ARCV 31-01-2008
        AddRegistro = nNumFilas
    End Function

    Public Sub CargaTabla(Optional ByVal psNomTab As String = "", Optional ByVal psWhere As String = "", _
                            Optional ByVal psCadStore As String = "", _
                            Optional ByVal pnConexion As Integer = 1, _
                            Optional ByVal pbSinTrim As Boolean = False, _
                            Optional ByVal pnTiempoEspera As Integer = -1)

        'pnConexion =1 (Negocio),   pnConexion =2 (Financiero)

        Dim objCon As ComConecta.clsConecta
        Dim i As Integer
        Dim dr As SqlClient.SqlDataReader
        Dim oComm As SqlCommand

        If psCadStore = "" Then

            cNomTab = psNomTab

            'Carga Nombres de Columnas de Tablas
            objCon = New ComConecta.clsConecta

            If pnConexion = 1 Then
                oComm = New SqlCommand("RecuperaColumnas", objCon.DameConexion)
                'ElseIf pnConexion = 2 Then
                '    oComm = New SqlCommand("RecuperaColumnas", objCon.DameConexionFinanc)
            End If

            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@NomTabla", SqlDbType.VarChar)
            oComm.Parameters("@NomTabla").Value = psNomTab

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            ReDim Columna(0)
            i = 0
            While dr.Read
                i = i + 1
                ReDim Preserve Columna(i)
                Columna(i) = Trim(dr.Item("name"))
            End While
            nNumColumn = i
            dr.Close()
            dr = Nothing
            objCon.CierraConexion()
            objCon = Nothing
            oComm.Dispose()
            oComm = Nothing

            'Carga Data de Tabla

            objCon = New ComConecta.clsConecta

            'oComm = New SqlCommand("RecuperaDatosCmdSql", objCon.DameConexion)
            If pnConexion = 1 Then
                oComm = New SqlCommand("RecuperaDatosCmdSql", objCon.DameConexion)
                'ElseIf pnConexion = 2 Then
                '    oComm = New SqlCommand("RecuperaDatosCmdSql", objCon.DameConexionFinanc)
            End If
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@NomTabla", SqlDbType.VarChar)
            oComm.Parameters("@NomTabla").Value = psNomTab
            oComm.Parameters.Add("@Where", SqlDbType.VarChar)
            oComm.Parameters("@Where").Value = psWhere

            oComm.Parameters.Add("@NumRegs", SqlDbType.Int)
            oComm.Parameters.Item("@NumRegs").Direction = ParameterDirection.Output
            oComm.ExecuteNonQuery()
            nNumFilas = CInt(oComm.Parameters.Item("@NumRegs").Value.ToString)
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

            ReDim Dato(nNumFilas, nNumColumn)
            nCursor = 0
            While dr.Read
                nCursor = nCursor + 1
                For i = 1 To nNumColumn
                    If IsDBNull(dr.Item(Columna(i))) Then
                        Dato(nCursor, i) = ""
                    Else
                        Dato(nCursor, i) = Trim(dr.Item(Columna(i)))
                    End If
                Next i
            End While
            nNumFilas = nCursor

            dr.Close()
            dr = Nothing
            objCon.CierraConexion()
            objCon = Nothing
            oComm.Dispose()
            oComm = Nothing
        Else

            'Carga Nombres de Columnas de Tablas
            objCon = New ComConecta.clsConecta
            'oComm = New SqlCommand("EjecutaStoreSql", objCon.DameConexion)
            If pnConexion = 1 Then
                oComm = New SqlCommand("EjecutaStoreSql", objCon.DameConexion)
                'ElseIf pnConexion = 2 Then
                '    oComm = New SqlCommand("EjecutaStoreSql", objCon.DameConexionFinanc)
            End If

            If pnTiempoEspera <> -1 Then
                oComm.CommandTimeout = pnTiempoEspera
            End If
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@sCadena", SqlDbType.VarChar).Value = psCadStore
            oComm.Parameters.Add("@NumRegs", SqlDbType.Int)
            oComm.Parameters.Item("@NumRegs").Direction = ParameterDirection.Output
            oComm.ExecuteNonQuery()
            nNumFilas = CInt(oComm.Parameters.Item("@NumRegs").Value.ToString)
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

            'Carga Columnas
            nNumColumn = dr.FieldCount
            ReDim Preserve Columna(nNumColumn)
            For i = 1 To nNumColumn
                Columna(i) = Trim(dr.GetName(i - 1))
            Next

            'Carga Datos
            ReDim Dato(nNumFilas, nNumColumn)
            nCursor = 0

            While dr.Read
                nCursor = nCursor + 1
                For i = 1 To nNumColumn
                    Dato(nCursor, i) = Trim(IIf(IsDBNull(dr.Item(Columna(i))), "", dr.Item(Columna(i))))
                Next i
            End While
            nNumFilas = nCursor

            dr.Close()
            dr = Nothing
            objCon.CierraConexion()
            objCon = Nothing
            oComm.Dispose()
            oComm = Nothing
        End If
    End Sub

    Public Function ObtenerDato(ByVal psNomCampo As String, _
                                Optional ByVal nFila As Integer = 0) As String
        Dim i As Integer
        ObtenerDato = "@#err"
        For i = 0 To nNumColumn
            If UCase(Columna(i)) = UCase(psNomCampo) Then
                'ARCV 28-02-2008
                'ObtenerDato = Dato(nCursor, i)
                If nFila = 0 Then
                    ObtenerDato = Dato(nCursor, i)
                Else
                    ObtenerDato = Dato(nFila, i)
                End If
                '-------------
                Exit For
            End If

        Next
        If ObtenerDato <> "@#err" Then
            If ObtenerDato = Nothing Then
                If Mid(psNomCampo, 1, 1) = "n" Then
                    ObtenerDato = "0"
                Else
                    ObtenerDato = ""
                End If

            End If
        End If
    End Function

    Public Sub Inicio()
        nCursor = 1
    End Sub

    Public Sub Fin()
        nCursor = nNumFilas
    End Sub

    Public Sub Previo()
        If nCursor > 0 Then nCursor -= 1
    End Sub

    Public Sub Siguiente()
        If Not EOF() Then nCursor += 1
    End Sub

    Public Function EOF() As Boolean
        If nCursor > nNumFilas Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub AsignarDato(ByVal sDato As String, _
                            ByVal sNomColumna As String, _
                            Optional ByVal nNumFila As Integer = 0)

        Dim i As Integer
        For i = 1 To nNumColumn
            If Columna(i) = sNomColumna Then
                If nNumFila = 0 Then
                    Dato(nCursor, i) = sDato
                Else
                    Dato(nNumFila, i) = sDato
                End If
                Exit For
            End If
        Next

    End Sub

    Public Function DevuelveXML() As String
        Dim oXml As String = ""
        Dim oXmlTemp As String = _
                "<GENESYS>"
        Dim i As Integer
        Inicio()
        While Not EOF()
            oXmlTemp = oXmlTemp + "<DATA "
            For i = 1 To nNumColumn
                oXmlTemp = oXmlTemp + " " + Columna(i) + "=""" + ObtenerDato(Columna(i)) + """"
                '"<" + oData.cNomTab + " " + .Columna(i) + "=" + .ObtenerDato(.Columna(i)) + "></" + oData.cNomTab + ">"

                If Len(oXmlTemp) > 1000 Then
                    oXml = oXml + oXmlTemp
                    oXmlTemp = ""
                End If
            Next
            oXmlTemp = oXmlTemp + "></DATA>"
            Siguiente()
        End While


        If oXmlTemp <> "" Then oXml = oXml + oXmlTemp
        oXml = oXml + "</GENESYS>"
        Return oXml
    End Function

    Public Function DevuelveXMLLinQ() As String
        Dim oXml As String = ""
        Dim oXmlTemp As String = _
                "<GENESYS>"
        Dim i As Integer
        Inicio()

        If nNumFilas = 0 Then
            oXmlTemp = oXmlTemp + "<DATA>"
            For i = 1 To nNumColumn
                oXmlTemp = oXmlTemp + "<" + Columna(i) + "> </" + Columna(i) + "> "
                '"<" + oData.cNomTab + " " + .Columna(i) + "=" + .ObtenerDato(.Columna(i)) + "></" + oData.cNomTab + ">"
                If Len(oXmlTemp) > 1000 Then
                    oXml = oXml + oXmlTemp
                    oXmlTemp = ""
                End If
            Next
            oXmlTemp = oXmlTemp + "</DATA>"
        Else
            While Not EOF()
                oXmlTemp = oXmlTemp + "<DATA>"
                For i = 1 To nNumColumn
                    oXmlTemp = oXmlTemp + "<" + Columna(i) + ">" + ObtenerDato(Columna(i)) + "</" + Columna(i) + "> "
                    '"<" + oData.cNomTab + " " + .Columna(i) + "=" + .ObtenerDato(.Columna(i)) + "></" + oData.cNomTab + ">"

                    If Len(oXmlTemp) > 1000 Then
                        oXml = oXml + oXmlTemp
                        oXmlTemp = ""
                    End If
                Next
                oXmlTemp = oXmlTemp + "</DATA>"
                Siguiente()
            End While
        End If
        If oXmlTemp <> "" Then oXml = oXml + oXmlTemp
        oXml = oXml + "</GENESYS>"

        Return fReemplazaCaracterEspecial_Cloud(oXml)

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
        'lsNomMod = Replace(lsNomMod, "Ñ", "#", , , vbTextCompare)
        'lsNomMod = Replace(lsNomMod, "ñ", "#", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "-", "", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "|", "", , , vbTextCompare)
        'lsNomMod = Replace(lsNomMod, ".", "", , , vbTextCompare)
        'al final
        lsNomMod = Replace(lsNomMod, "   ", " ", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "  ", " ", , , vbTextCompare)
        lsNomMod = Trim(lsNomMod)

        'Ultimas observaciones
        lsNomMod = Replace(lsNomMod, "Ñ", "N", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "ñ", "n", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "_", "-", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "Á", "A", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "É", "E", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "Í", "I", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "Ó", "O", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "Ú", "U", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "á", "a", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "é", "e", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "í", "i", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "ó", "o", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "ú", "u", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "°", "o", , , vbTextCompare)
        '--------

        If lsNomMod Is Nothing Then lsNomMod = ""

        fReemplazaCaracterEspecial = lsNomMod
    End Function

    Public Function fValidaCaracteres(ByVal psNombre As String) As String
        Dim lsNomMod As String
        Dim letra As String = ""
        Dim i As Integer

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
        'Return lsNomMod
    End Function

    Public Function fReemplazaCaracterEspecial_Cloud(ByVal psNombre As String) As String
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
        'lsNomMod = Replace(lsNomMod, "Ñ", "#", , , vbTextCompare)
        'lsNomMod = Replace(lsNomMod, "ñ", "#", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "-", "", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "|", "", , , vbTextCompare)
        'lsNomMod = Replace(lsNomMod, ".", "", , , vbTextCompare)
        'al final
        lsNomMod = Replace(lsNomMod, "   ", " ", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "  ", " ", , , vbTextCompare)
        lsNomMod = Trim(lsNomMod)

        'Ultimas observaciones
        lsNomMod = Replace(lsNomMod, "Ñ", "&#209;", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "ñ", "&#241;", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "_", "-", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "Á", "A", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "É", "E", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "Í", "I", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "Ó", "O", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "Ú", "U", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "á", "a", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "é", "e", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "í", "i", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "ó", "o", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "ú", "u", , , vbTextCompare)
        lsNomMod = Replace(lsNomMod, "°", "o", , , vbTextCompare)
        '--------

        If lsNomMod Is Nothing Then lsNomMod = ""

        fReemplazaCaracterEspecial_Cloud = lsNomMod
    End Function

    Public Sub CrearVacio(ByVal oColumnas As Object)
        Dim i As Integer

        ReDim Columna(UBound(oColumnas))
        nNumColumn = UBound(oColumnas)

        For i = 1 To UBound(oColumnas)
            Columna(i) = Trim(oColumnas(i))
        Next

    End Sub

    Public Function DevolverSumaColumna(ByVal sNomColumna As String, _
                                        Optional ByVal sColumnaEval As String = "", _
                                        Optional ByVal sColumnaValor As String = "") As Double
        Dim nSuma As Double = 0

        Inicio()
        While Not EOF()
            If sColumnaEval <> "" Then
                If sColumnaValor = ObtenerDato(sColumnaEval) Then
                    nSuma = nSuma + CDbl(ObtenerDato(sNomColumna))
                End If
            Else
                nSuma = nSuma + CDbl(ObtenerDato(sNomColumna))
            End If
            Siguiente()
        End While

        Return nSuma
    End Function

    Public Sub EliminaColumna(ByVal sNomColumna As String)
        Dim i As Integer, j As Integer
        Dim nColumnElim As Integer

        For i = 1 To nNumColumn
            If UCase(Columna(i)) = UCase(sNomColumna) Then
                nColumnElim = i
                Exit For
            End If
        Next

        For i = nColumnElim To nNumColumn - 1
            Columna(i) = Columna(i + 1)
        Next
        ReDim Preserve Columna(nNumColumn - 1)

        For j = 1 To nNumFilas
            For i = nColumnElim To nNumColumn - 1
                Dato(j, i) = Dato(j, i + 1)
            Next i
        Next j
        ReDim Preserve Dato(nNumFilas, nNumColumn - 1)
        nNumColumn -= 1

    End Sub

    'ARCV 05-12-2007
    Public Function DevolverMaximoValorColum(ByVal sNomColumna As String) As Double
        Dim nMaximo As Double = 0

        'Asignar el primer valor
        Inicio()
        nMaximo = CDbl(ObtenerDato(sNomColumna))
        Siguiente()
        While Not EOF()
            If nMaximo > CDbl(ObtenerDato(sNomColumna)) Then
                nMaximo = CDbl(ObtenerDato(sNomColumna))
            End If
            Siguiente()
        End While

        Return nMaximo

    End Function

    Public Function DevuelveIndiceColumna(ByVal sNomColumna As String) As Integer

        Dim i As Integer
        Dim nTemp As Integer = 0

        For i = 1 To nNumColumn
            If UCase(Columna(i)) = UCase(sNomColumna) Then
                nTemp = i
                Exit For
            End If
        Next

        Return nTemp
    End Function

    Public Sub AgregaColumna(ByVal sNomColumna As String, _
                            Optional ByVal sValorDefecto As String = "")
        Dim j As Integer
        Dim nIndiceCol As Integer

        ReDim Preserve Dato(nNumFilas, nNumColumn + 1)
        ReDim Preserve Columna(nNumColumn + 1)
        Columna(nNumColumn + 1) = sNomColumna
        nNumColumn += 1

        If sValorDefecto <> "" Then
            nIndiceCol = DevuelveIndiceColumna(sNomColumna)
            For j = 1 To nNumFilas
                Dato(j, nIndiceCol) = sValorDefecto
            Next j
        End If

    End Sub

    Public Sub FormateaColumnaFecha(ByVal sNomColumna As String)
        Dim i As Integer, j As Integer
        Dim nColumnFormat As Integer

        For i = 1 To nNumColumn
            If UCase(Columna(i)) = UCase(sNomColumna) Then
                nColumnFormat = i
                Exit For
            End If
        Next

        For j = 1 To nNumFilas
            Dato(j, nColumnFormat) = CDate(Dato(j, nColumnFormat)).ToString("yyyyMMdd")
        Next j
    End Sub

    Public Function DevolverCopia() As clsDataset
        Dim oDataSet As New clsDataset
        Dim i As Integer, j As Integer
        oDataSet.CrearVacio(Columna)
        Inicio()
        While Not EOF()
            j = oDataSet.AddRegistro()
            oDataSet.Seek(j)
            For i = 1 To nNumColumn
                oDataSet.AsignarDato(ObtenerDato(Columna(i)), Columna(i))
            Next
            Siguiente()
            oDataSet.Siguiente()
        End While
        Inicio()
        Return oDataSet
    End Function

    Public Function Find(ByVal sNomColumna As String, _
                        ByVal sDatoBusq As String, _
                        Optional ByVal bPrimerasLetras As Boolean = False) As Integer

        Dim i As Integer = 0

        Inicio()
        While Not EOF()
            If bPrimerasLetras = False Then
                If sDatoBusq = ObtenerDato(sNomColumna) Then
                    i = nCursor
                End If
            Else
                If InStr(ObtenerDato(sNomColumna), sDatoBusq) = 1 Then 'Busca solo las primeras letras
                    i = nCursor
                End If
            End If
            Siguiente()
        End While
        Inicio()
        Return i
    End Function

    Public Sub ReemplazarDato(ByVal sNomColumna As String, _
                    ByVal sDatoBusq As String, _
                    ByVal sDatoReemp As String)

        Dim i As Integer
        Inicio()

        i = Find(sNomColumna, sDatoBusq)
        If i > 0 Then
            Seek(i)
            AsignarDato(sDatoReemp, sNomColumna)
        End If

    End Sub

    Public Sub GeneraColumnaItem(Optional ByVal sColumnaItem As String = "Item")
        Dim nItem As Integer = 0

        AgregaColumna(sColumnaItem)
        Inicio()
        While Not EOF()
            nItem += 1
            AsignarDato(nItem, sColumnaItem)
            Siguiente()
        End While
        Inicio()
    End Sub

    Public Sub LlenaColumnaVacios(ByVal sNomColumna As String)
        Inicio()
        While Not EOF()
            AsignarDato("", sNomColumna)
            Siguiente()
        End While
        Inicio()
    End Sub

    Public Sub LlenaColumnaCeros(ByVal sNomColumna As String)
        Inicio()
        While Not EOF()
            AsignarDato(0, sNomColumna)
            Siguiente()
        End While
        Inicio()
    End Sub

    Public Sub LlenarConSqlDataReader(ByVal dr As SqlDataReader, _
                                ByVal nNumRegistros As Integer)
        Dim i As Integer

        'Carga Filas
        nNumFilas = nNumRegistros

        'Carga Columnas
        nNumColumn = dr.FieldCount

        ReDim Preserve Columna(nNumColumn)
        For i = 1 To nNumColumn
            Columna(i) = Trim(dr.GetName(i - 1))
        Next

        'Carga Datos
        ReDim Dato(nNumFilas, nNumColumn)
        nCursor = 0
        While dr.Read
            nCursor = nCursor + 1
            For i = 1 To nNumColumn
                If IsDBNull(dr.Item(Columna(i))) Then
                    Dato(nCursor, i) = ""
                Else
                    Dato(nCursor, i) = Trim(dr.Item(Columna(i)))
                End If
            Next i

        End While
        nNumFilas = nCursor
    End Sub

    Public Function DevuelveColumnaEnLinea(ByVal sNomColumna As String) As String
        Dim sTemp As String = ""
        Inicio()
        While Not EOF()
            sTemp = IIf(sTemp = "", "", sTemp & ",") & ObtenerDato(sNomColumna)
            Siguiente()
        End While
        Inicio()
        Return sTemp
    End Function

    Public Function nFilaActual() As Integer
        Return nCursor
    End Function

    Public Sub CopiaColumna(ByVal bNuevo As Boolean, _
                        ByVal sNomColumCopy As String, _
                        ByVal sNomColumPaste As String)

        If bNuevo Then AgregaColumna(sNomColumPaste)
        Inicio()
        While Not EOF()
            AsignarDato(ObtenerDato(sNomColumCopy), sNomColumPaste)
            Siguiente()
        End While
        Inicio()
    End Sub

    Public Sub CopiarDataSet(ByRef oDataSet As ComTipos.clsDataset)

        Dim i As Integer, j As Integer

        Inicio()
        While Not EOF()
            j = oDataSet.AddRegistro()
            oDataSet.Seek(j)
            For i = 1 To nNumColumn
                oDataSet.AsignarDato(ObtenerDato(Columna(i)), Columna(i))
            Next
            Siguiente()
            oDataSet.Siguiente()
        End While

    End Sub

    Public Function ObtenerDato_x_NroCol(ByVal pnNroCol As Integer, _
                            Optional ByVal nFila As Integer = 0) As String

        ObtenerDato_x_NroCol = Dato(IIf(nFila = 0, nCursor, nFila), pnNroCol)

        If ObtenerDato_x_NroCol = Nothing Then
            ObtenerDato_x_NroCol = "@#err"
        End If

    End Function
End Class


