Imports System.Math

Module ModValidaciones
    Dim unidad(0 To 9) As String
    Dim decena(0 To 9) As String
    Dim centena(0 To 10) As String
    Dim deci(0 To 9) As String
    Dim otros(0 To 15) As String
    Public Function ValidRucSunat(ByVal lcNroRuc As String) As String
        If Len(Trim(lcNroRuc)) <> 11 Then
            ValidRucSunat = ""
            Exit Function
        End If

        Dim i As Integer
        Dim lnResiduo As Integer
        Dim lnUltDigito As Integer

        Dim aArrayRuc(3, 11) As Integer

        For i = 1 To 11
            aArrayRuc(1, i) = Val(Mid(lcNroRuc, i, 1))
        Next

        aArrayRuc(2, 1) = 5
        aArrayRuc(2, 2) = 4
        aArrayRuc(2, 3) = 3
        aArrayRuc(2, 4) = 2
        aArrayRuc(2, 5) = 7
        aArrayRuc(2, 6) = 6
        aArrayRuc(2, 7) = 5
        aArrayRuc(2, 8) = 4
        aArrayRuc(2, 9) = 3
        aArrayRuc(2, 10) = 2
        aArrayRuc(3, 11) = 0

        For i = 1 To 10
            aArrayRuc(3, i) = aArrayRuc(1, i) * aArrayRuc(2, i)
            aArrayRuc(3, 11) = aArrayRuc(3, 11) + aArrayRuc(3, i)
        Next

        lnResiduo = aArrayRuc(3, 11) Mod 11
        lnUltDigito = 11 - lnResiduo

        If lnUltDigito = 11 Or lnUltDigito = 1 Then
            lnUltDigito = 1
        End If

        If lnUltDigito = 10 Or lnUltDigito = 0 Then
            lnUltDigito = 0
        End If

        ValidRucSunat = Trim(Str(lnUltDigito))

    End Function

    Public Function DevuelveHoraSistema() As String
        Return Right("0" + CStr(Now.Hour), 2) + ":" + Right("0" + CStr(Now.Minute), 2) + ":" + Right("0" + CStr(Now.Second), 2)
    End Function
    Function ValidaFecha(ByVal cadfec As String) As String
        Dim i As Integer

        ValidaFecha = ""

        If Len(cadfec) <> 10 Then
            ValidaFecha = "Fecha No Valida"
            Exit Function
        End If
        For i = 1 To 10
            If i = 3 Or i = 6 Then
                If Mid(cadfec, i, 1) <> "/" Then
                    ValidaFecha = "Fecha No Valida"
                    Exit Function
                End If
            Else
                If AscW(Mid(cadfec, i, 1)) < 48 Or AscW(Mid(cadfec, i, 1)) > 57 Then
                    ValidaFecha = "Fecha No Valida"
                    Exit Function
                End If
            End If
        Next i
        'validando dia
        If Val(Mid(cadfec, 1, 2)) < 1 Or Val(Mid(cadfec, 1, 2)) > 31 Then
            ValidaFecha = "Dia No Valido"
            Exit Function
        End If
        'validando mes
        If Val(Mid(cadfec, 4, 2)) < 1 Or Val(Mid(cadfec, 4, 2)) > 12 Then
            ValidaFecha = "Mes No Valido"
            Exit Function
        End If
        'validando año
        If Val(Mid(cadfec, 7, 4)) < 1900 Or Val(Mid(cadfec, 7, 4)) > 9972 Then
            ValidaFecha = "Año No Valido"
            Exit Function
        End If
        'validando con isdate
        If IsDate(cadfec) = False Then
            ValidaFecha = "Mes o Dia No Valido"
            Exit Function
        End If

    End Function
    Public Function DevuelveEdad(ByVal dFecNac As Date, ByVal dFecHoy As Date) As String
        Dim cAño As String
        Dim cMes As String

        cAño = Year(dFecHoy) - Year(dFecNac)

        If Month(dFecNac) > Month(dFecHoy) Then
            cAño = cAño - 1
            cMes = 12 - (Month(dFecNac) - Month(dFecHoy))
        End If

        If Month(dFecNac) < Month(dFecHoy) Then
            cMes = (Month(dFecHoy) - Month(dFecNac))
        End If

        If Month(dFecNac) = Month(dFecHoy) Then
            If Format(dFecNac, "dd") <= Format(dFecHoy, "dd") Then
                cMes = 0
            End If
            If Format(dFecNac, "dd") > Format(dFecHoy, "dd") Then
                cAño = cAño - 1
                cMes = 11
            End If
        End If

        DevuelveEdad = cAño
    End Function

    Public Function ConvierteNumALetras(ByVal nNumero As Double, _
                                    Optional ByVal lSoloText As Boolean = True, _
                                    Optional ByVal lSinMoneda As Boolean = False, _
                                    Optional ByVal pnMoneda As Integer = 1, _
                                    Optional ByVal pbSinDecimales As Boolean = False) As String
        Dim sCent As String
        Dim xValor As Single
        Dim vMoneda As String
        Dim cNumero As String
        ConvierteNumALetras = ""

        cNumero = Format(nNumero, "##,###,##0.00")
        xValor = nNumero - Int(nNumero)
        If xValor = 0 Then
            sCent = " Y 00/100 "
        Else
            sCent = " Y " & Right(Trim(cNumero), 2) & "/100 "
        End If

        If pbSinDecimales Then sCent = ""

        If pnMoneda <> 0 Then
            vMoneda = IIf(pnMoneda = 1, "NUEVOS SOLES", "DOLARES AMERICANOS")
        End If
        If Not lSoloText Then
            ConvierteNumALetras = Trim(IIf(pnMoneda = 1, "S/.", "US $")) & " " & Trim(Format(nNumero, "###,###,##0.00#")) & " ("
        End If
        ConvierteNumALetras = ConvierteNumALetras & Trim(UCase(NumLet(CStr(nNumero), 0))) & sCent & " " & IIf(lSinMoneda, "", Trim(vMoneda)) & IIf(lSoloText, "", ")")
    End Function

    Public Function NumLet(ByVal strNum As String, Optional ByVal vLo As Object = Nothing)   '  , Optional ByVal vMoneda, Optional ByVal vCentimos) As String
        '----------------------------------------------------------
        ' Convierte el número strNum en letras          (28/Feb/91)
        '----------------------------------------------------------
        Dim I As Integer
        Dim Lo As Integer
        Dim iHayDecimal As Integer          'Posición del signo decimal
        Dim sDecimal As String              'Signo decimal a usar
        Dim sEntero As String
        Dim sFraccion As String
        Dim fFraccion As Single
        Dim sNumero As String
        '
        Dim sMoneda As String
        Dim sCentimos As String

        'Averiguar el signo decimal
        sNumero = Format$(25.5, "#.#")
        If InStr(sNumero, ".") Then
            sDecimal = "."
        Else
            sDecimal = ","
        End If
        'Si no se especifica el ancho...
        If vLo Is Nothing Then
            Lo = 0
        Else
            Lo = vLo
        End If
        '
        If Lo Then
            sNumero = Space$(Lo)
        Else
            sNumero = ""
        End If
        'Quitar los espacios que haya por medio
        Do
            I = InStr(strNum, " ")
            If I = 0 Then Exit Do
            strNum = Left$(strNum, I - 1) & Mid$(strNum, I + 1)
        Loop
        'Comprobar si tiene decimales
        iHayDecimal = InStr(strNum, sDecimal)
        If iHayDecimal Then
            sEntero = Left$(strNum, iHayDecimal - 1)
            sFraccion = Mid$(strNum, iHayDecimal + 1) & "00"
            'obligar a que tenga dos cifras
            sFraccion = Left$(sFraccion, 2)
            fFraccion = Val(sFraccion)

            'Si no hay decimales... no agregar nada...
            If fFraccion < 1 Then
                strNum = RTrim$(UnNumero(sEntero) & sMoneda)
                If Lo Then
                    'LSet(sNumero = strNum)
                    sNumero = strNum
                Else
                    sNumero = strNum
                End If
                NumLet = sNumero
                Exit Function
            End If

            sEntero = UnNumero(sEntero)
            sFraccion = sFraccion & "/100"
            strNum = sEntero
            If Lo Then
                'LSet(sNumero = RTrim$(strNum))
                sNumero = RTrim$(strNum)
            Else
                sNumero = RTrim$(strNum)
            End If
            NumLet = sNumero
        Else
            strNum = RTrim$(UnNumero(strNum) & sMoneda)
            If Lo Then
                'LSet(sNumero = strNum)
                sNumero = strNum
            Else
                sNumero = strNum
            End If
            NumLet = sNumero
        End If
    End Function

    Public Function UnNumero(ByVal strNum As String) As String
        '----------------------------------------------------------
        'Esta es la rutina principal                    (10/Jul/97)
        'Está separada para poder actuar con decimales
        '----------------------------------------------------------

        Dim lngA As Double
        Dim Negativo As Boolean
        Dim L As Integer
        Dim Una As Boolean
        Dim Millon As Boolean
        Dim Millones As Boolean
        Dim vez As Integer
        Dim MaxVez As Integer
        Dim K As Integer
        Dim strQ As String
        Dim strB As String
        Dim strU As String
        Dim strD As String
        Dim strC As String
        Dim iA As Integer
        '
        Dim strN() As String

        'Si se amplia este valor... no se manipularán bien los números
        Const cAncho = 12
        Const cGrupos = cAncho \ 3
        '
        If unidad(1) <> "una" Then
            InicializarArrays()
        End If
        'Si se produce un error que se pare el mundo!!!
        'On Local Error GoTo 0

        lngA = IIf(CDbl(strNum) > 0, CDbl(strNum), CDbl(strNum) * -1) 'Abs(CDbl(strNum))
        Negativo = (lngA <> CDbl(strNum))
        strNum = LTrim$(RTrim$(Str$(lngA)))
        L = Len(strNum)

        If lngA < 1 Then
            UnNumero = "cero"
            Exit Function
        End If
        '
        Una = True
        Millon = False
        Millones = False
        If L < 4 Then Una = False
        If lngA > 999999 Then Millon = True
        If lngA > 1999999 Then Millones = True
        strB = ""
        strQ = strNum
        vez = 0

        'ReDim strN(1 To cGrupos)
        ReDim strN(cGrupos)

        'strQ = Right$(String$(cAncho, "0") & strNum, cAncho)
        strQ = Right(RellenaCaracter("", "0", cAncho) & strNum, cAncho)

        For K = Len(strQ) To 1 Step -3
            vez = vez + 1
            strN(vez) = Mid$(strQ, K - 2, 3)
        Next
        MaxVez = cGrupos
        For K = cGrupos To 1 Step -1
            If strN(K) = "000" Then
                MaxVez = MaxVez - 1
            Else
                Exit For
            End If
        Next
        For vez = 1 To MaxVez
            strU = "" : strD = "" : strC = ""
            strNum = strN(vez)
            L = Len(strNum)
            K = Val(Right$(strNum, 2))
            If Right$(strNum, 1) = "0" Then
                K = K \ 10
                strD = decena(K)
            ElseIf K > 10 And K < 16 Then
                K = Val(Mid$(strNum, L - 1, 2))
                strD = otros(K)
            Else
                strU = unidad(Val(Right$(strNum, 1)))
                If L - 1 > 0 Then
                    K = Val(Mid$(strNum, L - 1, 1))
                    strD = deci(K)
                End If
            End If

            If L - 2 > 0 Then
                K = Val(Mid$(strNum, L - 2, 1))
                'Con esto funcionará bien el 100100, por ejemplo...
                If K = 1 Then
                    If Val(strNum) = 100 Then
                        K = 10
                    End If
                End If
                strC = centena(K) & " "
            End If
            '------
            If strU = "uno" And Left$(strB, 4) = " mil" Then strU = ""
            strB = strC & strD & strU & " " & strB

            If (vez = 1 Or vez = 3) Then
                If strN(vez + 1) <> "000" Then strB = " mil " & strB
            End If
            If vez = 2 And Millon Then
                If Millones Then
                    strB = " millones " & strB
                Else
                    strB = "un millón " & strB
                End If
            End If
        Next
        strB = Trim$(strB)
        If Right$(strB, 3) = "uno" Then strB = Left$(strB, Len(strB) - 1) & "a"
        Do                              'Quitar los espacios que haya por medio
            iA = InStr(strB, "  ")
            If iA = 0 Then Exit Do
            strB = Left$(strB, iA - 1) & Mid$(strB, iA + 1)
        Loop
        If Left$(strB, 6) = "un  un" Then strB = Mid$(strB, 5)
        If Left$(strB, 7) = "un  mil" Then strB = Mid$(strB, 5)
        If Right$(strB, 16) <> "millones mil un " Then
            iA = InStr(strB, "millones mil un ")
            If iA Then strB = Left$(strB, iA + 8) & Mid$(strB, iA + 13)
        End If
        If Right$(strB, 6) = "ciento" Then strB = Left$(strB, Len(strB) - 2)
        If Negativo Then strB = "menos " & strB

        UnNumero = Trim$(strB)
    End Function

    Public Sub InicializarArrays()
        'Asignar los valores
        unidad(1) = "un"
        unidad(2) = "dos"
        unidad(3) = "tres"
        unidad(4) = "cuatro"
        unidad(5) = "cinco"
        unidad(6) = "seis"
        unidad(7) = "siete"
        unidad(8) = "ocho"
        unidad(9) = "nueve"
        '
        decena(1) = "diez"
        decena(2) = "veinte"
        decena(3) = "treinta"
        decena(4) = "cuarenta"
        decena(5) = "cincuenta"
        decena(6) = "sesenta"
        decena(7) = "setenta"
        decena(8) = "ochenta"
        decena(9) = "noventa"
        '
        centena(1) = "ciento"
        centena(2) = "doscientos"
        centena(3) = "trescientos"
        centena(4) = "cuatrocientos"
        centena(5) = "quinientos"
        centena(6) = "seiscientos"
        centena(7) = "setecientos"
        centena(8) = "ochocientos"
        centena(9) = "novecientos"
        centena(10) = "cien"                'Parche
        '
        deci(1) = "dieci"
        deci(2) = "veinti"
        deci(3) = "treinta y "
        deci(4) = "cuarenta y "
        deci(5) = "cincuenta y "
        deci(6) = "sesenta y "
        deci(7) = "setenta y "
        deci(8) = "ochenta y "
        deci(9) = "noventa y "
        '
        otros(1) = "1"
        otros(2) = "2"
        otros(3) = "3"
        otros(4) = "4"
        otros(5) = "5"
        otros(6) = "6"
        otros(7) = "7"
        otros(8) = "8"
        otros(9) = "9"
        otros(10) = "10"
        otros(11) = "once"
        otros(12) = "doce"
        otros(13) = "trece"
        otros(14) = "catorce"
        otros(15) = "quince"
    End Sub

    Public Function DevuelveValorITF(ByVal pnMonto As Double) As Double
        Dim nResultado As Double = 0
        Dim nEntero As Integer = 0
        Dim sCadena As String = ""
        Dim sDigito2 As String

        nResultado = CDbl(pnMonto * oLogin.nValorITF * 100)

        If InStr(CStr(nResultado), ".") = 0 Then
            sCadena = CStr(nResultado) & ".00"
        Else
            sCadena = CStr(nResultado)
        End If

        nResultado = CDbl(Mid(sCadena, 1, InStr(sCadena, ".") - 1))

        nEntero = CInt(nResultado)
        sCadena = CStr(nEntero)
        sDigito2 = Right(sCadena, 1)

        If CInt(sDigito2) < 5 Then
            sDigito2 = "0"
        Else
            sDigito2 = "5"
        End If

        sCadena = Mid(sCadena, 1, Len(sCadena) - 1) & sDigito2

        nResultado = CDbl(sCadena) / 100

        Return nResultado
    End Function

End Module
