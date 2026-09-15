Attribute VB_Name = "mod_ExportF0"
Option Explicit

' ============================================================
' mod_ExportF0  -  Exporta la data del Fondo 0 a un Excel nuevo
'
' Fuentes:
'   1) Archivo diario de Duracion  (F0_Duracion_AAAAMMDD.xlsx, hoja "Base F0")
'      -> portafolio ya valorizado. Posiciones a fecha FMS (t-2).
'   2) Vector de precios SBS (RFL)
'      -> universo de corporativos de corto plazo. Precios a fecha t-1.
'
' Salida: RV_F0_<fechaFMS>.xlsx, sin macros ni formato, hojas "Data F0" y "Config".
'
' Macro a ejecutar: ExportarDataF0
' ============================================================

Private Const SH_CFG        As String = "Config F0"  ' hoja de ajustes en el libro que aloja la macro
Private Const MONEDA_UNIV   As String = "PEN"   ' todo el universo en soles
Private Const DIAS_MAX_CORP As Long = 366       ' (solo aplica si UNIV_TODO_PLAZO = False)
Private Const SOB_TODO_PLAZO As Boolean = True  ' soberanos: sin tope de plazo (para tener la curva completa)
Private Const UNIV_TODO_PLAZO As Boolean = True ' universo COMPLETO: todo lo que este en soles entra,
                                                ' sin tope de plazo (CD seriados, papeles, titulizados,
                                                ' tit. c/d crediticio, CD BCRP, soberanos, bonos 3-4a, etc.)

Sub ExportarDataF0()
    Dim rutaD As Variant, rutaV As Variant
    Dim wbD As Workbook, wbV As Workbook, wbN As Workbook
    Dim shD As Worksheet, shV As Worksheet, ws As Worksheet, cfg As Worksheet
    Dim fFMS As String, fVec As String, fCal As String
    Dim r As Long, ultD As Long, ultV As Long, fila As Long
    Dim nPort As Long, nUniv As Long, nDup As Long
    Dim total As Double, monto As Double
    Dim dic As Object, clave As String, cat As String
    Dim entra As Boolean
    Dim cCod As Long, cIsin As Long, cNemo As Long, cTipo As Long, cEmi As Long
    Dim cMon As Long, cTir As Long, cSpr As Long, cDur As Long, cRat As Long, cVto As Long
    Dim fVecDate As Date, vto As Variant, dias As Double

    rutaD = Application.GetOpenFilename("Excel,*.xls;*.xlsx;*.xlsm", , "1) Archivo diario de Duracion (F0_Duracion_...)")
    If rutaD = False Then Exit Sub
    rutaV = Application.GetOpenFilename("Excel,*.xls;*.xlsx;*.xlsm", , "2) Vector de precios SBS (RFL)")
    If rutaV = False Then Exit Sub

    Dim carpetaSal As String
    carpetaSal = RutaSalida()

    Application.ScreenUpdating = False
    Set wbD = Workbooks.Open(CStr(rutaD), ReadOnly:=True)
    Set shD = Nothing
    On Error Resume Next
    Set shD = wbD.Worksheets("Base F0")
    On Error GoTo 0
    If shD Is Nothing Then
        MsgBox "No encontre la hoja 'Base F0' en " & wbD.Name, vbExclamation
        wbD.Close False: Application.ScreenUpdating = True: Exit Sub
    End If

    ' --- fechas desde la leyenda de la fila 2 ---
    Dim leyenda As String, i As Long
    For i = 1 To 12
        leyenda = leyenda & " " & CStr(shD.Cells(2, i).Value)
    Next i
    fFMS = SacarFecha(leyenda, "FMS")
    fVec = SacarFecha(leyenda, "Vector")
    fCal = SacarFecha(leyenda, "Calculo")
    If fFMS = "" Then fFMS = SoloDigitos(wbD.Name)

    If Len(carpetaSal) = 0 Then carpetaSal = wbD.Path
    If LCase$(Left$(carpetaSal, 4)) = "http" Then
        MsgBox "El archivo de Duracion esta en OneDrive/SharePoint, no en disco local." & vbCrLf & _
               "Al final te voy a pedir donde guardar el export.", vbInformation
    End If

    Set wbV = Workbooks.Open(CStr(rutaV), ReadOnly:=True)
    Set shV = wbV.Worksheets(1)

    cCod = ColV(shV, "digo sb"): If cCod = 0 Then cCod = ColV(shV, "codigo")
    cIsin = ColV(shV, "isin")
    cNemo = ColV(shV, "nem")
    cTipo = ColV(shV, "tipo de instrumento")
    cEmi = ColV(shV, "emisor")
    cMon = ColV(shV, "moneda")
    cTir = ColV(shV, "tir %"): If cTir = 0 Then cTir = ColV(shV, "tir")
    cSpr = ColV(shV, "spread")
    cDur = ColV(shV, "duracion")
    cRat = ColV(shV, "rating")
    cVto = ColV(shV, "fecha v")

    If cCod * cTipo * cVto = 0 Then
        MsgBox "No pude mapear el vector. Codigo=" & cCod & " Tipo=" & cTipo & " FechaVcto=" & cVto, vbExclamation
        wbD.Close False: wbV.Close False: Application.ScreenUpdating = True: Exit Sub
    End If

    fVecDate = Date
    If Len(fVec) = 8 Then fVecDate = DateSerial(CLng(Left$(fVec, 4)), CLng(Mid$(fVec, 5, 2)), CLng(Right$(fVec, 2)))

    ' --- libro nuevo ---
    Set wbN = Workbooks.Add
    Do While wbN.Worksheets.Count > 1
        Application.DisplayAlerts = False
        wbN.Worksheets(wbN.Worksheets.Count).Delete
        Application.DisplayAlerts = True
    Loop
    Set ws = wbN.Worksheets(1): ws.Name = "Data F0"
    Set cfg = wbN.Worksheets.Add(After:=ws): cfg.Name = "Config"
    Dim shC As Worksheet
    Set shC = wbN.Worksheets.Add(After:=cfg): shC.Name = "Curva Sob"
    Dim shT As Worksheet
    Set shT = wbN.Worksheets.Add(After:=shC): shT.Name = "Tasa Ref BCRP"

    Dim enc As Variant
    enc = Array("EN_PORT", "FUENTE_FILA", "CATEGORIA", "ASSET_CLASS", "TIPO_INSTRUMENTO", "EMISOR", "NEMONICO", "ISIN", _
                "COD_SBS", "MONEDA", "RATING", "VCTO", "DIAS_VCTO", "ANIOS_VCTO", "DUR", "YTW_PCT", _
                "SPREAD_PBS", "ORIGEN_VALOR", "MONTO_MM", "PESO_PCT", "FECHA_FMS", "FECHA_VECTOR", "COD_SBS_VECTOR")
    For i = 0 To UBound(enc)
        ws.Cells(1, i + 1).Value = enc(i)
    Next i
    ws.Rows(1).Font.Bold = True

    ' --- portafolio ---
    ultD = shD.Cells(shD.Rows.Count, 2).End(xlUp).Row
    For r = 4 To ultD
        If IsNumeric(shD.Cells(r, 7).Value) Then total = total + CDbl(shD.Cells(r, 7).Value)
    Next r

    Set dic = CreateObject("Scripting.Dictionary")
    fila = 2
    For r = 4 To ultD
        If Len(Trim$(CStr(shD.Cells(r, 2).Value))) > 0 Then
            monto = 0
            If IsNumeric(shD.Cells(r, 7).Value) Then monto = CDbl(shD.Cells(r, 7).Value)
            ws.Cells(fila, 1).Value = "SI"
            ws.Cells(fila, 2).Value = "Portafolio"
            ws.Cells(fila, 3).Value = CategoriaFila(CStr(shD.Cells(r, 1).Value), CStr(shD.Cells(r, 2).Value), CStr(shD.Cells(r, 1).Value))
            ws.Cells(fila, 4).Value = shD.Cells(r, 1).Value          ' Asset Class
            ws.Cells(fila, 5).Value = shD.Cells(r, 1).Value
            ws.Cells(fila, 6).Value = shD.Cells(r, 2).Value          ' Emisor
            ws.Cells(fila, 7).Value = shD.Cells(r, 4).Value          ' Nemonico
            ws.Cells(fila, 8).Value = shD.Cells(r, 3).Value          ' ISIN
            ws.Cells(fila, 9).Value = LimpiarCod(CStr(shD.Cells(r, 5).Value))   ' Codigo SBS
            ws.Cells(fila, 10).Value = shD.Cells(r, 6).Value         ' Moneda
            ws.Cells(fila, 11).Value = shD.Cells(r, 12).Value        ' Rating
            ws.Cells(fila, 12).Value = shD.Cells(r, 8).Value         ' Vcto
            If IsDate(shD.Cells(r, 8).Value) Then
                ws.Cells(fila, 13).Value = CLng(CDate(shD.Cells(r, 8).Value) - fVecDate)
                ws.Cells(fila, 14).Value = CLng(CDate(shD.Cells(r, 8).Value) - fVecDate) / 365
            End If
            ws.Cells(fila, 15).Value = shD.Cells(r, 10).Value        ' Duracion
            ws.Cells(fila, 16).Value = shD.Cells(r, 9).Value         ' YTW
            ws.Cells(fila, 17).Value = shD.Cells(r, 11).Value        ' Spread pb
            ws.Cells(fila, 18).Value = shD.Cells(r, 13).Value        ' Origen
            ws.Cells(fila, 19).Value = monto
            If total > 0 Then ws.Cells(fila, 20).Value = monto / total * 100
            ws.Cells(fila, 21).Value = fFMS
            ws.Cells(fila, 22).Value = fVec

            clave = LimpiarCod(CStr(shD.Cells(r, 5).Value))
            If Len(clave) > 0 Then If Not dic.Exists(clave) Then dic.Add clave, 1
            clave = LimpiarCod(CStr(shD.Cells(r, 3).Value))
            If Len(clave) > 0 Then If Not dic.Exists(clave) Then dic.Add clave, 1

            nPort = nPort + 1
            fila = fila + 1
        End If
    Next r

    ' --- universo del vector ---
    ultV = shV.Cells(shV.Rows.Count, cCod).End(xlUp).Row
    For r = 2 To ultV
        vto = shV.Cells(r, cVto).Value
        If IsDate(vto) Then
            dias = CDate(vto) - fVecDate
            cat = CategoriaFila(CStr(shV.Cells(r, cTipo).Value), CStr(shV.Cells(r, cEmi).Value), "")
            entra = False
            If dias > 0 Then
                If UNIV_TODO_PLAZO Then
                    entra = True                          ' universo completo en soles
                ElseIf cat = "SOBERANO" And SOB_TODO_PLAZO Then
                    entra = True
                ElseIf dias <= DIAS_MAX_CORP Then
                    entra = True
                End If
            End If
            If entra And cMon > 0 Then
                If Not EsSoles(CStr(shV.Cells(r, cMon).Value)) Then entra = False
            End If
            If entra Then
                clave = LimpiarCod(CStr(shV.Cells(r, cCod).Value))
                If Len(clave) = 0 Then clave = LimpiarCod(CStr(shV.Cells(r, cIsin).Value))
                If dic.Exists(clave) Or dic.Exists(LimpiarCod(CStr(shV.Cells(r, cIsin).Value))) Then
                    nDup = nDup + 1
                Else
                    ws.Cells(fila, 1).Value = "NO"
                    ws.Cells(fila, 2).Value = "Universo"
                    ws.Cells(fila, 3).Value = cat
                    ws.Cells(fila, 5).Value = shV.Cells(r, cTipo).Value
                    If cEmi > 0 Then ws.Cells(fila, 6).Value = shV.Cells(r, cEmi).Value
                    If cNemo > 0 Then ws.Cells(fila, 7).Value = shV.Cells(r, cNemo).Value
                    If cIsin > 0 Then ws.Cells(fila, 8).Value = shV.Cells(r, cIsin).Value
                    ws.Cells(fila, 9).Value = LimpiarCod(CStr(shV.Cells(r, cCod).Value))
                    ws.Cells(fila, 23).Value = shV.Cells(r, cCod).Value
                    If cMon > 0 Then ws.Cells(fila, 10).Value = shV.Cells(r, cMon).Value
                    If cRat > 0 Then ws.Cells(fila, 11).Value = shV.Cells(r, cRat).Value
                    ws.Cells(fila, 12).Value = CDate(vto)
                    ws.Cells(fila, 13).Value = CLng(dias)
                    ws.Cells(fila, 14).Value = dias / 365
                    If cDur > 0 Then ws.Cells(fila, 15).Value = shV.Cells(r, cDur).Value
                    If cTir > 0 Then ws.Cells(fila, 16).Value = shV.Cells(r, cTir).Value
                    If cSpr > 0 Then ws.Cells(fila, 17).Value = shV.Cells(r, cSpr).Value
                    ws.Cells(fila, 18).Value = "Vector"
                    ws.Cells(fila, 21).Value = fFMS
                    ws.Cells(fila, 22).Value = fVec
                    nUniv = nUniv + 1
                    fila = fila + 1
                End If
            End If
        End If
    Next r

    ws.Columns("L").NumberFormat = "dd/mm/yyyy"
    ws.Columns("A:W").AutoFit
    ws.Rows(1).AutoFilter

    ' --- Curva soberana (se llena desde Bloomberg) ---
    Dim encC As Variant, nCurva As Long
    encC = Array("FECHA", "TICKER", "DESCRIPCION", "TENOR_ANIOS", "VCTO", "DUR", "YTW_PCT", "FUENTE", "NOTA")
    For i = 0 To UBound(encC)
        shC.Cells(1, i + 1).Value = encC(i)
    Next i
    shC.Rows(1).Font.Bold = True
    nCurva = HeredarHoja(shC, carpetaSal, "RV_F0_" & fFMS & ".xlsx", "Curva Sob", 9)
    If nCurva = 0 Then
        shC.Cells(2, 1).Value = "Llenar desde Bloomberg. Una fila por punto de curva o por bono soberano."
        shC.Cells(3, 1).Value = "FECHA = fecha del dato. TENOR_ANIOS para puntos de curva; VCTO y DUR para bonos."
        shC.Cells(4, 1).Value = "Ejemplo BDH: =BDH(""GRPE10YR Index"";""PX_LAST"";fecha;fecha)"
        shC.Range("A2:A4").Font.Italic = True
        shC.Range("A2:A4").Font.Color = RGB(140, 140, 140)
    End If
    shC.Columns("A:I").AutoFit

    ' --- Historico de la tasa de referencia BCRP ---
    Dim encT As Variant, nTasa As Long
    encT = Array("FECHA", "TASA_PCT", "FUENTE", "NOTA")
    For i = 0 To UBound(encT)
        shT.Cells(1, i + 1).Value = encT(i)
    Next i
    shT.Rows(1).Font.Bold = True
    nTasa = HeredarHoja(shT, carpetaSal, "RV_F0_" & fFMS & ".xlsx", "Tasa Ref BCRP", 4)
    If nTasa = 0 Then
        shT.Cells(2, 1).Value = "Historico de la tasa de referencia del BCRP. Una fila por fecha de vigencia o por mes."
        shT.Cells(3, 1).Value = "Fuente: BCRPData o Bloomberg. TASA_PCT en porcentaje (4.50, no 0.045)."
        shT.Range("A2:A3").Font.Italic = True
        shT.Range("A2:A3").Font.Color = RGB(140, 140, 140)
    End If
    shT.Columns("A:D").AutoFit

    ' --- Benchmark del F0: numero flat, se hereda del export anterior ---
    Dim bmk As String
    bmk = HeredarParam(carpetaSal, "RV_F0_" & fFMS & ".xlsx", "Benchmark F0 (% flat)")

    ' --- Config ---
    Dim k As Long
    k = 0
    CfgFila cfg, k, "Fondo", "Fondo 0"
    CfgFila cfg, k, "Fecha FMS (posiciones, t-2)", fFMS
    CfgFila cfg, k, "Fecha vector (precios, t-1)", fVec
    CfgFila cfg, k, "Fecha de calculo", fCal
    CfgFila cfg, k, "Carpeta de salida", carpetaSal
    CfgFila cfg, k, "Archivo Duracion", Dir(CStr(rutaD))
    CfgFila cfg, k, "Archivo vector", Dir(CStr(rutaV))
    CfgFila cfg, k, "Filas portafolio", nPort
    CfgFila cfg, k, "Filas universo", nUniv
    CfgFila cfg, k, "Del universo ya en portafolio (no duplicados)", nDup
    CfgFila cfg, k, "Criterio universo: moneda", "Soles (PEN / S/ / SOLES / 1)"
    CfgFila cfg, k, "Criterio universo: SIN TOPE DE PLAZO (todo PEN entra)", UNIV_TODO_PLAZO
    CfgFila cfg, k, "Criterio universo: dias max al vcto (solo si tope activo)", DIAS_MAX_CORP
    CfgFila cfg, k, "Criterio universo: soberanos sin tope de plazo", SOB_TODO_PLAZO
    CfgFila cfg, k, "Vector - col codigo SBS", ColLetra(cCod)
    CfgFila cfg, k, "Vector - col tipo", ColLetra(cTipo)
    CfgFila cfg, k, "Vector - col TIR", ColLetra(cTir)
    CfgFila cfg, k, "Vector - col spread", ColLetra(cSpr)
    CfgFila cfg, k, "Vector - col duracion", ColLetra(cDur)
    CfgFila cfg, k, "Vector - col rating", ColLetra(cRat)
    CfgFila cfg, k, "Vector - col fecha vcto", ColLetra(cVto)
    CfgFila cfg, k, "Curva soberana", "Hoja 'Curva Sob', se llena a mano desde Bloomberg"
    CfgFila cfg, k, "Filas de curva heredadas del export anterior", nCurva
    CfgFila cfg, k, "Tasa de referencia BCRP", "Hoja 'Tasa Ref BCRP', historico manual"
    CfgFila cfg, k, "Filas de tasa heredadas del export anterior", nTasa
    CfgFila cfg, k, "Benchmark F0 (% flat)", bmk
    CfgFila cfg, k, "REVISAR", "Unidades de TIR y SPREAD del vector: confirmar si vienen en % o en pb"
    cfg.Columns("A:B").AutoFit
    cfg.Columns("A").Font.Bold = True

    Dim destino As String, alt As Variant
    destino = carpetaSal & Application.PathSeparator & "RV_F0_" & fFMS & ".xlsx"

    Application.DisplayAlerts = False

    ' crea la carpeta si no existe (solo si es ruta local o UNC accesible)
    On Error Resume Next
    If Len(Dir(carpetaSal, vbDirectory)) = 0 Then MkDir carpetaSal
    Err.Clear
    On Error GoTo 0

    ' intento 1: guardar directo
    Dim ok As Boolean, tmp As String, msgErr As String
    On Error Resume Next
    wbN.SaveAs Filename:=destino, FileFormat:=xlOpenXMLWorkbook
    ok = (Err.Number = 0)
    If Not ok Then
        msgErr = "Error " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    ' intento 2: guardar en local y copiar a destino (rutas de red UNC)
    If Not ok Then
        tmp = Environ$("TEMP") & Application.PathSeparator & "RV_F0_" & fFMS & ".xlsx"
        On Error Resume Next
        If Len(Dir(tmp)) > 0 Then Kill tmp
        Err.Clear
        wbN.SaveAs Filename:=tmp, FileFormat:=xlOpenXMLWorkbook
        If Err.Number = 0 Then
            Err.Clear
            If Len(Dir(destino)) > 0 Then Kill destino
            Err.Clear
            FileCopy tmp, destino
            If Err.Number = 0 Then
                ok = True
            Else
                msgErr = msgErr & vbCrLf & "Copia a red: Error " & Err.Number & ": " & Err.Description
                destino = tmp & "   (quedo en la carpeta temporal)"
                ok = True
            End If
            Err.Clear
        End If
        On Error GoTo 0
    End If

    ' intento 3: preguntar donde guardar
    If Not ok Then
        MsgBox "No pude guardar en:" & vbCrLf & destino & vbCrLf & vbCrLf & msgErr, vbExclamation
        alt = Application.GetSaveAsFilename("RV_F0_" & fFMS & ".xlsx", "Excel,*.xlsx", 1, "Elige donde guardar")
        If alt <> False Then
            On Error Resume Next
            wbN.SaveAs Filename:=CStr(alt), FileFormat:=xlOpenXMLWorkbook
            If Err.Number = 0 Then destino = CStr(alt) Else destino = "(no se pudo guardar)"
            Err.Clear
            On Error GoTo 0
        Else
            destino = "(no guardado; el libro quedo abierto)"
        End If
    End If

    Application.DisplayAlerts = True

    wbD.Close False
    wbV.Close False
    Application.ScreenUpdating = True
    ws.Activate
    ws.Range("A1").Select

    Dim msg As String
    msg = "Listo: " & destino & vbCrLf & vbCrLf
    msg = msg & "Portafolio: " & nPort & " instrumentos" & vbCrLf
    msg = msg & "Universo del vector: " & nUniv & vbCrLf
    msg = msg & "Ya estaban en el portafolio (no se duplicaron): " & nDup & vbCrLf & vbCrLf
    msg = msg & "FMS " & fFMS & "  |  Vector " & fVec & vbCrLf
    msg = msg & "Curva soberana: "
    If nCurva > 0 Then
        msg = msg & nCurva & " filas heredadas" & vbCrLf
    Else
        msg = msg & "hoja vacia, llenar desde Bloomberg" & vbCrLf
    End If
    msg = msg & "Tasa BCRP: "
    If nTasa > 0 Then
        msg = msg & nTasa & " filas heredadas" & vbCrLf
    Else
        msg = msg & "hoja vacia, llenar historico" & vbCrLf
    End If
    msg = msg & "Benchmark F0: "
    If Len(bmk) > 0 Then
        msg = msg & bmk & " (heredado)" & vbCrLf
    Else
        msg = msg & "vacio, cargar en Config" & vbCrLf
    End If
    msg = msg & vbCrLf & "REVISAR en Config las unidades de TIR y SPREAD del vector."
    MsgBox msg, vbInformation
End Sub

Private Sub CfgFila(cfg As Worksheet, k As Long, etiqueta As String, valor As Variant)
    k = k + 1
    cfg.Cells(k, 1).Value = etiqueta
    cfg.Cells(k, 2).Value = valor
End Sub

Private Function RutaSalida() As String
    ' Lee la carpeta de salida de la hoja "Config F0" del libro que aloja la macro.
    ' Si la hoja no existe la crea; si la celda esta vacia, se usa la carpeta del
    ' archivo diario de Duracion.
    Dim ws As Worksheet, ruta As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_CFG)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = SH_CFG
        ws.Range("A1").Value = "AJUSTES DEL EXPORT F0"
        ws.Range("A1").Font.Bold = True
        ws.Range("A3").Value = "Carpeta de salida"
        ws.Range("A4").Value = "(pegar aqui la ruta; si se deja vacia se guarda junto al archivo de Duracion)"
        ws.Range("A4").Font.Italic = True
        ws.Range("A4").Font.Color = RGB(140, 140, 140)
        With ws.Range("B3")
            .Interior.Color = RGB(255, 252, 232)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(217, 212, 181)
        End With
        ws.Columns("A").ColumnWidth = 30
        ws.Columns("B").ColumnWidth = 70
    End If

    ruta = CStr(ws.Range("B3").Value)
    ruta = Replace(ruta, Chr(34), "")          ' comillas de "Copiar como ruta de acceso"
    ruta = Replace(ruta, Chr(39), "")
    ruta = Replace(ruta, ChrW(8220), "")       ' comillas tipograficas
    ruta = Replace(ruta, ChrW(8221), "")
    ruta = Replace(ruta, ChrW(160), " ")       ' espacio duro
    ruta = Trim$(ruta)
    Do While Len(ruta) > 0
        If Right$(ruta, 1) = Application.PathSeparator Or Right$(ruta, 1) = " " Then
            ruta = Left$(ruta, Len(ruta) - 1)
        Else
            Exit Do
        End If
    Loop
    If Len(ruta) > 0 Then
        If Len(Dir(ruta, vbDirectory)) = 0 Then
            MsgBox "La carpeta de salida de " & SH_CFG & " no existe:" & vbCrLf & ruta & vbCrLf & vbCrLf & _
                   "Se guardara junto al archivo de Duracion.", vbExclamation
            ruta = ""
        End If
    End If
    RutaSalida = ruta
End Function

Private Function EsSoles(ByVal mon As String) As Boolean
    ' El vector puede escribir la moneda como PEN, S/, S/., SOLES o el codigo 1.
    Dim m As String
    m = UCase$(Trim$(mon))
    EsSoles = False
    If Len(m) = 0 Then Exit Function
    If InStr(m, "PEN") > 0 Then EsSoles = True: Exit Function
    If InStr(m, "SOL") > 0 Then EsSoles = True: Exit Function
    If Left$(m, 2) = "S/" Then EsSoles = True: Exit Function
    If m = "S" Or m = "1" Then EsSoles = True
End Function

Private Function CategoriaFila(tipo As String, emisor As String, assetClass As String) As String
    ' Tres categorias, segun lo pedido: soberanos, corporativos de corta duracion
    ' y CD del BCRP. El detalle fino queda en TIPO_INSTRUMENTO y ASSET_CLASS.
    '   BONOS GOB.CEN. / GDN GOB CEN  -> SOBERANO
    '   CD BCRP                        -> CD BCRP
    '   todo lo demas                  -> CORPORATIVO
    Dim t As String, e As String, a As String
    t = UCase$(Trim$(tipo))
    e = UCase$(Trim$(emisor))
    a = UCase$(Trim$(assetClass))

    If InStr(e, "CENTRAL DE RESERVA") > 0 Or InStr(t, "BCRP") > 0 Then
        CategoriaFila = "CD BCRP"
    ElseIf InStr(t, "GOB") > 0 Or InStr(a, "SOBERAN") > 0 Then
        CategoriaFila = "SOBERANO"
    Else
        CategoriaFila = "CORPORATIVO"
    End If
End Function

Private Function HeredarParam(carpeta As String, excluir As String, etiqueta As String) As String
    ' lee un valor de la hoja Config del export RV_F0 mas reciente
    Dim f As String, ultimo As String, wbP As Workbook, shP As Worksheet, r As Long
    f = Dir(carpeta & Application.PathSeparator & "RV_F0_*.xlsx")
    Do While Len(f) > 0
        If StrComp(f, excluir, vbTextCompare) <> 0 Then
            If f > ultimo Then ultimo = f
        End If
        f = Dir
    Loop
    If Len(ultimo) = 0 Then Exit Function

    On Error Resume Next
    Set wbP = Workbooks.Open(carpeta & Application.PathSeparator & ultimo, ReadOnly:=True)
    On Error GoTo 0
    If wbP Is Nothing Then Exit Function

    Set shP = Nothing
    On Error Resume Next
    Set shP = wbP.Worksheets("Config")
    On Error GoTo 0
    If Not shP Is Nothing Then
        For r = 1 To 40
            If StrComp(Trim$(CStr(shP.Cells(r, 1).Value)), etiqueta, vbTextCompare) = 0 Then
                HeredarParam = CStr(shP.Cells(r, 2).Value)
                Exit For
            End If
        Next r
    End If
    wbP.Close False
End Function

Private Function HeredarHoja(shC As Worksheet, carpeta As String, excluir As String, _
                             nombreHoja As String, nCols As Long) As Long
    ' copia una hoja del export RV_F0 mas reciente de la carpeta (solo filas con fecha valida)
    Dim f As String, ultimo As String, wbP As Workbook, shP As Worksheet
    Dim r As Long, c As Long, ult As Long

    f = Dir(carpeta & Application.PathSeparator & "RV_F0_*.xlsx")
    Do While Len(f) > 0
        If StrComp(f, excluir, vbTextCompare) <> 0 Then
            If f > ultimo Then ultimo = f
        End If
        f = Dir
    Loop
    If Len(ultimo) = 0 Then Exit Function

    On Error Resume Next
    Set wbP = Workbooks.Open(carpeta & Application.PathSeparator & ultimo, ReadOnly:=True)
    On Error GoTo 0
    If wbP Is Nothing Then Exit Function

    Set shP = Nothing
    On Error Resume Next
    Set shP = wbP.Worksheets(nombreHoja)
    On Error GoTo 0
    If Not shP Is Nothing Then
        ult = shP.Cells(shP.Rows.Count, 1).End(xlUp).Row
        For r = 2 To ult
            If IsDate(shP.Cells(r, 1).Value) Then
                HeredarHoja = HeredarHoja + 1
                For c = 1 To nCols
                    shC.Cells(HeredarHoja + 1, c).Value = shP.Cells(r, c).Value
                Next c
            End If
        Next r
    End If
    wbP.Close False
End Function

Private Function SacarFecha(txt As String, etiqueta As String) As String
    Dim p As Long, i As Long, d As String, ch As String
    p = InStr(1, txt, etiqueta, vbTextCompare)
    If p = 0 Then Exit Function
    For i = p To Len(txt)
        ch = Mid$(txt, i, 1)
        If ch >= "0" And ch <= "9" Then
            d = d & ch
            If Len(d) = 8 Then
                SacarFecha = d
                Exit Function
            End If
        ElseIf Len(d) > 0 And ch <> "-" And ch <> "/" Then
            d = ""
        End If
    Next i
    SacarFecha = d
End Function

Private Function SoloDigitos(s As String) As String
    Dim i As Long, ch As String, d As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then d = d & ch
    Next i
    If Len(d) >= 8 Then SoloDigitos = Right$(d, 8) Else SoloDigitos = d
End Function

Private Function ColV(sh As Worksheet, txt As String) As Long
    Dim c As Long, s As String
    For c = 1 To 60
        s = LimpiarEnc(CStr(sh.Cells(1, c).Value))
        If Len(s) > 0 And InStr(1, s, txt, vbTextCompare) > 0 Then
            ColV = c
            Exit Function
        End If
    Next c
End Function

Private Function LimpiarCod(s0 As String) As String
    ' el codigo SBS del vector viene con guiones (01-1000-1-12M29); se quitan para cruzar
    Dim s As String
    s = UCase$(Trim$(s0))
    s = Replace(s, "-", "")
    s = Replace(s, " ", "")
    s = Replace(s, ".", "")
    LimpiarCod = s
End Function

Private Function LimpiarEnc(s0 As String) As String
    Dim s As String
    s = LCase$(Trim$(s0))
    s = Replace(s, "&oacute;", "o")
    s = Replace(s, "&aacute;", "a")
    s = Replace(s, "&eacute;", "e")
    s = Replace(s, "&iacute;", "i")
    s = Replace(s, "&uacute;", "u")
    s = Replace(s, "&ntilde;", "n")
    s = Replace(s, ChrW(243), "o")
    s = Replace(s, ChrW(225), "a")
    s = Replace(s, ChrW(233), "e")
    s = Replace(s, ChrW(237), "i")
    s = Replace(s, ChrW(250), "u")
    s = Replace(s, ChrW(241), "n")
    s = Replace(s, Chr(10), " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    LimpiarEnc = s
End Function

Private Function ColLetra(i As Long) As String
    If i < 1 Then
        ColLetra = "n/d"
        Exit Function
    End If
    ColLetra = Split(Cells(1, i).Address(True, False), "$")(0)
End Function
