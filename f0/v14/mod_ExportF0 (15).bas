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
Private Const DIAS_MAX_CORP As Long = 1830      ' corporativos y demas: hasta 5 anos al vencimiento
Private Const SOB_TODO_PLAZO As Boolean = True  ' soberanos: TODOS, sin tope de plazo (curva completa)
Private Const UNIV_TODO_PLAZO As Boolean = False ' False = tope de plazo activo para lo no soberano

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
    Dim shX As Worksheet, filaX As Long
    Set shX = wbN.Worksheets.Add(After:=shT): shX.Name = "Excluidos"
    shX.Range("A1:H1").Value = Array("MOTIVO", "TIPO_INSTRUMENTO", "EMISOR", "NEMONICO", "ISIN", "COD_SBS", "MONEDA", "VCTO")
    shX.Rows(1).Font.Bold = True
    filaX = 2

    Dim enc As Variant
    enc = Array("EN_PORT", "FUENTE_FILA", "CATEGORIA", "ASSET_CLASS", "TIPO_INSTRUMENTO", "EMISOR", "NEMONICO", "ISIN", _
                "COD_SBS", "MONEDA", "RATING", "VCTO", "DIAS_VCTO", "ANIOS_VCTO", "DUR", "YTW_PCT", _
                "SPREAD_PBS", "ORIGEN_VALOR", "MONTO_MM", "PESO_PCT", "FECHA_FMS", "FECHA_VECTOR", "COD_SBS_VECTOR", "ASSET_CLASS_N")
    For i = 0 To UBound(enc)
        ws.Cells(1, i + 1).Value = enc(i)
    Next i
    ws.Rows(1).Font.Bold = True

    ' --- mapa codigo -> ISIN del vector (para uniformar el ISIN del portafolio) ---
    Dim dISIN As Object, uVec As Long, kIsin As String
    Set dISIN = CreateObject("Scripting.Dictionary")
    If cIsin > 0 Then
        uVec = shV.Cells(shV.Rows.Count, cCod).End(xlUp).Row
        For r = 2 To uVec
            clave = LimpiarCod(TxtSeg(shV.Cells(r, cCod).Value))
            kIsin = TxtSeg(shV.Cells(r, cIsin).Value)
            If Len(clave) > 0 And Len(kIsin) > 0 Then
                If Not dISIN.Exists(clave) Then dISIN.Add clave, kIsin
            End If
        Next r
    End If

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
            ' ISIN uniforme: preferir el del vector (via codigo SBS); si no, el del FMS
            kIsin = ""
            clave = LimpiarCod(TxtSeg(shD.Cells(r, 5).Value))
            If Len(clave) > 0 Then If dISIN.Exists(clave) Then kIsin = CStr(dISIN(clave))
            If Len(kIsin) = 0 Then kIsin = TxtSeg(shD.Cells(r, 3).Value)
            ws.Cells(fila, 8).Value = kIsin
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
            ws.Cells(fila, 24).Value = ACNorm(TxtSeg(shD.Cells(r, 1).Value))
            ws.Cells(fila, 19).Value = monto
            If total > 0 Then ws.Cells(fila, 20).Value = monto / total * 100
            ws.Cells(fila, 21).Value = fFMS
            ws.Cells(fila, 22).Value = fVec

            clave = LimpiarCod(CStr(shD.Cells(r, 5).Value))
            If Len(clave) > 0 Then If Not dic.Exists(clave) Then dic.Add clave, 1
            clave = LimpiarCod(CStr(shD.Cells(r, 3).Value))
            If Len(clave) > 0 Then If Not dic.Exists(clave) Then dic.Add clave, 1
            clave = LimpiarCod(kIsin)
            If Len(clave) > 0 Then If Not dic.Exists(clave) Then dic.Add clave, 1

            nPort = nPort + 1
            fila = fila + 1
        End If
    Next r

    ' --- universo del vector ---
    ' ultima fila real: el maximo entre codigo, nemonico y fecha de vcto
    ' (los CD seriados a veces vienen sin codigo SBS y quedaban fuera del barrido)
    ultV = shV.Cells(shV.Rows.Count, cCod).End(xlUp).Row
    If cNemo > 0 Then If shV.Cells(shV.Rows.Count, cNemo).End(xlUp).Row > ultV Then ultV = shV.Cells(shV.Rows.Count, cNemo).End(xlUp).Row
    If shV.Cells(shV.Rows.Count, cVto).End(xlUp).Row > ultV Then ultV = shV.Cells(shV.Rows.Count, cVto).End(xlUp).Row

    Dim nSinFecha As Long, nVencidos As Long, nNoSoles As Long
    Dim nSinRat As Long, nIsinUS As Long, nRatExcl As Long
    Dim dVto As Date, hayInstr As Boolean
    On Error GoTo FalloUniv
    For r = 2 To ultV
        ' la fila cuenta como instrumento si trae CUALQUIER identidad:
        ' codigo, nemonico, ISIN o emisor (los CD seriados a veces solo traen ISIN/emisor)
        hayInstr = (Len(TxtSeg(shV.Cells(r, cCod).Value)) > 0)
        If Not hayInstr And cNemo > 0 Then hayInstr = (Len(TxtSeg(shV.Cells(r, cNemo).Value)) > 0)
        If Not hayInstr And cIsin > 0 Then hayInstr = (Len(TxtSeg(shV.Cells(r, cIsin).Value)) > 0)
        If Not hayInstr And cEmi > 0 Then hayInstr = (Len(TxtSeg(shV.Cells(r, cEmi).Value)) > 0)
        If Not hayInstr Then GoTo SigV
        vto = shV.Cells(r, cVto).Value
        dVto = FechaVal(vto)
        If dVto = 0 Then
            nSinFecha = nSinFecha + 1
            AnotarExcl shX, filaX, "Sin fecha vcto legible", shV, r, cTipo, cEmi, cNemo, cIsin, cCod, cMon, cVto
        End If
        If dVto > 0 Then
            dias = dVto - fVecDate
            cat = CategoriaFila(TxtSeg(shV.Cells(r, cTipo).Value), TxtSeg(shV.Cells(r, cEmi).Value), "")
            entra = False
            If dias > 0 Then
                If UNIV_TODO_PLAZO Then
                    entra = True                          ' universo completo en soles
                ElseIf cat = "SOBERANO" And SOB_TODO_PLAZO Then
                    entra = True
                ElseIf dias <= DIAS_MAX_CORP Then
                    entra = True
                End If
            Else
                nVencidos = nVencidos + 1
                AnotarExcl shX, filaX, "Vencido", shV, r, cTipo, cEmi, cNemo, cIsin, cCod, cMon, cVto
            End If
            If entra And cMon > 0 Then
                If Not EsSoles(TxtSeg(shV.Cells(r, cMon).Value)) Then
                    entra = False
                    nNoSoles = nNoSoles + 1
                    AnotarExcl shX, filaX, "Otra moneda", shV, r, cTipo, cEmi, cNemo, cIsin, cCod, cMon, cVto
                End If
            End If
            ' solo con rating (soberanos y CD BCRP exentos: vienen sin clasificacion)
            If entra And cRat > 0 Then
                If cat <> "SOBERANO" And cat <> "CD BCRP" Then
                    If Len(TxtSeg(shV.Cells(r, cRat).Value)) = 0 Then
                        entra = False
                        nSinRat = nSinRat + 1
                        AnotarExcl shX, filaX, "Sin rating", shV, r, cTipo, cEmi, cNemo, cIsin, cCod, cMon, cVto
                    ElseIf RatingExcluido(TxtSeg(shV.Cells(r, cRat).Value)) Then
                        entra = False
                        nRatExcl = nRatExcl + 1
                        AnotarExcl shX, filaX, "Rating excluido (A / A-)", shV, r, cTipo, cEmi, cNemo, cIsin, cCod, cMon, cVto
                    End If
                End If
            End If
            ' fuera los ISIN que empiezan en US (emisiones extranjeras)
            If entra And cIsin > 0 Then
                If UCase$(Left$(TxtSeg(shV.Cells(r, cIsin).Value), 2)) = "US" Then
                    entra = False
                    nIsinUS = nIsinUS + 1
                    AnotarExcl shX, filaX, "ISIN US (extranjero)", shV, r, cTipo, cEmi, cNemo, cIsin, cCod, cMon, cVto
                End If
            End If
            If entra Then
                clave = LimpiarCod(TxtSeg(shV.Cells(r, cCod).Value))
                If Len(clave) = 0 Then clave = LimpiarCod(TxtSeg(shV.Cells(r, cIsin).Value))
                If dic.Exists(clave) Or dic.Exists(LimpiarCod(TxtSeg(shV.Cells(r, cIsin).Value))) Then
                    nDup = nDup + 1
                    AnotarExcl shX, filaX, "Ya en portafolio (fila Portafolio)", shV, r, cTipo, cEmi, cNemo, cIsin, cCod, cMon, cVto
                Else
                    ws.Cells(fila, 1).Value = "NO"
                    ws.Cells(fila, 2).Value = "Universo"
                    ws.Cells(fila, 3).Value = cat
                    ws.Cells(fila, 5).Value = shV.Cells(r, cTipo).Value
                    If cEmi > 0 Then ws.Cells(fila, 6).Value = shV.Cells(r, cEmi).Value
                    If cNemo > 0 Then ws.Cells(fila, 7).Value = shV.Cells(r, cNemo).Value
                    If cIsin > 0 Then ws.Cells(fila, 8).Value = shV.Cells(r, cIsin).Value
                    ws.Cells(fila, 9).Value = LimpiarCod(TxtSeg(shV.Cells(r, cCod).Value))
                    ws.Cells(fila, 23).Value = shV.Cells(r, cCod).Value
                    If cMon > 0 Then ws.Cells(fila, 10).Value = shV.Cells(r, cMon).Value
                    If cRat > 0 Then ws.Cells(fila, 11).Value = shV.Cells(r, cRat).Value
                    ws.Cells(fila, 12).Value = dVto
                    ws.Cells(fila, 13).Value = CLng(dias)
                    ws.Cells(fila, 14).Value = dias / 365
                    If cDur > 0 Then ws.Cells(fila, 15).Value = shV.Cells(r, cDur).Value
                    If cTir > 0 Then ws.Cells(fila, 16).Value = shV.Cells(r, cTir).Value
                    If cSpr > 0 Then ws.Cells(fila, 17).Value = shV.Cells(r, cSpr).Value
                    ws.Cells(fila, 18).Value = "Vector"
                    ws.Cells(fila, 24).Value = ACNorm(TxtSeg(shV.Cells(r, cTipo).Value))
                    ws.Cells(fila, 21).Value = fFMS
                    ws.Cells(fila, 22).Value = fVec
                    nUniv = nUniv + 1
                    fila = fila + 1
                End If
            End If
        End If
SigV:
    Next r
    On Error GoTo 0

    shX.Columns("A:H").AutoFit
    shX.Columns("H").NumberFormat = "dd/mm/yyyy"
    ws.Columns("L").NumberFormat = "dd/mm/yyyy"
    ws.Columns("A:X").AutoFit
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
    CfgFila cfg, k, "Criterio universo: dias max al vcto (corp y demas)", DIAS_MAX_CORP & " (5 anos)"
    CfgFila cfg, k, "Criterio universo: soberanos sin tope de plazo (todos)", SOB_TODO_PLAZO
    CfgFila cfg, k, "Criterio universo: solo con rating (exentos soberano y CD BCRP)", True
    CfgFila cfg, k, "Criterio universo: ratings A y A- excluidos (piso: A+)", True
    CfgFila cfg, k, "Criterio universo: ISIN que empieza en US se excluye", True
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

    ' --- hoja RV: referencias, filtros y los 4 scatters ---
    ConstruirRV wbN, ws, shT, fila - 1, bmk

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
    msg = msg & "Ya estaban en el portafolio (no se duplicaron): " & nDup & vbCrLf
    msg = msg & "Excluidos: vencidos " & nVencidos & " | sin fecha vcto legible " & nSinFecha & _
          " | otra moneda " & nNoSoles & " | sin rating " & nSinRat & " | rating A/A- " & nRatExcl & " | ISIN US " & nIsinUS & vbCrLf
    msg = msg & "AUDITORIA: hoja 'Excluidos' lista cada fila del vector que no entro y su motivo." & vbCrLf
    msg = msg & "GRAFICOS: hoja 'RV' al frente con los 4 scatters y sus filtros." & vbCrLf
    msg = msg & "Para contar por clase usa la columna ASSET_CLASS_N (unifica portafolio y universo)." & vbCrLf & vbCrLf
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
    Exit Sub

FalloUniv:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Ocurrio procesando la FILA " & r & " del vector (bloque universo)." & vbCrLf & _
           "Abre el vector y revisa esa fila (busca celdas con #N/A o similar)." & vbCrLf & _
           "Pasale a Claude el numero de fila y lo que ves en ella.", vbCritical
End Sub

Private Sub AnotarExcl(shX As Worksheet, ByRef filaX As Long, motivo As String, _
    shV As Worksheet, r As Long, cTipo As Long, cEmi As Long, cNemo As Long, _
    cIsin As Long, cCod As Long, cMon As Long, cVto As Long)
    On Error Resume Next
    If filaX > 4001 Then Exit Sub
    If filaX = 4001 Then
        shX.Cells(filaX, 1).Value = "(lista truncada en 4000 filas)"
        filaX = filaX + 1
        Exit Sub
    End If
    shX.Cells(filaX, 1).Value = motivo
    If cTipo > 0 Then shX.Cells(filaX, 2).Value = TxtSeg(shV.Cells(r, cTipo).Value)
    If cEmi > 0 Then shX.Cells(filaX, 3).Value = TxtSeg(shV.Cells(r, cEmi).Value)
    If cNemo > 0 Then shX.Cells(filaX, 4).Value = TxtSeg(shV.Cells(r, cNemo).Value)
    If cIsin > 0 Then shX.Cells(filaX, 5).Value = TxtSeg(shV.Cells(r, cIsin).Value)
    shX.Cells(filaX, 6).Value = TxtSeg(shV.Cells(r, cCod).Value)
    If cMon > 0 Then shX.Cells(filaX, 7).Value = TxtSeg(shV.Cells(r, cMon).Value)
    If Not IsError(shV.Cells(r, cVto).Value) Then shX.Cells(filaX, 8).Value = shV.Cells(r, cVto).Value
    filaX = filaX + 1
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

Private Function RatingExcluido(ByVal rat As String) As Boolean
    ' Ratings que NO entran al universo: A y A- (el piso es A+).
    ' Cubre ambas escrituras: "A"/"A-" y "LP A"/"LP A-".
    Dim s As String
    s = UCase$(Trim$(rat))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    RatingExcluido = (s = "A" Or s = "LP A" Or s = "A-" Or s = "LP A-")
End Function

Private Function TxtSeg(ByVal v As Variant) As String
    ' Texto seguro: las celdas con #N/A, #REF! u otro error devuelven vacio
    If IsError(v) Then TxtSeg = "" Else TxtSeg = Trim$(CStr(v))
End Function

Private Function FechaVal(ByVal v As Variant) As Date
    ' Acepta: fecha real, texto de fecha (dd/mm/yyyy), yyyymmdd o serial de Excel como texto.
    Dim s As String
    FechaVal = 0
    If IsError(v) Then Exit Function
    If IsDate(v) Then FechaVal = CDate(v): Exit Function
    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function
    If IsNumeric(s) Then
        If Len(s) = 8 Then
            On Error Resume Next
            FechaVal = DateSerial(CLng(Left$(s, 4)), CLng(Mid$(s, 5, 2)), CLng(Right$(s, 2)))
            On Error GoTo 0
            Exit Function
        ElseIf CDbl(s) > 20000 And CDbl(s) < 80000 Then
            FechaVal = CDate(CDbl(s))
            Exit Function
        End If
    End If
    On Error Resume Next
    FechaVal = CDate(s)
    On Error GoTo 0
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

Private Function ACNorm(ByVal s0 As String) As String
    ' Traduce el asset class del FMS y el tipo de instrumento del vector
    ' a una MISMA familia, para poder filtrar/contar ambas fuentes juntas.
    Dim s As String
    s = UCase$(Trim$(s0))
    s = Replace(s, ChrW(193), "A"): s = Replace(s, ChrW(201), "E")
    s = Replace(s, ChrW(205), "I"): s = Replace(s, ChrW(211), "O")
    s = Replace(s, ChrW(218), "U")
    If Len(s) = 0 Then ACNorm = "": Exit Function
    If InStr(s, "BCRP") > 0 Then
        ACNorm = "CD BCRP"
    ElseIf InStr(s, "GOB") > 0 Or InStr(s, "SOBERAN") > 0 Or InStr(s, "TESORO") > 0 Or InStr(s, "LETRA") > 0 Then
        ACNorm = "BONO SOBERANO"
    ElseIf InStr(s, "PAP") > 0 And InStr(s, "TITUL") > 0 Then
        ACNorm = "PAPEL COM. TITULIZADO"
    ElseIf (InStr(s, "TIT") > 0 And InStr(s, "CREDIT") > 0) Or InStr(s, "DERECHO CREDIT") > 0 Then
        ACNorm = "TITULIZACION"
    ElseIf InStr(s, "PAP") > 0 Then
        ACNorm = "PAPEL COMERCIAL"
    ElseIf InStr(s, "CD") > 0 Or InStr(s, "CERTIFICADO") > 0 Then
        ACNorm = "CD SERIADO"
    ElseIf InStr(s, "DEPOSITO") > 0 Then
        ACNorm = "DEPOSITO"
    ElseIf InStr(s, "BONO") > 0 Then
        ACNorm = "BONO CORPORATIVO"
    Else
        ACNorm = "OTROS"
    End If
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


'============================================================
' HOJA RV: referencias, filtros y 4 scatters sobre Data F0
' (YTW vs Dur y Spread vs Dur; par de mercado + issuer view)
'============================================================
Private Sub ConstruirRV(wbN As Workbook, ws As Worksheet, shT As Worksheet, _
                        ByVal uf As Long, ByVal bmk As String)
    Dim wsR As Worksheet, i As Long, r As Long, g As Long, n As Long
    Dim clases(1 To 9) As String, colores(1 To 9) As Long, hayClase(1 To 9) As Boolean
    Dim dEm As Object, dRat As Object, k As Variant
    Dim filaEm As Long, filaRat As Long
    Dim tBCRP As Double, tTgt As Double, bmkV As Double
    Dim durF As Double, ytwF As Double, sprF As Double
    Dim swD As Double, swY As Double, swS As Double, sD As Double, sY As Double, sS As Double, w As Double
    Dim ch As Chart, cb As CheckBox, xPos As Double, yPos As Double, nBox As Long
    Const B1 As Long = 30      ' AD: bloque auxiliar par 1 (9 clases x 3)
    Const B2 As Long = 58      ' BF: bloque auxiliar par 2 (issuer view)
    Const CEM As Long = 86, CRAT As Long = 87, CENP As Long = 88

    If uf < 2 Then Exit Sub

    clases(1) = "BONO SOBERANO":        colores(1) = RGB(212, 12, 12)
    clases(2) = "CD BCRP":              colores(2) = RGB(0, 0, 0)
    clases(3) = "CD SERIADO":           colores(3) = RGB(191, 168, 128)
    clases(4) = "PAPEL COMERCIAL":      colores(4) = RGB(138, 138, 138)
    clases(5) = "PAPEL COM. TITULIZADO": colores(5) = RGB(110, 110, 110)
    clases(6) = "TITULIZACION":         colores(6) = RGB(166, 124, 82)
    clases(7) = "BONO CORPORATIVO":     colores(7) = RGB(89, 89, 89)
    clases(8) = "DEPOSITO":             colores(8) = RGB(59, 109, 17)
    clases(9) = "OTROS":                colores(9) = RGB(120, 120, 60)

    ' clases presentes, emisores y ratings; agregado del Fondo 0 (ponderado por monto)
    Set dEm = CreateObject("Scripting.Dictionary")
    Set dRat = CreateObject("Scripting.Dictionary")
    For r = 2 To uf
        For g = 1 To 9
            If CStr(ws.Cells(r, 24).Value) = clases(g) Then hayClase(g) = True
        Next g
        If Len(TxtSeg(ws.Cells(r, 6).Value)) > 0 Then
            If Not dEm.Exists(TxtSeg(ws.Cells(r, 6).Value)) Then dEm.Add TxtSeg(ws.Cells(r, 6).Value), 1
        End If
        If Len(TxtSeg(ws.Cells(r, 11).Value)) > 0 Then
            If Not dRat.Exists(TxtSeg(ws.Cells(r, 11).Value)) Then dRat.Add TxtSeg(ws.Cells(r, 11).Value), 1
        End If
        If CStr(ws.Cells(r, 1).Value) = "SI" Then
            w = 0
            If IsNumeric(ws.Cells(r, 19).Value) Then w = CDbl(ws.Cells(r, 19).Value)
            If w > 0 Then
                If IsNumeric(ws.Cells(r, 15).Value) And Len(TxtSeg(ws.Cells(r, 15).Value)) > 0 Then
                    swD = swD + w: sD = sD + w * CDbl(ws.Cells(r, 15).Value)
                End If
                If IsNumeric(ws.Cells(r, 16).Value) And Len(TxtSeg(ws.Cells(r, 16).Value)) > 0 Then
                    swY = swY + w: sY = sY + w * CDbl(ws.Cells(r, 16).Value)
                End If
                If IsNumeric(ws.Cells(r, 17).Value) And Len(TxtSeg(ws.Cells(r, 17).Value)) > 0 Then
                    swS = swS + w: sS = sS + w * CDbl(ws.Cells(r, 17).Value)
                End If
            End If
        End If
    Next r
    If swD > 0 Then durF = sD / swD
    If swY > 0 Then ytwF = sY / swY
    If swS > 0 Then sprF = sS / swS

    ' referencias: BCRP (ultima tasa), Target (Config F0 B4), Benchmark heredado
    r = shT.Cells(shT.Rows.Count, 2).End(xlUp).Row
    If r >= 2 And IsNumeric(shT.Cells(r, 2).Value) Then tBCRP = CDbl(shT.Cells(r, 2).Value)
    With ThisWorkbook.Worksheets(SH_CFG)
        If Len(Trim$(CStr(.Range("A4").Value))) = 0 Then
            .Range("A4").Value = "Tasa target F0 en % (CD 1 anio inicio de anio + 40pbs)"
            .Range("A4").Font.Bold = True
        End If
        If IsNumeric(.Range("B4").Value) And Len(Trim$(CStr(.Range("B4").Value))) > 0 Then tTgt = CDbl(.Range("B4").Value)
    End With
    If IsNumeric(bmk) And Len(Trim$(bmk)) > 0 Then bmkV = Val(bmk)

    ' listas para desplegables (en Data F0, columnas ocultas)
    ws.Cells(1, CEM).Value = "(All)"
    filaEm = 2
    For Each k In dEm.Keys
        ws.Cells(filaEm, CEM).Value = k
        filaEm = filaEm + 1
    Next k
    If filaEm > 3 Then ws.Range(ws.Cells(2, CEM), ws.Cells(filaEm - 1, CEM)).Sort _
        Key1:=ws.Cells(2, CEM), Order1:=xlAscending, Header:=xlNo
    ws.Cells(1, CRAT).Value = "(All)"
    filaRat = 2
    For Each k In dRat.Keys
        ws.Cells(filaRat, CRAT).Value = k
        filaRat = filaRat + 1
    Next k
    If filaRat > 3 Then ws.Range(ws.Cells(2, CRAT), ws.Cells(filaRat - 1, CRAT)).Sort _
        Key1:=ws.Cells(2, CRAT), Order1:=xlAscending, Header:=xlNo
    ws.Cells(1, CENP).Value = "(All)": ws.Cells(2, CENP).Value = "SI": ws.Cells(3, CENP).Value = "NO"

    ' hoja RV al frente: DEBE existir antes de escribir las formulas
    ' que la referencian (si no, Excel pide "Actualizar valores: RV")
    Set wsR = wbN.Worksheets.Add(Before:=ws)
    wsR.Name = "RV"
    wsR.Range("B8").Value = "(All)": wsR.Range("B9").Value = "(All)": wsR.Range("B10").Value = "(All)"
    wsR.Range("B58").Value = "(All)": wsR.Range("B59").Value = "(All)": wsR.Range("B60").Value = "(All)"

    ' bloques auxiliares (motor NA) con FillDown por velocidad
    RVX_Aux ws, clases, hayClase, B1, uf, "$B$8", "$B$9", "$B$10", True
    RVX_Aux ws, clases, hayClase, B2, uf, "$B$58", "$B$59", "$B$60", False
    For g = 1 To 9
        If hayClase(g) Then ws.Cells(1, B2 + (g - 1) * 3).Formula = "=" & ColLetra(B1 + (g - 1) * 3) & "1"
    Next g
    ws.Range(ws.Columns(B1), ws.Columns(CENP)).EntireColumn.Hidden = True

    ' contenido de la hoja RV (ya creada arriba)
    wsR.Range("A1").Value = "RELATIVE VALUE - F0 PEN - " & Format(Date, "dd/mm/yyyy")
    wsR.Range("A1").Font.Bold = True
    wsR.Range("A1").Font.Size = 12

    wsR.Range("A3:D3").Value = Array("REFERENCE", "DUR", "YTW", "SPREAD")
    With wsR.Range("A3:D3")
        .Interior.Color = RGB(212, 12, 12)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    wsR.Range("A4").Value = "BCRP ref rate": wsR.Range("B4").Value = 0
    If tBCRP > 0 Then wsR.Range("C4").Value = tBCRP
    wsR.Range("A5").Value = "Fund 0": wsR.Range("B5").Value = durF
    If ytwF > 0 Then wsR.Range("C5").Value = ytwF
    If sprF > 0 Then wsR.Range("D5").Value = sprF
    wsR.Range("A6").Value = "Target": wsR.Range("B6").Value = 1
    If tTgt > 0 Then wsR.Range("C6").Value = tTgt
    wsR.Range("D6").Value = 40
    wsR.Range("A7").Value = "Benchmark": wsR.Range("B7").Value = durF
    If bmkV > 0 Then wsR.Range("C7").Value = bmkV
    wsR.Range("B4:C7").NumberFormat = "0.00"
    wsR.Range("D4:D7").NumberFormat = "0.0"

    RVX_Control wsR, ws, "A8", "B8", "ISSUER:", "(All)", CEM, filaEm - 1
    RVX_Control wsR, ws, "A9", "B9", "RATING:", "(All)", CRAT, filaRat - 1
    RVX_Control wsR, ws, "A10", "B10", "IN PORT:", "(All)", CENP, 3

    nBox = 0
    For g = 1 To 9
        If hayClase(g) Then
            xPos = 280 + (nBox Mod 5) * 152
            yPos = wsR.Range("A8").Top + Int(nBox / 5) * 18
            Set cb = wsR.CheckBoxes.Add(xPos, yPos, 150, 15)
            cb.Caption = clases(g)
            cb.LinkedCell = "'Data F0'!" & ColLetra(B1 + (g - 1) * 3) & "1"
            cb.Value = xlOn
            nBox = nBox + 1
        End If
    Next g

    Set ch = RVX_Scatter(wsR, 10, 175, "F0 PEN - YTW vs Duration")
    RVX_Series ch, ws, clases, colores, hayClase, B1, uf, 1
    RVX_Ref ch, wsR.Range("B4"), wsR.Range("C4"), "BCRP ref rate", RGB(0, 0, 0), xlMarkerStyleSquare
    RVX_Ref ch, wsR.Range("B5"), wsR.Range("C5"), "Fund 0", RGB(212, 12, 12), xlMarkerStyleDiamond
    RVX_Ref ch, wsR.Range("B6"), wsR.Range("C6"), "Target", RGB(138, 138, 138), xlMarkerStyleTriangle
    RVX_Ref ch, wsR.Range("B7"), wsR.Range("C7"), "Benchmark", RGB(59, 109, 17), xlMarkerStyleX
    RVX_Ejes ch, "Duration (yrs)", "YTW (%)"

    Set ch = RVX_Scatter(wsR, 10, 525, "F0 PEN - Spread vs Duration")
    RVX_Series ch, ws, clases, colores, hayClase, B1, uf, 2
    RVX_Ref ch, wsR.Range("B5"), wsR.Range("D5"), "Fund 0", RGB(212, 12, 12), xlMarkerStyleDiamond
    RVX_Ref ch, wsR.Range("B6"), wsR.Range("D6"), "Target", RGB(138, 138, 138), xlMarkerStyleTriangle
    RVX_Ejes ch, "Duration (yrs)", "Spread (bps)"

    wsR.Range("A57").Value = "ISSUER VIEW"
    wsR.Range("A57").Font.Bold = True
    wsR.Range("A57").Font.Size = 12
    RVX_Control wsR, ws, "A58", "B58", "ISSUER:", CStr(ws.Cells(2, CEM).Value), CEM, filaEm - 1
    RVX_Control wsR, ws, "A59", "B59", "RATING:", "(All)", CRAT, filaRat - 1
    RVX_Control wsR, ws, "A60", "B60", "IN PORT:", "(All)", CENP, 3

    Set ch = RVX_Scatter(wsR, 10, 930, "Issuer view - YTW vs Duration")
    RVX_Series ch, ws, clases, colores, hayClase, B2, uf, 1
    RVX_Ref ch, wsR.Range("B4"), wsR.Range("C4"), "BCRP ref rate", RGB(0, 0, 0), xlMarkerStyleSquare
    RVX_Ref ch, wsR.Range("B5"), wsR.Range("C5"), "Fund 0", RGB(212, 12, 12), xlMarkerStyleDiamond
    RVX_Ref ch, wsR.Range("B6"), wsR.Range("C6"), "Target", RGB(138, 138, 138), xlMarkerStyleTriangle
    RVX_Ejes ch, "Duration (yrs)", "YTW (%)"

    Set ch = RVX_Scatter(wsR, 10, 1280, "Issuer view - Spread vs Duration")
    RVX_Series ch, ws, clases, colores, hayClase, B2, uf, 2
    RVX_Ref ch, wsR.Range("B5"), wsR.Range("D5"), "Fund 0", RGB(212, 12, 12), xlMarkerStyleDiamond
    RVX_Ref ch, wsR.Range("B6"), wsR.Range("D6"), "Target", RGB(138, 138, 138), xlMarkerStyleTriangle
    RVX_Ejes ch, "Duration (yrs)", "Spread (bps)"

    wsR.Columns("A").ColumnWidth = 16
    wsR.Columns("B:D").ColumnWidth = 11
    wsR.Activate
    wsR.Range("A1").Select
End Sub

Private Sub RVX_Aux(ws As Worksheet, ByRef clases() As String, ByRef hayClase() As Boolean, _
    ByVal base As Long, ByVal uf As Long, ByVal cIss As String, ByVal cRat As String, _
    ByVal cEnp As String, ByVal conSwitch As Boolean)
    Dim g As Long, cX As Long, f As String
    For g = 1 To 9
        If hayClase(g) Then
            cX = base + (g - 1) * 3
            If conSwitch Then ws.Cells(1, cX).Value = True
            f = "=IF(AND(" & ColLetra(cX) & "$1,$X2=""" & clases(g) & """," & _
                "OR('RV'!" & cIss & "=""(All)"",$F2='RV'!" & cIss & ")," & _
                "OR('RV'!" & cRat & "=""(All)"",$K2='RV'!" & cRat & ")," & _
                "OR('RV'!" & cEnp & "=""(All)"",$A2='RV'!" & cEnp & ")," & _
                "$O2<>"""",$P2<>""""),$O2,NA())"
            ws.Cells(2, cX).Formula = f
            ws.Cells(2, cX + 1).Formula = "=IF(ISNA(" & ColLetra(cX) & "2),NA(),$P2)"
            ws.Cells(2, cX + 2).Formula = "=IF(OR(ISNA(" & ColLetra(cX) & "2),$Q2=""""),NA(),$Q2)"
            If uf > 2 Then ws.Range(ws.Cells(2, cX), ws.Cells(uf, cX + 2)).FillDown
        End If
    Next g
End Sub

Private Sub RVX_Control(wsR As Worksheet, ws As Worksheet, ByVal cLbl As String, _
    ByVal cVal As String, ByVal etiqueta As String, ByVal valorIni As String, _
    ByVal colLista As Long, ByVal nLista As Long)
    wsR.Range(cLbl).Value = etiqueta
    wsR.Range(cLbl).Font.Bold = True
    With wsR.Range(cVal).Validation
        .Delete
        .Add Type:=xlValidateList, _
            Formula1:="='Data F0'!" & ws.Cells(1, colLista).Resize(nLista, 1).Address
    End With
    wsR.Range(cVal).Value = valorIni
    wsR.Range(cVal).Interior.Color = RGB(251, 243, 243)
End Sub

Private Function RVX_Scatter(wsR As Worksheet, ByVal x As Double, ByVal y As Double, _
    ByVal titulo As String) As Chart
    Dim ch As Chart
    Set ch = wsR.Shapes.AddChart2(240, xlXYScatter, x, y, 560, 330).Chart
    ch.HasTitle = True
    ch.ChartTitle.Text = titulo
    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop
    ch.PlotVisibleOnly = False
    Set RVX_Scatter = ch
End Function

Private Sub RVX_Series(ch As Chart, ws As Worksheet, ByRef clases() As String, _
    ByRef colores() As Long, ByRef hayClase() As Boolean, ByVal base As Long, _
    ByVal uf As Long, ByVal yOfs As Long)
    Dim g As Long, cX As Long, s As Series
    For g = 1 To 9
        If hayClase(g) Then
            cX = base + (g - 1) * 3
            Set s = ch.SeriesCollection.NewSeries
            s.Name = clases(g)
            s.XValues = ws.Range(ws.Cells(2, cX), ws.Cells(uf, cX))
            s.Values = ws.Range(ws.Cells(2, cX + yOfs), ws.Cells(uf, cX + yOfs))
            s.MarkerStyle = xlMarkerStyleCircle
            s.MarkerSize = 7
            s.MarkerBackgroundColor = colores(g)
            s.MarkerForegroundColor = RGB(255, 255, 255)
        End If
    Next g
End Sub

Private Sub RVX_Ref(ch As Chart, cX As Range, cY As Range, ByVal nom As String, _
    ByVal col As Long, ByVal estilo As Long)
    Dim s As Series
    If Not IsNumeric(cY.Value) Then Exit Sub
    If Len(Trim$(CStr(cY.Value))) = 0 Then Exit Sub
    If CDbl(cY.Value) <= 0 Then Exit Sub
    Set s = ch.SeriesCollection.NewSeries
    s.Name = nom
    s.XValues = cX
    s.Values = cY
    s.MarkerStyle = estilo
    s.MarkerSize = 10
    s.MarkerBackgroundColor = col
    s.MarkerForegroundColor = col
    s.HasDataLabels = True
    s.DataLabels.ShowSeriesName = True
    s.DataLabels.ShowValue = False
    s.DataLabels.Font.Size = 8
End Sub

Private Sub RVX_Ejes(ch As Chart, ByVal tx As String, ByVal ty As String)
    On Error Resume Next
    With ch.ChartArea
        .Format.Line.Visible = msoFalse
        .Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Color = RGB(80, 80, 80)
    End With
    ch.PlotArea.Format.Line.Visible = msoFalse
    With ch.ChartTitle
        .Font.Name = "Arial"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(34, 34, 34)
    End With
    With ch.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = ty
        .AxisTitle.Font.Size = 9
        .AxisTitle.Font.Color = RGB(120, 120, 120)
        .TickLabels.Font.Size = 9
        .TickLabels.Font.Color = RGB(120, 120, 120)
        .Format.Line.ForeColor.RGB = RGB(200, 200, 200)
        .MajorGridlines.Format.Line.ForeColor.RGB = RGB(236, 234, 230)
        .MajorGridlines.Format.Line.Weight = 0.75
    End With
    With ch.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = tx
        .AxisTitle.Font.Size = 9
        .AxisTitle.Font.Color = RGB(120, 120, 120)
        .TickLabels.Font.Size = 9
        .TickLabels.Font.Color = RGB(120, 120, 120)
        .Format.Line.ForeColor.RGB = RGB(200, 200, 200)
        .HasMajorGridlines = False
        .MinimumScale = 0
        .TickLabels.NumberFormat = "0.0"
    End With
    ch.HasLegend = True
    With ch.Legend
        .Position = xlLegendPositionBottom
        .Font.Size = 8.5
        .Format.Line.Visible = msoFalse
    End With
    On Error GoTo 0
End Sub
