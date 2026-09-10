Attribute VB_Name = "mRV_F12"
Option Explicit

'======================================================================
' mRV_F12  -  Cuadros RV: YTW/Duración y Spread/Duración
'             Fondo 1 y Fondo 2, en PEN y USD
'             Vistas: Soberanos / Portafolio actual / Universo elegible
'
' Modulo AUTOCONTENIDO. Todos los helpers llevan prefijo RV12_ para no
' chocar con mod_MonitorDia, mMonitorF0 ni mRV_F0.
'
' Macros publicas:
'   RV_Construir   -> crea hojas RV F1 / RV F2, controles, listas y graficos
'   RV_Actualizar  -> recalcula y redibuja (lo llaman los controles)
'   RV_RecargarListas -> repuebla los desplegables de emisor y rating
'======================================================================

Private Const SH_CFG   As String = "Config"
Private Const SH_CALC  As String = "_RV_Calc"

' --- columnas de las hojas Data F1 / Data F2 ---
Private Const C_MON    As Long = 1
Private Const C_TIPO   As Long = 2
Private Const C_EMI    As Long = 3
Private Const C_NEM    As Long = 4
Private Const C_DES    As Long = 5
Private Const C_RAT    As Long = 6
Private Const C_VCTO   As Long = 7
Private Const C_DUR    As Long = 8
Private Const C_YTW    As Long = 9
Private Const C_SPR    As Long = 10
Private Const C_PORT   As Long = 11
Private Const C_PESO   As Long = 12

Private Const NTIP     As Long = 4
Private Const MAXLAB   As Long = 80      ' tope de puntos para mostrar etiquetas

'======================================================================
'  MACROS PUBLICAS
'======================================================================
Public Sub RV_Construir()
    Dim f As Long
    RV12_Speed True
    On Error GoTo fin

    RV12_HojaCalc
    For f = 1 To 2
        RV12_PrepDash f
        RV12_Controles f
        RV12_CargarListas f
    Next f
    For f = 1 To 2
        RV12_Calcular f
        RV12_Graficos f
    Next f

fin:
    RV12_Speed False
    If Err.Number <> 0 Then
        MsgBox "RV_Construir: " & Err.Description, vbExclamation, "RV F1/F2"
    Else
        MsgBox "Dashboards RV F1 y RV F2 generados.", vbInformation, "RV F1/F2"
    End If
End Sub

Public Sub RV_Actualizar()
    Dim f As Long
    RV12_Speed True
    On Error GoTo fin
    RV12_HojaCalc
    For f = 1 To 2
        RV12_Calcular f
        RV12_Graficos f
    Next f
fin:
    RV12_Speed False
    If Err.Number <> 0 Then MsgBox "RV_Actualizar: " & Err.Description, vbExclamation, "RV F1/F2"
End Sub

Public Sub RV_RecargarListas()
    Dim f As Long
    RV12_Speed True
    On Error GoTo fin
    For f = 1 To 2
        RV12_CargarListas f
    Next f
    RV_Actualizar
fin:
    RV12_Speed False
End Sub

'======================================================================
'  HOJAS Y LAYOUT
'======================================================================
Private Sub RV12_Speed(ByVal onOff As Boolean)
    Application.ScreenUpdating = Not onOff
    Application.EnableEvents = Not onOff
    Application.DisplayAlerts = Not onOff
    Application.Calculation = IIf(onOff, xlCalculationManual, xlCalculationAutomatic)
End Sub

Private Function RV12_Sheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:= _
                 ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nm
    End If
    Set RV12_Sheet = ws
End Function

Private Function RV12_Dash(ByVal f As Long) As Worksheet
    Set RV12_Dash = RV12_Sheet("RV F" & f)
End Function

Private Function RV12_Data(ByVal f As Long) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Data F" & f)
    On Error GoTo 0
    If ws Is Nothing Then Err.Raise 5, , "No existe la hoja 'Data F" & f & "'."
    Set RV12_Data = ws
End Function

Private Sub RV12_HojaCalc()
    Dim ws As Worksheet
    Set ws = RV12_Sheet(SH_CALC)
    ws.Visible = xlSheetHidden
End Sub

Private Sub RV12_PrepDash(ByVal f As Long)
    Dim ws As Worksheet
    Set ws = RV12_Dash(f)
    ws.Cells.Clear
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 8
    ws.Columns("A").ColumnWidth = 12
    ws.Cells.Interior.ColorIndex = xlNone

    ' banda institucional del encabezado y panel de filtros
    ws.Range("A1:P2").Interior.Color = RGB(212, 12, 12)
    ws.Range("A3:P8").Interior.Color = RGB(242, 242, 242)
    With ws.Range("A8:P8").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(191, 191, 191)
        .Weight = xlThin
    End With
End Sub

'======================================================================
'  CONTROLES (checkboxes + desplegables, estilo Dashboard CP)
'======================================================================
Private Sub RV12_Controles(ByVal f As Long)
    Dim ws As Worksheet
    Set ws = RV12_Dash(f)

    On Error Resume Next
    ws.CheckBoxes.Delete
    ws.DropDowns.Delete
    ws.Labels.Delete
    ws.Buttons.Delete
    On Error GoTo 0

    ' --- titulo ---
    With ws.Labels.Add(14, 6, 700, 20)
        .Caption = "FONDO " & f & "  -  Relative Value Perú  |  YTW/Duración y Spread/Duración  |  PEN y USD"
        .Font.Name = "Arial": .Font.Size = 11: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
    End With

    ' --- fila 1 de filtros: Vista / Emisor / Rating ---
    RV12_Lbl ws, "Vista:", 14, 38, 55
    With ws.DropDowns.Add(72, 36, 175, 18)
        .Name = "ddVista"
        .RemoveAllItems
        .AddItem "Universo elegible"
        .AddItem "Soberanos"
        .AddItem "Portafolio actual"
        .ListIndex = 1
        .OnAction = "RV_Actualizar"
    End With

    RV12_Lbl ws, "Emisor:", 265, 38, 55
    With ws.DropDowns.Add(320, 36, 210, 18)
        .Name = "ddEmisor"
        .OnAction = "RV_Actualizar"
    End With

    RV12_Lbl ws, "Rating:", 550, 38, 55
    With ws.DropDowns.Add(605, 36, 130, 18)
        .Name = "ddRating"
        .OnAction = "RV_Actualizar"
    End With

    ' --- fila 2 de filtros: tipos + etiquetas ---
    RV12_Lbl ws, "Tipo:", 14, 66, 55
    Dim tipos As Variant, i As Long
    tipos = Array("SOB", "CUASI", "CORP", "DP")
    For i = 0 To 3
        With ws.CheckBoxes.Add(72 + i * 78, 64, 74, 18)
            .Name = "cb" & tipos(i)
            .Caption = tipos(i)
            .Value = xlOn
            .OnAction = "RV_Actualizar"
            .Font.Name = "Arial": .Font.Size = 8
        End With
    Next i

    With ws.CheckBoxes.Add(400, 64, 110, 18)
        .Name = "cbEtiq"
        .Caption = "Etiquetas"
        .Value = xlOff
        .OnAction = "RV_Actualizar"
        .Font.Name = "Arial": .Font.Size = 8
    End With

    With ws.Buttons.Add(605, 63, 130, 20)
        .Name = "btnRV"
        .Caption = "Actualizar"
        .OnAction = "RV_Actualizar"
        .Font.Name = "Arial": .Font.Size = 8: .Font.Bold = True
    End With

    ' --- nota de referencias ---
    With ws.Labels.Add(14, 92, 900, 16)
        .Caption = "Referencias fijas (no dependen de los filtros):  tasa BCRP / SOFR  ·  portafolio (promedio ponderado por peso)  ·  benchmark del fondo (hoja Config)."
        .Font.Name = "Arial": .Font.Size = 8: .Font.Italic = True
    End With
End Sub

Private Sub RV12_Lbl(ws As Worksheet, ByVal txt As String, ByVal L As Double, _
                     ByVal T As Double, ByVal W As Double)
    With ws.Labels.Add(L, T, W, 16)
        .Caption = txt
        .Font.Name = "Arial": .Font.Size = 8: .Font.Bold = True
    End With
End Sub

Private Sub RV12_CargarListas(ByVal f As Long)
    Dim ws As Worksheet, wd As Worksheet
    Set ws = RV12_Dash(f): Set wd = RV12_Data(f)

    Dim emis As Variant, rats As Variant
    emis = RV12_Unicos(wd, C_EMI)
    rats = RV12_Unicos(wd, C_RAT)

    RV12_LlenaDD ws, "ddEmisor", emis
    RV12_LlenaDD ws, "ddRating", rats
End Sub

Private Sub RV12_LlenaDD(ws As Worksheet, ByVal nm As String, ByVal arr As Variant)
    Dim dd As DropDown, i As Long
    On Error Resume Next
    Set dd = ws.DropDowns(nm)
    On Error GoTo 0
    If dd Is Nothing Then Exit Sub
    dd.RemoveAllItems
    dd.AddItem "(Todos)"
    If IsArray(arr) Then
        For i = LBound(arr) To UBound(arr)
            If Len(arr(i)) > 0 Then dd.AddItem CStr(arr(i))
        Next i
    End If
    dd.ListIndex = 1
End Sub

Private Function RV12_Unicos(wd As Worksheet, ByVal col As Long) As Variant
    Dim n As Long, i As Long, j As Long, s As String
    Dim col2 As Collection, res() As String, tmp As String
    Set col2 = New Collection
    n = wd.Cells(wd.Rows.Count, C_MON).End(xlUp).Row
    For i = 2 To n
        s = Trim$(CStr(wd.Cells(i, col).Value))
        If Len(s) > 0 Then
            On Error Resume Next
            col2.Add s, UCase$(s)
            On Error GoTo 0
        End If
    Next i
    If col2.Count = 0 Then
        RV12_Unicos = Array()
        Exit Function
    End If
    ReDim res(0 To col2.Count - 1)
    For i = 1 To col2.Count
        res(i - 1) = col2(i)
    Next i
    For i = LBound(res) To UBound(res) - 1
        For j = i + 1 To UBound(res)
            If UCase$(res(j)) < UCase$(res(i)) Then
                tmp = res(i): res(i) = res(j): res(j) = tmp
            End If
        Next j
    Next i
    RV12_Unicos = res
End Function

'======================================================================
'  CALCULO: arma las series en la hoja _RV_Calc
'======================================================================
Private Function RV12_Col0(ByVal f As Long) As Long
    RV12_Col0 = IIf(f = 1, 1, 60)
End Function

Private Function RV12_Base(ByVal f As Long, ByVal mi As Long, ByVal ti As Long) As Long
    RV12_Base = RV12_Col0(f) + ((mi - 1) * NTIP + (ti - 1)) * 4
End Function

Private Function RV12_RefBase(ByVal f As Long, ByVal mi As Long, ByVal k As Long) As Long
    RV12_RefBase = RV12_Col0(f) + 32 + ((mi - 1) * 3 + (k - 1)) * 3
End Function

Private Sub RV12_Calcular(ByVal f As Long)
    Dim wd As Worksheet, wc As Worksheet, ws As Worksheet, cfg As Worksheet
    Set wd = RV12_Data(f)
    Set wc = ThisWorkbook.Worksheets(SH_CALC)
    Set ws = RV12_Dash(f)
    Set cfg = ThisWorkbook.Worksheets(SH_CFG)

    Dim vista As String, emi As String, rat As String
    vista = RV12_Combo(ws, "ddVista")
    emi = RV12_Combo(ws, "ddEmisor")
    rat = RV12_Combo(ws, "ddRating")

    Dim onT(1 To 4) As Boolean
    onT(1) = RV12_Chk(ws, "cbSOB")
    onT(2) = RV12_Chk(ws, "cbCUASI")
    onT(3) = RV12_Chk(ws, "cbCORP")
    onT(4) = RV12_Chk(ws, "cbDP")

    ' limpiar bloque del fondo
    Dim c0 As Long
    c0 = RV12_Col0(f)
    wc.Range(wc.Cells(1, c0), wc.Cells(5000, c0 + 49)).ClearContents

    Dim mi As Long, ti As Long, b As Long
    For mi = 1 To 2
        For ti = 1 To NTIP
            b = RV12_Base(f, mi, ti)
            wc.Cells(1, b).Value = RV12_Mon(mi) & "_" & RV12_Tipo(ti) & "_DUR"
            wc.Cells(1, b + 1).Value = "YTW"
            wc.Cells(1, b + 2).Value = "SPR"
            wc.Cells(1, b + 3).Value = "NEM"
        Next ti
    Next mi

    Dim cnt(1 To 2, 1 To 4) As Long
    Dim sw(1 To 2) As Double, sd(1 To 2) As Double
    Dim sy(1 To 2) As Double, ss(1 To 2) As Double

    Dim n As Long, i As Long
    Dim sMon As String, sTipo As String, pw As Double
    n = wd.Cells(wd.Rows.Count, C_MON).End(xlUp).Row

    For i = 2 To n
        sMon = UCase$(Trim$(CStr(wd.Cells(i, C_MON).Value)))
        sTipo = UCase$(Trim$(CStr(wd.Cells(i, C_TIPO).Value)))
        mi = 0
        If sMon = "PEN" Then mi = 1
        If sMon = "USD" Then mi = 2
        ti = RV12_TipoIdx(sTipo)

        If mi > 0 And ti > 0 Then

            ' --- punto de portafolio: siempre, sin filtros ---
            If RV12_EsSi(wd.Cells(i, C_PORT).Value) Then
                pw = RV12_Num(wd.Cells(i, C_PESO).Value)
                If pw > 0 Then
                    sw(mi) = sw(mi) + pw
                    sd(mi) = sd(mi) + pw * RV12_Num(wd.Cells(i, C_DUR).Value)
                    sy(mi) = sy(mi) + pw * RV12_Num(wd.Cells(i, C_YTW).Value)
                    ss(mi) = ss(mi) + pw * RV12_Num(wd.Cells(i, C_SPR).Value)
                End If
            End If

            ' --- filtros del dashboard ---
            If onT(ti) Then
                If RV12_PasaVista(vista, wd, i) Then
                    If RV12_Coincide(emi, wd.Cells(i, C_EMI).Value) Then
                        If RV12_Coincide(rat, wd.Cells(i, C_RAT).Value) Then
                            b = RV12_Base(f, mi, ti)
                            cnt(mi, ti) = cnt(mi, ti) + 1
                            wc.Cells(1 + cnt(mi, ti), b).Value = RV12_Num(wd.Cells(i, C_DUR).Value)
                            wc.Cells(1 + cnt(mi, ti), b + 1).Value = RV12_Num(wd.Cells(i, C_YTW).Value)
                            wc.Cells(1 + cnt(mi, ti), b + 2).Value = RV12_Num(wd.Cells(i, C_SPR).Value)
                            wc.Cells(1 + cnt(mi, ti), b + 3).Value = CStr(wd.Cells(i, C_NEM).Value)
                        End If
                    End If
                End If
            End If
        End If
    Next i

    ' --- puntos de referencia ---
    Dim rb As Long, filaCfg As Long
    For mi = 1 To 2
        ' 1) tasa de referencia (BCRP / SOFR): duracion 0, spread 0
        rb = RV12_RefBase(f, mi, 1)
        wc.Cells(1, rb).Value = "REF_TASA_" & RV12_Mon(mi)
        wc.Cells(2, rb).Value = 0
        wc.Cells(2, rb + 1).Value = RV12_Num(cfg.Range(IIf(mi = 1, "B5", "B6")).Value)
        wc.Cells(2, rb + 2).Value = 0

        ' 2) portafolio actual (promedio ponderado)
        rb = RV12_RefBase(f, mi, 2)
        wc.Cells(1, rb).Value = "REF_PORT_" & RV12_Mon(mi)
        If sw(mi) > 0 Then
            wc.Cells(2, rb).Value = sd(mi) / sw(mi)
            wc.Cells(2, rb + 1).Value = sy(mi) / sw(mi)
            wc.Cells(2, rb + 2).Value = ss(mi) / sw(mi)
        End If

        ' 3) benchmark del fondo (Config filas 12..15)
        rb = RV12_RefBase(f, mi, 3)
        wc.Cells(1, rb).Value = "REF_BMK_" & RV12_Mon(mi)
        filaCfg = 11 + (f - 1) * 2 + mi          ' F1 PEN=12, F1 USD=13, F2 PEN=14, F2 USD=15
        wc.Cells(2, rb).Value = RV12_Num(cfg.Cells(filaCfg, 3).Value)
        wc.Cells(2, rb + 1).Value = RV12_Num(cfg.Cells(filaCfg, 4).Value)
        wc.Cells(2, rb + 2).Value = RV12_Num(cfg.Cells(filaCfg, 5).Value)
    Next mi
End Sub

'======================================================================
'  GRAFICOS
'======================================================================
Private Sub RV12_Graficos(ByVal f As Long)
    Dim ws As Worksheet, co As ChartObject
    Set ws = RV12_Dash(f)
    For Each co In ws.ChartObjects
        co.Delete
    Next co

    Dim mi As Long, met As Long
    For mi = 1 To 2
        For met = 1 To 2
            RV12_UnGrafico ws, f, mi, met
        Next met
    Next mi
End Sub

Private Sub RV12_UnGrafico(ws As Worksheet, ByVal f As Long, ByVal mi As Long, ByVal met As Long)
    Dim wc As Worksheet, co As ChartObject, ch As Chart, sr As Series
    Dim ti As Long, b As Long, n As Long, colY As Long, tot As Long, p As Long
    Dim bLab As Boolean

    Set wc = ThisWorkbook.Worksheets(SH_CALC)
    colY = IIf(met = 1, 1, 2)

    tot = 0
    For ti = 1 To NTIP
        tot = tot + RV12_NFilas(wc, RV12_Base(f, mi, ti))
    Next ti
    bLab = RV12_Chk(ws, "cbEtiq", False) And (tot > 0) And (tot <= MAXLAB)

    Set co = ws.ChartObjects.Add(15 + (mi - 1) * 455, 125 + (met - 1) * 270, 445, 262)
    co.Name = "chRV_" & mi & "_" & met
    Set ch = co.Chart
    ch.ChartType = xlXYScatter
    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop

    ' --- series por tipo ---
    Dim srSOB As Series, nSOB As Long
    For ti = 1 To NTIP
        b = RV12_Base(f, mi, ti)
        n = RV12_NFilas(wc, b)
        If n > 0 Then
            Set sr = ch.SeriesCollection.NewSeries
            If ti = 1 Then Set srSOB = sr: nSOB = n
            sr.Name = RV12_Tipo(ti)
            sr.XValues = wc.Range(wc.Cells(2, b), wc.Cells(1 + n, b))
            sr.Values = wc.Range(wc.Cells(2, b + colY), wc.Cells(1 + n, b + colY))
            sr.MarkerStyle = RV12_Mk(ti)
            sr.MarkerSize = 7
            sr.MarkerBackgroundColor = RV12_Color(ti)
            sr.MarkerForegroundColor = RGB(255, 255, 255)
            If bLab Then
                On Error Resume Next
                sr.HasDataLabels = True
                For p = 1 To n
                    sr.Points(p).DataLabel.Text = CStr(wc.Cells(1 + p, b + 3).Value)
                    sr.Points(p).DataLabel.Font.Size = 6
                    sr.Points(p).DataLabel.Font.Name = "Arial"
                    sr.Points(p).DataLabel.Position = xlLabelPositionRight
                Next p
                On Error GoTo 0
            End If
        End If
    Next ti

    ' --- curva soberana ajustada (solo YTW): es la linea contra la que se lee el RV ---
    If met = 1 And nSOB >= 4 Then
        On Error Resume Next
        With srSOB.Trendlines.Add(Type:=xlPolynomial, Order:=2)
            .Name = "Curva soberana ajustada"
            .DisplayEquation = False
            .DisplayRSquared = False
            .Format.Line.ForeColor.RGB = RGB(212, 12, 12)
            .Format.Line.Weight = 1
            .Format.Line.DashStyle = msoLineDash
        End With
        On Error GoTo 0
    End If

    ' --- referencias ---
    RV12_SerieRef ch, wc, f, mi, 1, colY, RV12_EtiqTasa(mi), RGB(0, 0, 0), _
                  xlMarkerStyleX, 9
    RV12_SerieRef ch, wc, f, mi, 2, colY, "Portafolio", RGB(0, 0, 0), _
                  xlMarkerStyleDiamond, 11
    RV12_SerieRef ch, wc, f, mi, 3, colY, "Benchmark F" & f, RGB(255, 255, 255), _
                  xlMarkerStyleCircle, 11

    ' --- formato institucional ---
    On Error Resume Next
    With ch
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .ChartArea.Format.Line.Visible = msoFalse
        .ChartArea.Font.Name = "Arial"
        .ChartArea.Font.Size = 8
        .ChartArea.Font.Color = RGB(0, 0, 0)
        .PlotArea.Format.Fill.Visible = msoFalse
        .PlotArea.Format.Line.Visible = msoFalse

        .HasTitle = True
        Dim sT1 As String, sT2 As String
        sT1 = "FONDO " & f & "  |  " & RV12_Mon(mi) & "  |  " & _
              IIf(met = 1, "YTW", "Spread") & " vs Duración"
        sT2 = RV12_Subtitulo()
        .ChartTitle.Text = sT1 & Chr(10) & sT2
        .ChartTitle.Font.Name = "Arial"
        .ChartTitle.Characters(1, Len(sT1)).Font.Size = 9
        .ChartTitle.Characters(1, Len(sT1)).Font.Bold = True
        .ChartTitle.Characters(1, Len(sT1)).Font.Color = RGB(212, 12, 12)
        .ChartTitle.Characters(Len(sT1) + 2, Len(sT2)).Font.Size = 7
        .ChartTitle.Characters(Len(sT1) + 2, Len(sT2)).Font.Bold = False
        .ChartTitle.Characters(Len(sT1) + 2, Len(sT2)).Font.Color = RGB(89, 89, 89)
        .ChartTitle.Format.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
        .ChartTitle.Left = 6
        .ChartTitle.Top = 4

        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "Duración modificada (años)"
            .AxisTitle.Font.Size = 8
            .AxisTitle.Font.Bold = False
            .MinimumScale = 0
            .TickLabels.NumberFormat = "0.0"
            .TickLabels.Font.Size = 8
            .Format.Line.ForeColor.RGB = RGB(128, 128, 128)
            .Format.Line.Weight = 0.75
            .MajorTickMark = xlTickMarkOutside
            .MinorTickMark = xlTickMarkNone
            .HasMajorGridlines = False
        End With

        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = IIf(met = 1, "YTW (%)", "Spread (pbs)")
            .AxisTitle.Font.Size = 8
            .AxisTitle.Font.Bold = False
            .TickLabels.NumberFormat = IIf(met = 1, "0.0", "#,##0")
            .TickLabels.Font.Size = 8
            .HasMajorGridlines = True
            .MajorGridlines.Format.Line.ForeColor.RGB = RGB(217, 217, 217)
            .MajorGridlines.Format.Line.Weight = 0.5
            .Format.Line.Visible = msoFalse
            .MajorTickMark = xlTickMarkNone
        End With

        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 8
        .Legend.Format.Line.Visible = msoFalse
    End With

    ' marco gris claro del objeto grafico
    co.Format.Line.Visible = msoTrue
    co.Format.Line.ForeColor.RGB = RGB(191, 191, 191)
    co.Format.Line.Weight = 0.75
    On Error GoTo 0
End Sub

Private Sub RV12_SerieRef(ch As Chart, wc As Worksheet, ByVal f As Long, ByVal mi As Long, _
                          ByVal k As Long, ByVal colY As Long, ByVal nm As String, _
                          ByVal clr As Long, ByVal mk As Long, ByVal sz As Long)
    Dim rb As Long, sr As Series
    rb = RV12_RefBase(f, mi, k)
    If Len(Trim$(CStr(wc.Cells(2, rb).Value))) = 0 Then Exit Sub
    If Len(Trim$(CStr(wc.Cells(2, rb + colY).Value))) = 0 Then Exit Sub

    Set sr = ch.SeriesCollection.NewSeries
    sr.Name = nm
    sr.XValues = wc.Range(wc.Cells(2, rb), wc.Cells(2, rb))
    sr.Values = wc.Range(wc.Cells(2, rb + colY), wc.Cells(2, rb + colY))
    sr.MarkerStyle = mk
    sr.MarkerSize = sz
    sr.MarkerBackgroundColor = clr
    sr.MarkerForegroundColor = RGB(0, 0, 0)
End Sub

'======================================================================
'  HELPERS
'======================================================================
Private Function RV12_NFilas(wc As Worksheet, ByVal b As Long) As Long
    Dim r As Long
    r = wc.Cells(wc.Rows.Count, b).End(xlUp).Row
    If r < 2 Then
        RV12_NFilas = 0
    Else
        RV12_NFilas = r - 1
    End If
End Function

Private Function RV12_Mon(ByVal mi As Long) As String
    RV12_Mon = IIf(mi = 1, "PEN", "USD")
End Function

Private Function RV12_Subtitulo() As String
    Dim cfg As Worksheet, sF As String, sFu As String
    Set cfg = ThisWorkbook.Worksheets(SH_CFG)
    On Error Resume Next
    If IsDate(cfg.Range("B3").Value) Then sF = Format$(cfg.Range("B3").Value, "dd/mm/yyyy")
    sFu = Trim$(CStr(cfg.Range("B9").Value))
    On Error GoTo 0
    RV12_Subtitulo = "Cifras al " & IIf(Len(sF) = 0, "-", sF)
    If Len(sFu) > 0 Then RV12_Subtitulo = RV12_Subtitulo & "   ·   Fuente: " & sFu
End Function

Private Function RV12_EtiqTasa(ByVal mi As Long) As String
    Dim cfg As Worksheet, s As String
    Set cfg = ThisWorkbook.Worksheets(SH_CFG)
    s = Trim$(CStr(cfg.Range(IIf(mi = 1, "B7", "B8")).Value))
    If Len(s) = 0 Then s = IIf(mi = 1, "BCRP", "SOFR")
    RV12_EtiqTasa = "Tasa " & s
End Function

Private Function RV12_Tipo(ByVal ti As Long) As String
    Select Case ti
        Case 1: RV12_Tipo = "SOB"
        Case 2: RV12_Tipo = "CUASI"
        Case 3: RV12_Tipo = "CORP"
        Case 4: RV12_Tipo = "DP"
    End Select
End Function

Private Function RV12_TipoIdx(ByVal s As String) As Long
    Select Case UCase$(Trim$(s))
        Case "SOB", "SOBERANO": RV12_TipoIdx = 1
        Case "CUASI", "CUASISOBERANO": RV12_TipoIdx = 2
        Case "CORP", "CORPORATIVO": RV12_TipoIdx = 3
        Case "DP", "DEUDA PRIVADA": RV12_TipoIdx = 4
        Case Else: RV12_TipoIdx = 0
    End Select
End Function

Private Function RV12_Mk(ByVal ti As Long) As Long
    Select Case ti
        Case 1: RV12_Mk = xlMarkerStyleCircle
        Case 2: RV12_Mk = xlMarkerStyleSquare
        Case 3: RV12_Mk = xlMarkerStyleTriangle
        Case 4: RV12_Mk = xlMarkerStyleDiamond
    End Select
End Function

Private Function RV12_Color(ByVal ti As Long) As Long
    Select Case ti
        Case 1: RV12_Color = RGB(212, 12, 12)      ' soberano      - rojo institucional
        Case 2: RV12_Color = RGB(31, 78, 121)      ' cuasisoberano - azul institucional
        Case 3: RV12_Color = RGB(89, 89, 89)       ' corporativo   - gris oscuro
        Case 4: RV12_Color = RGB(191, 143, 0)      ' deuda privada - ocre
    End Select
End Function

Private Function RV12_Num(ByVal v As Variant) As Double
    On Error Resume Next
    If IsNumeric(v) Then RV12_Num = CDbl(v) Else RV12_Num = 0
    On Error GoTo 0
End Function

Private Function RV12_EsSi(ByVal v As Variant) As Boolean
    Dim s As String
    s = UCase$(Trim$(CStr(v)))
    RV12_EsSi = (s = "SI" Or s = ChrW$(83) & ChrW$(205) Or s = "S" Or s = "YES" Or s = "Y" Or s = "TRUE" Or s = "1")
End Function

Private Function RV12_Coincide(ByVal filtro As String, ByVal v As Variant) As Boolean
    If Len(Trim$(filtro)) = 0 Then RV12_Coincide = True: Exit Function
    If filtro = "(Todos)" Then RV12_Coincide = True: Exit Function
    RV12_Coincide = (UCase$(Trim$(CStr(v))) = UCase$(Trim$(filtro)))
End Function

Private Function RV12_PasaVista(ByVal vista As String, wd As Worksheet, ByVal i As Long) As Boolean
    Select Case UCase$(Trim$(vista))
        Case "SOBERANOS"
            RV12_PasaVista = (RV12_TipoIdx(CStr(wd.Cells(i, C_TIPO).Value)) = 1)
        Case "PORTAFOLIO ACTUAL"
            RV12_PasaVista = RV12_EsSi(wd.Cells(i, C_PORT).Value)
        Case "UNIVERSO ELEGIBLE"
            RV12_PasaVista = True          ' todo lo cargado en la hoja es el universo
        Case Else
            RV12_PasaVista = True
    End Select
End Function

Private Function RV12_Chk(ws As Worksheet, ByVal nm As String, _
                          Optional ByVal dflt As Boolean = True) As Boolean
    Dim v As Long
    RV12_Chk = dflt
    Err.Clear
    On Error Resume Next
    v = ws.CheckBoxes(nm).Value
    If Err.Number = 0 Then RV12_Chk = (v = xlOn)
    Err.Clear
    On Error GoTo 0
End Function

Private Function RV12_Combo(ws As Worksheet, ByVal nm As String) As String
    Dim dd As DropDown, k As Long
    RV12_Combo = "(Todos)"
    On Error Resume Next
    Set dd = ws.DropDowns(nm)
    On Error GoTo 0
    If dd Is Nothing Then Exit Function
    k = dd.ListIndex
    If k > 0 Then RV12_Combo = CStr(dd.List(k))
End Function
