' ============================================================================
' FV GRIFERÍA — MACRO EXPORTAR JSON
' ============================================================================

Option Explicit

Const SP_BASE As String = "https://fvwebmail.sharepoint.com/sites/PromocionyRelacionesInstitucionales/Documentos compartidos/PDV"
Const JSON_FILENAME As String = "fv_data.json"

Sub ExportarJSONParaApp()

    Dim wb As Workbook
    Dim wsClientes As Worksheet
    Dim wsProductos As Worksheet
    Dim wsExhibidores As Worksheet
    Dim json As String
    Dim savePath As String
    Dim fNum As Integer

    Set wb = ThisWorkbook

    ' Buscar cada hoja y avisar exactamente cual falta
    Dim listaHojas As String
    Dim sh As Worksheet
    For Each sh In wb.Sheets
        listaHojas = listaHojas & " - " & sh.Name & Chr(10)
    Next sh

    Dim errMsg As String
    errMsg = ""

    On Error Resume Next
    Set wsClientes = wb.Sheets("BASE_CLIENTES")
    If wsClientes Is Nothing Then errMsg = errMsg & "No encontrada: BASE_CLIENTES" & Chr(10)
    On Error GoTo 0

    On Error Resume Next
    Set wsProductos = wb.Sheets("BASE-PROD-FV")
    If wsProductos Is Nothing Then errMsg = errMsg & "No encontrada: BASE-PROD-FV" & Chr(10)
    On Error GoTo 0

    On Error Resume Next
    Set wsExhibidores = wb.Sheets("BASE_EXHIBIDORES")
    If wsExhibidores Is Nothing Then errMsg = errMsg & "No encontrada: BASE_EXHIBIDORES" & Chr(10)
    On Error GoTo 0

    If errMsg <> "" Then
        MsgBox "Hojas con problema:" & Chr(10) & errMsg & Chr(10) & _
               "Hojas disponibles:" & Chr(10) & listaHojas, _
               vbCritical, "FV - Diagnostico"
        Exit Sub
    End If

    Application.StatusBar = "FV: Generando JSON..."
    Application.ScreenUpdating = False

    Dim today As String
    today = Format(Now(), "YYYY-MM-DD")

    json = "{"
    json = json & Chr(10) & "  ""lastUpdated"": """ & today & """"

    ' CLIENTES
    json = json & "," & Chr(10) & "  ""clients"": ["
    Dim lastRowCl As Long
    lastRowCl = wsClientes.Cells(wsClientes.Rows.Count, 1).End(xlUp).Row
    Dim firstCl As Boolean: firstCl = True
    Dim r As Long

    For r = 3 To lastRowCl
        Dim numCl As String
        numCl = Trim(CStr(wsClientes.Cells(r, 1).Value))
        If numCl = "" Then GoTo NextClient
        Dim estadoCl As String
        estadoCl = UCase(Trim(CStr(wsClientes.Cells(r, 10).Value)))
        If estadoCl <> "ACTIVO" Then GoTo NextClient
        If Not firstCl Then json = json & ","
        firstCl = False
        json = json & Chr(10) & "    {"
        json = json & """num"":" & numCl & ","
        json = json & """razon"":""" & SanitizeJSON(CStr(wsClientes.Cells(r, 2).Value)) & ""","
        json = json & """fantasia"":""" & SanitizeJSON(CStr(wsClientes.Cells(r, 3).Value)) & ""","
        json = json & """dir"":""" & SanitizeJSON(CStr(wsClientes.Cells(r, 4).Value)) & ""","
        json = json & """zona"":""" & SanitizeJSON(CStr(wsClientes.Cells(r, 5).Value)) & ""","
        json = json & """seg"":""" & SanitizeJSON(CStr(wsClientes.Cells(r, 6).Value)) & ""","
        json = json & """vendedor"":""" & SanitizeJSON(CStr(wsClientes.Cells(r, 7).Value)) & ""","
        json = json & """estado"":""ACTIVO"""
        json = json & "}"
NextClient:
    Next r
    json = json & Chr(10) & "  ]"

    ' PRODUCTOS FV
    json = json & "," & Chr(10) & "  ""products"": ["
    Dim lastRowPr As Long
    lastRowPr = wsProductos.Cells(wsProductos.Rows.Count, 1).End(xlUp).Row
    Dim firstPr As Boolean: firstPr = True

    For r = 3 To lastRowPr
        Dim cod As String
        cod = Trim(CStr(wsProductos.Cells(r, 1).Value))
        If cod = "" Then GoTo NextProduct
        Dim costoNum As String
        costoNum = CleanNumber(CStr(wsProductos.Cells(r, 7).Value))
        If costoNum = "" Then costoNum = "0"
        If Not firstPr Then json = json & ","
        firstPr = False
        json = json & Chr(10) & "    {"
        json = json & """cod"":""" & SanitizeJSON(cod) & ""","
        json = json & """desc"":""" & SanitizeJSON(CStr(wsProductos.Cells(r, 2).Value)) & ""","
        json = json & """linea"":""" & SanitizeJSON(CStr(wsProductos.Cells(r, 3).Value)) & ""","
        json = json & """colores"":""" & SanitizeJSON(CStr(wsProductos.Cells(r, 4).Value)) & ""","
        json = json & """estado"":""" & SanitizeJSON(CStr(wsProductos.Cells(r, 5).Value)) & ""","
        json = json & """costo"":" & costoNum & ","
        json = json & """lanzamiento"":" & IIf(UCase(Trim(CStr(wsProductos.Cells(r, 8).Value))) = "SI", "true", "false") & ","
        json = json & """reemplazaPor"":""" & SanitizeJSON(CStr(wsProductos.Cells(r, 9).Value)) & ""","
        json = json & """seg"":{""HC"":" & BoolJSON(CStr(wsProductos.Cells(r, 12).Value))
        json = json & ",""DIST"":" & BoolJSON(CStr(wsProductos.Cells(r, 13).Value))
        json = json & ",""A"":" & BoolJSON(CStr(wsProductos.Cells(r, 14).Value))
        json = json & ",""B"":" & BoolJSON(CStr(wsProductos.Cells(r, 15).Value))
        json = json & ",""C"":" & BoolJSON(CStr(wsProductos.Cells(r, 16).Value))
        json = json & ",""BAJA"":" & BoolJSON(CStr(wsProductos.Cells(r, 17).Value)) & "}}"
NextProduct:
    Next r
    json = json & Chr(10) & "  ]"

    ' PRODUCTOS FRANZ
    json = json & "," & Chr(10) & "  ""productsFranz"": ["
    Dim wsFranz As Worksheet
    Dim firstFz As Boolean: firstFz = True
    On Error Resume Next
    Set wsFranz = wb.Sheets("BASE-PROD-FRANZ")
    On Error GoTo 0
    If Not wsFranz Is Nothing Then
        Dim lastRowFz As Long
        lastRowFz = wsFranz.Cells(wsFranz.Rows.Count, 1).End(xlUp).Row
        For r = 3 To lastRowFz
            Dim codFz As String
            codFz = Trim(CStr(wsFranz.Cells(r, 1).Value))
            If codFz = "" Then GoTo NextFranz
            Dim costoFz As String
            costoFz = CleanNumber(CStr(wsFranz.Cells(r, 7).Value))
            If costoFz = "" Then costoFz = "0"
            If Not firstFz Then json = json & ","
            firstFz = False
            json = json & Chr(10) & "    {"
            json = json & """cod"":""" & SanitizeJSON(codFz) & ""","
            json = json & """desc"":""" & SanitizeJSON(CStr(wsFranz.Cells(r, 2).Value)) & ""","
            json = json & """linea"":""" & SanitizeJSON(CStr(wsFranz.Cells(r, 3).Value)) & ""","
            json = json & """colores"":""" & SanitizeJSON(CStr(wsFranz.Cells(r, 4).Value)) & ""","
            json = json & """estado"":""" & SanitizeJSON(CStr(wsFranz.Cells(r, 5).Value)) & ""","
            json = json & """costo"":" & costoFz & ","
            json = json & """lanzamiento"":" & IIf(UCase(Trim(CStr(wsFranz.Cells(r, 8).Value))) = "SI", "true", "false") & ","
            json = json & """reemplazaPor"":""" & SanitizeJSON(CStr(wsFranz.Cells(r, 9).Value)) & """}"
NextFranz:
        Next r
    End If
    json = json & Chr(10) & "  ]"

    ' EXHIBIDORES
    json = json & "," & Chr(10) & "  ""exhibidores"": ["
    Dim lastRowEx As Long
    lastRowEx = wsExhibidores.Cells(wsExhibidores.Rows.Count, 1).End(xlUp).Row
    Dim firstEx As Boolean: firstEx = True

    For r = 3 To lastRowEx
        Dim codEx As String
        codEx = Trim(CStr(wsExhibidores.Cells(r, 1).Value))
        If codEx = "" Then GoTo NextExhib
        Dim costoEx As String
        costoEx = CleanNumber(CStr(wsExhibidores.Cells(r, 5).Value))
        If costoEx = "" Then costoEx = "0"
        If Not firstEx Then json = json & ","
        firstEx = False
        json = json & Chr(10) & "    {"
        json = json & """cod"":""" & SanitizeJSON(codEx) & ""","
        json = json & """desc"":""" & SanitizeJSON(CStr(wsExhibidores.Cells(r, 2).Value)) & ""","
        json = json & """tipo"":""" & SanitizeJSON(CStr(wsExhibidores.Cells(r, 3).Value)) & ""","
        json = json & """costo"":" & costoEx & ","
        json = json & """estado"":""" & SanitizeJSON(CStr(wsExhibidores.Cells(r, 6).Value)) & """}"
NextExhib:
    Next r
    json = json & Chr(10) & "  ]"

    ' PRODUCTOS SIMPLE
    json = json & "," & Chr(10) & "  ""productosSimple"": ["
    Dim wsSimple As Worksheet
    Dim firstSm As Boolean: firstSm = True
    On Error Resume Next
    Set wsSimple = wb.Sheets("BASE-PROD-SIMPLE")
    On Error GoTo 0
    If Not wsSimple Is Nothing Then
        Dim lastRowSm As Long
        lastRowSm = wsSimple.Cells(wsSimple.Rows.Count, 1).End(xlUp).Row
        For r = 2 To lastRowSm
            Dim ordenSm As String
            ordenSm = Trim(CStr(wsSimple.Cells(r, 1).Value))
            If ordenSm = "" Then GoTo NextSimple
            Dim lineaSm As String
            Dim nombreSm As String
            Dim coloresSm As String
            Dim tecnSm As String
            Dim imagenSm As String
            Dim estadoSm As String
            Dim detalleSm As String
            lineaSm  = SanitizeJSON(CStr(wsSimple.Cells(r, 2).Value))
            nombreSm = SanitizeJSON(CStr(wsSimple.Cells(r, 3).Value))
            coloresSm = SanitizeJSON(CStr(wsSimple.Cells(r, 4).Value))
            tecnSm   = SanitizeJSON(CStr(wsSimple.Cells(r, 5).Value))
            Dim imgRaw2 As String
            imgRaw2 = Trim(CStr(wsSimple.Cells(r, 6).Value))
            If imgRaw2 <> "" And imgRaw2 <> "URL de imagen SharePoint" Then
                imagenSm = "https://raw.githubusercontent.com/luisiered83-max/fv-auditoria/main/imagenes/" & imgRaw2 & ".jpg"
            Else
                imagenSm = ""
            End If
            estadoSm = SanitizeJSON(CStr(wsSimple.Cells(r, 7).Value))
            detalleSm = SanitizeJSON(CStr(wsSimple.Cells(r, 8).Value))
            Dim tipoSm As String
            tipoSm = Left(ordenSm, 1)
            If Not firstSm Then json = json & ","
            firstSm = False
            json = json & Chr(10) & "    {"
            json = json & """orden"":""" & SanitizeJSON(ordenSm) & ""","
            json = json & """tipo"":""" & tipoSm & ""","
            json = json & """linea"":""" & lineaSm & ""","
            json = json & """nombre"":""" & nombreSm & ""","
            json = json & """colores"":""" & coloresSm & ""","
            json = json & """tecnologia"":""" & tecnSm & ""","
            json = json & """imagen"":""" & imagenSm & ""","
            json = json & """estado"":""" & estadoSm & ""","
            json = json & """lanzamiento"":" & IIf(UCase(detalleSm) = "LANZAMIENTO", "true", "false") & ","
            json = json & """oportunidad"":" & IIf(UCase(detalleSm) = "OPORTUNIDAD", "true", "false") & "}"
NextSimple:
        Next r
    End If
    json = json & Chr(10) & "  ]"
    json = json & Chr(10) & "}"

    ' GUARDAR
    savePath = Environ("USERPROFILE") & "\Documents\" & JSON_FILENAME
    fNum = FreeFile()
    Open savePath For Output As #fNum
        Print #fNum, json
    Close #fNum

    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "JSON generado correctamente." & Chr(10) & Chr(10) & _
           "Guardado en:" & Chr(10) & savePath & Chr(10) & Chr(10) & _
           "Subilo a GitHub: github.com/luisiered83-max/fv-auditoria", _
           vbInformation, "FV Griferia - Exportacion JSON"

    ' Abrir carpeta sin Shell (evita error 429)
    Dim carpetaDocs As String
    carpetaDocs = Environ("USERPROFILE") & "\Documents\"
    Application.FollowHyperlink carpetaDocs

    Exit Sub

End Sub

' FUNCIONES AUXILIARES
Private Function SanitizeJSON(s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "'")
    s = Replace(s, Chr(10), " ")
    s = Replace(s, Chr(13), " ")
    SanitizeJSON = Trim(s)
End Function

Private Function CleanNumber(s As String) As String
    Dim i As Integer, clean As String
    s = Replace(Replace(Replace(s, "$", ""), ".", ""), ",", "")
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Then
            clean = clean & Mid(s, i, 1)
        End If
    Next i
    CleanNumber = clean
End Function

Private Function BoolJSON(s As String) As String
    If Len(Trim(s)) > 0 Then
        BoolJSON = "true"
    Else
        BoolJSON = "false"
    End If
End Function
