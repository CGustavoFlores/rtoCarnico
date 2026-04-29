Attribute VB_Name = "ModGFlowRC"
'==============================================================================
' ModGFlowRC.bas — Módulo de comunicación con el servicio GFlowRC
'
' Reemplaza el uso de C:\WsRe\VALIDAR.BAT y la lectura de OUTPUT.TXT
' por llamadas HTTP directas al servicio Node local (puerto 3001).
'
' Requiere: WinHttp.WinHttpRequest (disponible en Windows sin instalar nada)
'==============================================================================

Option Explicit

Private Const GFLOWRC_URL As String = "http://127.0.0.1:3001"

' ─── Estructura del resultado ─────────────────────────────────────────────────

Public Type TResultadoRemito
    Ok               As Boolean
    YaExistia        As Boolean  ' True = ARCA ya tenía este remito (evitó duplicado)
    Aviso            As String   ' Mensaje de advertencia si YaExistia = True
    CodRemito        As String
    CRE              As String
    QRUrl            As String
    QRPath           As String   ' Path local del PNG descargado
    Estado           As String
    MensajeError     As String
End Type

' ─── Verificar que el servicio está corriendo ─────────────────────────────────

Public Function GFlowRC_Disponible() As Boolean
On Error GoTo PROCERROR
    Dim http As Object
    Dim respuesta As String
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", GFLOWRC_URL & "/health", False
    http.SetTimeouts 3000, 3000, 3000, 3000  ' 3 segundos para health check
    http.Send
    
    If http.Status = 200 Then
        GFlowRC_Disponible = True
    Else
        GFlowRC_Disponible = False
    End If
    
    Set http = Nothing
    Exit Function
PROCERROR:
    GFlowRC_Disponible = False
End Function

' ─── Generar Remito Electrónico Cárnico ───────────────────────────────────────

Public Function GFlowRC_GenerarRemito(xNroAuxRemito As Long) As TResultadoRemito
On Error GoTo PROCERROR

    Dim resultado As TResultadoRemito
    resultado.Ok = False
    
    ' Verificar que el servicio está disponible
    If Not GFlowRC_Disponible() Then
        resultado.MensajeError = "El servicio GFlowRC no está disponible." & vbCrLf & _
                                 "Verificá que el servicio está iniciado en Servicios de Windows (GFlowRC)."
        GFlowRC_GenerarRemito = resultado
        Exit Function
    End If
    
    ' Cargar datos del remito desde SQL
    Dim datos As Object
    Set datos = GFlowRC_CargarDatosRemito(xNroAuxRemito)
    If datos Is Nothing Then
        resultado.MensajeError = "No se encontraron datos del remito N° " & xNroAuxRemito
        GFlowRC_GenerarRemito = resultado
        Exit Function
    End If
    
    ' Construir JSON del body
    Dim jsonBody As String
    jsonBody = GFlowRC_ConstruirJSON(datos, xNroAuxRemito)
    
    ' Llamar al servicio
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", GFLOWRC_URL & "/remito/generar", False
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetTimeouts 5000, 5000, 90000, 90000  ' 90 seg para la llamada a ARCA
    http.Send jsonBody
    
    ' Parsear respuesta
    Dim jsonResp As String
    jsonResp = http.ResponseText
    Set http = Nothing
    
    resultado = GFlowRC_ParsearResultado(jsonResp)
    
    ' Si hubo duplicado, mostrar aviso especial
    If resultado.Ok And resultado.YaExistia Then
        MsgBox "ATENCIÓN: " & vbCrLf & vbCrLf & _
               resultado.Aviso & vbCrLf & vbCrLf & _
               "CRE: " & resultado.CRE & vbCrLf & _
               "Código Remito: " & resultado.CodRemito, _
               vbExclamation, "Remito ya existente en ARCA"
    End If
    
    GFlowRC_GenerarRemito = resultado
    Exit Function
    
PROCERROR:
    resultado.MensajeError = "Error al comunicarse con GFlowRC: " & Err.Description
    GFlowRC_GenerarRemito = resultado
End Function

' ─── Consultar último remito emitido ──────────────────────────────────────────
' Si generarQR = True, genera el archivo PNG del QR localmente

Public Function GFlowRC_ConsultarUltimo(Optional generarQR As Boolean = False) As TResultadoRemito
On Error GoTo PROCERROR

    Dim resultado As TResultadoRemito
    resultado.Ok = False
    
    If Not GFlowRC_Disponible() Then
        resultado.MensajeError = "El servicio GFlowRC no está disponible."
        GFlowRC_ConsultarUltimo = resultado
        Exit Function
    End If
    
    Dim url As String
    url = GFLOWRC_URL & "/remito/ultimo"
    If generarQR Then url = url & "?generarQR=1"
    
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.SetTimeouts 5000, 5000, 30000, 30000
    http.Send
    
    Dim jsonResp As String
    jsonResp = http.ResponseText
    Set http = Nothing
    
    ' Extraer datos del último remito
    resultado.Ok        = (GFlowRC_ExtraerCampo(jsonResp, "ok") = "true")
    resultado.CodRemito = GFlowRC_ExtraerCampo(jsonResp, "codRemito")
    resultado.CRE       = GFlowRC_ExtraerCampo(jsonResp, "cre")
    resultado.Estado    = GFlowRC_ExtraerCampo(jsonResp, "estado")
    resultado.QRUrl     = GFlowRC_ExtraerCampo(jsonResp, "qrUrl")
    resultado.QRPath    = GFlowRC_ExtraerCampo(jsonResp, "qrPath")
    
    GFlowRC_ConsultarUltimo = resultado
    Exit Function
    
PROCERROR:
    resultado.MensajeError = "Error al consultar último remito: " & Err.Description
    GFlowRC_ConsultarUltimo = resultado
End Function

' ─── Consultar remito por código (para reimpresión) ───────────────────────────
' Si generarQR = True (default), genera el archivo PNG del QR localmente

Public Function GFlowRC_ConsultarRemito(codRemito As String, Optional generarQR As Boolean = True) As TResultadoRemito
On Error GoTo PROCERROR

    Dim resultado As TResultadoRemito
    resultado.Ok = False
    
    If Not GFlowRC_Disponible() Then
        resultado.MensajeError = "El servicio GFlowRC no está disponible."
        GFlowRC_ConsultarRemito = resultado
        Exit Function
    End If
    
    If Len(Trim(codRemito)) = 0 Then
        resultado.MensajeError = "Debe especificar un código de remito."
        GFlowRC_ConsultarRemito = resultado
        Exit Function
    End If
    
    Dim url As String
    url = GFLOWRC_URL & "/remito/consultar/" & Trim(codRemito)
    If generarQR Then url = url & "?generarQR=1"
    
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.SetTimeouts 5000, 5000, 30000, 30000
    http.Send
    
    Dim jsonResp As String
    jsonResp = http.ResponseText
    Set http = Nothing
    
    ' Verificar si se encontró
    Dim encontrado As String
    encontrado = GFlowRC_ExtraerCampo(jsonResp, "encontrado")
    
    If encontrado = "false" Then
        resultado.Ok = False
        resultado.MensajeError = "No se encontró el remito " & codRemito & " en ARCA."
        GFlowRC_ConsultarRemito = resultado
        Exit Function
    End If
    
    ' Extraer datos del remito
    resultado.Ok        = (GFlowRC_ExtraerCampo(jsonResp, "ok") = "true")
    resultado.CodRemito = GFlowRC_ExtraerCampo(jsonResp, "codRemito")
    resultado.CRE       = GFlowRC_ExtraerCampo(jsonResp, "cre")
    resultado.Estado    = GFlowRC_ExtraerCampo(jsonResp, "estado")
    resultado.QRUrl     = GFlowRC_ExtraerCampo(jsonResp, "qrUrl")
    resultado.QRPath    = GFlowRC_ExtraerCampo(jsonResp, "qrPath")
    
    If Not resultado.Ok Then
        resultado.MensajeError = GFlowRC_ExtraerCampo(jsonResp, "mensajeError")
    End If
    
    GFlowRC_ConsultarRemito = resultado
    Exit Function
    
PROCERROR:
    resultado.MensajeError = "Error al consultar remito: " & Err.Description
    GFlowRC_ConsultarRemito = resultado
End Function

' ─── Parsear resultado JSON ───────────────────────────────────────────────────

Public Function GFlowRC_ParsearResultado(jsonStr As String) As TResultadoRemito
    Dim r As TResultadoRemito
    
    r.Ok           = (GFlowRC_ExtraerCampo(jsonStr, "ok") = "true")
    r.YaExistia    = (GFlowRC_ExtraerCampo(jsonStr, "yaExistia") = "true")
    r.Aviso        = GFlowRC_ExtraerCampo(jsonStr, "aviso")
    r.CodRemito    = GFlowRC_ExtraerCampo(jsonStr, "codRemito")
    r.CRE          = GFlowRC_ExtraerCampo(jsonStr, "cre")
    r.QRUrl        = GFlowRC_ExtraerCampo(jsonStr, "qrUrl")
    r.QRPath       = GFlowRC_ExtraerCampo(jsonStr, "qrPath")
    r.Estado       = GFlowRC_ExtraerCampo(jsonStr, "estado")
    r.MensajeError = GFlowRC_ExtraerCampo(jsonStr, "mensajeError")
    
    GFlowRC_ParsearResultado = r
End Function

' ─── Extractor simple de campos JSON ─────────────────────────────────────────
' Parseo manual sin dependencias externas
' Funciona para campos de tipo string, boolean y number simples

Public Function GFlowRC_ExtraerCampo(jsonStr As String, campo As String) As String
    Dim buscar As String
    Dim pos As Integer
    Dim posInicio As Integer
    Dim posFin As Integer
    Dim valor As String
    
    ' Buscar "campo":
    buscar = """" & campo & """"
    pos = InStr(1, jsonStr, buscar, vbTextCompare)
    If pos = 0 Then
        GFlowRC_ExtraerCampo = ""
        Exit Function
    End If
    
    ' Avanzar hasta después de los ":"
    pos = pos + Len(buscar)
    Do While Mid(jsonStr, pos, 1) = " " Or Mid(jsonStr, pos, 1) = ":"
        pos = pos + 1
    Loop
    
    ' ¿Es string (empieza con ")?
    If Mid(jsonStr, pos, 1) = """" Then
        posInicio = pos + 1
        posFin = InStr(posInicio, jsonStr, """")
        If posFin > posInicio Then
            valor = Mid(jsonStr, posInicio, posFin - posInicio)
            ' Desescapar barras invertidas (paths Windows)
            valor = Replace(valor, "\\", "\")
            GFlowRC_ExtraerCampo = valor
        End If
    Else
        ' Booleano o número — leer hasta coma, } o ]
        posInicio = pos
        posFin = posInicio
        Do While posFin <= Len(jsonStr)
            Dim c As String
            c = Mid(jsonStr, posFin, 1)
            If c = "," Or c = "}" Or c = "]" Then Exit Do
            posFin = posFin + 1
        Loop
        GFlowRC_ExtraerCampo = Trim(Mid(jsonStr, posInicio, posFin - posInicio))
    End If
End Function

' ─── Cargar datos del remito desde SQL ───────────────────────────────────────

Public Function GFlowRC_CargarDatosRemito(xNroAuxRemito As Long) As Object
    ' Retorna Nothing si no encuentra el remito
    ' Usamos Scripting.Dictionary como contenedor de datos
    
    Dim datos As Object
    Set datos = CreateObject("Scripting.Dictionary")
    
    ModSQL.VtaTraeRemito xNroAuxRemito
    ModSQL.AbrirRSRead
    
    With RsRead
        If .RecordCount = 0 Then
            .Close
            Set GFlowRC_CargarDatosRemito = Nothing
            Exit Function
        End If
        
        datos("nroAuxRemito")      = xNroAuxRemito
        datos("nroRemito")         = Trim(!NroRemito)
        datos("cuitCliente")       = Trim(!Cuit)
        datos("tipoReceptorRM")    = Trim(!TipoReceptorRM)
        datos("codigoDestinoRM")   = Trim(!CodigoDestinoRM)
        datos("fechaInicioViaje")  = Format(IIf(IsNull(!FechaInicioViaje), Date, !FechaInicioViaje), "YYYY-MM-DD")
        datos("kmARecorrer")       = IIf(IsNull(!KmARecorrerRM), 0, !KmARecorrerRM)
        datos("patente")           = Left(Trim(!PatenteyOtros), 7)
        datos("cuitTransporte")    = Trim(!CuitTransporte)
        datos("cuitChofer")        = Trim(!CuilChofer)
        datos("tipoMovimiento")    = IIf(ModVariables.xTipoRemitoElectronico = "PLA", "PLA", "ENV")
        datos("ptoEmision")        = Left(!NroRemito, 4)
        
        .Close
    End With
    
    ' Calcular idReq = MAX(NROINTERNOELECTRONICO) + 1
    ModVariables.sTrSQL = "SELECT MAX(NROINTERNOELECTRONICO) + 1 AS NRO FROM VTAREMITOS"
    ModSQL.AbrirRSRead5
    Dim idReq As Long
    idReq = IIf(IsNull(RsRead5!Nro), 1, RsRead5!Nro)
    If xAumentarRequest Then
        idReq = idReq + 5
    End If
    RsRead5.Close
    Set RsRead5 = Nothing
    datos("idReq") = idReq
    ModVariables.xNroInternoRCarnico = idReq

    ' Reservar idReq en la DB ANTES de llamar al servicio.
    ' Si el servicio falla o no responde, el próximo intento calculará idReq+1
    ' y no habrá drift con ARCA.
    ModVariables.sTrSQL = "UPDATE VTAREMITOS SET NroInternoElectronico = " & idReq & " WHERE NroAuxRemito = " & xNroAuxRemito
    ModSQL.AbrirRSRead5
    RsRead5.Close
    Set RsRead5 = Nothing

    Set GFlowRC_CargarDatosRemito = datos
End Function

' ─── Construir JSON del body ──────────────────────────────────────────────────

Public Function GFlowRC_ConstruirJSON(datos As Object, xNroAuxRemito As Long) As String
    Dim json As String
    Dim mercaderias As String
    
    ' Categoría receptor
    Dim catReceptor As Integer
    catReceptor = 3  ' abastecedor por defecto
    If ModVariables.XCATEGOCLIENTEREMCARNI > 0 Then
        catReceptor = ModVariables.XCATEGOCLIENTEREMCARNI
    End If
    
    ' Leer mercaderías
    ModSQL.VtaLeeDetaRemitoCarnico xNroAuxRemito
    ModSQL.AbrirRSRead2
    
    mercaderias = ""
    Dim i As Integer
    i = 0
    Do While Not RsRead2.EOF
        Dim kilos As String
        kilos = Format(RsRead2!CantidadRemitada, "##0.00")
        kilos = Replace(kilos, ",", ".")
        
        Dim tropa As Long
        tropa = IIf(IsNull(RsRead2!Tropa) Or RsRead2!Tropa = 0, 1, RsRead2!Tropa)
        
        If Len(mercaderias) > 0 Then mercaderias = mercaderias & ","
        mercaderias = mercaderias & "{"
        mercaderias = mercaderias & """codTipoProd"":""" & Trim(RsRead2!CodSenasa) & ""","
        mercaderias = mercaderias & """tropa"":" & tropa & ","
        mercaderias = mercaderias & """kilos"":" & kilos & ","
        mercaderias = mercaderias & """unidades"":" & (IIf(IsNull(RsRead2!Bultos), 0, RsRead2!Bultos))
        mercaderias = mercaderias & "}"
        
        RsRead2.MoveNext
        i = i + 1
    Loop
    RsRead2.Close
    Set RsRead2 = Nothing
    
    ' Construir JSON completo
    json = "{"
    json = json & """idReq"":" & datos("idReq") & ","
    json = json & """tipoMovimiento"":""" & datos("tipoMovimiento") & ""","
    json = json & """categoriaEmisor"":1,"
    json = json & """tipoReceptor"":""" & Trim(Left(datos("tipoReceptorRM"), 2)) & ""","
    json = json & """categoriaReceptor"":" & catReceptor & ","
    json = json & """cuitReceptor"":""" & datos("cuitCliente") & ""","
    json = json & """codDomDestino"":""" & datos("codigoDestinoRM") & ""","
    json = json & """cuitTransportista"":""" & datos("cuitTransporte") & ""","
    json = json & """cuitConductor"":""" & datos("cuitChofer") & ""","
    json = json & """fechaInicioViaje"":""" & datos("fechaInicioViaje") & ""","
    json = json & """distanciaKm"":" & datos("kmARecorrer") & ","
    json = json & """dominio"":""" & datos("patente") & ""","
    json = json & """mercaderias"":[" & mercaderias & "]"
    json = json & "}"
    
    GFlowRC_ConstruirJSON = json
End Function
