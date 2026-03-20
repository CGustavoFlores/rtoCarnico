VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArpRemitoCarnico 
   Caption         =   "ArpRemitoCarnico"
   ClientHeight    =   9540
   ClientLeft      =   -180
   ClientTop       =   855
   ClientWidth     =   13890
   WindowState     =   2  'Maximized
   _ExtentX        =   24500
   _ExtentY        =   16828
   SectionData     =   "ArpRemitoCarnico.dsx":0000
End
Attribute VB_Name = "ArpRemitoCarnico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim pngClass As New LoadPNG

Dim xNroFactura As String, xTransporteRemito As String, xProvincia
Dim xObserva As String, xEstado As String, xDeposito As String
Dim xRSocial As String, xCodCte As Long, xDomicilio As String, xLocalidad As String
Dim xPosicionFiscal As String, xCuit As String, xCpbte As String, xLetra As String, xCuilChofer As String
Dim xPto As String, xNro As String, xFecha As Date, xCondiVta As String, xLugarDeDescarga As String
Dim xCodProducto As String, xCantidad As Double, xDescripcion As String, xPUnitario As Double
Dim xImporte As Double, xIInterno As Double, xNetoExento As Double, xSubTotal As Double, XNROINTERNORC As Long, XarchivoQR As String
Dim xDescuento As Double, xSubTotalConDto As Double, xNetoGravado As Double, xTasaIva As Double
Dim xIva As Double, xTasaIvaNI As Double, xIvaNI As Double, xTotal As Double, xFechaViaje As Date, xFechaVenceViaje As Date
Dim xNroAuxiliar As Long, xPercepIB As Double, xPercepIVA As Double, xTituloTotal  As String, xPatenteTransporte As String
Dim xEsdelExterior As Boolean, xChofer As String, xLeyendaremito As String, xDescarga As String, xNroOCompra As String
Dim xCodSuc As Integer, XLEYENDASUC  As String, xLeyendaAlFinal As String, xCuitCte As String, xNROCRE As String
Const xLeyenda = "Por Cuenta y Orden de CTI"


Private Sub ActiveReport_Initialize()

Me.Toolbar.Tools.Insert 21, "Exportar"
Me.Toolbar.Tools.Insert 22, "Enviar por mail"
Me.Toolbar.Tools.Insert 23, "CRE"
Me.Toolbar.Tools.Item(19).Visible = False
Me.Toolbar.Tools.Item(20).Visible = False

End Sub

Private Sub ActiveReport_ToolbarClick(ByVal tool As DDActiveReports2.DDTool)
On Error GoTo PROCERROR

If tool.Caption = "CRE" Then
    Call GenerarRemitoCarnico
End If

   If tool.Caption = "Exportar" Then
        Set ModVariables.xListadoAExportar = Me
        frmOpcionesExportacion.Show vbModal
        Unload frmOpcionesExportacion
        If ModVariables.xEstadoComprobante <> "CANCELA" Then
            Dialogo.FileName = "" 'para que limpie el nombre, sino queda el último usado
            Dialogo.Flags = cdlOFNOverwritePrompt
            Dialogo.FilterIndex = 0 'XLS
            Dialogo.ShowSave
            ModVariables.xnombrearchivo = Dialogo.FileName
            ModVariables.GenerarExportacion
        End If
   End If
Exit Sub
PROCERROR:
    If Err.Number = cdlCancel Then 'se presionó CANCELAR
        Exit Sub
    Else
        MsgBox "Ha ocurrido un error en la generación de la exportación", vbExclamation, "Atención"
    End If
End Sub



Private Sub ActiveReport_DataInitialize()
'
Fields.Add ("Nro")
Fields.Add ("Fecha")
Fields.Add ("CondiVta")
Fields.Add ("FechaViaje")
Fields.Add ("FechaVencimiento")

'origen
Fields.Add ("OrigenRemito")
Fields.Add ("RSocialOrigen")
Fields.Add ("NroFactura")
Fields.Add ("CuitOrigen")
Fields.Add ("DomicilioOrigen")
'destino
Fields.Add ("RSocial")
Fields.Add ("CodCte")
Fields.Add ("Cuit")
Fields.Add ("DepositoDestino")

' transporte
Fields.Add ("Transporte")
Fields.Add ("CuitTransporte")
Fields.Add ("PatenteVehiculos")
Fields.Add ("Conductor")
Fields.Add ("CuitConductor")
'productos
Fields.Add ("CodProducto")
Fields.Add ("Descripcion")
Fields.Add ("Tropa")
Fields.Add ("Cantidad")
Fields.Add ("Bultos")

'Fields.Add ("LeyendaRemito")
'Fields.Add ("LugarDescarga")

End Sub


Private Sub ActiveReport_ReportStart()

Me.Caption = "Remitos"

xNroFactura = ""
xNroOCompra = ""
If ModVariables.xNroAuxiliar = 0 Then
    ModVariables.xNroAuxiliar = ModVariables.xNroAuxRemito
End If
ModSQL.VtaTraeRemito ModVariables.xNroAuxiliar
ModSQL.AbrirRSRead
With RsRead
    If .RecordCount > 0 Then
        xRSocial = Trim(![RazonSocial])
        xCodCte = ![CodCliente]
        xDomicilio = IIf(IsNull(![Domicilio]), "", Trim(![Domicilio]))
        xLocalidad = IIf(IsNull(![Localidad]), "", Trim(![Localidad]))
        xProvincia = IIf(IsNull(![NombrePcia]), "", Trim(![NombrePcia]))
        xPosicionFiscal = IIf(IsNull(![NombrePosFiscal]), "", ![NombrePosFiscal])
        xCuit = IIf(IsNull(![Cuit]), "", ![Cuit])
        xCuitCte = xCuit
        If ![delexterior] = True Then
            xEsdelExterior = ![delexterior]
            xCuit = ![CuitExterior]
        End If
        xNroFactura = Trim(![NroFactura])
        xNroOCompra = Trim(![OCompraCliente])
        xFecha = ![Fecha]
        xEstado = ![Estado]
        xObserva = Trim(![OBSERVACIONES])
        xDeposito = ![CodDeposito] & " " & Trim(![DescripDeposito])
        xNro = Format(![NroRemito], "0000-00000000")
        xTransporteRemito = "(" & ![CodTransporte] & ") " & Trim(![NombreTransporte])
        xCuitTransporte = Trim(![CuitTransporte])
        xPatenteTransporte = Trim(![PatenteyOtros])
        xChofer = Trim(![NombreyDnichofer])
        xCuilChofer = Trim(![CuilChofer])
        xFechaViaje = IIf(IsNull(![FechaInicioViaje]), 0, ![FechaInicioViaje])
        xFechaVenceViaje = IIf(IsNull(![FechaVencimiento]), 0, ![FechaVencimiento])
        xNROCRE = ![CRE]
        XNROINTERNORC = ![NroInternoElectronico]
        XarchivoQR = "C:\WsRe\imagenes\QR-" & Trim(Str(XNROINTERNORC)) & ".png"
        
        'xLeyendaAlFinal = Trim(![leyendaenremito])
        xLugarDeDescarga = Trim(![LugardeDescarga])
        If Trim(![LocalidadDescarga]) <> "" Then
            xLocalidad = Trim(![LocalidadDescarga])
        End If
        If ModVariables.xUtilizarSucursalCliente Then
            xCodSuc = ![codsucursal]
        End If
    Else
        .Close
        MsgBox "No se encontró el Nº de Remito", vbInformation, "Atención"
        Unload Me
    End If
End With
'Set Rs = VEN.rsVtaTraeRemito
Set Rs = ModSQL.RsRead
ModSQL.VtaLeeDetaRemitoCarnico ModVariables.xNroAuxiliar
'
'If ModVariables.xCuitEmpresa = "30636764155" And xCodCte = 5 And ModVariables.XIMPRIMECODBARRACONGELADOS Then
'    ModSQL.VtaLeeDetaRemitoInternoParaCongelado ModVariables.xNroAuxiliar
'End If

ModSQL.AbrirRSRead5

If VEN.rsVtaLeerRelCpbtesFACTURA.state = adStateOpen Then
    VEN.rsVtaLeerRelCpbtesFACTURA.Close
End If
VEN.VtaLeerRelCpbtesFACTURA ModVariables.xNroAuxiliar ' VariablesCpbteVta.xNroAuxCpbteVta
If VEN.rsVtaLeerRelCpbtesFACTURA.RecordCount > 0 Then
    xNroFactura = VEN.rsVtaLeerRelCpbtesFACTURA![Tipo] & " " & VEN.rsVtaLeerRelCpbtesFACTURA![PtoVta] & "-" & VEN.rsVtaLeerRelCpbtesFACTURA![NroCpbte]
Else
    'xNroFactura = ""
End If
VEN.rsVtaLeerRelCpbtesFACTURA.Close

XLEYENDASUC = ""
If ModVariables.xUtilizarSucursalCliente And xCodSuc <> 0 Then
    ModVariables.sTrSQL = "SELECT CLIENTESUCURSAL.*,  LOCALIDADES.LOCALIDAD AS NOMBRELOCALIDAD , PROVINCIAS.NombrePcia " & vbCrLf & _
        " FROM CLIENTESUCURSAL " & vbCrLf & _
        " LEFT JOIN LOCALIDADES ON CLIENTESUCURSAL.LOCALIDAD = LOCALIDADES.codlocali " & vbCrLf & _
        " LEFT JOIN PROVINCIAS ON LOCALIDADES.CodProvincia= PROVINCIAS.IdProvincia " & vbCrLf & _
        " WHERE CLIENTESUCURSAL.IDCLIENTE= " & xCodCte & " AND CLIENTESUCURSAL.idsucursal = " & xCodSuc
    ModSQL.AbrirRSRead4
    If RsRead4.RecordCount = 1 Then
        XLEYENDASUC = "MERCADERIA A ENTREGAR EN SUC :" & RsRead4![Idsucursal] & "-" & Trim(RsRead4![Direccion]) & " - " & Trim(RsRead4![NombreLocalidad]) & " (" & Trim(RsRead4![NombrePcia]) & ")"
    Else
        XLEYENDASUC = ""
    End If
    RsRead4.Close
    Set RsRead4 = Nothing
    xLeyendaAlFinal = XLEYENDASUC & " //// " & xLeyendaAlFinal
End If



End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)
On Error GoTo PROCERROR
CodigoBarra.Style = ddbcEAN_13

With RsRead5
    If .EOF Then
        Exit Sub
    End If
    Fields("CodProducto").Value = IIf(IsNull(![CodSenasa]), "", Trim(![CodSenasa]))
    Fields("Descripcion") = IIf(IsNull(![Descripcion]), "", Trim(![Descripcion])) & " (" & Trim(![CodArticulo]) & ")"
    Fields("Tropa") = IIf(IsNull(![Tropa]), "0", ![Tropa])
    Fields("Cantidad") = IIf(IsNull(![CantidadRemitada]), 0, (![CantidadRemitada] - ![CantidadDevuelta]))
    Fields("Bultos") = IIf(IsNull(![Bultos]), 0, ![Bultos])
    If ModVariables.XIMPRIMECODBARRACONGELADOS = True Then
        CodigoBarra.Caption = IIf(IsNull(![CodigoBarraCongelado]), "", Trim(![CodigoBarraCongelado]))
    Else
        CodigoBarra.Caption = IIf(IsNull(![CodigoBarra]), "", Trim(![CodigoBarra]))
    End If
   .MoveNext
    EOF = False
End With
Exit Sub
PROCERROR:
    ModSQL.MensajeError "PageHeader_Format"
Resume Next
End Sub


Private Sub PageFooter_Format()
CAE.DataValue = xNROCRE
End Sub

Private Sub PageHeader_Format()
On Error GoTo PROCERROR

Nro.DataValue = xNro
Fecha.DataValue = Format(xFecha, "dd/mm/yyyy")


FechaViaje.DataValue = Format(xFechaViaje, "dd/mm/yy")
FechaVencimiento.DataValue = Format(xFechaVenceViaje, "dd/mm/yy")

'origen
OrigenRemito.DataValue = "ORIGEN DEL TRASLADO - Remision de mercaderia propia desde deposito propio"
RSocialOrigen.DataValue = ModVariables.xRSocialEmpresa
'Fields.Add ("NroFactura")
CuitOrigen.DataValue = ModVariables.xCuitEmpresa
DomicilioOrigen.DataValue = ModVariables.xDomicilioEmpresa
'destino
RSocial.DataValue = xRSocial
CodCte.DataValue = xCodCte
Cuit.DataValue = xCuitCte
DepositoDestino.DataValue = xLugarDeDescarga
NroOrden.DataValue = xNroOCompra
Factura.DataValue = xNroFactura
' transporte
Transporte.DataValue = xTransporteRemito
CuitTransporte.DataValue = xCuitTransporte
PatenteVehiculo.DataValue = xPatenteTransporte
CuitConductor.DataValue = xCuilChofer
Conductor.DataValue = xChofer

CodigoBarra.Visible = False
If xCodCte = 5 Or xCteImprimeCodBarraEnArticu = True Then
    CodigoBarra.Visible = True
End If

Call ModImprimirPNG.PngImageLoad(XarchivoQR, VtaPrnComprobantes.Image1)        'para cargarlo en el image
'Call ModImprimirPNG.PngImageLoad("C:\WsRe\imagenes\QR-7.png", VtaPrnComprobantes.Image1)      'para cargarlo en el image
ImageQR.Picture = VtaPrnComprobantes.Image1.Picture
'Call ModImprimirPNG.PngPictureLoad("C:\WsRe\imagenes\QR-7.png", Picture1, True) 'para cargarlo en el picture

'If File1.filename <> "" Then
    
'    ImageQR.Picture = LoadPicture("")
'    pngClass.PicBox = ImageQR
'    pngClass.SetToBkgrnd False, 0, 0  'set to Background (True or false), x and y
'    pngClass.BackgroundPicture = ImageQR 'same Backgroundpicture
'    pngClass.SetAlpha = True 'when Alpha then alpha
'    pngClass.SetTrans = True 'when transparent Color then transparent Color
'
'
'    ' Visualiza el Archivo en Picture1
'    'pngClass.OpenPNG File1.Path & "\" & File1.filename
'     pngClass.OpenPNG "CAMERA.PNG"

'End If
'If RsRead Is Nothing Then
''If RsRead.EOF Then
'    ModSQL.VtaLeeDetaRemitoCarnico ModVariables.xNroAuxiliar
'    ModSQL.AbrirRSRead
'End If
''
'If ModVariables.xCuitEmpresa = "30636764155" And xCodCte = 5 And ModVariables.XIMPRIMECODBARRACONGELADOS Then
'    ModSQL.VtaLeeDetaRemitoInternoParaCongelado ModVariables.xNroAuxiliar
'End If

Exit Sub
PROCERROR:
    ModSQL.MensajeError "PageHeader_Format"
    'Resume Next
End Sub

Private Sub ReportFooter_Format()
'If VEN.rsVtaLeeCpbteDetaVta.State = adStateOpen Then
'LeyendaRemito.DataValue = xLeyendaAlFinal
'LeyendaRemito.Visible = True


End Sub

Private Sub ReportHeader_Format()
'FechaEmision.DataValue = Format(Date, "short date")
'HoraEmision.DataValue = Format(Now, "hh:mm AM/PM")
'txtComprobante.DataValue = "REMITO"

End Sub

Private Sub GenerarRemitoCarnico()
ModVariables.XCRE = ""

'ModVariables.xProcesoRemElectronico = "RMGenerarRCarnico"
frmXMLRemitoCarnico.GenerarXMLRemitoCarnico (ModVariables.xNroAuxiliar)
If ModVariables.XCRE <> "" Then
    MsgBox ModVariables.XCRE, vbInformation, "NRO. REMITO ELECTRONICO"
Else
    MsgBox "No se obtuvo el Nro. de Remito Electronico"
End If



End Sub
