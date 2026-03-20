VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{F09A78C8-7814-11D2-8355-4854E82A9183}#1.1#0"; "CUIT32.OCX"
Begin VB.Form frmXMLRemitoCarnico 
   BackColor       =   &H00D1ECFC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REMITO CARNICO XML DESTINOS"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkCamara 
      Caption         =   "Camara o Deposito"
      Height          =   225
      Left            =   2535
      TabIndex        =   19
      Top             =   420
      Width           =   1935
   End
   Begin VB.CommandButton CmdCorregirNroInterno 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Corregir Nro. Interno Remito"
      Height          =   510
      Left            =   7830
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2085
      Width           =   1590
   End
   Begin VB.CommandButton cmdConsultarUltimoRemito 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consultar Nro.Ultimo Remito"
      Height          =   510
      Left            =   7815
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1110
      Width           =   1590
   End
   Begin VB.CheckBox ChkRetiroEnPlanta 
      Caption         =   "Retiro en Planta"
      Height          =   225
      Left            =   4485
      TabIndex        =   16
      Top             =   435
      Width           =   1515
   End
   Begin VB.CommandButton CmdAnularRC 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Anular Rem. Pte Recepcion"
      Height          =   510
      Left            =   7815
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   510
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.CheckBox CHKCMD 
      Caption         =   "CMD"
      Height          =   285
      Left            =   7830
      TabIndex        =   14
      Top             =   4485
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CommandButton cmdGrabarDestinos 
      Caption         =   "Grabar Destinos"
      Height          =   420
      Left            =   6000
      TabIndex        =   13
      Top             =   15
      Width           =   1500
   End
   Begin VB.CommandButton CMDLEEROUTPUTXML 
      BackColor       =   &H0080FFFF&
      Caption         =   "LeerOutPutXML"
      Height          =   510
      Left            =   7665
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   -30
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.CommandButton CmdLeerDesinos 
      Caption         =   "Leer Destinos "
      Height          =   420
      Left            =   4500
      TabIndex        =   11
      Top             =   15
      Width           =   1500
   End
   Begin VB.CommandButton CmdBuscarEnAFIP 
      Caption         =   "Buscar destinos en AFIP"
      Height          =   420
      Left            =   2550
      TabIndex        =   10
      Top             =   15
      Width           =   1950
   End
   Begin VB.CommandButton CmdCancelar 
      BackColor       =   &H00C5D5E2&
      Caption         =   "Ca&ncelar"
      Height          =   495
      Left            =   8235
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3615
      Width           =   1110
   End
   Begin VB.CommandButton CmdConfirmar 
      BackColor       =   &H00C5D5E2&
      Caption         =   "Confirmar"
      Height          =   495
      Left            =   8235
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3105
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.CommandButton cmdLeerOutPut 
      BackColor       =   &H0080FFFF&
      Caption         =   "LeerOutPut"
      Height          =   510
      Left            =   3135
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4275
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton CmdObtenerDestino 
      BackColor       =   &H0080FFFF&
      Caption         =   "Obtener Destino"
      Height          =   510
      Left            =   1875
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4275
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.CommandButton CmdRemitoStockPropio 
      BackColor       =   &H0080FFFF&
      Caption         =   "Remito Comun Stock Propio"
      Height          =   510
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4275
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Generar XML"
      Height          =   480
      Left            =   6555
      TabIndex        =   2
      Top             =   4305
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton CmdGeneraXmlInput 
      BackColor       =   &H0080FFFF&
      Caption         =   "Generar XML (input.xml)"
      Height          =   510
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4275
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VSFlex7Ctl.VSFlexGrid Grid 
      Height          =   3555
      Left            =   30
      TabIndex        =   0
      Top             =   645
      Width           =   7755
      _cx             =   13679
      _cy             =   6271
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   14737632
      ForeColor       =   -2147483640
      BackColorFixed  =   32768
      ForeColorFixed  =   16777215
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   14737632
      BackColorAlternate=   14737632
      GridColor       =   -2147483633
      GridColorFixed  =   16711680
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmXMLRemitoCarnico.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin Control_CUIT.CUIT TxtCuitCliente 
      Bindings        =   "frmXMLRemitoCarnico.frx":0062
      DataField       =   "cuit"
      Height          =   315
      Left            =   945
      TabIndex        =   8
      Top             =   30
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      ConSeparador    =   0   'False
      Text            =   ""
      Enabled         =   0   'False
      MensajeErr      =   "Cuit Incorrecto"
   End
   Begin VB.Label label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuit Cliente:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   33
      Left            =   30
      TabIndex        =   9
      Top             =   135
      Width           =   885
   End
End
Attribute VB_Name = "frmXMLRemitoCarnico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Option Explicit
Dim strRutaArchivo As String


'Funciones del api
'''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function OpenProcess Lib "kernel32" _
() '  (ByVal dwDesiredAccess As Long, _
   ByVal bInheritHandle As Long, _
   ByVal dwProcessId As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
() '  (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" _
() '  (ByVal hObject As Long) As Long

'Constantes
'''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const STATUS_PENDING = &H103&



' Recibe el argumento de la línea de comandos para pasarle al Shell
Private Sub Ejecutar_shell(programa As String)
On Error GoTo PROCERROR
    Dim handle_Process As Long
    Dim id_process As Long
    Dim lp_ExitCode As Long

    ' Abre el proceso con el shell
    id_process = Shell(programa, 1)

    ' handle del proceso
    handle_Process = OpenProcess(PROCESS_QUERY_INFORMATION, False, id_process)

    ' Mientras lp_ExitCode = STATUS_PENDING, se ejecuta el do
    Do

        Call GetExitCodeProcess(handle_Process, lp_ExitCode)

        DoEvents

    Loop While lp_ExitCode = STATUS_PENDING

    ' fin
    ' Cierra
    Call CloseHandle(handle_Process)

    'MsgBox "Se cerró el " & programa, vbInformation
Exit Sub
PROCERROR:
    Resume Next
End Sub


Private Sub CmdAnularRC_Click()
On Error GoTo PROCERROR

strRutaArchivo = "C:\WsRe\"
 
Set FSO = New Scripting.FileSystemObject
Set FXML = FSO.CreateTextFile("c:\wsre\input.xml", True)
FXML.WriteLine "<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/' xmlns:rem='http://ar.gob.afip.wsremcarne/RemCarneService/' >"
FXML.WriteLine "<soapenv:Header/>"
FXML.WriteLine "<soapenv:Body>"
FXML.WriteLine "<rem:informarContingenciaRequest>"
FXML.WriteLine "<authRequest>"
FXML.WriteLine "<token>@TOKEN</token>"
FXML.WriteLine "<sign>@SIGN</sign>"
FXML.WriteLine "<cuitRepresentada>" & Trim(ModVariables.xCuitEmpresa) & "</cuitRepresentada>"
FXML.WriteLine "</authRequest>"
FXML.WriteLine "<codRemito>" & ModVariables.xNROCAERC & "</codRemito>"
FXML.WriteLine "<contingencia>"
FXML.WriteLine "<tipoContingencia>" & ModVariables.XCODIGOCONTIGENCIA & "</tipoContingencia>"  '8=ANULACION TOTAL Y REGRESO A PLANTA, 4=ACCIDENTE VIAL, 5=ROBO, 6=DESTRUCCION, 7=OTROS SINIENTROS
FXML.WriteLine "<observacion>" & ModVariables.XOBSERCONTINGENCIA & "</observacion>"
FXML.WriteLine "</contingencia>"
FXML.WriteLine "</rem:informarContingenciaRequest>"
FXML.WriteLine "</soapenv:Body>"
FXML.WriteLine "</soapenv:Envelope>"


strRutaArchivo = strRutaArchivo & "VALIDAR.BAT"
' Call Ejecutar_shell(strRutaArchivo)
If CHKCMD.Value = 0 Then
    Shell strRutaArchivo, vbMinimizedNoFocus
Else
    Shell ("cmd.exe /k" & strRutaArchivo), vbMinimizedNoFocus
End If

strRutaArchivo = strRutaArchivo & "VALIDAR.BAT"
' Call Ejecutar_shell(strRutaArchivo)
If CHKCMD.Value = 0 Then
    Shell strRutaArchivo, vbMinimizedNoFocus
Else
    Shell ("cmd.exe /k" & strRutaArchivo), vbMinimizedNoFocus
End If
Exit Sub
PROCERROR:
    ModSQL.MensajeError "CmdObtenerDestino"
    Resume Next
End Sub

Private Sub CmdBuscarEnAFIP_Click()
On Error GoTo PROCERROR
Select Case ModVariables.xProcesoRemElectronico
    Case "RMObtenerDestinos"
        If ModVariables.xEmiteRemitoCarnico Then
            GRid.Rows = 1
            If Trim(TxtCuitCliente.Text) <> "" Then
                Call CmdObtenerDestino_Click
                MsgBox "Proceso finalizado. Pulse el boton 'Leer Destinos'", vbInformation, "AVISO"
                'Call Sleep(2000) 'espera por 2 segundos
    
                'Call cmdLeerOutPut_Click
            End If
        End If
        If ModVariables.xEmiteRemitoHarinero Then
            If Fe.rcConsultarCodigoDomicilio(TxtCuitCliente.Text) Then
                Dim nContador As Integer
                Dim nTope As Integer
                  nTope = Val(Fe.rcLeerPropiedad("rcConsultarCodigoDomicilio", "itemcantidad", "", 0, 0))
                  GRid.Rows = 1
                For nContador = 1 To nTope
                    xrenglon = Fe.rcLeerPropiedad("rcConsultarCodigoDomicilio", "codigo", "", nContador - 1, 0) & vbTab & Fe.rcLeerPropiedad("rcConsultarCodigoDomicilio", "descripcion", "", nContador - 1, 0)
                    GRid.AddItem xrenglon
                    If Fe.UltimoMensajeError = "" Then
                    Else
                        MsgBox ("fallo al leer resultado " + Chr(10) + Fe.UltimoMensajeError + Fe.rcErrorDescripcion1)
                        Exit For
                    End If
                Next
'                For nContador = nTope + 1 To 50
'                    hoja.Cells(nContador + 1, 1).Value = ""
'                    hoja.Cells(nContador + 1, 2).Value = ""
'                Next
            Else
                MsgBox ("fallo al obtener lista de domicilio" + Chr(10) + gFe.UltimoMensajeError + gFe.rcErrorDescripcion1)
            End If
        End If
End Select

Exit Sub
PROCERROR:
    ModSQL.MensajeError "CmdBuscarEnAFIP"
    Resume Next
End Sub

Private Sub cmdConsultarUltimoRemito_Click()
On Error GoTo PROCERROR


strRutaArchivo = "C:\WsRe\"
 
Dim xLineaRepresentada As String
Dim xLineaCliente As String
'txtcuitempresa.text = ModVariables.xCuitEmpresa
xLineaRepresentada = "            <cuitRepresentada>" & Trim(ModVariables.xCuitEmpresa) & "</cuitRepresentada>"
xLineaCliente = "        <cuitTitularDomicilio>" & Trim(TxtCuitCliente.Text) & "</cuitTitularDomicilio>"
Set FSO = New Scripting.FileSystemObject
Set FXML = FSO.CreateTextFile("c:\wsre\input.xml", True)
FXML.WriteLine "<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/' xmlns:rem='http://ar.gob.afip.wsremcarne/RemCarneService/'> "
FXML.WriteLine "<soapenv:Header/>"
FXML.WriteLine "<soapenv:Body>"
FXML.WriteLine "<rem:consultarUltimoRemitoEmitidoRequest>"
FXML.WriteLine "    <authRequest>"
FXML.WriteLine "        <token>@TOKEN</token>"
FXML.WriteLine "            <sign>@SIGN</sign>"
FXML.WriteLine xLineaRepresentada
FXML.WriteLine "        </authRequest>"
FXML.WriteLine "<tipoComprobante>995</tipoComprobante>"
FXML.WriteLine "<puntoEmision>" & "0014" & "</puntoEmision>"
FXML.WriteLine "</rem:consultarUltimoRemitoEmitidoRequest>"
FXML.WriteLine "</soapenv:Body>"
FXML.WriteLine "</soapenv:Envelope>"

'Shell strRutaArchivo & "VALIDAR.BAT", vbMaximizedFocus

strRutaArchivo = strRutaArchivo & "VALIDAR.BAT"
' Call Ejecutar_shell(strRutaArchivo)
If CHKCMD.Value = 0 Then
    Shell strRutaArchivo, vbMinimizedNoFocus
Else
    Shell ("cmd.exe /k" & strRutaArchivo), vbMinimizedNoFocus
End If
Exit Sub
PROCERROR:
    ModSQL.MensajeError "CmdObtenerDestino"
    Resume Next

End Sub

Private Sub CmdCorregirNroInterno_Click()
On Error GoTo PROCERROR
ModVariables.sTrSQL = "SELECT     TOP (1) NROINTERNOELECTRONICO From VTAREMITOS Where NROINTERNOELECTRONICO > 0 ORDER BY  Fecha DESC , NroRemito DESC"
ModSQL.AbrirRSOptimista4
With RsOpt4
    If .RecordCount = 1 Then
        RsOpt4![NROINTERNOELECTRONICO] = RsOpt4![NROINTERNOELECTRONICO] + 20
        RsOpt4.Update
    End If
    RsOpt4.Close
    Set RsOpt4 = Nothing
End With
Exit Sub
PROCERROR:
    ModSQL.MensajeError "CmdCorregirNroInterno_ClickS"
End Sub

Private Sub CmdLeerDesinos_Click()
    Call cmdLeerOutPut_Click
End Sub

Private Sub CMDLEEROUTPUTXML_Click()
    Call LeerOutPutRCEMITIDO
    'Call LeerOutPutXML
End Sub

Private Sub CmdObtenerDestino_Click()
On Error GoTo PROCERROR

strRutaArchivo = "C:\WsRe\"
 
Dim xLineaRepresentada As String
Dim xLineaCliente As String
'txtcuitempresa.text = ModVariables.xCuitEmpresa
xLineaRepresentada = "            <cuitRepresentada>" & Trim(ModVariables.xCuitEmpresa) & "</cuitRepresentada>"
xLineaCliente = "        <cuitTitularDomicilio>" & Trim(TxtCuitCliente.Text) & "</cuitTitularDomicilio>"
Set FSO = New Scripting.FileSystemObject
Set FXML = FSO.CreateTextFile("c:\wsre\input.xml", True)
FXML.WriteLine "<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/' xmlns:rem='http://ar.gob.afip.wsremcarne/RemCarneService/'> "
FXML.WriteLine "<soapenv:Header/>"
FXML.WriteLine "<soapenv:Body>"
FXML.WriteLine "<rem:consultarCodigosDomicilioRequest>"
FXML.WriteLine "    <authRequest>"
FXML.WriteLine "        <token>@TOKEN</token>"
FXML.WriteLine "            <sign>@SIGN</sign>"
'fXML.WriteLine "            <cuitRepresentada>CUIT DE LA EMPRESA</cuitRepresentada>"
FXML.WriteLine xLineaRepresentada
FXML.WriteLine "        </authRequest>"
'fXML.WriteLine "        <cuitTitularDomicilio>CUIT DEL CLIENTE</cuitTitularDomicilio>"
FXML.WriteLine xLineaCliente
FXML.WriteLine "    </rem:consultarCodigosDomicilioRequest>"
FXML.WriteLine "</soapenv:Body>"
FXML.WriteLine "</soapenv:Envelope>"

'Shell strRutaArchivo & "VALIDAR.BAT", vbMaximizedFocus

strRutaArchivo = strRutaArchivo & "VALIDAR.BAT"
' Call Ejecutar_shell(strRutaArchivo)
If CHKCMD.Value = 0 Then
    Shell strRutaArchivo, vbMinimizedNoFocus
Else
    Shell ("cmd.exe /k" & strRutaArchivo), vbMinimizedNoFocus
End If
Exit Sub
PROCERROR:
    ModSQL.MensajeError "CmdObtenerDestino"
    Resume Next
End Sub

Private Sub cmdLeerOutPut_Click()
On Error GoTo PROCERROR

Dim xNroFile As Integer, xArchivo As String, xRuta As String, xEstado As String, XLINEA As String
Dim xStringTotal As String, xrenglon As String, xINicio As Integer, X As Integer

Me.MousePointer = 11
Call Sleep(2000) 'espera por 2 segundos
Me.MousePointer = 0
xStringTotal = ""
xRuta = "C:\WsRe"
If Right(xRuta, 1) = "\" Then
    xRuta = xRuta
Else
    xRuta = xRuta & "\"
End If
xArchivo = "OUTPUT.TXT"
xNroFile = FreeFile()
Open xRuta & xArchivo For Input As xNroFile
Do While Not EOF(xNroFile)
    Input #xNroFile, XLINEA
    xStringTotal = xStringTotal & " " & XLINEA
    'MsgBox XLINEA, vbInformation, "Continuar"

Loop
Close xNroFile
'x = InStr(1, xStringTotal, "WAIT")  'SI DEVUELVE WAIT, HAY QUE DEMORAR UN MOMENTO EL PROGRAMA Y PROBAR DE NUEVO
'If x <> 0 Then
'    MsgBox "Presione una tecla para continuar", vbInformation, "WAIT"
'End If
X = InStr(1, xStringTotal, "RESPUESTAAFIP:")  'busca cadena2 en cadena1 partiendo de inicio en cadena1, si no se pone inicio empieza en 1, devuelve la primera posicion donde encuentre cadena2
If X = 0 Then
    MsgBox xStringTotal
    MsgBox "NO se pudo obtener los destinos para este cliente", vbInformation, "Validar"
    Exit Sub
End If
xrenglon = ""
xINicio = False
GRid.Rows = 1
    For I = X To Len(xStringTotal)
        If UCase(Mid(xStringTotal, I, 1)) = ":" Then
            xINicio = True
        End If
    If xINicio Then
'        If Mid(xStringTotal, I, 1) = "," Then
'           ' I = I
'        End If
        If Mid(xStringTotal, I, 1) <> ":" And Mid(xStringTotal, I, 1) <> "#" And Mid(xStringTotal, I, 1) <> "@" Then
            xrenglon = xrenglon & Mid(xStringTotal, I, 1)
        End If
        Select Case UCase(Mid(xStringTotal, I, 1))
            Case "#"
                If IsNumeric(xrenglon) Then
                    xrenglon = Trim(xrenglon)
                End If
                xrenglon = xrenglon & vbTab
            Case "@"
                GRid.AddItem xrenglon
                xrenglon = ""
        End Select
    End If
Next
GRid.AddItem xrenglon
Exit Sub
PROCERROR:
    ModSQL.MensajeError "CmdLeerOutPut.txt"
   ' Resume Next
End Sub



Private Sub LeerOutPutRM()
On Error GoTo PROCERROR

Dim xNroFile As Integer, xArchivo As String, xRuta As String, xEstado As String, XLINEA As String
Dim xStringTotal As String, xrenglon As String, xINicio As Integer, X As Integer

xRuta = "C:\WsRe"
If Right(xRuta, 1) = "\" Then
    xRuta = xRuta
Else
    xRuta = xRuta & "\"
End If
xArchivo = "OUTPUT.TXT"
xNroFile = FreeFile()
Open xRuta & xArchivo For Input As xNroFile
Do While Not EOF(xNroFile)
    Input #xNroFile, XLINEA
    xStringTotal = xStringTotal & XLINEA
    'MsgBox XLINEA, vbInformation, "Continuar"
Loop
Close xNroFile
X = InStr(1, xStringTotal, "RESPUESTAAFIP:")  'busca cadena2 en cadena1 partiendo de inicio en cadena1, si no se pone inicio empieza en 1, devuelve la primera posicion donde encuentre cadena2
If X = 0 Then
    MsgBox xStringTotal
    MsgBox "NO se pudo obtener Remito Carnico (SP)", vbInformation, "Validar"
    Exit Sub
End If
xrenglon = ""
xINicio = False
GRid.Rows = 1
For I = X To Len(xStringTotal)
    If UCase(Mid(xStringTotal, I, 1)) = ":" Then
        xINicio = True
    End If
    If xINicio Then
        If Mid(xStringTotal, I, 1) = "," Then
            I = I
        End If
        If Mid(xStringTotal, I, 1) <> ":" And Mid(xStringTotal, I, 1) <> "#" And Mid(xStringTotal, I, 1) <> "@" Then
            xrenglon = xrenglon & Mid(xStringTotal, I, 1)
        End If
        Select Case UCase(Mid(xStringTotal, I, 1))
            Case "#"
                If IsNumeric(xrenglon) Then
                    xrenglon = Trim(xrenglon)
                End If
                xrenglon = xrenglon & vbTab
            Case "@"
                GRid.AddItem xrenglon
                xrenglon = ""
        End Select
    End If
Next
GRid.AddItem xrenglon
Exit Sub
PROCERROR:
    ModSQL.MensajeError "CmdLeerOutPut.txt"
    Resume Next
End Sub

Private Sub LeerOutPutXML()
On Error GoTo PROCERROR

Dim MiXml As New DOMDocument60, xRuta As String, xArchivo As String
xRuta = "C:\WsRe"
If Right(xRuta, 1) = "\" Then
    xRuta = xRuta
Else
    xRuta = xRuta & "\"
End If
xArchivo = "OUTPUT.XML"

'Para este ejemplo tenemos ya creado un archivo XML con lo siguiente ( recordar que un archivo XML se puede abrir con sólo arrastrar el archivo a un bloc de notas ).
'
'<?xml version="1.0" encoding="utf-8"?>
'<Config>
'<RutaBD>prueba</RutaBD>
'<Ruta2 prueba="prueba atributo">prueba2</Ruta2>
'</Config>
'
'Los nodos o etiquetas los llamo a ( <Ruta2></Ruta2> ) y los atributos son aquellos que están dentro de ese nodo ( prueba="prueba atributo" )
'A este archivo lo hemos llamado "archivo.xml" dentro de la carpeta "c:ruta_del_archivo"
'Para cargarlo o leerlo en visual basic es sencillo
'MiXml.Load ("C:ruta_del_archivoarchivo.xml")
'Ahora ya podemos hacer con él lo que queramos.. por ejemplo para saber el número de nodos o etiquetas.
'MsgBox MiXml.childNodes.Length
'Ahora para capturar o leer cada nodo hay dos métodos:
'Uno sería más o menos de tipo index o array es decir leeme lo que tenga el nodo uno.
'' item(1)=<config> childNodes(0)=prueba childNodes(1)=prueba2
'MsgBox MiXml.childNodes.Item(1).childNodes(1).Text
'y el segundo método es con el nombre de ese nodo o etiqueta nodo ( RutaBD ).
'MsgBox MiXml.selectSingleNode("Config/RutaBD").Text
'No voy a poner todos los métodos de caputrar o leer el xml pero si los más importantes que son:
'' capturo atributo
'MsgBox MiXml.selectSingleNode("Config/Ruta2").Attributes(0).Text
'' nombre del nodo
'MsgBox MiXml.childNodes.Item(1).childNodes(1).nodeName
'' muestra todo del nodo 1
'MsgBox MiXml.childNodes.Item(1).Text
'MiXml.Load ("C:\WsRe\OUTPUT.XML")
'ModVariables.XNROCRE = MiXml.selectSingleNode("GenerarRemitoReturn/codRemito").Text
'MsgBox ModVariables.XNROCRE, vbInformation, "CRE"



'Private Sub LeerXMLOutput()

Dim FicheroXML As String
FicheroXML = "C:\WsRe\input.xml"
FicheroXML = "C:\enviado.xml"
Dim oXML As New DOMDocument
oXML.Load sXmlFile
If oXML.parseError.reason <> "" Then
   MsgBox oXML.parseError.reason
   
   Exit Sub
End If

Dim oNode As IXMLDOMNode
Dim oAttr As IXMLDOMAttribute

For Each oNode In oXML.documentElement.childNodes
   Set oAttr = oNode.Attributes.getNamedItem("vconfirm")
   MsgBox oNode.nodeName & ": " & _
   oNode.Text
Next oNode


''También puedes obtener el valor de un nodo en concreto:
'
'Set oNode = oXML.selectSingleNode("//NODO//VALOR")
'If Not oNode Is Nothing Then
'   valorBuscado = oNode.Text
'End If

Exit Sub
PROCERROR:
    MsgBox "LeerOutPutXML"
End Sub




Public Sub LeerOutPutRCEMITIDO()
On Error GoTo PROCERROR

Dim xNroFile As Integer, xArchivo As String, xRuta As String, xEstado As String, XLINEA As String
Dim xStringTotal As String, xrenglon As String, xINicio As Boolean, X As Integer
Dim xCae As String, xVencimiento As Date, XCODIGOQR As String, xx As Integer
Dim xVence As String, XRESPUESTAAFIP As String, XAUTORIZADO As String

xRuta = "C:\WsRe"
If Right(xRuta, 1) = "\" Then
    xRuta = xRuta
Else
    xRuta = xRuta & "\"
End If
xArchivo = "OUTPUT.TXT"
xNroFile = FreeFile()
Open xRuta & xArchivo For Input As xNroFile
Do While Not EOF(xNroFile)
    Input #xNroFile, XLINEA
    xStringTotal = xStringTotal & XLINEA
    'MsgBox XLINEA, vbInformation, "Continuar"

Loop
Close xNroFile

X = InStr(1, xStringTotal, "ESTADO")  'busca cadena2 en cadena1 partiendo de inicio en cadena1, si no se pone inicio empieza en 1, devuelve la primera posicion donde encuentre cadena2
If X = 0 Then
    MsgBox xStringTotal
    MsgBox "NO se pudo obtener Remito Carnico (SP)", vbInformation, "Validar"
    ModVariables.XSEGENERORCOK = False
    Exit Sub
Else
    xEstado = ""
    For I = X To Len(xStringTotal)
        If UCase(Mid(xStringTotal, I, 1)) = ":" Then
            xINicio = True
        End If
        If xINicio Then
            If Mid(xStringTotal, I, 1) = "," Then
                I = I
            End If
            If Mid(xStringTotal, I, 1) <> " " And Mid(xStringTotal, I, 1) <> ":" Then 'And XESTADO <> "" Then
                xEstado = xEstado & Mid(xStringTotal, I, 1)
            End If
            If Mid(xStringTotal, I, 1) = " " And xEstado <> "" Then
                 Exit For
            End If
        End If
    Next
End If

xAumentarRequest = False
X = InStr(1, xStringTotal, "CAE")  'busca cadena2 en cadena1 partiendo de inicio en cadena1, si no se pone inicio empieza en 1, devuelve la primera posicion donde encuentre cadena2
If X = 0 Then
    MsgBox xStringTotal
    MsgBox "NO se pudo obtener Remito Carnico (SP)", vbInformation, "Validar"
    ModVariables.XSEGENERORCOK = False
    ' aca si en string total aparece "El ID de request"
    xx = InStr(1, xStringTotal, "El ID de request")
    If xx > 0 Then
        xAumentarRequest = True
    End If

    Exit Sub
Else
    xCae = ""
    For I = X To Len(xStringTotal)
        If UCase(Mid(xStringTotal, I, 1)) = ":" Then
            xINicio = True
        End If
        If xINicio Then
            If Mid(xStringTotal, I, 1) = "," Then
                I = I
            End If
            If Mid(xStringTotal, I, 1) <> " " And IsNumeric(Mid(xStringTotal, I, 1)) Then 'And XESTADO <> "" Then
                xCae = xCae & Mid(xStringTotal, I, 1)
            End If
            If Mid(xStringTotal, I, 1) = " " And xCae <> "" Then
                 Exit For
            End If
        End If
    Next
End If

    
xAumentarRequest = False
X = InStr(1, xStringTotal, "AUTORIZADO")  'busca cadena2 en cadena1 partiendo de inicio en cadena1, si no se pone inicio empieza en 1, devuelve la primera posicion donde encuentre cadena2
If X = 0 Then
    MsgBox xStringTotal
    ' aca si en string total aparece "El ID de request"
    xx = InStr(1, xStringTotal, "El ID de request")
    If xx > 0 Then
        xAumentarRequest = True
    End If
    MsgBox "NO se pudo obtener Remito Carnico (SP)", vbInformation, "Validar"
    Exit Sub
Else
    XAUTORIZADO = ""
    For I = X To Len(xStringTotal)
        If UCase(Mid(xStringTotal, I, 1)) = ":" Then
            xINicio = True
        End If
        If xINicio Then
            If Mid(xStringTotal, I, 1) = "," Then
                I = I
            End If
            If Mid(xStringTotal, I, 1) <> " " And Mid(xStringTotal, I, 1) <> ":" Then 'And XESTADO <> "" Then
                XAUTORIZADO = XAUTORIZADO & Mid(xStringTotal, I, 1)
            End If
            If Mid(xStringTotal, I, 1) = " " And (XAUTORIZADO <> "" And XAUTORIZADO <> "AUTORIZADO") Then
                 Exit For
            End If
        End If
    Next
End If

X = InStr(1, xStringTotal, "RESPUESTAAFIP")  'busca cadena2 en cadena1 partiendo de inicio en cadena1, si no se pone inicio empieza en 1, devuelve la primera posicion donde encuentre cadena2
If X = 0 Then
    MsgBox xStringTotal
    MsgBox "NO se pudo obtener Remito Carnico (SP)", vbInformation, "Validar"
    Exit Sub
Else
    XRESPUESTAAFIP = ""
    For I = X To Len(xStringTotal)
        If UCase(Mid(xStringTotal, I, 1)) = ":" Then
            xINicio = True
        End If
        If xINicio Then
            If Mid(xStringTotal, I, 1) = "," Then
                I = I
            End If
            If Mid(xStringTotal, I, 1) <> " " And Mid(xStringTotal, I, 1) <> ":" Then 'And XRESPUESTAAFIP <> "" Then
                XRESPUESTAAFIP = XRESPUESTAAFIP & Mid(xStringTotal, I, 1)
            End If
            If Mid(xStringTotal, I, 1) = " " And (XRESPUESTAAFIP <> "" And XRESPUESTAAFIP <> "RESPUESTAAFIP") Then
                 Exit For
            End If
        End If
    Next
End If

If xCae <> "" And ModVariables.xNroInternoRCarnico <> 0 Then
    ModVariables.XSEGENERORCOK = True
    ModVariables.sTrSQL = " UPDATE VTAREMITOS SET CRE= '" & xCae & "' , NroInternoElectronico =" & ModVariables.xNroInternoRCarnico & "  WHERE VTAREMITOS.nroauxremito= " & ModVariables.xNroAuxiliar
    ModSQL.AbrirRSRead2
    Set RsRead2 = Nothing
    ModVariables.sTrSQL = "SELECT * FROM VTAREMITOELECTRONICO WHERE NROAUXREMITO= " & ModVariables.xNroAuxiliar
    ModSQL.AbrirRSOptimista4
    If RsOpt4.RecordCount = 0 Then
        RsOpt4.AddNew
        RsOpt4![NroAuxRemito] = ModVariables.xNroAuxiliar
        RsOpt4![DATOSREC] = Left(xStringTotal, 250)
        RsOpt4.Update
    End If
    RsOpt4.Close
    Set RsOpt4 = Nothing
End If



xINicio = False

MsgBox "Estado: " & xEstado & " CAE: " & xCae & VCRLF & " Autorizado:" & XAUTORIZADO & vbCrLf & " RESPUESTA AFIP: " & XRESPUESTAAFIP, vbInformation, "Resultado"
'
'        If Mid(xStringTotal, I, 1) <> ":" And Mid(xStringTotal, I, 1) <> "#" And Mid(xStringTotal, I, 1) <> "@" Then
'            xrenglon = xrenglon & Mid(xStringTotal, I, 1)
'        End If
'        Select Case UCase(Mid(xStringTotal, I, 1))
'            Case "#"
'                If IsNumeric(xrenglon) Then
'                    xrenglon = Trim(xrenglon)
'                End If
'                xrenglon = xrenglon & vbTab
'            Case "@"
'                Grid.AddItem xrenglon
'                xrenglon = ""
'        End Select
'    End If
'Next
'Grid.AddItem xrenglon
Exit Sub
PROCERROR:
    ModSQL.MensajeError "CmdLeerOutPut.txt"
    Resume Next
End Sub






''''Dim documento As MSXML2.DOMDocument60  '   MSXML2.DOMDocument30
''''
''''
'''''Esta última instrucción en especial me es útil para crear un documento desde cero.
'''''También vamos a necesitar declarar objetos de tipo “nodo”
''''
'''''Dim nodo As MSXML2.IXMLDOMNode
'''''Y objetos tipo “lista de nodos”
'''''Dim listaNodos As MSXML2.IXMLDOMNodeList
'''''Estos objetos nos ayudarán a mantener referencias a la información que necesitemos dentro de la estructura del XML.
'''''Dentro de un XML podemos leer la información seleccionando datos, por ejemplo podemos hacer búsquedas y obtener el resultado en un objeto IXMLDOMNodeList, con esta instrucción:
'''''Set listaNodos = nodo.selectNodes("prueba")
'''''Esto nos devolverá una lista de nodos con los que podemos trabajar, en este caso todos los nodos que se llamen “prueba”.
'''''Existen otros comandos que nos permiten navegar por los nodos de un documento XML.
'''''Texto1.text = nodo.baseName  'Devuelve el nombre del nodo
'''''nodo.childNodes
'''''Permite acceder a los "hijos" de este nodo, de la misma manera se puede saber el numero de hijos con nodo.childNodes.length o acceder a un "hijo" con nodo.childNodes(x) (x es el numero del hijo, empezando desde cero).
'''''nodo.firstChild
'''''es equivalente a nodo.childNodes(0)
'''''nodo.parentNode
'''''apunta hacia el padre de este nodo
'''''nodo.text
'''''devuelve el contenido de texto de un nodo (por ej. <nodo> texto </nodo> )
'''''Set nodo2 = nodo.appendChild(documento.createElement("Datos"))
'''''Esta instruccion crea un elemento y lo agrega dentro del objeto "nodo", el "set" es para que al final de la instrucción nos devuelva en "nodo2" una referencia al nodo recién creado.
'''''Y para terminar la función más importante, con la que grabas un XML.
''''
'''''documento.Save ("archivo.xml")
'''''
'''''Con estos comandos básicos es suficiente para crear y guardar documentos, archivos de configuración y otras monerías para tus programas, si necesitas ayuda o tienes alguna duda, para eso son los comentarios!
''''Private Sub Form_Load()
''''Set documento = New DOMDocument60  ' DOMDocument30
'''''Ya creado, podremos por ejemplo abrir un archivo:
''''documento.Load ("archivo.xml")
''''
'''''O cargar XML desde un string
''''
''''documento.LoadXML ("<Documento></Documento>")
''''documento.Save ("archivo.xml")
''''End Sub
''*//////////////////////////////////////////////////////////////
'
'
'Option Explicit
'
'Private m_objDOMPeople As DOMDocument
'Dim temp As String
''Función api para desabilitar el repintado del control
'Private Declare Function LockWindowUpdate Lib "user32" ( _
'    ByVal hwndLock As Long) As Long
'
'Private Sub Command1_Click()
'    'Cargamos el Xml en el control RichTextBox
'    Call Cargar_XML(App.Path & "\Archivo-ejemplo.xml")
'End Sub
'
''***************************************************************************
''*  Cargar_XML         : Subrutina que carga el XML en el RichTextBox _
'                         (recibe el path del archivo)
''***************************************************************************
'
'Private Sub Cargar_XML(Path_XML As String)
'
'On Error GoTo Err_Sub
'
'    Dim objPeopleRoot As IXMLDOMElement
'    Dim objPersonElement As IXMLDOMElement
'    Dim ObjElement As IXMLDOMNode
'
'    Dim x As IXMLDOMNodeList
'
'    If Len(Dir(Path_XML)) = 0 Then
'       MsgBox "El archivo " & Path_XML & _
'               " No está en el directorio ." & vbNewLine & _
'               " Compruebe la ruta", vbCritical
'       Exit Sub
'    End If
'
'    'Seteamos la variable
'    Set m_objDOMPeople = New DOMDocument
'
'    m_objDOMPeople.resolveExternals = True
'
'    'Para que valide el documento xml
'    m_objDOMPeople.ValidateOnParse = True
'
'    'Carga el documento
'    m_objDOMPeople.async = False
'    Call m_objDOMPeople.Load(Path_XML)
'
'    'Comprobamos si se carga
'    If m_objDOMPeople.parseError.reason <> "" Then
'        ' si hay un error muestra el fallo
'        MsgBox m_objDOMPeople.parseError.reason
'        Exit Sub
'    End If
'
'
'    Set objPeopleRoot = m_objDOMPeople.documentElement
'
'    'nos devuelve el nombre del nodo
'    Debug.Print objPeopleRoot.nodeName
'    'con esto vemos el tipo de nodo
'    Debug.Print objPeopleRoot.nodeType
'    'nos devuelve el valor del nodo si es aplicable
'    Debug.Print objPeopleRoot.nodeValue
'    'Propiedad booleana que nos indica si un nodo tiene "hijos"
'    Debug.Print objPeopleRoot.hasChildNodes
'
'
'    Dim index As Integer
'    Dim lista As IXMLDOMNodeList
'
'    RichTextBox1.text = ""
'    'LockWindowUpdate Me.hWnd
'    Me.Enabled = False
'    'Le pasamos el Nodo a MuestraNodos
'    MuestraNodos m_objDOMPeople.childNodes
'    'LockWindowUpdate 0&
'    Me.Enabled = True
'Exit Sub
'
'Err_Sub:
'
'MsgBox Err.Description
''LockWindowUpdate 0&
'End Sub
'
''Rutina que carga los nodos
'Public Sub MuestraNodos(ByRef Nodos As MSXML2.IXMLDOMNodeList)
'
'
'On Error GoTo Err_Sub
'
'
'    Dim TitulilloA As String
'    Dim TitulilloB As String
'
'    Dim oNodo As MSXML.IXMLDOMNode
'
'    For Each oNodo In Nodos
'
'        If oNodo.nodeType = 1 Then
'
'            TitulilloA = UCase(oNodo.parentNode.nodeName)
'
'                If TitulilloA <> TitulilloB Then
'                    TitulilloB = TitulilloA
'                    temp = temp & vbCrLf & _
'                    UCase(oNodo.parentNode.nodeName) & vbCrLf & vbCrLf
'                End If
'
'        End If
'
'        If oNodo.nodeType = 4 Or oNodo.nodeType = 3 Then
'
'
'            temp = temp & oNodo.parentNode.nodeName & _
'                                                "=" & oNodo.nodeValue & vbCrLf
'
'        End If
'
'        'Si ese nodo tiene hijos (campos) se lo autopasa a la funcion
'        If oNodo.hasChildNodes Then
'            MuestraNodos oNodo.childNodes
'        End If
'    'DoEvents
'    Next oNodo
'
'    RichTextBox1.text = temp
'
'Exit Sub
'
'Err_Sub:
'
'MsgBox Err.Description, vbCritical
''LockWindowUpdate 0&
'
'End Sub
'
'
'Private Sub cmdGenerarXML_Click()
'Call Cargar_XML(App.Path & "\Archivo-ejemplo.xml")
'End Sub
'
'Private Sub Form_Load()
'    Command1.Caption = "Parsear Xml"
'    RichTextBox1.text = " en el App.Path, debe esta el archivo-ejemplo"
'End Sub
'
'Private Sub Form_Resize()
'    On Local Error Resume Next
'    Command1.Top = ScaleHeight - (Command1.Height + 50)
'    Command1.Left = ScaleWidth - (Command1.Width + 50)
'
'    RichTextBox1.Move 0, 0, ScaleWidth, ScaleHeight - (Command1.Height + 50)
'End Sub
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
''//////////////////////////////////////////OTRO EJEMPLO
'
'Este código hace lo que pides. Utiliza ADO para leer un archivo delimitado y genera 2 archivos XML: uno usando directamente el motor ADO y el otro usando el objeto MSXML de Microsoft. Saludos.
'
'Dim appPath As String
'Dim rs As New ADODB.Recordset
'Dim cnn As New ADODB.Connection
'Dim cString
'Dim fld As ADODB.field
'Dim xml As MSXML2.DOMDocument
'Dim record As MSXML2.IXMLDOMElement
'Dim field As MSXML2.IXMLDOMElement
'Dim cDataSecion As MSXML2.IXMLDOMCDATASection
'Set cnn = New ADODB.Connection
'appPath = App.Path
'If Right(appPath, 1) <> "" Then appPath = appPath & ""
'cString = "Provider=MSDASQL;Driver={Microsoft Text Driver (*.txt; *.csv)};DBQ=" & appPath
'cnn.ConnectionString = cString
'cnn.Open
'rs.Open "datos.txt", cnn
'If Dir(appPath & "datos1.xml") <> "" Then
'Kill appPath & "datos1.xml"
'End If
'rs.Save appPath & "datos1.xml", adPersistXML
'Set xml = New MSXML2.DOMDocument
'xml.LoadXML "<records/>"
'Do While Not rs.EOF
'Set record = xml.createElement("record")
'For Each fld In rs.Fields
'fldName = fld.Name
'fldValue = fld.Value & ""
'Set field = xml.createElement(fldName)
'Set cDataSection = xml.createCDATASection(fldValue)
'field.appendChild cDataSection
'record.appendChild field
'Next
'xml.documentElement.appendChild record
'rs.MoveNext
'Loop
'xml.Save appPath & "datos2.xml"
'rs.Close
'Set rs = Nothing
'cnn.Close
'Set cnn = Nothing
'
'
'
'
''////////////////////////////////////////// OTRO EJEMPLO
'Dim docXML As New MSXML2.DOMDocument60
'Private Sub Form_Load()
'If docXML.Load(App.Path & "\coco.xml") Then
'Me.Text1.text = docXML.childNodes.Item(1).childNodes(0).text
'Else
'MsgBox "Archivo de configuración no encontrado o con formato erroneo"
'End
'End If
'End Sub
'
'Private Sub Command1_Click()
'If MsgBox("¿Seguro que desea modificar este nodo?", vbQuestion + vbOKCancel, _
'docXML.childNodes.Item(1).childNodes(0).baseName) = vbOK Then
'docXML.childNodes.Item(1).childNodes(0).text = Me.Text1.text
'docXML.Save (App.Path & "\coco.xml")
'Else
'MsgBox "Cancelado por el usuario"
'End If
'End Sub
'
''Y el XML podria ser:
''
''Código:
''
''<?xml version="1.0" encoding="utf-8"?>
''<Config>
''   <RutaBD>\\Servidor\MiBD.mdb</RutaBD>
''</Config>




'Set fso = New Scripting.FileSystemObject
'Set fXML = fso.CreateTextFile("c:\libros.xml", True)
'
'fXML.WriteLine "<?xml version=""1.0"" encoding=""ISO-8859-1""?>"
'fXML.WriteLine "<libros>"
''Do While Not Rs.EOF
'fXML.WriteLine "<libro>"
'fXML.WriteLine "<id>" & Rs("title_id") & "</id>"
'fXML.WriteLine "<titulo>" & Rs("title") & "</titulo>"
'fXML.WriteLine "<precio>" & Rs("price") & "</precio>"
'fXML.WriteLine "</libro>"
'Rs.MoveNext
'Loop
'fXML.WriteLine "</libros>"



Private Sub CmdGeneraXmlInput_Click()


' original
'Set fso = New Scripting.FileSystemObject
'Set fXML = fso.CreateTextFile("c:\libros.xml", True)
'
'fXML.WriteLine "<?xml version=""1.0"" encoding=""ISO-8859-1""?>"
'fXML.WriteLine "<libros>"
'Do While Not Rs.EOF
'fXML.WriteLine "<libro>"
'fXML.WriteLine "<id>" & Rs("title_id") & "</id>"
'fXML.WriteLine "<titulo>" & Rs("title") & "</titulo>"
'fXML.WriteLine "<precio>" & Rs("price") & "</precio>"
'fXML.WriteLine "</libro>"
'Rs.MoveNext
'Loop
'fXML.WriteLine "</libros>"

End Sub






Private Sub Command2_Click()
Dim MiXml As New DOMDocument60
MiXml.Load ("C:\SISTEMAS\Sistema General\XML\enviado.xml")
MsgBox MiXml.childNodes.Length
MsgBox MiXml.childNodes.Item(1).childNodes(1).Text
MsgBox MiXml.selectSingleNode("solicitarCTGInicialRequestType/datosSolicitarCTGInicial").Text
' capturo atributo
'MsgBox MiXml.selectSingleNode("solicitarCTGInicialRequestType/datosSolicitarCTGInicial").Attributes(0).text
'MsgBox MiXml.selectSingleNode("solicitarCTGInicialRequestType/datosSolicitarCTGInicial").baseName

' nombre del nodo
MsgBox MiXml.childNodes.Item(1).childNodes(1).nodeName

' muestra todo del nodo 1
MsgBox MiXml.childNodes.Item(1).Text

'Comprendido esto pasamos a la modificación que no tendrá efecto hasta guardar el archivo... digamos que se guarda en memoria.
'MiXml.selectSingleNode("Config/RutaBD").text = "nueva prueba"  ' lo cambio
'MiXml.selectSingleNode("Config/Ruta2").Attributes(0).text = "cambio atributo2"  ' cambio atributo

'Podemos incluso crear nodos
'MiXml.selectSingleNode("Config/Ruta2").appendChild MiXml.createElement("prueba")
'MiXml.selectSingleNode("Config/Ruta2/prueba").text = "TEXTO1"

'Remplazarlo ( Importante Tiene que ser en el mismo nivel de Nodo )
'MiXml.childNodes.Item(1).replaceChild MiXml.selectSingleNode("Config/RutaBD"), MiXml.selectSingleNode("Config/Ruta2")

'y Eliminarlo ( El removechild tiene que estar en el mismo nivel que el nodo a borrar )
'MiXml.childNodes.Item(1).childNodes(1).removeChild MiXml.selectSingleNode("Config/Ruta2/prueba")

'Por último podemos guardarlo con el mismo nombre o con otro distinto.
MiXml.Save ("C:ruta_del_archivoarchivo2.xml")


End Sub



Private Sub Form_Load()
TxtCuitCliente.Text = ModVariables.xCuitCliente

strRutaArchivo = "C:\WsRe\"
CHKCMD.Value = 0

'Select Case ModVariables.xProcesoRemElectronico
'    Case "RMObtenerDestinos"
'        If Trim(TxtCuitCliente.Text) <> "" Then
'            Call CmdObtenerDestino_Click
'            Call cmdLeerOutPut_Click
'        End If
'    Case "RMGenerarRCarnico"
'        Call GenerarXMLRemitoCarnico(ModVariables.xNroAuxRemito)
'        'Call cmdLeerOutPut_Click
'End Select
End Sub


Public Sub GenerarXMLRemitoCarnico(xNroAuxRemito As Long)
On Error GoTo PROCERROR
Dim c As Double, xStrCantidad As String, SeparadorDecimal As String
strRutaArchivo = "C:\WsRe\"

'esto es para saber cuál es el separador decimal de la conf. regional
c = 1.11
SeparadorDecimal = Mid(c, 2, 1)


ModVariables.sTrSQL = " SELECT MAX(NROINTERNOELECTRONICO) +1 AS NRO FROM VTAREMITOS "
ModSQL.AbrirRSRead5

ModVariables.xNroInternoRCarnico = IIf(IsNull(RsRead5![Nro]), 1, RsRead5![Nro])
If xAumentarRequest Then
    ModVariables.xNroInternoRCarnico = ModVariables.xNroInternoRCarnico + 5
End If
RsRead5.Close
Set RsRead5 = Nothing
ModSQL.VtaTraeRemito ModVariables.xNroAuxiliar
ModSQL.AbrirRSRead
With RsRead
    If .RecordCount > 0 Then
'        ModVariables.XNROINTERNOREMITO = ![NroRMElectronico] ' se obtiene del nro mas alto +1
        ModVariables.XPTOVTAREMITOCARNICO = Left(![NroRemito], 4)
        ModVariables.xCuitCliente = ![Cuit]
        ModVariables.XCODTIPOCLIENTEREMCARNI = ![TipoReceptorRM]   ' VER SI SE GRABA EN TABLA CLIENTES O EN EL REMITO
        ModVariables.XCATEGOCLIENTEREMCARNI = 3 ' 3=abastecedor 1=frigorifico  7= camara o deposito
        If ChkCamara.Value = 1 Then
            ModVariables.XCATEGOCLIENTEREMCARNI = 7
        End If
        ModVariables.XCODDESTINOREMCARNI = ![CodigoDestinoRM] ' 3 ' VER SI SE GRABA EN TABLA CLIENTES O EN EL REMITO
        ModVariables.XFECHAVIAJEREMCARNI = Format(IIf(IsNull(![FechaInicioViaje]), Date, ![FechaInicioViaje]), "YYYY-MM-DD")   ' VALIDAR QUE SE GRABE
        ModVariables.XKMARECORRERREMCARNI = ![KmARecorrerRM]   ' GRABAR EN REMITO
        ModVariables.xPatente = Trim(![PatenteyOtros])   ' GRABAR EN REMITO
        ModVariables.xCuitTransporte = ![CuitTransporte]  ' GRABAR EN REMITO
        ModVariables.xCuitChofer = ![CuilChofer]  ' GRABAR EN REMITO
    Else
        .Close
        MsgBox "No se encontró el Nº de Remito", vbInformation, "Atención"
        Unload Me
    End If
End With

'txtcuitempresa.Text = ModVariables.xCuitEmpresa
xLineaRepresentada = "<cuitRepresentada>" & Trim(ModVariables.xCuitEmpresa) & "</cuitRepresentada>"
TxtCuitCliente.Text = ModVariables.xCuitCliente
xLineaCliente = "<cuitTitularDomicilio>" & Trim(TxtCuitCliente.Text) & "</cuitTitularDomicilio>"

Set FSO = New Scripting.FileSystemObject
Set FXML = FSO.CreateTextFile("c:\WsRe\input.xml", True)
FXML.WriteLine "<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/' xmlns:rem='http://ar.gob.afip.wsremcarne/RemCarneService/'>"
FXML.WriteLine "<soapenv:Header/>"
FXML.WriteLine "<soapenv:Body>"
FXML.WriteLine "<rem:generarRemitoRequest>"
FXML.WriteLine "<authRequest>"
FXML.WriteLine "<token>@TOKEN</token>"
FXML.WriteLine "<sign>@SIGN</sign>"
FXML.WriteLine "<cuitRepresentada>" & Trim(ModVariables.xCuitEmpresa) & "</cuitRepresentada>"
FXML.WriteLine "</authRequest>"
FXML.WriteLine "<idReq>" & ModVariables.xNroInternoRCarnico & "</idReq>"    ' NRO INTERNO DEL REMITO
FXML.WriteLine "<remito>"
FXML.WriteLine "<tipoComprobante>995</tipoComprobante>"
FXML.WriteLine "<tipoMovimiento>ENV</tipoMovimiento>"
FXML.WriteLine "<categoriaEmisor>1</categoriaEmisor>"
FXML.WriteLine "<puntoEmision>" & ModVariables.XPTOVTAREMITOCARNICO & "</puntoEmision>"
FXML.WriteLine "<cuitTitularMercaderia>" & Trim(ModVariables.xCuitEmpresa) & "</cuitTitularMercaderia>"
FXML.WriteLine "<tipoReceptor>" & Left(Trim(ModVariables.XCODTIPOCLIENTEREMCARNI), 2) & "</tipoReceptor>"
FXML.WriteLine "<categoriaReceptor>" & ModVariables.XCATEGOCLIENTEREMCARNI & "</categoriaReceptor>"
FXML.WriteLine "<cuitReceptor>" & ModVariables.xCuitCliente & "</cuitReceptor>"
FXML.WriteLine "<codDomDestino>" & ModVariables.XCODDESTINOREMCARNI & "</codDomDestino>"
FXML.WriteLine "<viaje>"
FXML.WriteLine "<cuitTransportista>" & ModVariables.xCuitTransporte & "</cuitTransportista>"
FXML.WriteLine "<cuitConductor>" & ModVariables.xCuitChofer & "</cuitConductor>"
FXML.WriteLine "<fechaInicioViaje>" & ModVariables.XFECHAVIAJEREMCARNI & "</fechaInicioViaje>"
FXML.WriteLine "<distanciaKm>" & ModVariables.XKMARECORRERREMCARNI & "</distanciaKm>"
FXML.WriteLine "<vehiculo>"
FXML.WriteLine "<dominioVehiculo>" & Left(Trim(ModVariables.xPatente), 7) & "</dominioVehiculo>" ' MAXIMO DE 7 CARACTERES
FXML.WriteLine "</vehiculo>"
FXML.WriteLine "</viaje>"
' hacer un bucle por cada linea de detalle del remito
ModSQL.VtaLeeDetaRemitoCarnico ModVariables.xNroAuxiliar
ModSQL.AbrirRSRead2
FXML.WriteLine "<arrayMercaderias>"
For I = 1 To RsRead2.RecordCount
    xStrCantidad = Format(RsRead2![CantidadRemitada], "##0.00")
    xStrCantidad = Replace(xStrCantidad, ",", ".")

    FXML.WriteLine "<mercaderia>"
    FXML.WriteLine "<orden>" & I & "</orden>"
    FXML.WriteLine "<codTipoProd>" & RsRead2![CodSenasa] & "</codTipoProd>"
    'FXML.WriteLine "<tropa>" & RsRead2![Tropa] & "</tropa>"
    FXML.WriteLine "<tropa>" & 1 & "</tropa>"  ' NO PUEDE SER 0
    FXML.WriteLine "<kilos>" & xStrCantidad & "</kilos>"   ' SEPARADOR DECIMAL DEBE SER PUNTO
    FXML.WriteLine "<unidades>" & RsRead2![Bultos] & "</unidades>"
    FXML.WriteLine "</mercaderia>"
RsRead2.MoveNext
Next
RsRead2.Close
Set RsRead2 = Nothing
FXML.WriteLine "</arrayMercaderias>"
FXML.WriteLine "</remito>"
FXML.WriteLine "</rem:generarRemitoRequest>"
FXML.WriteLine "</soapenv:Body>"
FXML.WriteLine "</soapenv:Envelope>"

strRutaArchivo = strRutaArchivo & "VALIDAR.BAT"
'Call Ejecutar_shell(strRutaArchivo)---

If CHKCMD.Value = 0 Then
    Shell strRutaArchivo, vbMinimizedFocus
Else
    Shell ("cmd.exe /k" & strRutaArchivo)
End If
Me.MousePointer = 11
Call Sleep(3000) 'espera por 10 segundos
Me.MousePointer = 0

'MsgBox "Pulse una tecla para continuar", vbInformation, "Generar Remito Electronico"

Exit Sub
PROCERROR:
    ModSQL.MensajeError "GenerarXMLREMITOCARNICO"
    Resume Next
End Sub

Public Sub GenerarXMLRemitoCarnicoAA(xNroAuxRemito As Long)
On Error GoTo PROCERROR
Dim c As Double, xStrCantidad As String, SeparadorDecimal As String
'esto es para saber cuál es el separador decimal de la conf. regional
c = 1.11
SeparadorDecimal = Mid(c, 2, 1)


ModSQL.VtaTraeRemito ModVariables.xNroAuxiliar
ModSQL.AbrirRSRead
With RsRead
    If .RecordCount > 0 Then
'        ModVariables.XNROINTERNOREMITO = ![NroRMElectronico] ' se obtiene del nro mas alto +1
        ModVariables.XPTOVTAREMITOCARNICO = Left(![NroRemito], 4)
        ModVariables.xCuitCliente = ![Cuit]
        ModVariables.XCODTIPOCLIENTEREMCARNI = ![TipoReceptorRM]   ' ' VER SI SE GRABA EN TABLA CLIENTES O EN EL REMITO
        ModVariables.XCODDESTINOREMCARNI = ![CodigoDestinoRM] ' 3 ' VER SI SE GRABA EN TABLA CLIENTES O EN EL REMITO
        ModVariables.XFECHAVIAJEREMCARNI = Format(IIf(IsNull(![FechaInicioViaje]), Date, ![FechaInicioViaje]), "YYYY-MM-DD")   ' VALIDAR QUE SE GRABE
        ModVariables.XKMARECORRERREMCARNI = ![KmARecorrerRM]   ' GRABAR EN REMITO
        ModVariables.xPatente = Trim(![PatenteyOtros])   ' GRABAR EN REMITO
        ModVariables.xCuitTransporte = ![CuitTransporte]  ' GRABAR EN REMITO
        ModVariables.xCuitChofer = ![CuilChofer]  ' GRABAR EN REMITO
    Else
        .Close
        MsgBox "No se encontró el Nº de Remito", vbInformation, "Atención"
        Unload Me
    End If
End With

'txtcuitempresa.Text = ModVariables.xCuitEmpresa
xLineaRepresentada = "<cuitRepresentada>" & Trim(ModVariables.xCuitEmpresa) & "</cuitRepresentada>"
TxtCuitCliente.Text = ModVariables.xCuitCliente
xLineaCliente = "<cuitTitularDomicilio>" & Trim(TxtCuitCliente.Text) & "</cuitTitularDomicilio>"

Set FSO = New Scripting.FileSystemObject
Set FXML = FSO.CreateTextFile("c:\input.xml", True)
FXML.WriteLine "<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/' xmlns:rem='http://ar.gob.afip.wsremcarne/RemCarneService/'>"
FXML.WriteLine "    <soapenv:Header/>"
FXML.WriteLine "    <soapenv:Body>"
FXML.WriteLine "        <rem:generarRemitoRequest>"
FXML.WriteLine "            <authRequest>"
FXML.WriteLine "                <token>@TOKEN</token>"
FXML.WriteLine "                <sign>@SIGN</sign>"
FXML.WriteLine "                <cuitRepresentada>" & Trim(ModVariables.xCuitEmpresa) & "</cuitRepresentada>"
FXML.WriteLine "            </authRequest>"
FXML.WriteLine "            <idReq>1</idReq>"   ' NRO INTERNO DEL REMITO
FXML.WriteLine "            <remito>"
FXML.WriteLine "                <tipoComprobante>995</tipoComprobante>"
FXML.WriteLine "                <tipoMovimiento>ENV</tipoMovimiento>"
FXML.WriteLine "                <categoriaEmisor>1</categoriaEmisor>"
FXML.WriteLine "                <puntoEmision>" & ModVariables.XPTOVTAREMITOCARNICO & "</puntoEmision>"
FXML.WriteLine "                <cuitTitularMercaderia>" & Trim(ModVariables.xCuitEmpresa) & "</cuitTitularMercaderia>"
FXML.WriteLine "                <tipoReceptor>" & ModVariables.XCODTIPOCLIENTEREMCARNI & "</tipoReceptor>"
FXML.WriteLine "                <categoriaReceptor>" & ModVariables.XCODTIPOCLIENTEREMCARNI & "</categoriaReceptor>"
FXML.WriteLine "                <cuitReceptor>" & ModVariables.xCuitCliente & "</cuitReceptor>"
FXML.WriteLine "                <codDomDestino>" & ModVariables.XCODDESTINOREMCARNI & "</codDomDestino>"
FXML.WriteLine "                <viaje>"
FXML.WriteLine "                    <cuitTransportista>" & ModVariables.xCuitTransporte & "</cuitTransportista>"
FXML.WriteLine "                    <cuitConductor>" & ModVariables.xCuitChofer & "</cuitConductor>"
FXML.WriteLine "                    <fechaInicioViaje>" & ModVariables.XFECHAVIAJEREMCARNI & "</fechaInicioViaje>"
FXML.WriteLine "                    <distanciaKm>" & ModVariables.XKMARECORRERREMCARNI & "</distanciaKm>"
FXML.WriteLine "                    <vehiculo>"
FXML.WriteLine "                        <dominioVehiculo>" & ModVariables.xPatente & "</dominioVehiculo>"
FXML.WriteLine "                    </vehiculo>"
FXML.WriteLine "                </viaje>"
' hacer un bucle por cada linea de detalle del remito
ModSQL.VtaLeeDetaRemitoCarnico ModVariables.xNroAuxiliar
ModSQL.AbrirRSRead2
FXML.WriteLine "                <arrayMercaderias>"
For I = 1 To RsRead2.RecordCount
    xStrCantidad = Format(RsRead2![CantidadRemitada], "##0.00")
    xStrCantidad = Replace(xStrCantidad, ",", ".")

    FXML.WriteLine "                    <mercaderia>"
    FXML.WriteLine "                        <orden>I</orden>"
    FXML.WriteLine "                        <codTipoProd>" & RsRead2![CodSenasa] & "</codTipoProd>"
    FXML.WriteLine "                        <tropa>" & RsRead2![Tropa] & "</tropa>"
    FXML.WriteLine "                        <kilos>" & xStrCantidad & "</kilos>"
    FXML.WriteLine "                        <unidades>" & RsRead2![Bultos] & "</unidades>"
    FXML.WriteLine "                    </mercaderia>"
RsRead2.MoveNext
Next
RsRead2.Close
Set RsRead2 = Nothing
FXML.WriteLine "                </arrayMercaderias>"
FXML.WriteLine "            </remito>"
FXML.WriteLine "        </rem:generarRemitoRequest>"
FXML.WriteLine "    </soapenv:Body>"
FXML.WriteLine "</soapenv:Envelope>"




'FXML.WriteLine "<token>@TOKEN</token>"  ' VER QUE VALORES VAN
'FXML.WriteLine "<sign>@SIGN</sign>"     ' VER QUE VALORES VAN
'FXML.WriteLine "<cuitRepresentada>" & Trim(ModVariables.xCuitEmpresa) & "</cuitRepresentada>"
'FXML.WriteLine "</authRequest>"
'FXML.WriteLine "<idReq>1</idReq>"  ' PONER NRO INTERNO CORRELATIVO CON EL ULTIMO AUTORIZADO
'FXML.WriteLine "<tipoComprobante>995</tipoComprobante>"
'FXML.WriteLine "<tipoMovimiento>ENV</tipoMovimiento>"
'FXML.WriteLine "<categoriaEmisor>1</categoriaEmisor>"
'FXML.WriteLine "<puntoEmision>" & ModVariables.XPTOVTAREMITOCARNICO & "</puntoEmision>"
'FXML.WriteLine "<cuitTitularMercaderia>" & Trim(ModVariables.xCuitEmpresa) & "</cuitTitularMercaderia>"
'FXML.WriteLine "<tipoReceptor>MI</tipoReceptor>"
'FXML.WriteLine "<categoriaReceptor>" & ModVariables.XCODTIPOCLIENTEREMCARNI & "</categoriaReceptor>"  ' PONER EN TABLA DE CLIENTES
'FXML.WriteLine "<cuitReceptor>" & ModVariables.xCuitCliente & "</cuitReceptor>"
'FXML.WriteLine "<codDomDestino>" & ModVariables.XCODDESTINOREMCARNI & "</codDomDestino>"  ' TENER GRABADO EN REMITO
'FXML.WriteLine "<viaje>"
'FXML.WriteLine "<cuitTransportista>" & ModVariables.xCuitTransporte & "</cuitTransportista>"  ' TENER GRABADO EN REMITO
'FXML.WriteLine "<cuitConductor>" & ModVariables.xCuitChofer & "</cuitConductor>"  ' TENER GRABADO EN REMITO
'FXML.WriteLine "<fechaInicioViaje>" & ModVariables.XFECHAVIAJEREMCARNI & "</fechaInicioViaje>" ' TENER GRABADO EN REMITO
'FXML.WriteLine "<distanciaKm>" & ModVariables.XKMARECORRERREMCARNI & "</distanciaKm>"  ' TENER GRABADO EN REMITO
'FXML.WriteLine "<vehiculo>"
'FXML.WriteLine "<dominioVehiculo>" & ModVariables.xPatente & "</dominioVehiculo>"  ' TENER GRABADO EN REMITO
'FXML.WriteLine "</vehiculo>"
'FXML.WriteLine "</viaje>"
'FXML.WriteLine "<arrayMercaderias>"
'' hacer un bucle por cada linea de detalle del remito
'ModSQL.VtaLeeDetaRemitoCarnico ModVariables.xNroAuxiliar
'ModSQL.AbrirRSRead2
'For I = 1 To RsRead2.RecordCount
'    FXML.WriteLine "<mercaderia>"
'    FXML.WriteLine "<orden>" & I & "</orden>"
'    FXML.WriteLine "<codTipoProd>" & RsRead2![CodSenasa] & "</codTipoProd>"   ' VALIDAR AL GRABAR REMITO
'    FXML.WriteLine "<tropa>" & RsRead2![Tropa] & "</tropa>"                   ' PEDIR AL ARMAR REMITO
'    FXML.WriteLine "<kilos>" & RsRead2![CantidadRemitada] & "</kilos>"
'    FXML.WriteLine "<unidades>" & RsRead2![Bultos] & "</unidades>"
'    FXML.WriteLine "</mercaderia>"
'    RsRead2.MoveNext
'Next
'RsRead2.Close
'Set RsRead2 = Nothing
'FXML.WriteLine "</arrayMercaderias>"
'FXML.WriteLine "</remito>"
'FXML.WriteLine "</rem:generarRemitoRequest>"
'FXML.WriteLine "</soapenv:Body>"
'FXML.WriteLine "</soapenv:Envelope>"

Exit Sub
PROCERROR:
    ModSQL.MensajeError "GenerarXMLREMITOCARNICO"
    Resume Next
End Sub

Private Sub CmdRemitoStockPropio_Click()
On Error GoTo PROCERROR
Dim xCuitCte As String, xNroRemito As Long, xPtoVenta As Integer, xTipoRemito As String
Dim xTipoFrigorifico As Integer, xTipoCliente As String, xCategoriaReceptor As String
Dim xCodDomiDestino As String, xCuitTransporte As String, xCuilChofer As String, SeparadorDecimal As String
Dim xFechadelViaje As Date, xKmARecorrer As Long, xPatente As String, c As Double, xStrCantidad As String
'esto es para saber cuál es el separador decimal de la conf. regional
c = 1.11
SeparadorDecimal = Mid(c, 2, 1)

ModSQL.VtaTraeRemito ModVariables.xNroAuxiliar
ModSQL.AbrirRSRead5
With RsRead5
If .RecordCount = 1 Then
    xCuitCte = IIf(IsNull(![Cuit]), "", ![Cuit])
    xNroRemito = Val(Right(![NroRemito], 8))
    xPtoVenta = Val(Left(![NroRemito], 4))
    xTipoRemito = "ENV" 'ENV=ENVIO COMUN o PLA: Retiro en Planta
    If ModVariables.xTipoRemitoElectronico = "PLA" Then
        xTipoRemito = "PLA"
    End If
    xTipoFrigorifico = 1 ' TIPO DE FRIGORIFICO : 1=FAENADOR o 4=ABASTECEDOR
    xTipoCliente = "MI" 'TIPO MI=CLIENTE o RP=REPARTO
    xCategoriaReceptor = ![TipoReceptorRM]  ' categoria receptor cliente
    xCodDomiDestino = ModVariables.XCODDESTINOREMCARNI  'comdomdestino
    xCuitTransporte = Trim(![CuitTransporte])
    xCuilChofer = Trim(![CuilChofer])
    xFechadelViaje = Format(IIf(IsNull(![FechaInicioViaje]), Date, ![FechaInicioViaje]), "YYYY-MM-DD")
    xKmARecorrer = ![KmARecorrerRM]
    xPatente = Trim(![PatenteyOtros])
End If
.Close
End With
Set RsRead5 = Nothing
    
ModSQL.VtaLeeDetaRemitoCarnico ModVariables.xNroAuxiliar
 
 
Set FSO = New Scripting.FileSystemObject
Set FXML = FSO.CreateTextFile("c:\input.xml", True)
FXML.WriteLine "<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/' xmlns:rem='http://ar.gob.afip.wsremcarne/RemCarneService/'>"
FXML.WriteLine "<soapenv:Header/>"
FXML.WriteLine "<soapenv:Body>"
FXML.WriteLine "<rem:generarRemitoRequest>"
FXML.WriteLine "<authRequest>"
FXML.WriteLine "<token>@TOKEN</token>"
FXML.WriteLine "<sign>@SIGN</sign>"
FXML.WriteLine "<cuitRepresentada>" & ModVariables.xCuitEmpresa & "</cuitRepresentada>"
FXML.WriteLine "</authRequest>"
FXML.WriteLine "<idReq>" & xNroRemito & "</idReq>"
FXML.WriteLine "<remito>"
FXML.WriteLine "<tipoComprobante>995</tipoComprobante>"
FXML.WriteLine "<tipoMovimiento>" & xTipoRemito & "</tipoMovimiento>"
FXML.WriteLine "<categoriaEmisor>" & xTipoFrigorifico & "</categoriaEmisor>"
FXML.WriteLine "<puntoEmision>" & xPtoVenta & "</puntoEmision>"
FXML.WriteLine "<cuitTitularMercaderia>" & ModVariables.xCuitEmpresa & "</cuitTitularMercaderia>"
FXML.WriteLine "<tipoReceptor>" & xTipoCliente & "</tipoReceptor>"
FXML.WriteLine "<categoriaReceptor>" & xCategoriaReceptor & "</categoriaReceptor>"
FXML.WriteLine "<cuitReceptor>" & xCuitCte & "</cuitReceptor>"
FXML.WriteLine "<codDomDestino>" & xCodDomiDestino & "</codDomDestino>"
FXML.WriteLine "<viaje>"
FXML.WriteLine "<cuitTransportista>" & xCuitTransporte & "</cuitTransportista>"
FXML.WriteLine "<cuitConductor>" & xCuilChofer & "</cuitConductor>"
FXML.WriteLine "<fechaInicioViaje>" & xFechadelViaje & "</fechaInicioViaje>"
FXML.WriteLine "<distanciaKm>" & xKmARecorrer & "</distanciaKm>"
FXML.WriteLine "<vehiculo>"
FXML.WriteLine "<dominioVehiculo>" & xPatente & "</dominioVehiculo>"
FXML.WriteLine "</vehiculo>"
FXML.WriteLine "</viaje>"
FXML.WriteLine "<arrayMercaderias>"
ModSQL.AbrirRSRead4
With RsRead4
    Do Until .EOF
        xOrden = xOrden + 1
        xStrCantidad = Format(![CantidadRemitada], "##0.00")
        xStrCantidad = Replace(xStrCantidad, ",", ".")
        FXML.WriteLine "<mercaderia>"
        FXML.WriteLine "<orden>" & xOrden & "</orden>"
        FXML.WriteLine "<codTipoProd>" & ![CodSenasa] & "</codTipoProd>"
        'FXML.WriteLine "<tropa>" & ![Tropa] & "</tropa>"
        FXML.WriteLine "<tropa>" & 1 & "</tropa>"
        FXML.WriteLine "<kilos>" & Format(![CantidadRemitada], "##0.00") & "</kilos>"
        FXML.WriteLine "<unidades>" & ![Bultos] & "</unidades>"
        FXML.WriteLine "</mercaderia>"
        .MoveNext
    Loop
    .Close
    Set RsRead4 = Nothing
End With
FXML.WriteLine "</arrayMercaderias>"
FXML.WriteLine "</remito>"
FXML.WriteLine "</rem:generarRemitoRequest>"
FXML.WriteLine "</soapenv:Body>"
FXML.WriteLine "</soapenv:Envelope>"
Exit Sub
PROCERROR:
    ModSQL.MensajeError "CmdRemitoStockPropio"
    Resume Next
End Sub


Private Sub Grid_DblClick()
On Error GoTo PROCERROR
With GRid
    If GRid.Row > 0 Then
        ModVariables.XCODDESTINOREMCARNI = Trim(.TextMatrix(GRid.Row, .ColIndex("Codigo")))
        ModVariables.xDestinoRemitoCarnico = Trim(.TextMatrix(GRid.Row, .ColIndex("Descripcion")))
        ModVariables.VuelveDeAyuda = True
        Unload Me
    End If
End With
Exit Sub
PROCERROR:
    ModSQL.MensajeError "GridBusqueda_DblClick"
    Resume Next
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn
    '    gridbusqueda_DblClick
    Case vbKeyEscape
        Sendkeys "%n"
End Select
End Sub

Private Sub cmdCancelar_Click()
    ModVariables.XCODDESTINOREMCARNI = 0
    ModVariables.xDestinoRemitoCarnico = ""
    Unload Me
End Sub




'<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:rem="http://ar.gob.afip.wsremcarne/RemCarneService/">
'    <soapenv:Header/>
'    <soapenv:Body>
'        <rem:generarRemitoRequest>
'            <authRequest>
'                <token>@TOKEN</token>
'                <sign>@SIGN</sign>
'                <cuitRepresentada>11111111111</cuitRepresentada>
'            </authRequest>
'            <idReq>410</idReq>
'            <remito>
'                <tipoComprobante>995</tipoComprobante>
'                <tipoMovimiento>ENV</tipoMovimiento>
'                <categoriaEmisor>1</categoriaEmisor>
'                <puntoEmision>9000</puntoEmision>
'                <cuitTitularMercaderia>11111111111</cuitTitularMercaderia>
'                <tipoReceptor>MI</tipoReceptor>
'                <categoriaReceptor>3</categoriaReceptor>
'                <cuitReceptor>22222222222</cuitReceptor>
'                <codDomDestino>3</codDomDestino>
'                <viaje>
'                    <cuitTransportista>33333333333</cuitTransportista>
'                    <cuitConductor>44444444444</cuitConductor>
'                    <fechaInicioViaje>2019-02-14</fechaInicioViaje>
'                    <distanciaKm>670</distanciaKm>
'                    <vehiculo>
'                        <dominioVehiculo>ARF458</dominioVehiculo>
'                    </vehiculo>
'                </viaje>
'                <arrayMercaderias>
'                    <mercaderia>
'                        <orden>1</orden>
'                        <codTipoProd>6.24</codTipoProd>
'                        <tropa>121</tropa>
'                        <kilos>50</kilos>
'                        <unidades>4</unidades>
'                    </mercaderia>
'                    <mercaderia>
'                        <orden>2</orden>
'                        <codTipoProd>6.9</codTipoProd>
'                        <tropa>121</tropa>
'                        <kilos>50</kilos>
'                        <unidades>1</unidades>
'                    </mercaderia>
'                </arrayMercaderias>
'            </remito>
'        </rem:generarRemitoRequest>
'    </soapenv:Body>
'</soapenv:Envelope>

Private Sub cmdGrabarDestinos_Click()
    Call GrabarDestinosRemitoElectronico
End Sub

Private Sub GrabarDestinosRemitoElectronico()
On Error GoTo PROCERROR
Me.MousePointer = 11
For I = 1 To GRid.Rows - 1
    ModVariables.sTrSQL = "SELECT * FROM RMELECTRONICODESTINOS WHERE CUIT = '" & TxtCuitCliente.Text & "' AND CODIGO= " & GRid.TextMatrix(I, GRid.ColIndex("Codigo"))
    ModSQL.AbrirRSOptimista3
    With RsOpt3
        If .RecordCount = 0 Then
            .AddNew
            ![Cuit] = TxtCuitCliente.Text
            ![Codigo] = GRid.TextMatrix(I, GRid.ColIndex("Codigo"))
        End If
        ![Cliente] = ModVariables.xCodCliente
        ![nombredestino] = Left(GRid.TextMatrix(I, GRid.ColIndex("Descripcion")), 100)
        .Update
        .Close
    End With
Next
Me.MousePointer = 0
MsgBox "Los destinos fueron grabados", vbInformation, "AViso"
Exit Sub
PROCERROR:
    ModSQL.MensajeError "GrabarDestinosREmitosElectronico"
End Sub


Private Sub CambiarTransporte()
On Error GoTo PROCERROR

strRutaArchivo = "C:\WsRe\"
 
Dim xLineaRepresentada As String
'txtcuitempresa.text = ModVariables.xCuitEmpresa

Set FSO = New Scripting.FileSystemObject
Set FXML = FSO.CreateTextFile("c:\wsre\input.xml", True)
FXML.WriteLine "<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/' xmlns:rem='http://ar.gob.afip.wsremcarne/RemCarneService/'> "
FXML.WriteLine "<soapenv:Header/>"
FXML.WriteLine "<soapenv:Body>"
FXML.WriteLine "<rem:modificarViajeRequest>"
FXML.WriteLine "<authRequest>"
FXML.WriteLine "<token>@TOKEN</token>"
FXML.WriteLine "<sign>@SIGN</sign>"
FXML.WriteLine "<cuitRepresentada>" & Trim(ModVariables.xCuitEmpresa) & "</cuitRepresentada>"
FXML.WriteLine "</authRequest>"
FXML.WriteLine "<codRemito>Cód CRE del Remito</codRemito>"
'<!--Optional:-->
FXML.WriteLine "<cuitTransportista>CUIT Transporte</cuitTransportista>"
'<!--Optional:-->
FXML.WriteLine "<cuitConductor>CUIT Chofer</cuitConductor>"
'<!--Optional:-->
FXML.WriteLine "<vehiculo>"
FXML.WriteLine "<dominioVehiculo>Patente</dominioVehiculo>"
'<!--Optional:-->
FXML.WriteLine "<dominioAcoplado>Patente Acoplado</dominioAcoplado>"
FXML.WriteLine "</vehiculo>"
FXML.WriteLine "</rem:modificarViajeRequest>"

FXML.WriteLine "</soapenv:Body>"
FXML.WriteLine "</soapenv:Envelope>"

'Shell strRutaArchivo & "VALIDAR.BAT", vbMaximizedFocus

strRutaArchivo = strRutaArchivo & "VALIDAR.BAT"
' Call Ejecutar_shell(strRutaArchivo)
If CHKCMD.Value = 0 Then
    Shell strRutaArchivo, vbMinimizedNoFocus
Else
    Shell ("cmd.exe /k" & strRutaArchivo), vbMinimizedNoFocus
End If
Exit Sub
PROCERROR:
    ModSQL.MensajeError "CmdObtenerDestino"
    Resume Next
End Sub


Public Sub GenerarRemitoHarinero(xNroAuxRemito As Long)
On Error GoTo PROCERROR
Dim bresultado  As Boolean, XNROREM As String
Dim c As Double, xStrCantidad As String, SeparadorDecimal As String, xNroRemito As String
'esto es para saber cuál es el separador decimal de la conf. regional
c = 1.11
SeparadorDecimal = Mid(c, 2, 1)

' TIPOS DE EMBALAJE
'1       a granel
'6   Bidón
'4   Big Bag
'3   Bolsa <= 1kg
'2   Bolsa > 1kg
'8   Botella
'18  Bultos
'11  Caja
'10  Cajón
'12  Contenedor
'14  Fardo
'5   Frasco
'7   Lata Cuñete
'15  Pallets
'13  Paquete
'17  Pieza
'16  Rollo
'9   Tambor
' UNIDADES DE MEDIDA
'1   Kilogramo
'1   Kilogramo
'4   Litro
'5   Metro cúbico
'2   Tonelada
'2   Unidad
'3   Unidad

ModVariables.sTrSQL = " SELECT MAX(NROINTERNOELECTRONICO) +1 AS NRO FROM VTAREMITOS "
ModSQL.AbrirRSRead5
ModVariables.xNroInternoRCarnico = IIf(IsNull(RsRead5![Nro]), 1, RsRead5![Nro])
'If xAumentarRequest Then
'    ModVariables.xNroInternoRCarnico = ModVariables.xNroInternoRCarnico + 1
'End If

XNROIDRH = ModVariables.xNroInternoRCarnico + ModVariables.xSumarANroInterno


RsRead5.Close
Set RsRead5 = Nothing
ModSQL.VtaTraeRemito xNroAuxRemito
ModSQL.AbrirRSRead
With RsRead
    If .RecordCount > 0 Then
'        ModVariables.XNROINTERNOREMITO = ![NroRMElectronico] ' se obtiene del nro mas alto +1
        ModVariables.XPTOVTAREMITOHARINERO = Left(![NroRemito], 4)
        XNROREM = Right(![NroRemito], 5)
        ModVariables.xCuitCliente = ![Cuit]
        ModVariables.XCODTIPOCLIENTEREMCARNI = Trim(![TipoReceptorRM])    ' VER SI SE GRABA EN TABLA CLIENTES O EN EL REMITO
        ModVariables.XCATEGOCLIENTEREMCARNI = 3 ' 3=abastecedor 1=frigorifico  7= camara o deposito
        If ChkCamara.Value = 1 Then
            ModVariables.XCATEGOCLIENTEREMCARNI = 7
        End If
        ModVariables.XCODDESTINOREMCARNI = ![CodigoDestinoRM] ' 3 ' VER SI SE GRABA EN TABLA CLIENTES O EN EL REMITO
        ModVariables.XFECHAVIAJEREMCARNI = Format(IIf(IsNull(![FechaInicioViaje]), Date, ![FechaInicioViaje]), "YYYYMMDD")   ' VALIDAR QUE SE GRABE
        ModVariables.XKMARECORRERREMCARNI = ![KmARecorrerRM]   ' GRABAR EN REMITO
        ModVariables.xPatente = Trim(![PatenteyOtros])   ' GRABAR EN REMITO
        ModVariables.xPatenteAcoplado = Trim(![Acoplado])   ' GRABAR EN REMITO
        ModVariables.xCuitTransporte = ![CuitTransporte]  ' GRABAR EN REMITO
        ModVariables.xCuitChofer = ![CuilChofer]  ' GRABAR EN REMITO
    Else
        .Close
        MsgBox "No se encontró el Nº de Remito", vbInformation, "Atención"
        Unload Me
    End If
End With

TxtCuitCliente.Text = ModVariables.xCuitCliente
'993 Remito Electrónico Harinero Automotor
'994 Remito Electrónico Harinero Ferroviario
'ModVariables.XNROESTABRUCA = 5073
' punto de venta iman=0009
'ModVariables.XPTOVTAREMITOHARINERO

Fe.ArchivoXMLEnviado = "C:\_txtsistema\rhenviado.xml"
Fe.ArchivoXMLRecibido = "C:\_txtsistemaN\rhrecibido.xml"
Fe.ArchivoCertificadoPassword = ""
'If Fe.rcObtenerTicketAcceso() Then
    bresultado = Fe.rcCargarPropiedad("LIMPIAR", "", 0, 0)

    bresultado = bresultado And Fe.rcCargarPropiedad("tipoMovimiento", "ENV", 0, 0)
    bresultado = bresultado And Fe.rcCargarPropiedad("esentregamostrador", "N", 0, 0)
    bresultado = bresultado And Fe.rcCargarPropiedad("esentregamostradorNO", "V", 0, 0)
    bresultado = bresultado And Fe.rcCargarPropiedad("reingresado", "", 0, 0)
    bresultado = bresultado And Fe.rcCargarPropiedad("depositario.codDomOrigen", "1", 0, 0) '1=fiscal 3=comercial
    bresultado = bresultado And Fe.rcCargarPropiedad("depositario.codDomOrigenNo", "V", 0, 0)

    ModVariables.xCuitReceptorNacionalJuridica = "55000002002"
    ModVariables.xCuitReceptorNacionalFisica = "50000002000"
    
    bresultado = bresultado And Fe.rcCargarPropiedad("receptor.cuitPaisReceptor", ModVariables.xCuitReceptorNacionalJuridica, 0, 0)
    bresultado = bresultado And Fe.rcCargarPropiedad("receptor.nacional.codDomReceptor", "1", 0, 0)
    bresultado = bresultado And Fe.rcCargarPropiedad("receptor.nacional.cuitReceptor", ModVariables.xCuitCliente, 0, 0)
    bresultado = bresultado And Fe.rcCargarPropiedad("receptor.nacional.tipodomreceptor", "1", 0, 0)
    
    Rem bresultado = bresultado And fe.rcCargarPropiedad("receptor.NacionalNoCat.documento", "123456", 0, 0)
    Rem bresultado = bresultado And fe.rcCargarPropiedad("receptor.NacionalNoCat.domDestinoCalle", "156", 0, 0)
    Rem bresultado = bresultado And fe.rcCargarPropiedad("receptor.NacionalNoCat.domDestinoCp", "3526", 0, 0)
    Rem bresultado = bresultado And fe.rcCargarPropiedad("receptor.NacionalNoCat.domDestinoIdPcia", "1", 0, 0)
    Rem bresultado = bresultado And fe.rcCargarPropiedad("receptor.NacionalNoCat.domDestinoLoc", "2513", 0, 0)
    Rem bresultado = bresultado And fe.rcCargarPropiedad("receptor.NacionalNoCat.domDestinoNumero", "235635", 0, 0)
    Rem bresultado = bresultado And fe.rcCargarPropiedad("receptor.NacionalNoCat.razonSocial", "numero", 0, 0)

    Rem bresultado = bresultado And fe.rcCargarPropiedad("receptor.extranjero.codigoAduana", "1", 0, 0)
    Rem bresultado = bresultado And fe.rcCargarPropiedad("receptor.extranjero.cuitDespachante", "2565465564", 0, 0)
    Rem bresultado = bresultado And fe.rcCargarPropiedad("receptor.extranjero.denominacionReceptor", "juan", 0, 0)
    Rem bresultado = bresultado And fe.rcCargarPropiedad("receptor.extranjero.domicilioReceptor", "street 3", 0, 0)

    bresultado = bresultado And Fe.rcCargarPropiedad("codRemRedestinar", "1", 0, 0)
    bresultado = bresultado And Fe.rcCargarPropiedad("codRemRedestinarNo", "V", 0, 0)
    'bresultado = bresultado And Fe.rcCargarPropiedad("depositario.cuitDepositario", ModVariables.xCuitCliente, 0, 0)
    bresultado = bresultado And Fe.rcCargarPropiedad("depositario.cuitDepositarioNo", "V", 0, 0)

    bresultado = bresultado And Fe.rcCargarPropiedad("cuitTitular", ModVariables.xCuitEmpresa, 0, 0)  ' cuit iman "20235709490"
    bresultado = bresultado And Fe.rcCargarPropiedad("puntoEmision", ModVariables.XPTOVTAREMITOHARINERO, 0, 0)
    bresultado = bresultado And Fe.rcCargarPropiedad("depositario.rucaEstDepositario", "1", 0, 0)
    bresultado = bresultado And Fe.rcCargarPropiedad("depositario.rucaEstDepositarioNO", "V", 0, 0)
                'bresultado = bresultado And Fe.rcCargarPropiedad("rucaEstEmisor", "1", 0, 0)
    bresultado = bresultado And Fe.rcCargarPropiedad("rucaEstEmisor", ModVariables.XNROESTABRUCA, 0, 0)
    bresultado = bresultado And Fe.rcCargarPropiedad("rucaEstEmisorNO", "V", 0, 0)
                
    bresultado = bresultado And Fe.rcCargarPropiedad("tipoCmp", "1", 0, 0)
    bresultado = bresultado And Fe.rcCargarPropiedad("tipoCmpNo", "f", 0, 0)
    Rem bresultado = bresultado And fe.rcCargarPropiedad("depositario.tipoDepositario", "D/E/I", 0, 0) ' D=DEPOSITARIO E=EMISOR I= INDUSTRIA MOLINERA
    bresultado = bresultado And Fe.rcCargarPropiedad("depositario.tipoDepositario", "E", 0, 0)
    Rem bresultado = bresultado And fe.rcCargarPropiedad("tipoDepositarioNO", "F", 0, 0)
    bresultado = bresultado And Fe.rcCargarPropiedad("depositario.tipoDomOrigen", "1", 0, 0)
    bresultado = bresultado And Fe.rcCargarPropiedad("depositario.tipoDomOrigenNo", "f", 0, 0)
    Rem bresultado = bresultado And fe.rcCargarPropiedad("tipoDomReceptor", "1", 0, 0)
    Rem bresultado = bresultado And fe.rcCargarPropiedad("tipoEmisor", "I/U", 0, 0)  I= INDUSTRIAL U= USUARIO DE MOLIENDA
    bresultado = bresultado And Fe.rcCargarPropiedad("tipoEmisor", "I", 0, 0)

    bresultado = bresultado And Fe.rcCargarPropiedad("importeCOT", "0", 0, 0)

Rem bresultado = bresultado And fe.rcCargarPropiedad("tipoEmisorNo", "f", 0, 0)
Rem bresultado = bresultado And fe.rcCargarPropiedad("viaje.cuitTransportista", "20235709490", 0, 0)
    bresultado = bresultado And Fe.rcCargarPropiedad("viaje.distanciaKm", ModVariables.XKMARECORRERREMCARNI, 0, 0)
    bresultado = bresultado And Fe.rcCargarPropiedad("viaje.fechaInicioViaje", ModVariables.XFECHAVIAJEREMCARNI, 0, 0)
    bresultado = bresultado And Fe.rcCargarPropiedad("viaje.transportista.codPaisTransportista", "200", 0, 0)
    bresultado = bresultado And Fe.rcCargarPropiedad("viaje.transportista.nacional.cuitConductor", ModVariables.xCuitChofer, 0, 0)
    bresultado = bresultado And Fe.rcCargarPropiedad("viaje.transportista.nacional.cuitConductorNO", "F", 0, 0)
    bresultado = bresultado And Fe.rcCargarPropiedad("viaje.transportista.nacional.cuitTransportista", ModVariables.xCuitTransporte, 0, 0)
    bresultado = bresultado And Fe.rcCargarPropiedad("viaje.transportista.nacional.cuitTransportistaNO", "F", 0, 0)
Rem bresultado = bresultado And fe.rcCargarPropiedad("viaje.transportista.extranjero.apellidoConductor", "smith", 0, 0)
Rem bresultado = bresultado And fe.rcCargarPropiedad("viaje.transportista.extranjero.cedulaConductor", "11325656", 0, 0)
Rem bresultado = bresultado And fe.rcCargarPropiedad("viaje.transportista.extranjero.denomTransportista", "prueba", 0, 0)
Rem bresultado = bresultado And fe.rcCargarPropiedad("viaje.transportista.extranjero.nombreConductor", "juan", 0, 0)
    'bresultado = bresultado And Fe.rcCargarPropiedad("viaje.vehiculo.ferroviario.arrayIdLocomotora.itemcantidad", "3", 0, 0)
    'bresultado = bresultado And Fe.rcCargarPropiedad("viaje.vehiculo.ferroviario.arrayIdLocomotora", "A", 0, 0)
    'bresultado = bresultado And Fe.rcCargarPropiedad("viaje.vehiculo.ferroviario.arrayIdLocomotora", "B", 1, 0)
    'bresultado = bresultado And Fe.rcCargarPropiedad("viaje.vehiculo.ferroviario.arrayIdLocomotora", "C", 2, 0)
    'bresultado = bresultado And Fe.rcCargarPropiedad("viaje.vehiculo.ferroviario.arrayIdVagon.itemcantidad", "2", 0, 0)
    'bresultado = bresultado And Fe.rcCargarPropiedad("viaje.vehiculo.ferroviario.arrayIdVagon", "A", 0, 0)
    'bresultado = bresultado And Fe.rcCargarPropiedad("viaje.vehiculo.ferroviario.arrayIdVagon", "B", 1, 0)
    'ModVariables.xPatenteAcoplado = "ABC111"
    Rem o vehiculo ferroviario o automotor no ambos
    bresultado = bresultado And Fe.rcCargarPropiedad("viaje.vehiculo.automotor.arrayDominioAcoplado.itemcantidad", "1", 0, 0)
    bresultado = bresultado And Fe.rcCargarPropiedad("viaje.vehiculo.automotor.arrayDominioAcoplado", ModVariables.xPatenteAcoplado, 0, 0)
   ' bresultado = bresultado And Fe.rcCargarPropiedad("viaje.vehiculo.automotor.arrayDominioAcoplado", "B", 1, 0)
    bresultado = bresultado And Fe.rcCargarPropiedad("viaje.vehiculo.automotor.cuitconductor", ModVariables.xCuitChofer, 0, 0)
    bresultado = bresultado And Fe.rcCargarPropiedad("viaje.vehiculo.automotor.dominiovehiculo", ModVariables.xPatente, 0, 0)
    If bresultado = False Then
        MsgBox "Error en armado de XML de Encabezado de Remito Electronico", vbInformation, "Validar"
        Exit Sub
    End If
    ' hacer un bucle por cada linea de detalle del remito
    ModSQL.VtaLeeDetaRemitoHarinero xNroAuxRemito
    ModSQL.AbrirRSRead2
    
    bresultado = bresultado And Fe.rcCargarPropiedad("arrayMercaderia.itemcantidad", RsRead2.RecordCount, 0, 0)

    For I = 0 To RsRead2.RecordCount - 1
        bresultado = bresultado And Fe.rcCargarPropiedad("arrayMercaderia.codComer", RsRead2![CodArticulo], I, 0)
        bresultado = bresultado And Fe.rcCargarPropiedad("arrayMercaderia.descComer", Trim(RsRead2![Descripcion]), I, 0)
        bresultado = bresultado And Fe.rcCargarPropiedad("arrayMercaderia.orden", I + 1, I, 0)
        bresultado = bresultado And Fe.rcCargarPropiedad("arrayMercaderia.codTipoUnidad", "1", I, 0)   ' 1=KILOS
        bresultado = bresultado And Fe.rcCargarPropiedad("arrayMercaderia.cantidadUnidad", Format(RsRead2![CantidadRemitada] * RsRead2![PesoEspecifico], "######0"), I, 0)
        bresultado = bresultado And Fe.rcCargarPropiedad("arrayMercaderia.codTipoEmb", RsRead2![CodEmbalaje], I, 0)
        bresultado = bresultado And Fe.rcCargarPropiedad("arrayMercaderia.cantidadEmb", Format(RsRead2![CantidadRemitada], "######0"), I, 0)
        bresultado = bresultado And Fe.rcCargarPropiedad("arrayMercaderia.codTipo", RsRead2![CodSenasa], I, 0)
        bresultado = bresultado And Fe.rcCargarPropiedad("arrayMercaderia.pesoNetoKg", Format(RsRead2![CantidadRemitada] * RsRead2![PesoEspecifico], "######0"), I, 0)
        bresultado = bresultado And Fe.rcCargarPropiedad("arrayMercaderia.pesoNetoKgNO", "f", I, 0) ' no se envia el dato
        
        bresultado = bresultado And Fe.rcCargarPropiedad("arrayMercaderia.pesoNetoPerKg", "1.1", I, 0)
        bresultado = bresultado And Fe.rcCargarPropiedad("arrayMercaderia.pesoNetoPerKgNo", "V", I, 0) ' no se envia
        bresultado = bresultado And Fe.rcCargarPropiedad("arrayMercaderia.pesoNetoRecKg", "1.1", I, 0)
        bresultado = bresultado And Fe.rcCargarPropiedad("arrayMercaderia.pesoNetoRecKgNo", "V", I, 0) ' no se envia
        bresultado = bresultado And Fe.rcCargarPropiedad("arrayMercaderia.pesoNetoRedKg", "1", I, 0)
        bresultado = bresultado And Fe.rcCargarPropiedad("arrayMercaderia.pesoNetoRedKgNO", "V", I, 0) ' no se envia
        bresultado = bresultado And Fe.rcCargarPropiedad("arrayMercaderia.pesoNetoReiKg", "1.1", I, 0)
        bresultado = bresultado And Fe.rcCargarPropiedad("arrayMercaderia.pesoNetoReiKgNO", "V", I, 0) ' no se envia+
        If bresultado = False Then
            MsgBox "Error al armar XML de detalle de remito electronico en artiuculo " & RsRead2![CodArticulo] & " " & Trim(RsRead2![Descripcion]) & ". Ver codigos senanasa y de embalaje", vbInformation, "validar"
        End If
        RsRead2.MoveNext
    Next
    RsRead2.Close
    Set RsRead2 = Nothing
    If bresultado Then
        ModVariables.XSEGENERORCOK = True
        bresultado = Fe.rcGenerarRemito(XNROIDRH)
        If Val(Fe.rcErrorCodigo1) = 151 Or Val(Fe.rcErrorCodigo1) = 1546 Then 'And Fe.rcErrorDescripcion1 = "El ID de request" Then
            ModVariables.xSumarANroInterno = ModVariables.xSumarANroInterno + 1
            MsgBox ("ERROR DEL SERVIDOR:  " + Fe.rcErrorDescripcion1) + " COD. ERROR: " & Fe.rcErrorCodigo1
            ModVariables.XSEGENERORCOK = False
        End If
                
        If ModVariables.XSEGENERORCOK = True Then
            ModVariables.xNROCAERC = Fe.rcLeerPropiedad("rcremitoGenerar", "datosemision.codAutorizacion", "", 0, 0)
            
            If Trim(ModVariables.xNROCAERC) <> "" Then   'OR Val(Fe.rcErrorCodigo1)) <> 1546) And (Fe.UltimoMensajeError = "" Or Val(Fe.rcErrorCodigo1) = 5002 Or Val(Fe.rcErrorCodigo1) = 151 Then
                    'MsgBox ("resultado " + Fe.rcLeerPropiedad("rcremitoGenerar", "resultado", "", 0, 0))
                    'MsgBox ("codrem " + Fe.rcLeerPropiedad("rcremitoGenerar", "codremito", "", 0, 0))
                   ' MsgBox "Error:   " & Fe.rcErrorCodigo1 & vbCrLf & " Ultimo Msje. de Error:   " & Fe.UltimoMensajeError
                    MsgBox ("Codigo CRE:  " + Fe.rcLeerPropiedad("rcremitoGenerar", "datosemision.codAutorizacion", "", 0, 0))
                    ModVariables.xNROCAERC = Fe.rcLeerPropiedad("rcremitoGenerar", "datosemision.codAutorizacion", "", 0, 0)
                    'MsgBox ("fecha " + Fe.rcLeerPropiedad("rcremitoGenerar", "datosemision.fechavencimiento", "", 0, 0))
                    XFECHACRE = Fe.rcLeerPropiedad("rcremitoGenerar", "datosemision.fechavencimiento", "", 0, 0)
                    If Trim(XFECHACRE) <> "" Then
                        ModVariables.xFechaVence = XFECHACRE
                    End If
                    'MsgBox ("imange " + Fe.rcLeerPropiedad("rcremitoGenerar", "qr", "C:\_txtsistema\remito.jpg", 0, 0))
                    xxx = Fe.rcLeerPropiedad("rcremitoGenerar", "qr", "C:\_txtsistema\remito.jpg", 0, 0)
                   xxx = Fe.rcLeerPropiedad("rcremitoGenerar", "qr", "C:\_txtsistema\REM" & XNROREM & ".jpg", 0, 0)
                    ModVariables.sTrSQL = "SELECT * FROM VTAREMITOS WHERE VTAREMITOS.NROAUXREMITO= " & xNroAuxRemito
                    ModSQL.AbrirRSOptimista4
                    With RsOpt4
                        ![CRE] = ModVariables.xNROCAERC
                        ![FechaVencimiento] = Format(ModVariables.xFechaVence, "dd/mm/yy")
                        ![NROINTERNOELECTRONICO] = XNROIDRH
                        .Update
                        .Close
                    End With
                    Set RsOpt4 = Nothing
                    'ModVariables.sTrSQL = "UPDATE VTAREMITOS SET CRE= " & ModVariables.xNROCAERC & ", FECHAVENCIMIENTO= '" & Format(ModVariables.xFechaVence, "dd/mm/yy") & "' , NROINTERNOELECTRONICO = " & xnroidrh & " WHERE VTAREMITOS.NROAUXREMITO= " & xNroAuxRemito
                    'ModSQL.AbrirRSRead4
                    'Set RsRead4 = Nothing
                    ModVariables.XSEGENERORCOK = True
                    If Val(Fe.rcErrorCodigo1) = 5002 Then
                        MsgBox ("ATENCION !! " + Fe.rcErrorDescripcion1)
                    End If
            Else
                    ModVariables.XSEGENERORCOK = False
                    If Val(Fe.rcErrorCodigo1) = 151 Then 'And Fe.rcErrorDescripcion1 = "El ID de request" Then
                        ModVariables.xSumarANroInterno = ModVariables.xSumarANroInterno + 1
                    End If
                    'MsgBox ("error leer" + Fe.UltimoMensajeError)
                    If Fe.rcErrorItemCantidad > 0 Then
                        MsgBox ("error del servidor " + Fe.rcErrorDescripcion1) + " Cod.Error " & Fe.rcErrorCodigo1
                    Else
                        MsgBox "Error al generar remito electronico. El servidor de AFIP no devuelve el codigo de error "
                    End If
            End If
        End If
    End If
Exit Sub
PROCERROR:
    ModSQL.MensajeError "GenerarRemitoHarinero " & ModVariables.xNROCAERC & ModVariables.xFechaVence
    
   Resume Next
End Sub
