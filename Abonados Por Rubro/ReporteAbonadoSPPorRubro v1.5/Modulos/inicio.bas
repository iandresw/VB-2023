Attribute VB_Name = "inicio"
Option Explicit
Public vFactVen As Long
Public BitIP, BitPc, StrBit As String ' para Bitacora
Public VTpoId As String
Public VGrid, VMasServicios, VrMillar, VrMillarR As Integer
Public vFFACT As Integer
Public VPeriodoJuridicoCJ As Integer
Public Linea As Integer
Public Const pv_ModoEdicion = 0 'Enter Data Mode
Public Const Pv_ModoLectura = 1 'Display Data Mode
Public Const EnBackColor = &H80000005
Public Const DisBackColor = &H80000005 '&H8000000F  '&HE0E0E0
Public pv_Datamode, VarJur, VPermiso, VarRuralUrb As Integer
Public VValpermiso As Currency
Public pv_Identidad As String
Public VExcel As Integer
Public VIdentidad2 As String
Public VFecPP1, VFecPP2 As Date
Public VxFechaPP As Integer
Public VNTipo As Integer
Public VDirImg, VDirDiagrama As String
Public UsrNivel As String
Public VFechaQuita As Date

Public VAmniIC, VAmniIp, VAmniBi, VAmniSP As Integer 'Para que Aplique Amnistia a impuestos selectivo
Public Itmx As ListItem
Public VaProp, VImp, VMillar As Currency
Public VTipoImpuesto As Integer
Public VMensualidad As Currency
Public VCuentaServicios As String
Public VFechaRecibo As Date
Public VImp22, VBotBanco As Integer
Public VFactNum, vFactNumX1 As Currency
Public Anio, Anio2, ResAnio, MesD, Anio3, VEstadoFact As Integer
Public VQuita, QAno As Integer
Public Qdia, QMes As String
Public pv_NombreAlcaldia, strSql, strSql2 As String
Public pv_CodigoAlcaldia As String
Public pMainConnStr As String
Public gsString As String
Public VDesCuenta As String
Public Cuenta As Integer
Public HelpWindow As String
Public vConteo As Integer
Public VCodAldea2, VAldea2  As String
Public VSexoHUrbano, VSexoHRural, VSexoMUrbano, VSexoMRural As String
Public VTotUrbano, VTotRural, VTotUrbanoR, VTotRuralR, VTotPendiente As Currency
Public VNombreCont As String
Public gRsEnc As Recordset
Public gRsDet As Recordset
Public PathFileName As String
Public VRVal2 As Currency
Public gsUsername As String
Public VUsuario As String
Public VUsuario2 As String
Public gsNombreModulo As String
Public rFormRs, VFact, TpoImp, Tpo2 As String
Public Date2 As Date
Public i, VPer, DateBI, Q2 As Integer
Public oCallingForm As Form
Public VFechaIC1, VFechaIC2 As Date
Public vNuFact As Long
Public vRTNMun As String
Public VxFecha As Integer
Public VCtaIngresoMultaOPSinPermiso, VCtaIngresoMultaDeclaraTarde, VCtaIngresoRecargoImp, VCtaIngresoRecargoServ, VCtaIngresoIntImp, VCtaIngresoIntServ As String 'Para Aministia
'Public cuurvalor, txtImpuesto, txtMultaDeclaraTarde, txtDescuento, txtRecargo1, txtRecargo2, txtRecargo3 As Currency
Public rsFacturaServicios As New ADODB.Recordset
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Sub Sleep Lib "kernel32" (ByVal dwmilliseconds As Long)
Public aCuotasMensuales(12, 9) As String 'aCuotasVencidas(Fecha,Impuesto,Multa,Otros,Meses,Intereses,Saldo Acumulado,Recargo,Total)
Public dal As New CapaBD
Public ConcertaImp, ConcertaTerreno, ConcertaEdificacion, VValTTT, VValEdif  As Currency
Public VPeriodoConcerta As Date
Sub Main()

    Dim rs As New Recordset
    On Error GoTo MainError

    gsNombreModulo = "Modulo de Administración Tributaria"
    DeRia.cmdParametro 'Load parameter data
    If CheckParametros = False Then
        MsgBox "Debe primero definir los parametros del modulo...!"
        Exit Sub
    End If
    pv_CodigoAlcaldia = DeRia.rscmdParametro!CodMuni
    pv_NombreAlcaldia = DeRia.rscmdParametro!NombreMuni
    VDirImg = DeRia.rscmdParametro!DirFoto
    VDirDiagrama = DeRia.rscmdParametro!DirFoto
    DeRia.rscmdParametro.Close
    
        Dim rsRTN As New ADODB.Recordset
    Set rsRTN = DeRia.CoRia.Execute("Select * from ParametroCont")
    If Not rsRTN.EOF Then
    vRTNMun = Trim(rsRTN!RtnEmpresa)
    End If
    
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open ("Select * from Parametro")
    FrmReporteMo.Show


    Exit Sub

MainError:
    MsgBox Err.Description
End Sub
