VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CapaBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"CapaBD"
Option Explicit
Private Conexion As New ADODB.Connection
Private mvarConnectionString As String
Private mvarDbProvider As String

'Constructor
Private Sub Class_Initialize()
    
    Dim strConeccion As String
    Dim strDataPath As String
    
 '   On Error GoTo ConectarBD_Error
   
    strDataPath = GetSetting("SAFT", "DBSetting", "DataPath")
   

    'con esta coneccion de arriba no daba el recordcount
    'strConeccion = "DRIVER=SQL Server;SERVER=(local);UID=sa;PWD=Excalibur2008;APP=2007 Microsoft Office system;WSID=SATELLITEMSx3;DATABASE=saft2" 'qwerty3025
    strConeccion = "DRIVER=SQL Server;SERVER=(local);UID=sa;PWD=Excalibur2008;APP=2007 Microsoft Office system;WSID=SATELLITEMSx3;DATABASE=saft" 'MSx3
' strConeccion = "PROVIDER=MSDataShape;DATA Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDataPath & ";Jet OLEDB:Database password=mgia730@lc@3b3s" 'qwerty3025
    strConeccion = DeRia.CoRia
    Conexion.ConnectionString = strConeccion
    mvarConnectionString = strConeccion
    'mvarDbProvider = "MSAccess"
    strConeccion = DeRia.CoRia
    Conexion.Open
    Exit Sub

ConectarBD_Error:
    MsgBox Err.Description
    End 'Finaliza ejecucion de programa.
End Sub
Public Property Get ConnectionString() As String
    ConnectionString = mvarConnectionString
End Property
Public Property Get DbProvider() As String
    DbProvider = mvarDbProvider
End Property

Public Function ObtenerRecordset(strSql As String) As ADODB.Recordset
    Set ObtenerRecordset = DeRia.CoRia.Execute(strSql)
End Function
Public Sub EjecutarSql(strSql As String)
    DeRia.CoRia.Execute (strSql)
End Sub

'Entidad Factura
Public Function ObtenerFacturasVencidasPorClaveCatastral(strClave As String) As Recordset
    Dim strSql As String
    '22 de enero 2014 que solo muestre las facturas Tipo 1
    strSql = "Select * from AvPgEnc A " & _
    "where AvPgEstado=1 and FechaVenceAvPg < '" & Format(DiaEnProcesoCT, "dd/mm/yyyy") & "' " & _
    "and RTRIM(LTRIM(A.ClaveCatastro)) ='" & strClave & "' and A.AvpgTipoImpuesto = 1 order by FechaVenceAvPg"
    Set ObtenerFacturasVencidasPorClaveCatastral = DeRia.CoRia.Execute(strSql)
    
End Function
Public Function ObtenerFacturasPorClaveCatastralParaAvisoPago(strClave As String) As ADODB.Recordset
'selecciona las facturas vencidas y las que vencen el ano actual.

    Dim strSql As String
    Dim RsCat As New ADODB.Recordset
    
    Set RsCat = DeRia.CoRia.Execute("SELECT SUBSTRING(AvPgDescripcion, 39, 25) AS ClaveCat, NumAvPg  FROM  AvPgEnc WHERE (AvPgTipoImpuesto = 1) AND (ClaveCatastro = '') AND (AvPgEstado = 1) ORDER BY FechaVenceAvPg")
    
    Do While Not RsCat.EOF
    DeRia.CoRia.Execute ("Update avpgenc set ClaveCatastro = '" & RsCat!ClaveCat & "' where numavpg = " & RsCat!NumAvPg & "")
    
    RsCat.MoveNext
    Loop
    

    strSql = "Select * from AvPgEnc A " & _
    "where AvPgEstado=1 and year(FechaVenceAvPg) <= " & Year(DiaEnProcesoCT) & " " & _
    "and LTRIM(RTRIM(A.ClaveCatastro))='" & Trim(strClave) & "' and AvPgTipoImpuesto = 1 order by FechaVenceAvPg"
    'strSql = "Select * from AvPgEnc A " & _
    '"where AvPgEstado=1 and " & _
    '"A.ClaveCatastro='" & strClave & "' order by FechaVenceAvPg"
    Set ObtenerFacturasPorClaveCatastralParaAvisoPago = DeRia.CoRia.Execute(strSql)

End Function


Public Function ObtenerFacturasVencidasPorIdentidad(strId As String) As Recordset
    Dim strSql As String
    Dim rsParametro As New ADODB.Recordset
    
    Set rsParametro = DeRia.CoRia.Execute("select * from Parametro")
    If rsParametro.RecordCount = 0 Then Exit Function
    
    If VxFecha = 1 Then
    strSql = "Select * from AvPgEnc A " & _
    "where AvPgEstado<> 3 " & _
    "and A.Identidad='" & strId & "' " & _
    "and FechaVenceAvPg between '" & Format(VFechaIC1, "dd/mm/yyyy") & "' and '" & Format(VFechaIC2, "dd/mm/yyyy") & "' " & _
    "order by FechaVenceAvPg"
    End If
    
    If VxFecha = 0 Then
    strSql = "Select * from AvPgEnc A " & _
    "where AvPgEstado<> 3 " & _
    "and A.Identidad='" & strId & "' " & _
    "and FechaVenceAvPg < '" & Format(rsParametro!DiaProcesoCT, "dd/mm/yyyy") & "' " & _
    "order by FechaVenceAvPg"
    End If
    
    
    Set ObtenerFacturasVencidasPorIdentidad = DeRia.CoRia.Execute(strSql)
End Function
Public Function ObtenerFacturasParaAvisoDePagoPorIdentidad(strId As String) As Recordset
'las facturas para aviso de pagos, especialmente de industria y comercio son las que vencen
'el mes que estamos, no solo las facturas vencidas, por eso esta function es diferente
'a ObtenerFacturasVencidasPorIdentidad.

    Dim strSql As String
    Dim rsParametro As New ADODB.Recordset
    Dim dDia As Date
    
    Set rsParametro = DeRia.CoRia.Execute("select * from Parametro")
    If rsParametro.RecordCount = 0 Then Exit Function
    dDia = rsParametro!DiaProcesoCT
    'dDia = FinDeMes(dDia)
    strSql = "Select * from AvPgEnc A " & _
    "where AvPgEstado=1 " & _
    "and A.Identidad='" & strId & "' " & _
    "and FechaVenceAvPg <= '" & Format(dDia, "dd/mm/yyyy") & "' " & _
    "order by FechaVenceAvPg"
    Set ObtenerFacturasParaAvisoDePagoPorIdentidad = DeRia.CoRia.Execute(strSql)
End Function

Public Sub ObtenerFacturasPorDeclaracion()
End Sub

Public Sub ObtenerFacturasPorClaveCatastral()
End Sub

Public Function ObtenerFacturasPorIdentidad(strIdentidad As String) As ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select * from AvPgEnc where " & _
    "Identidad='" & strIdentidad & "' " & _
    "order by FechaVenceAvPg"
    Set ObtenerFacturasPorIdentidad = DeRia.CoRia.Execute(strSql)

End Function
Public Function ObtenerFacturasPendientesPorIdentidad(strIdentidad) As ADODB.Recordset
    
    strSql = "select NumAvPg from AvPgEnc " & _
    "where Identidad='" & strIdentidad & "' and AvPgEstado=1 order by FechaVenceAvPg"
    Set ObtenerFacturasPendientesPorIdentidad = DeRia.CoRia.Execute(strSql)
    
End Function

Public Function ObtenerFacturasPorPrimerNombre(strNombre As String) As ADODB.Recordset
    Dim strSql As String
    Dim myrs As New Recordset
    
    strSql = "select * from AvPgEnc" & _
    "(select Identidad from Contribuyente where PNombre like '%" & strNombre & "%' )"
    'MsgBox strSql
    strSql = "select * from AvPgEnc where AvPgTipoImpuesto=1"
    Set myrs = DeRia.CoRia.Execute(strSql)
    'MsgBox myrs.RecordCount
End Function

Public Function ObtenerFacturaPorID(lngNumAvPg As Long) As ADODB.Recordset
    Dim strSql As String

    strSql = "Select * from AvPgEnc A, AvPgDetalle B " & _
    "where A.NumAvPg=" & lngNumAvPg & " " & _
    "and B.NumAvPg=A.NumAvPg " & _
    "order by FechaVenceAvPg"
    Set ObtenerFacturaPorID = DeRia.CoRia.Execute(strSql)
    
End Function
Public Function ObtenerDetalleDeFactura(lngSeq As Long) As ADODB.Recordset
'El detalle de una factura se extrae uno por uno, en base al sequencial no al numero de factura

    Dim strSql As String
    
    strSql = "Select * from AvPgDetalle " & _
    "where SeqAvPgDet=" & lngSeq & " "
    Set ObtenerDetalleDeFactura = DeRia.CoRia.Execute(strSql)

End Function

Public Function ObtenerDetallesDeFactura(lngNumero As Long) As ADODB.Recordset
'Devuelve todo el detalle de una factura

    Dim strSql As String
  
    strSql = "Select * from AvPgDetalle " & _
    "where NumAvPg=" & lngNumero & " "
    Set ObtenerDetallesDeFactura = DeRia.CoRia.Execute(strSql)
    'MsgBox ObtenerDetallesDeFactura.RecordCount
End Function
Public Function ObtenerFacturasPorRecibo(lngNumRecibo As Long) As ADODB.Recordset
    Dim strSql As String
    strSql = "select * from AvPgEnc where NumAvPg in " & _
    "(Select NumFactura from ReciboDet " & _
    "where NumRecibo=" & lngNumRecibo & ")"
    
    Set ObtenerFacturasPorRecibo = DeRia.CoRia.Execute(strSql)
End Function
Public Function ObtenerFacturaPorID_Detallada(strIdentidad As String)
    strSql = "select AvPgEnc.NumAvPg,FechaEmAvPg,FechaVenceAvPg,AvPgEstado,TipoAvPg,AvPgTipoImpuesto,AvPgDescripcion,CreadoPor,FechaCreado,AvPgEnc.ClaveCatastro," & _
    "ValorUnitAvPgDet,CantAvPgDet,AvPgDetalle.CtaIngreso,NombreCtaIngreso,RefAvPgDet,VisibleEnTesAvPgDet," & _
    "AvPgEnc.Identidad,PNombre,SNombre,PApellido,SApellido " & _
    "from AvPgEnc, AvPgDetalle, Contribuyente,CuentaIngreso " & _
    "where AvPgEnc.Identidad='" & strIdentidad & "' and " & _
    "AvPgDetalle.NumAvPg=AvPgEnc.NumAvPg and " & _
    "Contribuyente.Identidad=AvPgEnc.Identidad and " & _
    "CuentaIngreso.CtaIngreso=AvPgDetalle.CtaIngreso"
    
    Set ObtenerFacturaPorID_Detallada = DeRia.CoRia.Execute(strSql)
End Function

Public Sub AnularFactura(lngNumero As Long)
    Dim strSql As String
    
    strSql = "Update AvPgEnc set " & _
    "AvPgEstado=3, " & _
    "ModificadoPor= '" & gsUsername & "', " & _
    "FechaModificado= '" & Format(DiaEnProcesoCT, "dd/mm/yyyy") & "' " & _
    "where NumAvPg=" & lngNumero & ""
    DeRia.CoRia.Execute (strSql)
    
End Sub

'Entidad CuentaIngreso
Public Function ObtenerCuentasIngreso() As ADODB.Recordset

    Dim strSql As String
    
    strSql = "Select * from CuentaIngreso "
    Set ObtenerCuentasIngreso = DeRia.CoRia.Execute(strSql)

End Function
Public Function ObtenerCuentaIngresoPorID(strCta As String, strAnio As Integer) As ADODB.Recordset
    ', Anio As Integer
    Dim strSql As String
    
    strSql = "Select * FROM CuentaIngreso INNER JOIN CuentaIngreso_A ON CuentaIngreso.CtaIngreso = CuentaIngreso_A.CtaIngreso where CuentaIngreso.CtaIngreso='" & strCta & "' AND CUENTAINGRESO_A.ANIO = " & strAnio & " "
    Set ObtenerCuentaIngresoPorID = DeRia.CoRia.Execute(strSql)

End Function
Public Function CuentaIngresoEsPermisoOperacion(strCta As String) As Boolean
    Dim strSql As String
    Dim vAnio As Integer
    
    Dim rs As New ADODB.Recordset
    vAnio = Format(DiaEnProcesoCT, "yyyy")
    
  '  strSql = "select * from CuentaIngreso where CtaIngreso='" & strCta & "' " & _
  '  "and ctaingreso in (select CtaPermOP from CuentaIngreso) and tipo=1"
    
    strSql = "select * from CuentaIngreso_A  where CtaIngreso ='" & strCta & "' and ctaingreso in (select CtaPermOP from CuentaIngreso_A ) and tipo=1 AND Anio = " & vAnio & ""
    Set rs = DeRia.CoRia.Execute(strSql)
    If rs.RecordCount <= 0 Then
        CuentaIngresoEsPermisoOperacion = False
    Else
        CuentaIngresoEsPermisoOperacion = True
    End If
    
End Function
Public Function ObtenerCuentasParametroGeneral() As ADODB.Recordset
    Set ObtenerCuentasParametroGeneral = DeRia.CoRia.Execute("select * from SystemParam")
End Function
'Entidad Propiedad
Public Function ObtenerPropiedadPorClave(strClave As String) As ADODB.Recordset

    Set ObtenerPropiedadPorClave = DeRia.CoRia.Execute("select * from Catastro where ClaveCatastro='" & strClave & "'")
End Function
'Entidad Contribuyente
Public Function ObtenerContribuyentePorID(strId As String) As ADODB.Recordset
    Set ObtenerContribuyentePorID = DeRia.CoRia.Execute("select * from Contribuyente where Identidad='" & strId & "'")
End Function
Public Function ObtenerDeclaracionesIPporId(strId As String) As ADODB.Recordset
    Set ObtenerDeclaracionesIPporId = DeRia.CoRia.Execute("select * from DeclaraImpInd where Identidad='" & strId & "' order by PeriodoDeclara")
End Function
Public Function ObtenerDeclaracionesIPporNumero(strNumero As String) As ADODB.Recordset
    Set ObtenerDeclaracionesIPporNumero = DeRia.CoRia.Execute("select * from DeclaraImpInd where CodDeclaraIP='" & strNumero & "'")
End Function

Public Function ObtenerDeclaracionesICporID(strId As String) As ADODB.Recordset
    Set ObtenerDeclaracionesICporID = DeRia.CoRia.Execute("select * from DeclaraContJurid where Identidad='" & strId & "' order by Periodo")
End Function
Public Function ObtenerDeclaracionICPorNumero(strNumero As String) As ADODB.Recordset
    Set ObtenerDeclaracionICPorNumero = DeRia.CoRia.Execute("select * from DeclaraContJurid where CodDeclaraCJ='" & strNumero & "'")
End Function
Public Sub ActivarContribuyente(strIdentidad As String)
    DeRia.CoRia.Execute ("update Contribuyente set Activo=1 where Identidad='" & strIdentidad & "'")
    'DeRia.CoRia.Close
End Sub
Public Sub InactivarContribuyente(strIdentidad As String)
'inactiva el contribuyente y anula las facturas no pagadas

    Dim strSql As String
    Dim rs As New ADODB.Recordset
    
    Set rs = Me.ObtenerFacturasPendientesPorIdentidad(strIdentidad)
    DeRia.CoRia.BeginTrans
    'Conexion.BeginTrans
    On Error GoTo TrapError
    Do While rs.EOF = False
        Me.AnularFactura rs!NumAvPg
        rs.MoveNext
    Loop
    DeRia.CoRia.Execute ("update Contribuyente set Activo=0 where Identidad='" & strIdentidad & "'")
    DeRia.CoRia.CommitTrans
    
    Exit Sub
    
TrapError:
    DeRia.CoRia.RollbackTrans
End Sub

'Entidad Aldea
Public Function ObtenerAldeaPorCodigo(strCodigo As String) As ADODB.Recordset
    Set ObtenerAldeaPorCodigo = DeRia.CoRia.Execute("select * from Aldea where CodAldea='" & strCodigo & "'")
End Function
'Entidad Barrio / Caserio
Public Function ObtenerBarrioPorCodigo(strCodigo As String) As ADODB.Recordset
    Set ObtenerBarrioPorCodigo = DeRia.CoRia.Execute("select * from TablaBarrio where CodBarrio='" & strCodigo & "'")
End Function
Public Function ObtenerBarriosPorAldea(strCodAldea As String) As ADODB.Recordset
'MSx3 antes strCod
    Set ObtenerBarriosPorAldea = DeRia.CoRia.Execute("select * from TablaBarrio where CodAldea='" & strCodAldea & "'")
End Function
Public Function ObtenerAlcaldia() As ADODB.Recordset
    Set ObtenerAlcaldia = DeRia.CoRia.Execute("select * from Parametro")
    
End Function
Public Function ObtenerFacturaDeDeclaracion(strNumDecla) As ADODB.Recordset
    Set ObtenerFacturaDeDeclaracion = DeRia.CoRia.Execute("select  * from AvPgDetalle where RefAvPgDet='" & strNumDecla & "'")
End Function

Public Function ObtenerContribuyentesPorNombre(strNombre As String) As ADODB.Recordset
    Set ObtenerContribuyentesPorNombre = DeRia.CoRia.Execute("select * from Contribuyente where PNombre like '%" & Trim(strNombre) & "%' order by PNombre ")
End Function
Public Function ObtenerContribuyentesPorApellido(strApellido As String) As ADODB.Recordset
    Set ObtenerContribuyentesPorApellido = DeRia.CoRia.Execute("select * from Contribuyente where PApellido like '%" & Trim(strApellido) & "%' order by PNombre ")
End Function
Public Function ObtenerContribuyentesPorID(strApellido As String) As ADODB.Recordset
    Set ObtenerContribuyentesPorID = DeRia.CoRia.Execute("select * from Contribuyente where Identidad = '" & Trim(strApellido) & "' order by PNombre ")
End Function
Public Function ObtenerContribuyentes() As ADODB.Recordset
    Set ObtenerContribuyentes = DeRia.CoRia.Execute("select * from Contribuyente order by PNombre ")
End Function
Public Function ObtenerContribuyentesNaturales() As ADODB.Recordset
'Set ObtenerContribuyentesNaturales = DeRia.CoRia.Execute("select * from Contribuyente where Tipo= " & 0 & " order by PNombre ")
    Set ObtenerContribuyentesNaturales = DeRia.CoRia.Execute("select * from Contribuyente where Tipo= " & 0 & " order by PNombre ")
End Function
Public Function ObtenerEmpresas() As ADODB.Recordset
    Set ObtenerEmpresas = DeRia.CoRia.Execute("select * from Contribuyente where Tipo= 1")
End Function
Public Function ObtenerEmpresasPorNombre(strNombre As String) As ADODB.Recordset
    Set ObtenerEmpresasPorNombre = DeRia.CoRia.Execute("select * from Contribuyente where Tipo= 1 and PNombre like '%" & Trim(strNombre) & "%'")
End Function
Public Function ObtenerContribNaturalesPorNombre(strNombre As String) As ADODB.Recordset
    Set ObtenerContribNaturalesPorNombre = DeRia.CoRia.Execute("select * from Contribuyente where Tipo= " & 0 & " and PNombre like '%" & RTrim(LTrim(strNombre)) & "%'")
End Function
Public Function ObtenerParametrosCT() As ADODB.Recordset
        
    Set ObtenerParametrosCT = DeRia.CoRia.Execute("select * from parametro")
    
End Function


'Entidad Usuario
Public Function ObtenerUsuarios() As ADODB.Recordset
    Set ObtenerUsuarios = DeRia.CoRia.Execute("select * from Usuario order by UsuarioNombre")
End Function
Public Function ObtenerUsuarioPorID(strId As String) As ADODB.Recordset
    Set ObtenerUsuarioPorID = DeRia.CoRia.Execute("select * from Usuario where UsuarioCod='" & strId & "'")
End Function
Public Function ObtenerModulosPorUsuario(strId As String) As ADODB.Recordset
    Set ObtenerModulosPorUsuario = DeRia.CoRia.Execute("Select * from UsuarioModulo where UsuarioCod='" & strId & "'")
End Function
'Entidad Cuenta Contable
Public Function ObtenerCuentasContables() As ADODB.Recordset
    Set ObtenerCuentasContables = DeRia.CoRia.Execute("select * from Catalogo order by CtaContable")
    
End Function
Public Function ObtenerCuentaContablePorID(strId As String) As ADODB.Recordset
    Set ObtenerCuentaContablePorID = DeRia.CoRia.Execute("select * from Catalogo where CtaContable='" & strId & "'")
End Function

'Entidad Abonado de Servicio Publico
Public Function ObtenerAbonadoServPubPorId(lngID As Long) As ADODB.Recordset
    Set ObtenerAbonadoServPubPorId = DeRia.CoRia.Execute("select * from AbonadoSPEnc where ASPE_Seq=" & lngID & "")
End Function
Public Function ObtenerFacturasAbonadoSP(lngID As Long) As ADODB.Recordset
    Set ObtenerFacturasAbonadoSP = DeRia.CoRia.Execute("select * from AvPgEnc where CodDeclara='" & lngID & "'  and AvPgTipoImpuesto=5")
End Function
Public Function ObtenerFacturasNoPagadasAbonadoSP(lngID As Long) As ADODB.Recordset
    Set ObtenerFacturasNoPagadasAbonadoSP = DeRia.CoRia.Execute("select * from AvPgEnc where CodDeclara='" & lngID & "'  and AvPgTipoImpuesto=5 and AvPgEstado=1 order by FechaVenceAvPg")
    'MsgBox ObtenerFacturasNoPagadasAbonadoSP.RecordCount
End Function

'Entidad Recibo
'Public Function ObtenerReciboPorID(lngID As Long) As ADODB.Recordset
'    Set ObtenerReciboPorID = DeRia.CoRia.Execute("Select Recibo A, ReciboDet B where A.NumRecibo=" & lngID & " and B.NumRecibo=A.Numrecibo")
'End Function

Public Function ObtenerRecibosPorFactura(lngNumFactura As Long) As ADODB.Recordset
    Set ObtenerRecibosPorFactura = DeRia.CoRia.Execute("select * from  Recibo A, ReciboDet B where B.NumFactura =" & lngNumFactura & " and B.NumRecibo=A.NumRecibo")
End Function
Public Function ObtenerReciboPorID(lngNumRecibo As Long) As ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select A.Identidad,A.FechaRecibo,A.DescRecibo,A.NumRecibo," & _
    "SentToCont,ReciboAnulado,CreadoPor,FechaCreado, ValorUnitReciboDet,CtaIngreso,CantReciboDet " & _
    "from Recibo A, ReciboDet B where  A.NumRecibo=" & lngNumRecibo & " and B.NumRecibo=A.NumRecibo"
    Set ObtenerReciboPorID = DeRia.CoRia.Execute(strSql)
End Function
Public Function ObtenerDetalleReciboResumido(lngNumRecibo As Long) As ADODB.Recordset
    Dim strSql As String
    
    strSql = "select CtaIngreso, sum(ValorUnitReciboDet) as SubTotal," & _
    "sum(CantReciboDet) as Cantidad " & _
    "from ReciboDet where NumRecibo=" & lngNumRecibo & "  " & _
    "group by CtaIngreso"
    
    Set ObtenerDetalleReciboResumido = DeRia.CoRia.Execute(strSql)
    
End Function
'Public Function ObtenerReciboTotal(lngNumRecibo As Long) As adodb.Recordset
Public Function ObtenerReciboTotal(lngNumRecibo As Long) As Currency 'MSx3
    strSql = "select sum(ValorUnitReciboDet) as SubTotal " & _
    "from ReciboDet where NumRecibo=" & lngNumRecibo & "  "
    'ObtenerTotalRecibo = DeRia.CoRia.Execute(strSql)
    DeRia.CoRia.Execute (strSql)
    ObtenerReciboTotal = DeRia.rsAbonadoSP!Subtotal
    
End Function
Public Function ObtenerTablaParametro() As ADODB.Recordset

    Set ObtenerTablaParametro = DeRia.CoRia.Execute("select * from Parametro")
End Function
