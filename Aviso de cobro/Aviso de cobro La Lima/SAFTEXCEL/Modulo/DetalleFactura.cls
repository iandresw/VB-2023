VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DetalleFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarNumeroFactura As Long
Private mvarId As Long
Private mvarCuenta As String
Private mvarDescripcionCuenta As String
Private mvarValorUnitario As Currency
Private mvarCantidad As Integer
Private mvarVisibleEnTesoreria As Boolean
Private mvarTotal As Currency
Private mvarEsPermisoOperacion As Boolean
Private mvarEsImpuesto As Boolean
Private mvarEsICservicio As Boolean
Private mvarEsInteresPorImpuesto As Boolean
Private mvarEsInteresPorServicio As Boolean
Private mvarEsRecargoPorImpuesto As Boolean
Private mvarEsRecargoPorServicio As Boolean
Private mvarEsMultaPorOperarSinPermiso As Boolean
Private mvarEsMultaPorDeclararTarde As Boolean
Private mvarEsMulta As Boolean
Private mvarEsDescuentoTE As Boolean
Private mvarEsDescuentoPA As Boolean
Private Sub Class_Initialize()
    mvarCuenta = ""
    mvarNumeroFactura = 0
    mvarId = 0
End Sub
Public Property Get NumeroFactura() As Long
    NumeroFactura = mvarNumeroFactura
End Property
Public Property Let NumeroFactura(lngNumeroFactura As Long)
    mvarNumeroFactura = lngNumeroFactura
End Property
Public Property Let Id(lngID)
    mvarId = lngID
End Property

Public Property Get Cuenta() As String
    Cuenta = mvarCuenta
End Property
Public Property Let Cuenta(strCta As String)
    mvarCuenta = strCta
End Property
Public Property Get DescripcionCuenta() As String
    DescripcionCuenta = mvarDescripcionCuenta
End Property
Public Property Let DescripcionCuenta(strDescripcion As String)
    mvarDescripcion = strDescripcion
End Property
Public Property Get ValorUnitario() As Currency
    ValorUnitario = mvarValorUnitario
End Property
Public Property Let ValorUnitario(valor As Currency)
    mvarValorUnitario = valor
End Property
Public Property Get Cantidad() As Long
    Cantidad = mvarCantidad
End Property
Public Property Let Cantidad(lngCantidad As Long)
    mvarCantidad = lngCantidad
End Property
Public Property Get VisibleEnTesoreria() As Boolean
    VisibleEnTesoreria = mvarVisibleEnTesoreria
End Property
Public Property Let VisibleEnTesoreria(VisibleEnTes As Boolean)
    mvarVisibleEnTesoreria = VisibleEnTes
End Property
Public Property Get Total() As Currency
    Total = mvarTotal
End Property
Public Property Get EsImpuesto() As Boolean
    EsImpuesto = mvarEsImpuesto
End Property
Public Property Get EsInteresPorImpuesto() As Boolean
    EsInteresPorImpuesto = mvarEsInteresPorImpuesto
End Property
Public Property Get EsInteresPorServicio() As Boolean
    EsInteresPorServicio = mvarEsInteresPorServicio
End Property
Public Property Get EsRecargoPorImpuesto() As Boolean
    EsRecargoPorImpuesto = mvarEsRecargoPorImpuesto
End Property
Public Property Get EsRecargoPorServicio() As Boolean
    EsRecargoPorServicio = mvarEsRecargoPorServicio
End Property
Public Property Get EsDescuentoTerceraEdad() As Boolean
    EsDescuentoTerceraEdad = mvarEsDescuentoTE
End Property
Public Property Get EsDescuentoPagosAdelantados() As Boolean
    EsDescuentoPagosAdelantados = mvarEsDescuentoPA
End Property
Public Property Get EsMultaPorOperarSinPermiso() As Boolean
    EsMultaPorOperarSinPermiso = mvarEsMultaPorOperarSinPermiso
End Property
Public Property Get EsMultaPorDeclararTarde() As Boolean
    EsMultaPorDeclararTarde = mvarEsMultaPorDeclararTarde
End Property
Public Property Get EsPermisoOperacion() As Boolean
    EsPermisoOperacion = mvarEsPermisoOperacion
End Property
Public Property Get EsICservicio() As Boolean
    EsICservicio = mvarEsICservicio
End Property
Public Sub Load_data()
'cargamos un item del detalle, se entiende que se utilizara
'como un objeto collection

    'Dim capadb As New capabd
    Dim rs As New Recordset
    Dim Cuenta As New CuentaIngreso
    
    If mvarId = 0 Then Exit Sub
    
    Set rs = dal.ObtenerDetalleDeFactura(mvarId)
    mvarNumeroFactura = rs!numavpg
    mvarCuenta = rs!CtaIngreso
    mvarValorUnitario = rs!ValorUnitAvPgDet
    mvarCantidad = rs!CantAvPgDet
    mvarTotal = mvarValorUnitario * mvarCantidad
    Cuenta.CtaIngreso = rs!CtaIngreso
    Cuenta.CargarDatos
    mvarDescripcionCuenta = Cuenta.Nombre
    If Cuenta.EsImpuesto Then mvarEsImpuesto = True
    If Cuenta.EsInteresPorImpuesto Then mvarEsInteresPorImpuesto = True
    If Cuenta.EsInteresPorServicio Then mvarEsInteresPorServicio = True
    If Cuenta.EsRecargoPorImpuesto Then mvarEsRecargoPorImpuesto = True
    If Cuenta.EsRecargoPorServicio Then mvarEsRecargoPorServicio = True
    If Cuenta.EsDescuentoPorPagoAdelantado Then mvarEsDescuentoPA = True
    If Cuenta.EsDescuentoTerceraEdad Then mvarEsDescuentoTE = True
    If Cuenta.EsMultaPorOperarSinPermiso Then mvarEsMultaPorOperarSinPermiso = True
    If Cuenta.EsMultaPorDeclararTarde Then mvarEsMultaPorDeclararTarde = True
    If Cuenta.EsPermisoOperacion Then mvarEsPermisoOperacion = True
    If Cuenta.EsICservicio Then mvarEsICservicio = True
End Sub
Public Sub SetupData(rs As Recordset)
    Dim Cuenta As New CuentaIngreso

    mvarNumeroFactura = rs!numavpg
    mvarCuenta = rs!CtaIngreso
    mvarValorUnitario = rs!ValorUnitAvPgDet
    mvarCantidad = rs!CantAvPgDet
    If mvarCantidad = 0 Then mvarCantidad = 1 ' 14/02/2014
    mvarTotal = mvarValorUnitario * mvarCantidad
    Cuenta.CtaIngreso = rs!CtaIngreso
    Cuenta.CargarDatos
    mvarDescripcionCuenta = Cuenta.Nombre
    
    If Cuenta.EsImpuesto Then mvarEsImpuesto = True
    If Cuenta.EsInteresPorImpuesto Then mvarEsInteresPorImpuesto = True
    If Cuenta.EsInteresPorServicio Then mvarEsInteresPorServicio = True
    If Cuenta.EsRecargoPorImpuesto Then mvarEsRecargoPorImpuesto = True
    If Cuenta.EsRecargoPorServicio Then mvarEsRecargoPorServicio = True
    If Cuenta.EsDescuentoPorPagoAdelantado Then mvarEsDescuentoPA = True
    If Cuenta.EsDescuentoTerceraEdad Then mvarEsDescuentoTE = True
    If Cuenta.EsMultaPorOperarSinPermiso Then mvarEsMultaPorOperarSinPermiso = True
    If Cuenta.EsMultaPorDeclararTarde Then mvarEsMultaPorDeclararTarde = True
    If Cuenta.EsPermisoOperacion Then mvarEsPermisoOperacion = True
    If Cuenta.EsICservicio Then mvarEsICservicio = True
End Sub
