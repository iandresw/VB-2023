VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Factura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"DetalleFactura"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarNumero As Long 'local copy
Private mvarIdentidad As String 'local copy
Private mvarNombreBeneficiario As String
Private mvarFechaEmision As Date 'local copy
Private mvarFechaVence As Date 'local copy
Private mvarTipoFactura As Integer 'local copy
Private mvarTipoImpuesto As Integer 'local copy
Private mvarDescripcion As String 'local copy
Private mvarCodDeclara As String 'local copy
Private mvarClaveCatastro As String
'Private mvarSubTotal As Currency 'local copy
Private mvarImpuesto As Currency
Private mvarICservicio As Currency 'idustria y comercio y otros servicios
Private mvarSP As Currency
Private mvarPermisoOperacion As Currency
Private mvarInteres As Currency 'local copy
Private mvarRecargo As Currency 'local copy
Private mvarDescuentoTE As Currency 'local copy
Private mvarDescuentoPE As Currency 'local copy
Private mvarMultaPorOperarSinPermiso As Currency
Private mvarMultaPorDeclararTarde As Currency
Private mvarTotalFactura As Currency 'local copy
Private mvarEstadoFactura As String
Private mvarCodigoEstado As Integer
Private mvarDetalle As Collection
Private mvarUsuarioCreaFactura As String
Property Get UsuarioCreaFactura() As String
    UsuarioCreaFactura = mvarUsuarioCreaFactura
End Property

Public Property Get TotalFactura() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Descuento
    TotalFactura = mvarTotalFactura
End Property
Public Property Get DescuentoTE() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Descuento
    DescuentoTE = mvarDescuentoTE
End Property
Public Property Get DescuentoPE() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Descuento
    DescuentoPE = mvarDescuentoPE
End Property

Public Property Get Recargo() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Recargo
    Recargo = mvarRecargo
End Property

Public Property Get Interes() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Interes
    Interes = mvarInteres
End Property
Public Property Get MultaPorOperarSinPermiso() As Currency
    MultaPorOperarSinPermiso = mvarMultaPorOperarSinPermiso
End Property
Public Property Get MultaPorDeclararTarde() As Currency
    MultaPorDeclararTarde = mvarMultaPorDeclararTarde
End Property
Public Property Get Impuesto() As Currency
    Impuesto = mvarImpuesto
End Property
Public Property Get PermisoOperacion() As Currency
    PermisoOperacion = mvarPermisoOperacion
End Property
Public Property Get ICservicios() As Currency
    ICservicios = mvarICservicio
End Property
'Public Property Get SubTotal() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubTotal
    'SubTotal = mvarSubTotal
'End Property

Public Property Get CodDeclara() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CodDeclara
    CodDeclara = mvarCodDeclara
End Property

Public Property Get Descripcion() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AvPgDescripcion
    Descripcion = mvarDescripcion
End Property
Public Property Let Descripcion(strDescripcion As String)
    mvarDescripcion = strDescripcion
End Property
Public Property Get TipoImpuesto() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AvPgTipoImpuesto
'0=Undefined, 1=Bienes Inmuebles cuando la factura esta en mora tiene un recargo especial,2=Volumen de Ventas,3=Permiso Operacion, 4=Impuesto Personal, 5=Servicios Publicos,6= Incluida en un Plan de Pago,7=Plan Pago
    TipoImpuesto = mvarTipoImpuesto
End Property

Public Property Get TipoAvPg() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TipoAvPg
'1=Contado , 2=Pendiente
    TipoAvPg = mvarTipoFactura
End Property

Public Property Get EstadoFactura() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AvPgEstado
'1=No Pagada, 2=Pagada, 3=Anulada,4=En Tesoreria,5=Pagada Parcial, 6=Plan de Pago

'On Error Resume Next 'MSx3

    EstadoFactura = Switch(mvarCodigoEstado = 1, "NoPagada", mvarCodigoEstado = 2, "Pagada", mvarCodigoEstado = 3, "Anulada", mvarCodigoEstado = 4, "En Tesoreria", mvarCodigoEstado = 5, "Pago Parcial", mvarCodigoEstado = 6, "Plan de Pago")
End Property
Public Property Get FechaVence() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FechaVenceAvPg
    FechaVence = mvarFechaVence
End Property

Public Property Get FechaEmision() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FechaEmAvPg
    FechaEmision = mvarFechaEmision
End Property

Public Property Get Identidad() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Identidad
    Identidad = mvarIdentidad
End Property
Public Property Let Identidad(Identidad As String)
    Identidad = mvarIdentidad
End Property
Public Property Get Numero() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NumAvPg
    Numero = mvarNumero
End Property
Public Property Let Numero(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NombrePropietario = 5
    mvarNumero = vData
End Property

Public Property Get NombreBenefeciario() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NumAvPg
    NombreBenefeciario = mvarNombreBeneficiario
End Property
Public Property Get ClaveCatastro() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NumAvPg
    ClaveCatastro = mvarClaveCatastro
End Property
Public Property Get DetalleFactura() As Collection
    Set DetalleFactura = mvarDetalle
End Property
Public Property Get AvPgEstado() As Integer
    AvPgEstado = mvarCodigoEstado
End Property

Public Sub CargarDatos()
    'Cargue los datos locales
    'Calcule el valor del impuesto que se paga
    'calcule el valor del interes
    'calcule el valor del recargo
    'calcule el valor del descuento
    'Calcule el importe a pagar de la factura
    
    'Dim capabase As New capabd
    Dim cuentai As CuentaIngreso
    Dim rs As New ADODB.Recordset
    Dim Contrib As New Contribuyente
    Dim Detalle As New DetalleFactura
    Dim rsEm As New Recordset
    Dim mvarSubTotal, mvarDescuento As Integer
    Set rs = dal.ObtenerFacturaPorID(mvarNumero)

    If rs.RecordCount <= 0 Then
        mvarSubTotal = 0
        mvarInteres = 0
        mvarRecargo = 0
        mvarDescuento = 0
        mvarTotalFactura = 0
    Else
        mvarIdentidad = rs!Identidad
        mvarFechaEmision = rs!FechaEmAvPg
        mvarFechaVence = rs!FechaVenceAvPg
        mvarCodigoEstado = rs!AvPgEstado
        mvarTipoFactura = rs!TipoAvPg
        mvarTipoImpuesto = rs!AvPgTipoImpuesto
        mvarDescripcion = rs!AvPgDescripcion
        mvarCodDeclara = rs!CodDeclara
        If rs!AvPgTipoImpuesto = 5 Then
        Set rsEm = DeRia.CoRia.Execute("Select ClaveCatastro from AbonadoSPEnc where ASPE_Seq = " & rs!CodDeclara & "")
         If Not rsEm.EOF Then
           DeRia.CoRia.Execute ("update AvpgEnc set ClaveCatastro = '" & rsEm!ClaveCatastro & "' where numavpg = " & mvarNumero & " ")
         
        mvarClaveCatastro = IIf(IsNull(rsEm!ClaveCatastro), "", rsEm!ClaveCatastro)
         End If
        
        Else
            If rs!AvPgTipoImpuesto = 1 Then

            
               If mvarClaveCatastro = "" Then
                  mvarClaveCatastro = Mid(rs!AvPgDescripcion, 40, 25)
                  
         
                   DeRia.CoRia.Execute ("update AvpgEnc set ClaveCatastro = '" & mvarClaveCatastro & "' where numavpg = " & mvarNumero & " ")
                   DeRia.CoRia.Execute ("update AvPgDetalle  set ClaveCatastro = '" & mvarClaveCatastro & "' where numavpg = " & mvarNumero & " ")
               End If
           End If
           If Not IIf(IsNull(rs!ClaveCatastro), "", rs!ClaveCatastro) = "" Then
            mvarClaveCatastro = IIf(IsNull(rs!ClaveCatastro), "", rs!ClaveCatastro)
           End If
            
            Dim RsD As New ADODB.Recordset
            Set RsD = DeRia.CoRia.Execute("Select numavpg from AvPgDetalle where numavpg = " & mvarNumero & "  and clavecatastro = ''")
            If Not RsD.EOF Then
             DeRia.CoRia.Execute ("update AvPgDetalle  set ClaveCatastro = '" & mvarClaveCatastro & "' where numavpg = " & mvarNumero & " ")
            
            End If
            
            
            
        End If
        
       
       
        mvarUsuarioCreaFactura = rs!CreadoPor
        Set Contrib = New Contribuyente
        Contrib.Load_data (mvarIdentidad)
        mvarNombreBeneficiario = Contrib.NombreCompleto
    End If
    Set rs = dal.ObtenerDetallesDeFactura(mvarNumero)
    Set mvarDetalle = New Collection
    Do While rs.EOF = False
        Set Detalle = New DetalleFactura
        Detalle.Id = rs!SeqAvPgDet
        vFFACT = Format(mvarFechaVence, "yyyy")
        Detalle.SetupData rs
        mvarDetalle.Add Detalle
        
        mvarTotalFactura = mvarTotalFactura + (rs!CantAvPgDet * rs!ValorUnitAvPgDet)
        If Detalle.EsImpuesto Then
            'mvarSubTotal = mvarSubTotal + (rs!CantAvPgDet * rs!ValorUnitAvPgDet)
            mvarImpuesto = mvarImpuesto + (rs!CantAvPgDet * rs!ValorUnitAvPgDet)
        End If
        If Detalle.EsPermisoOperacion Then
            mvarPermisoOperacion = mvarPermisoOperacion + (rs!CantAvPgDet * rs!ValorUnitAvPgDet)
        End If
        If Detalle.EsICservicio Then
            mvarICservicio = mvarICservicio + (rs!CantAvPgDet * rs!ValorUnitAvPgDet)
        End If
        If Detalle.EsInteresPorImpuesto Or Detalle.EsInteresPorServicio Then
            mvarInteres = rs!CantAvPgDet * rs!ValorUnitAvPgDet
        End If
        If Detalle.EsRecargoPorImpuesto Or Detalle.EsRecargoPorServicio Then
            mvarRecargo = rs!CantAvPgDet * rs!ValorUnitAvPgDet
        End If
        If Detalle.EsDescuentoTerceraEdad Then
            mvarDescuentoTE = (rs!CantAvPgDet * rs!ValorUnitAvPgDet)
        End If
        If Detalle.EsDescuentoPagosAdelantados Then
            mvarDescuentoPE = (rs!CantAvPgDet * rs!ValorUnitAvPgDet)
        End If
        If Detalle.EsMultaPorOperarSinPermiso Then
            mvarMultaPorOperarSinPermiso = (rs!CantAvPgDet * rs!ValorUnitAvPgDet)
        End If
        If Detalle.EsMultaPorDeclararTarde Then
            mvarMultaPorDeclararTarde = (rs!CantAvPgDet * rs!ValorUnitAvPgDet)
        End If
        rs.MoveNext
    Loop
    'mvarTotalFactura = (mvarSubTotal + mvarInteres + mvarRecargo) + (mvarDescuentoTE + mvarDescuentoPE)
End Sub
Public Sub LimpiarDetalle()
    Set mvarDetalle = New Collection
End Sub

Public Function AgregarDetalle(det As DetalleFactura)
    Dim Detalle As New DetalleFactura
    
    Detalle.NumeroFactura = mvarNumero
    Detalle.Cuenta = det.Cuenta
    Detalle.Cantidad = det.Cantidad
    Detalle.ValorUnitario = det.ValorUnitario
    Detalle.VisibleEnTesoreria = det.VisibleEnTesoreria
    mvarDetalle.Add Detalle
    
End Function
Public Function TipoImpuestoEnLetras() As String
    Select Case mvarTipoImpuesto
        Case 0
            TipoImpuestoEnLetras = "OS"
        Case 1
            TipoImpuestoEnLetras = "BI"
        Case 3
            TipoImpuestoEnLetras = "IC"
        Case 2
            TipoImpuestoEnLetras = "IC"
        Case 4
            TipoImpuestoEnLetras = "IP"
        Case 5
            TipoImpuestoEnLetras = "SP"
        Case 7
            TipoImpuestoEnLetras = "PP"
    End Select
End Function

