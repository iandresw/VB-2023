VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Alcaldia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarNombre As String 'local copy
Private mvarDepartamento As String 'local copy
Private mvarMunicipio As String 'local copy
Private mvarNombreAlcalde As String 'local copy
Private mvarNombreViceAlcalde As String 'local copy
Private mvarNombreContador As String 'local copy
Private mvarNombreAuditor As String 'local copy
Private mvarCodigoAlcaldia As String 'local copy
Private mvarCodigoDepartamento As String 'local copy
Private mvarNombreJefeAdmonTributaria As String 'local copy
Private mvarNombreJefeServiciosPublicos As String 'local copy
Public VMuni As String
Private Sub Class_Initialize()
    'Dim CapaAD As New capabd
    Dim rs As New ADODB.Recordset
    
    Set rs = dal.ObtenerAlcaldia()
    
    If rs.RecordCount = 0 Then Exit Sub
    
    mvarNombre = Trim(rs!NombreMuni)
    mvarCodigoAlcaldia = rs!CodMuni
    mvarDepartamento = rs!NombreDepto
    mvarMunicipio = rs!NombreMuni
    
End Sub
Public Property Let NombreJefeServiciosPublicos(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NombreJefeServiciosPublicos = 5
    mvarNombreJefeServiciosPublicos = vData
End Property


Public Property Get NombreJefeServiciosPublicos() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NombreJefeServiciosPublicos
    NombreJefeServiciosPublicos = mvarNombreJefeServiciosPublicos
End Property



Public Property Let NombreJefeAdmonTributaria(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NombreJefeAdmonTributaria = 5
    mvarNombreJefeAdmonTributaria = vData
End Property


Public Property Get NombreJefeAdmonTributaria() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NombreJefeAdmonTributaria
    NombreJefeAdmonTributaria = mvarNombreJefeAdmonTributaria
End Property



Public Property Let CodigoDepartamento(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CodigoDepartamento = 5
    mvarCodigoDepartamento = vData
End Property


Public Property Set CodigoDepartamento(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.CodigoDepartamento = Form1
    mvarCodigoDepartamento = vData
End Property


Public Property Get CodigoDepartamento() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CodigoDepartamento
    If IsObject(mvarCodigoDepartamento) Then
        CodigoDepartamento = mvarCodigoDepartamento
    Else
        CodigoDepartamento = mvarCodigoDepartamento
    End If
End Property



Public Property Let CodigoAlcaldia(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CodigoAlcaldia = 5
    mvarCodigoAlcaldia = vData
End Property


Public Property Get CodigoAlcaldia() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CodigoAlcaldia
    CodigoAlcaldia = mvarCodigoAlcaldia
End Property



Public Property Let NombreAuditor(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NombreAuditor = 5
    mvarNombreAuditor = vData
End Property


Public Property Get NombreAuditor() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NombreAuditor
    NombreAuditor = mvarNombreAuditor
End Property



Public Property Let NombreContador(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NombreContador = 5
    mvarNombreContador = vData
End Property


Public Property Get NombreContador() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NombreContador
    NombreContador = mvarNombreContador
End Property



Public Property Let NombreViceAlcalde(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NombreViceAlcalde = 5
    mvarNombreViceAlcalde = vData
End Property


Public Property Get NombreViceAlcalde() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NombreViceAlcalde
    NombreViceAlcalde = mvarNombreViceAlcalde
End Property



Public Property Let NombreAlcalde(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NombreAlcalde = 5
    mvarNombreAlcalde = vData
End Property


Public Property Get NombreAlcalde() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NombreAlcalde
    NombreAlcalde = mvarNombreAlcalde
End Property



Public Property Let Municipio(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Municipio = 5
    mvarMunicipio = vData
End Property


Public Property Get Municipio() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Municipio
    Municipio = mvarMunicipio
End Property



Public Property Let Departamento(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Departamento = 5
    mvarDepartamento = vData
End Property


Public Property Get Departamento() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Departamento
    Departamento = mvarDepartamento
End Property



Public Property Let Nombre(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Nombre = 5
    mvarNombre = vData
End Property


Public Property Get Nombre() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Nombre
    Nombre = mvarNombre
End Property



