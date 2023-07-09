VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBarrio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'local variable(s) to hold property value(s)
Private mvarCodBarrio As String 'local copy
Private mvarNombreBarrio As String 'local copy
Private mvarCodAldea As String 'local copy
Private mvarCodDIMATELA As String 'local copy
Private Sub Class_Initialize()
    Dim rs As New ADODB.Recordset
    Set rs = dal.ObtenerBarrioPorCodigo
    If rs.RecordCount = 0 Then Exit Sub
    mvarCodBarrio = rs!codBarrio
    mvarNombreBarrio = Trim(rs!NombreBarrio)
    mvarCodAldea = rs!CodAldea
    mvarCodDIMATELA = rs!CodDIMATELA
End Sub
Public Property Get NombreBarrio() As String
    NombreBarrio = mvarNombreBarrio
End Property
Public Property Let NombreBarrio(ByVal vData As String)
    mvarNombreBarrio = vData
End Property

Public Property Get codBarrio() As String
    codBarrio = mvarCodBarrio
End Property
Public Property Let codBarrio(ByVal vData As String)
    mvarCodBarrio = vData
End Property

Public Property Get CodAldea() As String
    CodAldea = mvarCodAldea
End Property
Public Property Let CodAldea(ByVal vData As String)
    mvarCodAldea = vData
End Property

Public Property Get CodDIMATELA() As String
    CodDIMATELA = mvarCodDIMATELA
End Property
Public Property Let CodDIMATELA(ByVal vData As String)
    mvarCodDIMATELA = vData
End Property

