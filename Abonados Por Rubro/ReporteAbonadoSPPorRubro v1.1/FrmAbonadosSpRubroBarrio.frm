VERSION 5.00
Begin VB.Form FrmAbonadosSpRubroBarrio 
   Caption         =   "Reportes"
   ClientHeight    =   3450
   ClientLeft      =   4050
   ClientTop       =   3015
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   5250
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   3360
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   495
      Left            =   960
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ComboBox cboBarrio 
      Height          =   315
      Left            =   1800
      TabIndex        =   4
      Text            =   "010701003"
      Top             =   1920
      Width           =   2175
   End
   Begin VB.ComboBox cboAldea 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtRubroCtaIngreso 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Text            =   "1111180201"
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Barrio"
      Height          =   195
      Left            =   720
      TabIndex        =   6
      Top             =   2040
      Width           =   420
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Abonados de Servicios Publicos por Rubro y Barrio"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   150
      TabIndex        =   5
      Top             =   120
      Width           =   4995
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Aldea"
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Top             =   1440
      Width           =   420
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Rubro:"
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   480
   End
End
Attribute VB_Name = "FrmAbonadosSpRubroBarrio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdImprimir_Click()
    On Error GoTo FactEmitidas_Error
    Dim ExpDim As New clsAbonadoSPPorRubroBarrio
    ExpDim.CrearReporte Me.txtRubroCtaIngreso.Text, Me.cboBarrio.Text
    ExpDim.SendToExcel
    Unload Me
    Exit Sub
FactEmitidas_Error:
    MsgBox Err.Description
End Sub


Sub CargarAldea()
    Dim rsAbonadoSP As ADODB.Recordset
    Dim sql As String
    Set rsAbonadoSP = New ADODB.Recordset
    sql = "SELECT  CodAldea, NombreAldea FROM Aldea"
    cboAldea.Text = "[Seleccione una aldea]"
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open (sql)
    Do Until DeRia.rsAbonadoSP.EOF
        cboAldea.AddItem (rsAbonadoSP!NombreAldea)
        cboAldea.ItemData(cboAldea.NewIndex) = rsAbonadoSP.Collect("CodigoAldea")
        rsAbonadoSP.MoveNext
    Loop
End Sub

Private Sub Form_Load()
    CargarAldea
End Sub
