VERSION 5.00
Begin VB.Form FrmAbonadosSpRubroBarrio 
   Caption         =   "Reportes"
   ClientHeight    =   3480
   ClientLeft      =   1920
   ClientTop       =   8415
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
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
      Left            =   1800
      TabIndex        =   1
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

    Dim txtcodBarrio As String

    With DeRia.rscmdTablaBarrio
        .Requery
        .Find "NombreBarrio='" & Trim(cboBarrio.Text) & "'"
        If .EOF Then
            MsgBox "No se encontro una aldea en la tabla"
            Exit Sub
        Else
            If !NombreBarrio = Trim(cboBarrio.Text) Then
                txtcodBarrio = !codBarrio
            Else
                MsgBox ("Ingrese un barrio valido")
                cboBarrio = ""
                cboBarrio.SetFocus
                Exit Sub
            End If
        End If
    End With
    

    Dim ExpDim As New clsAbonadoSPPorRubroBarrio
    ExpDim.CrearReporte txtRubroCtaIngreso.Text, txtcodBarrio
    ExpDim.SendToExcel txtRubroCtaIngreso.Text, txtcodBarrio
    Unload Me
    Exit Sub
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Sub CargarAldea()
    If DeRia.rscmdAldeas.State = 1 Then DeRia.rscmdAldeas.Close
    DeRia.rscmdAldeas.Open
    cboAldea.Text = "[Seleccione una aldea]"
    Do Until DeRia.rscmdAldeas.EOF
    cboAldea.AddItem DeRia.rscmdAldeas!NombreAldea
    cboAldea.ItemData(cboAldea.NewIndex) = DeRia.rscmdAldeas.Collect("CodAldea")
    DeRia.rscmdAldeas.MoveNext
    Loop
End Sub
Sub CargarBarrio(CodAldea As String)
    Dim strCodBarrio As String
    strCodBarrio = "SELECT * FROM TablaBarrio WHERE (CodAldea = '" & CodAldea & "')"
    If DeRia.rscmdTablaBarrio.State = 1 Then DeRia.rscmdTablaBarrio.Close
    DeRia.rscmdTablaBarrio.Open strCodBarrio
    cboBarrio.Text = "[Seleccione un barrio]"
    Do Until DeRia.rscmdTablaBarrio.EOF
    cboBarrio.AddItem DeRia.rscmdTablaBarrio!NombreBarrio
    cboBarrio.ItemData(cboBarrio.NewIndex) = DeRia.rscmdTablaBarrio.Collect("CodBarrio")
    DeRia.rscmdTablaBarrio.MoveNext
    Loop
End Sub

Private Sub Form_Load()
Dim CodAldea As String
Dim codBarrio As String
CargarAldea
CodAldea = "010701"
CargarBarrio CodAldea
End Sub


