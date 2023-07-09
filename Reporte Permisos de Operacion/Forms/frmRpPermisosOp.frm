VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmRpPermisosOp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rpr Permisos de Operacion"
   ClientHeight    =   4365
   ClientLeft      =   5085
   ClientTop       =   2385
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4680
   Begin VB.ComboBox cboAldea 
      Height          =   315
      ItemData        =   "frmRpPermisosOp.frx":0000
      Left            =   2040
      List            =   "frmRpPermisosOp.frx":0007
      TabIndex        =   13
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   570
      Left            =   2400
      Picture         =   "frmRpPermisosOp.frx":0012
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3360
      Width           =   1455
   End
   Begin VB.OptionButton optTodos 
      Caption         =   "Todos"
      Height          =   195
      Left            =   3360
      TabIndex        =   4
      Top             =   2640
      Width           =   800
   End
   Begin VB.OptionButton optRenovacion 
      Caption         =   "Renovacion"
      Height          =   195
      Left            =   1800
      TabIndex        =   3
      Top             =   2640
      Width           =   1200
   End
   Begin VB.OptionButton optApertura 
      Caption         =   "Apertura"
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   2640
      Width           =   900
   End
   Begin VB.CommandButton cmdImprimirExel 
      Caption         =   "&Enviar a Exel "
      Height          =   570
      Left            =   600
      TabIndex        =   1
      Top             =   3360
      Width           =   1455
   End
   Begin MSMask.MaskEdBox txtFecha1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   300
      Left            =   2040
      TabIndex        =   5
      Top             =   720
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFecha2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   300
      Left            =   2040
      TabIndex        =   6
      Top             =   1320
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Aldea"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   -240
      TabIndex        =   12
      Top             =   1920
      Width           =   1665
   End
   Begin VB.Label lblArrow 
      BackStyle       =   0  'Transparent
      Caption         =   "ç"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   10
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblArrow 
      BackStyle       =   0  'Transparent
      Caption         =   "ç"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   9
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Final:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Inicial:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "REPORTE PERMISOS DE OPERACION "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmRpPermisosOp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSalir_Click()
        Unload Me
End Sub

Private Sub cmdImprimirExel_Click()
    On Error GoTo FactEmitidas_Error
    Dim AptoRen As String
    Dim CodAldea As String
    Dim sqlAldea As String
    
    
     If Not IsDate(txtFecha1) Or Not IsDate(txtFecha2) Then
        MsgBox "Debe ingresar fechas validas..!"
        Exit Sub
    End If

    If optApertura.Value = True Then
        AptoRen = " AND Tra_PermOP.Observacion = 'APERTURA' "
    ElseIf optRenovacion.Value = True Then
        AptoRen = " AND Tra_PermOP.Observacion = 'RENOVACION' "
    ElseIf optTodos.Value = True Then
        AptoRen = ""
    End If
    
     With DeRia.rscmdAldeas
        .Requery
        .Find "NombreAldea='" & Trim(cboAldea.Text) & "'"
        If .EOF Then
            sqlAldea = ""
        Else
            If !NombreAldea = Trim(cboAldea.Text) Then
                CodAldea = !CodAldea
                sqlAldea = " AND Tra_PermOP.CodAldea = '" & CodAldea & "'"
            Else
                MsgBox ("Ingrese Aldea valida")
                cboAldea = ""
                cboAldea.SetFocus
                Exit Sub
            End If
        End If
    End With
    
    
    Dim ExpDim As New clsPermisoOperacion
    ExpDim.CrearReporte Me.txtFecha1.Text, Me.txtFecha2.Text, AptoRen, sqlAldea
    ExpDim.SendToExcel

    Exit Sub
FactEmitidas_Error:
    MsgBox Err.Description
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


Private Sub Form_Load()
CargarAldea
End Sub

