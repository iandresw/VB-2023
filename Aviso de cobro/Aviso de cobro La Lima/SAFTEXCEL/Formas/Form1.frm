VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmReportePP 
   Caption         =   "Reporte Planes de Pago"
   ClientHeight    =   4620
   ClientLeft      =   9885
   ClientTop       =   2940
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   4515
   Begin VB.OptionButton optTodos 
      Caption         =   "Todos"
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   1680
      Width           =   1695
   End
   Begin VB.OptionButton optMeseVencidos 
      Caption         =   "Meses Vencidos"
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   1200
      Width           =   1695
   End
   Begin VB.OptionButton optEntreFecha 
      Caption         =   "Entre Fechas"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton btnSalir 
      Caption         =   "Salir"
      Height          =   570
      Left            =   2520
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   1230
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "Imprimir"
      Height          =   570
      Left            =   720
      Picture         =   "Form1.frx":00FA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   1230
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
      Left            =   1920
      TabIndex        =   0
      Top             =   2160
      Visible         =   0   'False
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
      Left            =   1920
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtMeses 
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
      Left            =   1920
      TabIndex        =   13
      Top             =   2400
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   3
      Mask            =   "###"
      PromptChar      =   "_"
   End
   Begin VB.Label lblArrow 
      BackStyle       =   0  'Transparent
      Caption         =   "ç"
      Enabled         =   0   'False
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
      Index           =   2
      Left            =   3480
      TabIndex        =   14
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad de meses"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   480
      TabIndex        =   12
      Top             =   2400
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label lblArrow 
      BackStyle       =   0  'Transparent
      Caption         =   "ç"
      Enabled         =   0   'False
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
      TabIndex        =   8
      Top             =   2640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblArrow 
      BackStyle       =   0  'Transparent
      Caption         =   "ç"
      Enabled         =   0   'False
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
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Final:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   840
      TabIndex        =   6
      Top             =   2640
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Inicial:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estado planes de pago emitios"
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
      Left            =   930
      TabIndex        =   4
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "FrmReportePP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAceptar_Click()
Dim VanioI As Integer
    On Error GoTo FactEmitidas_Error
    
    If optEntreFecha.Value = True Then
        If Not IsDate(txtFecha1) Or Not IsDate(txtFecha2) Then
            MsgBox "Debe ingresar fechas validas..!"
            Exit Sub
        End If
    End If
    
    Dim Osp As New clsRptPlanPago
    Osp.CrearReporte
    Osp.SendToExcel
    Exit Sub
FactEmitidas_Error:
    MsgBox Err.Description
End Sub

Private Sub btnSalir_Click()
Unload Me
End Sub

Private Sub optEntreFecha_Click()
    txtFecha1.Visible = True
    txtFecha2.Visible = True
    Label2.Visible = True
    Label3.Visible = True
    Label4.Visible = False
    txtMeses.Visible = False
End Sub

Private Sub optMeseVencidos_Click()
    txtMeses.Visible = True
    txtFecha1.Visible = False
    txtFecha2.Visible = False
    Label2.Visible = False
    Label3.Visible = False
    Label4.Visible = True
    
End Sub

Private Sub optTodos_Click()
    txtMeses.Visible = False
    txtFecha1.Visible = False
    txtFecha2.Visible = False
    Label2.Visible = False
    Label3.Visible = False
    Label4.Visible = False
End Sub

