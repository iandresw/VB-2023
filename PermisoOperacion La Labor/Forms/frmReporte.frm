VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmReporte 
   Caption         =   "Reporte Permisos Operacion"
   ClientHeight    =   3435
   ClientLeft      =   2130
   ClientTop       =   3300
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   4770
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin MSMask.MaskEdBox txtFechaInicial 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-MM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   18442
         SubFormatType   =   3
      EndProperty
      Height          =   420
      Left            =   2280
      TabIndex        =   1
      Top             =   600
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   741
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##-##-####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFechaFinal 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-MM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   18442
         SubFormatType   =   3
      EndProperty
      Height          =   420
      Left            =   2280
      TabIndex        =   3
      Top             =   1320
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   741
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Final:"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   660
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Inicial:"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   645
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "frmReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdImprimir_Click()
    Dim expdim As New CrearReporte
    expdim.CrearReporte Me.txtFechaInicial, Me.txtFechaFinal
    expdim.SendToExcel
    Unload Me
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub







Private Sub Form_Load()
 frmReporte.BackColor = RGB(29, 127, 146)
 Me.cmdImprimir.BackColor = RGB(231, 73, 54)
 Me.cmdSalir.BackColor = RGB(231, 73, 54)
 
End Sub
