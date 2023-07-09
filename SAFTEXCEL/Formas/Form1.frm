VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmReporteMo 
   Caption         =   "Form1"
   ClientHeight    =   3096
   ClientLeft      =   6876
   ClientTop       =   3336
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   ScaleHeight     =   3096
   ScaleWidth      =   5100
   Begin VB.CommandButton btnSalir 
      Caption         =   "Salir"
      Height          =   570
      Left            =   3108
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   1230
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "Imprimir"
      Height          =   570
      Left            =   1800
      Picture         =   "Form1.frx":00FA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   1230
   End
   Begin VB.ComboBox CmdTipo 
      Height          =   288
      ItemData        =   "Form1.frx":062C
      Left            =   1920
      List            =   "Form1.frx":066D
      TabIndex        =   2
      Text            =   "Seleccione"
      Top             =   1440
      Width           =   1935
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
      Left            =   1860
      TabIndex        =   0
      Top             =   600
      Width           =   1248
      _ExtentX        =   2201
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
      Left            =   1860
      TabIndex        =   1
      Top             =   960
      Width           =   1248
      _ExtentX        =   2201
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
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
      Height          =   252
      Index           =   1
      Left            =   3960
      TabIndex        =   11
      Top             =   960
      Width           =   732
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
      Height          =   252
      Index           =   0
      Left            =   3960
      TabIndex        =   10
      Top             =   600
      Width           =   732
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Final:"
      ForeColor       =   &H00C00000&
      Height          =   252
      Left            =   696
      TabIndex        =   9
      Top             =   960
      Width           =   1092
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Inicial:"
      ForeColor       =   &H00C00000&
      Height          =   252
      Left            =   696
      TabIndex        =   8
      Top             =   600
      Width           =   1092
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cuadro Resumen Impuesto Personal"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   5052
   End
   Begin VB.Label LblFlecha 
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
      Height          =   252
      Left            =   3960
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Impuesto:"
      ForeColor       =   &H00C00000&
      Height          =   252
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1668
   End
End
Attribute VB_Name = "FrmReporteMo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAceptar_Click()
Dim VanioI As Integer
    On Error GoTo FactEmitidas_Error
    

 
    If Not IsDate(txtFecha1) Or Not IsDate(txtFecha2) Then
        MsgBox "Debe ingresar fechas validas..!"
        Exit Sub
    End If
    
  
    
    Dim Osp As New Exporta
    Osp.CrearReporte Me.txtFecha1, Me.txtFecha2
    Osp.SendToExcel
    
    Exit Sub
FactEmitidas_Error:
    MsgBox Err.Description
End Sub

Private Sub btnSalir_Click()
Unload Me
End Sub
