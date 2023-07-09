VERSION 5.00
Begin VB.Form frmFichaRegistro 
   Caption         =   "Form1"
   ClientHeight    =   2415
   ClientLeft      =   7875
   ClientTop       =   3870
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   4230
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtRTM 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   480
      Width           =   1635
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese el RTM:"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   225
      Left            =   300
      TabIndex        =   0
      Top             =   480
      Width           =   1515
   End
End
Attribute VB_Name = "frmFichaRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim crearFicha As New clsGenerarFicha
Private Sub cmdImprimir_Click()
    Dim rtm As String
    rtm = Me.txtRTM
    crearFicha.crearFicha rtm
End Sub


Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    frmFichaRegistro.BackColor = RGB(29, 127, 146)
    Me.cmdImprimir.BackColor = RGB(231, 73, 54)
    Me.cmdSalir.BackColor = RGB(231, 73, 54)
    
End Sub
