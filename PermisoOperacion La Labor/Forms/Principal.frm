VERSION 5.00
Begin VB.Form Principal 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Permisos de Operacion"
   ClientHeight    =   2820
   ClientLeft      =   5355
   ClientTop       =   2355
   ClientWidth     =   6240
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   6240
   Begin VB.CommandButton cmdImprimirFicha 
      Appearance      =   0  'Flat
      Caption         =   "Imprimir Ficha"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdReporte 
      Caption         =   "Reporte"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdPOTaxi 
      Appearance      =   0  'Flat
      Caption         =   "Permido de operacion Moto Taxi"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdPO 
      Appearance      =   0  'Flat
      Caption         =   "Permido de operacion Establecimientos Comerciales"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdImprimirFicha_Click()
    frmFichaRegistro.Show
End Sub



Private Sub cmdPO_Click()
    fmrPerimisosOp.Show
End Sub

Private Sub cmdPOTaxi_Click()
    frmInfoTaxi.Show
End Sub

Private Sub cmdReporte_Click()
    frmReporte.Show
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Principal.BackColor = RGB(29, 127, 146)
    Me.cmdImprimirFicha.BackColor = RGB(231, 73, 54)
    Me.cmdPO.BackColor = RGB(231, 73, 54)
    Me.cmdPOTaxi.BackColor = RGB(231, 73, 54)
    Me.cmdReporte.BackColor = RGB(231, 73, 54)
    Me.cmdSalir.BackColor = RGB(231, 73, 54)
    
End Sub
