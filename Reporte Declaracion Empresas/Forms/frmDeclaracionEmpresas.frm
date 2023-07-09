VERSION 5.00
Begin VB.Form frmDeclaracionEmpresas 
   Caption         =   "Form1"
   ClientHeight    =   2505
   ClientLeft      =   8490
   ClientTop       =   4020
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   4260
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtPeriodo 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtNombreEmpresa 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Periodo"
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   960
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombre Empresa"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmDeclaracionEmpresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimir_Click()
    Dim ExpDim As New clsRptDeclaraEmpresa
    ExpDim.CrearReporte Me.txtNombreEmpresa, Me.txtPeriodo
    ExpDim.SendToExcel
    Unload Me
End Sub
