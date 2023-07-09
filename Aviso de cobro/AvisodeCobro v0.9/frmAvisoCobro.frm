VERSION 5.00
Begin VB.Form frmAvisoCobro 
   Caption         =   "Form1"
   ClientHeight    =   3480
   ClientLeft      =   6870
   ClientTop       =   3330
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   5115
   Begin VB.TextBox txtIdentidad 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   2535
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Identidad"
      Height          =   195
      Left            =   705
      TabIndex        =   4
      Top             =   1080
      Width           =   660
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Impresion de Avisos de Cobro Por Contribuyente"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4785
   End
End
Attribute VB_Name = "frmAvisoCobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private AbonadoSPDet As New ADODB.Recordset
Private Sub cmdImprimir_Click()
Dim nombre As String
   Dim expdim As New rpAvisoDeCobro
   expdim.Show
   
End Sub


Private Sub cmdSalir_Click()
 Unload Me
End Sub
