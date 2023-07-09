VERSION 5.00
Begin VB.Form FrmRepContibuyente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes"
   ClientHeight    =   2535
   ClientLeft      =   2700
   ClientTop       =   2820
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4935
   Begin VB.CommandButton smdSalir 
      Caption         =   "&Salir"
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de Contribuyentes"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "FrmRepContibuyente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimir_Click()
    On Error GoTo FactEmitidas_Error
    Dim ExpDim As New clsReporteContribuyentes
    ExpDim.CrearReporte
    ExpDim.SendToExcel
    Unload Me
    Exit Sub
FactEmitidas_Error:
    MsgBox Err.Description
End Sub


Private Sub smdSalir_Click()
     Unload Me
End Sub
