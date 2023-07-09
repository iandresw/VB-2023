VERSION 5.00
Begin VB.Form BotonExelBanco 
   Caption         =   "Form1"
   ClientHeight    =   675
   ClientLeft      =   4950
   ClientTop       =   5625
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   ScaleHeight     =   675
   ScaleWidth      =   3585
   Begin VB.CommandButton cmdEnviarExel 
      Caption         =   "Mostar en Exel"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "BotonExelBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnviarExel_Click()
    Dim ExpDim As New clsEnviarExel
    ExpDim.CrearReporte
    ExpDim.SendToExcel
    Unload Me
End Sub

