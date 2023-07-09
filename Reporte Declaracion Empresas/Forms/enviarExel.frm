VERSION 5.00
Begin VB.Form enviarExel 
   Caption         =   "Form1"
   ClientHeight    =   1635
   ClientLeft      =   6165
   ClientTop       =   2715
   ClientWidth     =   2970
   LinkTopic       =   "Form1"
   ScaleHeight     =   1635
   ScaleWidth      =   2970
   Begin VB.CommandButton cmdEnviarExel 
      Caption         =   "Command1"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "enviarExel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnviarExel_Click()
    Dim ExpDim As New clsRprDeclaraEmpresasTemps
    ExpDim.CrearReporte
    ExpDim.SendToExcel
    Unload Me
End Sub

