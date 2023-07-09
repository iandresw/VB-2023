VERSION 5.00
Begin VB.Form FrmBuscaDima 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes"
   ClientHeight    =   3105
   ClientLeft      =   11565
   ClientTop       =   3420
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4935
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtFecha2 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox txtFecha1 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Final"
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   1500
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Inicial"
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   900
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Reporte de Mora Servicios Publicos"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "FrmBuscaDima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimir_Click()
 On Error GoTo FactEmitidas_Error
    If Not IsDate(txtFecha1) Or Not IsDate(txtFecha2) Then
        MsgBox "Debe ingresar fechas validas..!"
        Exit Sub
    End If
    Dim ExpDim As New ClsMoraDimaTela
    ExpDim.CrearReporte Me.txtFecha1.Text, Me.txtFecha2.Text
    ExpDim.SendToExcel
    Unload Me
    Exit Sub
FactEmitidas_Error:
    MsgBox Err.Description
End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub


Private Sub Form_Load()

End Sub
