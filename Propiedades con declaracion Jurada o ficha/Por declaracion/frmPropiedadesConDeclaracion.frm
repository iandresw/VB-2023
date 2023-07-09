VERSION 5.00
Begin VB.Form frmPropiedadesConDeclaracion 
   Caption         =   "Busca Propiedades por declaracion"
   ClientHeight    =   3645
   ClientLeft      =   5940
   ClientTop       =   3045
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   4695
   Begin VB.TextBox txtFechaFinal 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtFechaInicio 
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton cmdImprimirExel 
      Caption         =   "&Enviar a Exel "
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha final"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Top             =   1800
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha inicial"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Width           =   1230
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Propiedades con declaracion jurada por fecha "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmPropiedadesConDeclaracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimirExel_Click()
    On Error GoTo FactEmitidas_Error
    
    If Not IsDate(txtFechaInicio) Or Not IsDate(txtFechaFinal) Then
        MsgBox "Debe ingresar fechas validas"
        Exit Sub
    End If
        
    Dim ExpDim As New clsPropiedadesPorDeclaracion
    ExpDim.CrearReporte Me.txtFechaInicio.Text, Me.txtFechaFinal.Text
    ExpDim.SendToExcel

    Exit Sub
FactEmitidas_Error:
    MsgBox Err.Description
End Sub

Private Sub cmdSalir_Click()
 Unload Me
End Sub

Private Sub Form_Load()

End Sub
