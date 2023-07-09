VERSION 5.00
Begin VB.Form frmPropiedadesConDeclaracion 
   Caption         =   "Busca Propiedades por declaracion"
   ClientHeight    =   3105
   ClientLeft      =   8820
   ClientTop       =   4200
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   4590
   Begin VB.TextBox txtFechaFinal 
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtFechaInicio 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton cmdImprimirExel 
      Caption         =   "&Enviar a Exel "
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   2280
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
      Left            =   360
      TabIndex        =   4
      Top             =   1680
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
      Left            =   360
      TabIndex        =   3
      Top             =   960
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

Private Sub cmdsalir_Click()
 Unload Me
End Sub

