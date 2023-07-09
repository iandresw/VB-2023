VERSION 5.00
Begin VB.Form frmPropConFichasUrb 
   Caption         =   "Propiedades con Ficha Catastral"
   ClientHeight    =   4080
   ClientLeft      =   9135
   ClientTop       =   2940
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   4695
   Begin VB.OptionButton optRurales 
      Caption         =   "Rurales"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin VB.OptionButton optUrbanas 
      Caption         =   "Urbanas"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdImprimirExel 
      Caption         =   "&Enviar a Exel "
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox txtFechaFinal 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtFechaInicio 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   840
      Width           =   1815
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   960
      Width           =   1230
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Popriedades con Ficha Catastral ingresada por fecha"
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
Attribute VB_Name = "frmPropConFichasUrb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimirExel_Click()
    On Error GoTo FactEmitidas_Error
    Dim UrbOrural As Integer
    
    
  
        
        
    If optUrbanas.Value = True Then
        UrbOrural = 0
    ElseIf optRurales.Value = True Then
        UrbOrural = 1
    End If
    
    Dim ExpDim As New clsPropConFichasUrb
    ExpDim.CrearReporte
    ExpDim.SendToExcel

    Exit Sub
FactEmitidas_Error:
    MsgBox Err.Description
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


