VERSION 5.00
Begin VB.Form RegistroAbonados 
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   3975
   ClientTop       =   2400
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   7200
   Begin VB.PictureBox Picture1 
      Height          =   5535
      Left            =   240
      ScaleHeight     =   5475
      ScaleWidth      =   6675
      TabIndex        =   0
      Top             =   240
      Width           =   6735
      Begin VB.TextBox txtValor 
         Height          =   375
         Left            =   2760
         TabIndex        =   14
         Top             =   3720
         Width           =   2415
      End
      Begin VB.TextBox txtcuenta 
         Height          =   375
         Left            =   2760
         TabIndex        =   13
         Top             =   3000
         Width           =   2415
      End
      Begin VB.TextBox txtFecha 
         Height          =   375
         Left            =   2760
         TabIndex        =   12
         Top             =   2400
         Width           =   2415
      End
      Begin VB.TextBox txtDireccion 
         Height          =   375
         Left            =   2760
         TabIndex        =   11
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox txtClavecatastro 
         Height          =   375
         Left            =   2760
         TabIndex        =   10
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtDNI 
         Height          =   375
         Left            =   2760
         TabIndex        =   9
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtCodAbonado 
         Height          =   375
         Left            =   2760
         TabIndex        =   8
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "VALOR"
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
         Left            =   720
         TabIndex        =   7
         Top             =   3720
         Width           =   675
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "CUENTA DE INGRESO"
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
         Left            =   480
         TabIndex        =   6
         Top             =   3120
         Width           =   2040
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ULTIMO PERIODO"
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
         Left            =   480
         TabIndex        =   5
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "DIRECCION FACTURA"
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
         Left            =   480
         TabIndex        =   4
         Top             =   1920
         Width           =   2040
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "CODIGO ABONADO"
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
         Left            =   480
         TabIndex        =   3
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "CLAVE CATASTRO"
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
         Left            =   480
         TabIndex        =   2
         Top             =   1440
         Width           =   1770
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DNI"
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
         Left            =   480
         TabIndex        =   1
         Top             =   1080
         Width           =   330
      End
   End
End
Attribute VB_Name = "RegistroAbonados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Dim Osp As New Registro
    Osp.IngresoAbonados

End Sub
