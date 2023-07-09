VERSION 5.00
Begin VB.Form fmrPerimisosOp 
   Caption         =   "Permisos de Operacion"
   ClientHeight    =   9405
   ClientLeft      =   3825
   ClientTop       =   1290
   ClientWidth     =   8355
   Icon            =   "frmPermisoOp.frx":0000
   ScaleHeight     =   9405
   ScaleWidth      =   8355
   Begin VB.Frame Frame2 
      Caption         =   "Busqueda por Recibo"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   7680
         Width           =   2175
      End
      Begin VB.TextBox txtRTM 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   2400
         TabIndex        =   17
         Top             =   900
         Width           =   1400
      End
      Begin VB.Frame Frame3 
         Caption         =   "Apertura o Renovacion"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   480
         TabIndex        =   14
         Top             =   6120
         Width           =   7095
         Begin VB.OptionButton optApertura 
            Caption         =   "Apertura"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4560
            MaskColor       =   &H80000018&
            TabIndex        =   16
            Top             =   600
            UseMaskColor    =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optRenovacion 
            Caption         =   "Renovacion"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   15
            Top             =   600
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.TextBox txtTipo 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   500
         TabIndex        =   13
         Top             =   5640
         Width           =   7100
      End
      Begin VB.TextBox txtFechaNac 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   500
         TabIndex        =   12
         Top             =   4840
         Width           =   7100
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   8520
         Width           =   2175
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         Height          =   495
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   7680
         Width           =   2175
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   495
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   7680
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command1"
         Height          =   375
         Left            =   -840
         TabIndex        =   8
         Top             =   1440
         Width           =   75
      End
      Begin VB.TextBox txtActividad 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   500
         TabIndex        =   7
         Top             =   4040
         Width           =   7100
      End
      Begin VB.TextBox txtUbicacion 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   500
         TabIndex        =   6
         Top             =   3240
         Width           =   7100
      End
      Begin VB.TextBox txtPropietario 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   500
         TabIndex        =   5
         Top             =   2500
         Width           =   7100
      End
      Begin VB.TextBox txtNombreEstablecimiento 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   500
         TabIndex        =   4
         Top             =   1700
         Width           =   7100
      End
      Begin VB.TextBox txtPeriodo 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   6200
         TabIndex        =   3
         Top             =   900
         Width           =   1400
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   4300
         TabIndex        =   2
         Top             =   900
         Width           =   1400
      End
      Begin VB.TextBox txtnumrecibo 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   900
         Width           =   1400
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CODIGO"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4305
         TabIndex        =   28
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RTM"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2400
         TabIndex        =   27
         Top             =   600
         Width           =   420
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. RECIBO"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   465
         TabIndex        =   26
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO DE ESTABLECIMIENTO"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   435
         TabIndex        =   25
         Top             =   5340
         Width           =   2775
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA ESTABLECIDO"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   450
         TabIndex        =   24
         Top             =   4545
         Width           =   2175
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ACTIVIDAD ECONOMICA"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   420
         TabIndex        =   23
         Top             =   3735
         Width           =   2415
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UBICACION"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   450
         TabIndex        =   22
         Top             =   2940
         Width           =   1155
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOMBRE DEL PROPIETARIO"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   465
         TabIndex        =   21
         Top             =   2205
         Width           =   2745
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOMBRE ESTABLECIMIENTO COMERCIAL"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   420
         TabIndex        =   20
         Top             =   1395
         Width           =   4095
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PERIODO"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6180
         TabIndex        =   19
         Top             =   600
         Width           =   945
      End
   End
End
Attribute VB_Name = "fmrPerimisosOp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim crearPermiso As New clsCrearPermiso
Private rsAbonadoSP As New ADODB.Recordset
Private Sub cmdBuscar_Click()
    Dim numRecibo As Long
    numRecibo = CLng(Me.txtnumrecibo)
    crearPermiso.CrearPermisoOP numRecibo, Me
End Sub
Private Sub cmdGuardar_Click()
    crearPermiso.Guardar Me
End Sub

Private Sub cmdImprimir_Click()
    Dim rpt As New PermisoOperacion
    crearPermiso.Imprimir Me, rpt

End Sub
Private Sub cmdSalir_Click()
 Unload Me
End Sub

Private Sub Form_Load()
    Me.Frame2.BackColor = RGB(29, 127, 146)
    fmrPerimisosOp.BackColor = RGB(29, 127, 146)
    Me.Frame3.BackColor = RGB(29, 127, 146)
    Me.optApertura.BackColor = RGB(29, 127, 146)
    Me.optRenovacion.BackColor = RGB(29, 127, 146)
    Me.cmdBuscar.BackColor = RGB(231, 73, 54)
    Me.cmdGuardar.BackColor = RGB(231, 73, 54)
    Me.cmdImprimir.BackColor = RGB(231, 73, 54)
    Me.cmdSalir.BackColor = RGB(231, 73, 54)
    
  
End Sub

