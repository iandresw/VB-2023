VERSION 5.00
Begin VB.Form frmInfoTaxi 
   Caption         =   "PO TAXI"
   ClientHeight    =   8370
   ClientLeft      =   165
   ClientTop       =   1905
   ClientWidth     =   15000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   15000
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
      Height          =   7935
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   14535
      Begin VB.CommandButton cmdActualizar 
         Appearance      =   0  'Flat
         Caption         =   "Actualizar Datos Moto Taxi"
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   495
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Frame Frame1 
         Caption         =   "INFORMACION MOTO TAXI"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   8280
         TabIndex        =   29
         Top             =   840
         Width           =   5535
         Begin VB.TextBox txtNumeroMoto 
            Height          =   405
            Left            =   2880
            TabIndex        =   32
            Top             =   1320
            Width           =   1935
         End
         Begin VB.TextBox txtColor 
            Height          =   405
            Left            =   2880
            TabIndex        =   31
            Top             =   1920
            Width           =   1935
         End
         Begin VB.TextBox txtPlaca 
            Height          =   405
            Left            =   2880
            TabIndex        =   30
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COLOR:"
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
            Left            =   240
            TabIndex        =   35
            Top             =   1920
            Width           =   765
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NUMERO DE MOTO TAXI"
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
            Left            =   240
            TabIndex        =   34
            Top             =   1440
            Width           =   2325
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NUMERO DE PLACA:"
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
            Left            =   240
            TabIndex        =   33
            Top             =   840
            Width           =   1995
         End
      End
      Begin VB.TextBox txtnumrecibo 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   480
         TabIndex        =   18
         Top             =   900
         Width           =   1400
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   4300
         TabIndex        =   17
         Top             =   900
         Width           =   1400
      End
      Begin VB.TextBox txtPeriodo 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   6200
         TabIndex        =   16
         Top             =   900
         Width           =   1400
      End
      Begin VB.TextBox txtNombreEstablecimiento 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   500
         TabIndex        =   15
         Top             =   1700
         Width           =   7100
      End
      Begin VB.TextBox txtPropietario 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   500
         TabIndex        =   14
         Top             =   2500
         Width           =   7100
      End
      Begin VB.TextBox txtUbicacion 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   500
         TabIndex        =   13
         Top             =   3240
         Width           =   7100
      End
      Begin VB.TextBox txtActividad 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   500
         TabIndex        =   12
         Top             =   4040
         Width           =   7100
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command1"
         Height          =   375
         Left            =   -840
         TabIndex        =   11
         Top             =   1440
         Width           =   75
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   495
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   5040
         Width           =   2175
      End
      Begin VB.CommandButton cmdImprimir 
         Appearance      =   0  'Flat
         Caption         =   "Imprimir"
         Height          =   495
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   6000
         Width           =   2175
      End
      Begin VB.CommandButton cmdSalir 
         Appearance      =   0  'Flat
         Caption         =   "Salir"
         Height          =   495
         Left            =   11400
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6000
         Width           =   2175
      End
      Begin VB.TextBox txtFechaNac 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   500
         TabIndex        =   7
         Top             =   4840
         Width           =   7100
      End
      Begin VB.TextBox txtTipo 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   500
         TabIndex        =   6
         Top             =   5640
         Width           =   7100
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
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
         TabIndex        =   3
         Top             =   6120
         Width           =   7095
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
            Left            =   1200
            TabIndex        =   5
            Top             =   480
            Width           =   1815
         End
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
            Left            =   4200
            TabIndex        =   4
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.TextBox txtRTM 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   900
         Width           =   1400
      End
      Begin VB.CommandButton cmdGuardar 
         Appearance      =   0  'Flat
         Caption         =   "Guardar"
         Height          =   495
         Left            =   11400
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5040
         Width           =   2175
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
         TabIndex        =   28
         Top             =   600
         Width           =   945
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
         TabIndex        =   27
         Top             =   1395
         Width           =   4095
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
         TabIndex        =   26
         Top             =   2205
         Width           =   2745
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
         TabIndex        =   25
         Top             =   2940
         Width           =   1155
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
         TabIndex        =   24
         Top             =   3735
         Width           =   2415
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
         TabIndex        =   23
         Top             =   4545
         Width           =   2175
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
         TabIndex        =   22
         Top             =   5340
         Width           =   2775
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
         TabIndex        =   21
         Top             =   600
         Width           =   1155
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
         TabIndex        =   20
         Top             =   600
         Width           =   420
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
         Left            =   4290
         TabIndex        =   19
         Top             =   600
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmInfoTaxi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim crearPermiso As New clsCrearPermiso
Private rsAbonadoSP As New ADODB.Recordset

Private Sub ActualizarPlaca(txtRTM As String)
    sql = "UPDATE Contribuyente SET PlacaMotoTaxi = '" + Me.txtPlaca.Text + "', NumMotoTaxi = '" + Me.txtNumeroMoto + "', ColorMotoTaxi = '" + Me.txtColor + "' Where identidad = '" + txtRTM + "' and PlacaMotoTaxi is NULL"
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open (sql)
End Sub

Private Sub cmdActualizar_Click()
        ActualizarPlaca Me.txtRTM.Text
        MsgBox "Datos Ingresados.", vbInformation
        Me.cmdActualizar.Enabled = False
End Sub

Private Sub cmdGuardar_Click()
    crearPermiso.Guardar Me
End Sub
Private Sub cmdBuscar_Click()
    Dim numRecibo As Long
    numRecibo = CLng(Me.txtnumrecibo)
    crearPermiso.CrearPermisoOP numRecibo, Me
    If Not (DatosMoto(Me.txtRTM)) Then
        MsgBox "Ingrese los datos de la Moto taxi.", vbInformation
        Me.txtPlaca.SetFocus
        Me.cmdActualizar.Enabled = True
    Else
        cargaDatosMotoTaxi (Me.txtRTM.Text)
    End If
End Sub


Private Sub cmdImprimir_Click()
    Dim rpt As New PermisoOperacion
    rpt.lbColor.Visible = True
    rpt.lbPlaca.Visible = True
    rpt.lbNumeroMoto.Visible = True
    rpt.txtColor = Me.txtColor
    rpt.txtPlaca = Me.txtPlaca
    rpt.txtNumMoto = Me.txtNumeroMoto
    crearPermiso.Imprimir Me, rpt
End Sub


Private Function DatosMoto(rtm As String) As Boolean
    Dim sql As String
    sql = "SELECT COUNT(*) From Contribuyente WHERE (Identidad = '" + rtm + "') AND (NumMotoTaxi IS NOT NULL) AND (PlacaMotoTaxi IS NOT NULL)"
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open sql
    DatosMoto = (DeRia.rsAbonadoSP.Fields(0).Value > 0)
End Function



Private Sub cargaDatosMotoTaxi(rtm As String)
    Dim sql As String
    sql = "SELECT PlacaMotoTaxi, NumMotoTaxi, ColorMotoTaxi FROM Contribuyente WHERE (Identidad = '" + rtm + "')"
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open sql
    Me.txtColor = DeRia.rsAbonadoSP!ColorMotoTaxi
    Me.txtPlaca = DeRia.rsAbonadoSP!PlacaMotoTaxi
    Me.txtNumeroMoto = DeRia.rsAbonadoSP!NumMotoTaxi
    
End Sub












Private Sub cmdSalir_Click()
 Unload Me
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Form_Load()
    frmInfoTaxi.BackColor = RGB(29, 127, 146)
    Me.Frame1.BackColor = RGB(29, 127, 146)
    Me.Frame2.BackColor = RGB(29, 127, 146)
    Me.Frame3.BackColor = RGB(29, 127, 146)
    Me.optApertura.BackColor = RGB(29, 127, 146)
    Me.optRenovacion.BackColor = RGB(29, 127, 146)
    Me.cmdBuscar.BackColor = RGB(231, 73, 54)
    Me.cmdGuardar.BackColor = RGB(231, 73, 54)
    Me.cmdImprimir.BackColor = RGB(231, 73, 54)
    Me.cmdSalir.BackColor = RGB(231, 73, 54)
    Me.cmdActualizar.BackColor = RGB(231, 73, 54)
End Sub

