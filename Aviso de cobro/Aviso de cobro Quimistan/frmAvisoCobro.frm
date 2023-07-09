VERSION 5.00
Begin VB.Form frmAvisoCobro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aviso de Cobro"
   ClientHeight    =   4200
   ClientLeft      =   11010
   ClientTop       =   3270
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   4920
   Begin VB.ComboBox cboBarrio 
      Height          =   315
      ItemData        =   "frmAvisoCobro.frx":0000
      Left            =   1920
      List            =   "frmAvisoCobro.frx":0007
      TabIndex        =   4
      Top             =   2280
      Width           =   2895
   End
   Begin VB.ComboBox cboAldea 
      Height          =   315
      Left            =   1920
      TabIndex        =   3
      Top             =   1800
      Width           =   2895
   End
   Begin VB.ComboBox cboTipoImpuesto 
      Height          =   315
      ItemData        =   "frmAvisoCobro.frx":0012
      Left            =   1920
      List            =   "frmAvisoCobro.frx":0014
      TabIndex        =   5
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox txtDias 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtMeses 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Barrio"
      Height          =   195
      Left            =   360
      TabIndex        =   12
      Top             =   2280
      Width           =   420
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Aldea"
      Height          =   195
      Left            =   360
      TabIndex        =   11
      Top             =   1800
      Width           =   420
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Impuesto"
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   2760
      Width           =   1230
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Cantidad de Dias"
      Height          =   195
      Left            =   360
      TabIndex        =   9
      Top             =   1200
      Width           =   1230
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Meses Adeudados"
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   600
      Width           =   1320
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Impresion de Avisos de Cobro Por Contribuyente"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4785
   End
End
Attribute VB_Name = "frmAvisoCobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rsCont As New ADODB.Recordset
Private AbonadoSPDet As New ADODB.Recordset
Dim rpt As New rpAvisoDeCobro



Private Sub cboTipoImpuesto_Click()
    If cboTipoImpuesto.Text = "Bienes inmuebles" Then
        tipoImpuesto = 1
    ElseIf cboTipoImpuesto.Text = "Industria comercio y servicio" Then
         tipoImpuesto = 2
    ElseIf cboTipoImpuesto.Text = "Servicios públicos" Then
    

    tipoImpuesto = 5
    End If
End Sub

Private Sub cmdImprimir_Click()


    With DeRia.rscmdAldeas
        .Requery
        .Find "NombreAldea='" & Trim(cboAldea.Text) & "'"
        If .EOF Then
            MsgBox "No se encontro la Aldea"
            Exit Sub
        Else
            If !NombreAldea = Trim(cboAldea.Text) Then
                CodAldea = !CodAldea
            Else
                MsgBox ("Ingrese Aldea valida")
                cboRubro = ""
                cboRubro.SetFocus
                Exit Sub
            End If
        End If
    End With
    
    
    With DeRia.rscmdTablaBarrio
        .Requery
        .Find "NombreBarrio='" & Trim(cboBarrio.Text) & "'"
        If .EOF Then
            sqlBarrio = ""
            codBarrio = "Todos"
        Else
            If !NombreBarrio = Trim(cboBarrio.Text) Then
                codigoBarrio = Trim(!codBarrio)
                sqlBarrio = "AND (TablaBarrio.CodBarrio = '" & codBarrio & "')"
            Else
                MsgBox ("Ingrese un barrio valido")
                cboBarrio = ""
                cboBarrio.SetFocus
                Exit Sub
            End If
        End If
    End With
    mesesAdeudados = Me.txtMeses.Text
    
    diasParaPago = Me.txtDias.Text
    
    


            Dim rpt As New rpAvisoDeCobro
            rpt.Run (True)
            rpt.Show
     
End Sub
Sub CargarBarrio()
    With DeRia.rscmdAldeas
        .Requery
        .Find "NombreAldea='" & Trim(cboAldea.Text) & "'"
        If .EOF Then
            MsgBox "No se encontro la Aldea"
            Exit Sub
        Else
            If !NombreAldea = Trim(cboAldea.Text) Then
                CodAldea = !CodAldea
                Dim strCodBarrio As String
                strCodBarrio = "SELECT * FROM TablaBarrio WHERE (CodAldea = '" & CodAldea & "')"
                If DeRia.rscmdTablaBarrio.State = 1 Then DeRia.rscmdTablaBarrio.Close
                DeRia.rscmdTablaBarrio.Open strCodBarrio
                cboBarrio.Text = "[Seleccione un barrio]"
                cboBarrio.AddItem "Todos"
                Do Until DeRia.rscmdTablaBarrio.EOF
                cboBarrio.AddItem DeRia.rscmdTablaBarrio!NombreBarrio
                cboBarrio.ItemData(cboBarrio.NewIndex) = DeRia.rscmdTablaBarrio.Collect("CodBarrio")
                DeRia.rscmdTablaBarrio.MoveNext
                Loop
            Else
                MsgBox ("Ingrese Aldea valida")
                cboRubro = ""
                cboRubro.SetFocus
                Exit Sub
            End If
        End If
    End With
End Sub
Sub CargarAldea()
    If DeRia.rscmdAldeas.State = 1 Then DeRia.rscmdAldeas.Close
    DeRia.rscmdAldeas.Open
    cboAldea.Text = "[Seleccione una aldea]"
    Do Until DeRia.rscmdAldeas.EOF
    cboAldea.AddItem DeRia.rscmdAldeas!NombreAldea
    cboAldea.ItemData(cboAldea.NewIndex) = DeRia.rscmdAldeas.Collect("CodAldea")
    DeRia.rscmdAldeas.MoveNext
    Loop
End Sub
Private Sub cmdSalir_Click()
 Unload Me
End Sub
Private Sub cboAldea_Click()
    cboBarrio.Clear
    CargarBarrio
End Sub

Private Sub Form_Load()
    CargarAldea
    cboAldea.Text = "[Seleccione un Barrio]"
    cboBarrio.Text = "[Seleccione una Aldea]"
    cboTipoImpuesto = "[Seleccione el Impuesto]"
    cboTipoImpuesto.AddItem "Bienes inmuebles"
    cboTipoImpuesto.AddItem "Industria comercio y servicio"
    cboTipoImpuesto.AddItem "Servicios públicos"
End Sub


