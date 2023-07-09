VERSION 5.00
Begin VB.Form FrmAbonadosSpRubroBarrio 
   Caption         =   "Reportes"
   ClientHeight    =   4875
   ClientLeft      =   -15555
   ClientTop       =   4575
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4875
   ScaleWidth      =   5205
   Begin VB.ComboBox cboServicio 
      Height          =   315
      ItemData        =   "FrmAbonadosSpRubroBarrio.frx":0000
      Left            =   1560
      List            =   "FrmAbonadosSpRubroBarrio.frx":002E
      TabIndex        =   11
      Top             =   960
      Width           =   2895
   End
   Begin VB.ComboBox cboRubro 
      Height          =   315
      Left            =   1560
      TabIndex        =   10
      Top             =   1560
      Width           =   2895
   End
   Begin VB.OptionButton optInactivo 
      Caption         =   "Inactivos"
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   3360
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton optActivo 
      Caption         =   "Activos"
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   3960
      Width           =   1215
   End
   Begin VB.ComboBox cboBarrio 
      Height          =   315
      ItemData        =   "FrmAbonadosSpRubroBarrio.frx":00B0
      Left            =   1560
      List            =   "FrmAbonadosSpRubroBarrio.frx":00B7
      TabIndex        =   3
      Top             =   2760
      Width           =   2895
   End
   Begin VB.ComboBox cboAldea 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Servicio:"
      Height          =   195
      Left            =   720
      TabIndex        =   12
      Top             =   960
      Width           =   630
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Barrio"
      Height          =   195
      Left            =   720
      TabIndex        =   5
      Top             =   2760
      Width           =   420
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Abonados de Servicios Publicos por Rubro y Barrio"
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
      Left            =   150
      TabIndex        =   4
      Top             =   120
      Width           =   4995
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Aldea"
      Height          =   195
      Left            =   720
      TabIndex        =   1
      Top             =   2160
      Width           =   420
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Rubro:"
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   1560
      Width           =   480
   End
End
Attribute VB_Name = "FrmAbonadosSpRubroBarrio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboAldea_Click()
    cboBarrio.Clear
    CargarBarrio
End Sub
Private Sub cboServicio_Click()
    cboRubro.Clear
    CargarRubro
End Sub

Private Sub cmdImprimir_Click()
    Dim sqlBarrio As String
    Dim codBarrio As String
    Dim rubroCtaIngreso As String
    Dim CodAldea As String
    Dim Estado As Integer
    Dim txtRubroCtaIngreso As String

    
    With DeRia.rscmdCuentasIngreso
        .Requery
        .Find "NombreCtaIngreso='" & Trim(cboRubro.Text) & "'"
        If .EOF Then
           rubroCtaIngreso = " AND (SUBSTRING(CuentaIngreso_A.CtaIngreso, 4, 5) IN ('" & cboServicio.ItemData(cboServicio.ListIndex) & "')) OR (SUBSTRING(CuentaIngreso_A.CtaIngreso, 1, 9) IN ('152190210'))"
        Else
            If !NombreCtaIngreso = Trim(cboRubro.Text) Then
                txtRubroCtaIngreso = !CtaIngreso
                rubroCtaIngreso = "AND (CuentaIngreso_A.CtaIngreso = '" & txtRubroCtaIngreso & "')"
            Else
               MsgBox ("Ingrese una cuenta de ingreso valida para servicios publicos")
                cboRubro = ""
               cboRubro.SetFocus
                Exit Sub
            End If
        End If
    End With
    
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
                codBarrio = !codBarrio
                sqlBarrio = "AND (TablaBarrio.CodBarrio = '" & codBarrio & "')"
            Else
                MsgBox ("Ingrese un barrio valido")
                cboBarrio = ""
                cboBarrio.SetFocus
                Exit Sub
            End If
        End If
    End With
    
    If optActivo.Value = True Then
        Estado = 0
    ElseIf optInactivo.Value = True Then
        Estado = 1
    End If

    Dim expdim As New clsAbonadoSPPorRubroBarrio
    expdim.CrearReporte rubroCtaIngreso, Estado, sqlBarrio
    expdim.SendToExcel rubroCtaIngreso, codBarrio, CodAldea
    Unload Me
    'Exit Sub
End Sub
Private Sub cmdSalir_Click()
    Unload Me
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
Sub CargarRubro()
    If cboServicio.ItemData(cboServicio.ListIndex) = "11801" Then
        rubro = "AND (SUBSTRING(CtaIngreso, 4, 5) IN ('11801')) OR (SUBSTRING(CtaIngreso, 1, 9) IN ('152190210'))"
    ElseIf cboServicio.Text = "Alcantarillado Sanitario" Then
        rubro = "AND (SUBSTRING(CtaIngreso, 4, 5) IN ('11802')) OR (SUBSTRING(CtaIngreso, 1, 9) IN ('152190220'))"
    ElseIf cboServicio.Text = "Alumbrado Publico" Then
        rubro = "AND (SUBSTRING(CtaIngreso, 4, 5) IN ('11803')) OR (SUBSTRING(CtaIngreso, 1, 9) IN ('152190230'))"
    ElseIf cboServicio.Text = "Tren de Aseo" Then
        rubro = "AND (SUBSTRING(CtaIngreso, 4, 5) IN ('11804')) OR (SUBSTRING(CtaIngreso, 1, 9) IN ('152190250'))"
    ElseIf cboServicio.Text = "Conexiones y Reconexiones de agua Potable" Then
        rubro = "AND (SUBSTRING(CtaIngreso, 4, 5) IN ('11805')) "
    ElseIf cboServicio.Text = "Bomberos" Then
        rubro = "AND (SUBSTRING(CtaIngreso, 4, 5) IN ('11806')) OR (SUBSTRING(CtaIngreso, 1, 9) IN ('152190260'))"
    Else
        MsgBox ("Ingrese un Servicio")
        cboServicio = ""
        cboServicio.SetFocus
        Exit Sub
     End If

    strCodigo = "SELECT NombreCtaIngreso, CtaIngreso  From CuentaIngreso_A "
    strCodigo = strCodigo & "WHERE DATEPART(YEAR, GETDATE())= Anio " & rubro
    strCodigo = strCodigo & " ORDER BY NombreCtaIngreso "
    If DeRia.rscmdCuentasIngreso.State = 1 Then DeRia.rscmdCuentasIngreso.Close
    DeRia.rscmdCuentasIngreso.Open strCodigo
    cboRubro.Text = "[Seleccione un Rubro]"
    cboRubro.AddItem "Todos"
    Do Until DeRia.rscmdCuentasIngreso.EOF
    cboRubro.AddItem DeRia.rscmdCuentasIngreso!NombreCtaIngreso
    DeRia.rscmdCuentasIngreso.MoveNext
    Loop
End Sub

Private Sub Form_Load()
Dim rubro As String
CargarAldea
cboServicio.Text = "[Seleccione un Servicio]"
cboRubro.Text = "[Seleccione un Rubro]"
cboBarrio.Text = "[Seleccione un Barrio]"
End Sub


