VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rpAvisoDeCobro 
   Caption         =   "SAFT - rpAvisoDeCobro (ActiveReport)"
   ClientHeight    =   10305
   ClientLeft      =   5610
   ClientTop       =   1590
   ClientWidth     =   10155
   _ExtentX        =   17912
   _ExtentY        =   18177
   SectionData     =   "rpAvisoDeCobro.dsx":0000
End
Attribute VB_Name = "rpAvisoDeCobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rReportRs As New ADODB.Recordset
Public Function DatosGenerales(Identidad As String)
    lbMunicipio = ""
    Dim sql As String
    sql = "SELECT NombreMuni FROM Parametro"
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open sql
    lbMunicipio = Trim(DeRia.rsAbonadoSP!NombreMuni)
    txtMontAdeuCalHastMes = Format(Now, "mmmm yyyy")
    strSql = "SELECT Pnombre, SNombre, PApellido, SApellido, direccion FROM Contribuyente WHERE (Identidad = '" & Identidad & "')"
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open strSql
    txtMontAdeuNombreCont = Trim(DeRia.rsAbonadoSP!PNombre) & " " & Trim(DeRia.rsAbonadoSP!sNombre) & " " & Trim(DeRia.rsAbonadoSP!pApellido) & " " & Trim(DeRia.rsAbonadoSP!sApellido)
    txtMontAdeuDireccion = Trim(DeRia.rsAbonadoSP!Direccion)
    txtFechaAviso = Format(Now, "dddd, dd  mmmm  yyyy")
End Function

Public Function Interes(ByVal Identidad As String, ByVal anio As String, ByVal cuenta As String) As Double
    Dim sql As String
    Dim Cantidad As String
    Dim cuenta2 As String
    Dim tipoInteres As String
    cuenta2 = Left(cuenta, 6)
    If cuenta2 = "112122" Then ' cuenta recuperacion impuestos
        tipoInteres = "11212601"
    ElseIf cuenta2 = "112123" Then ' cuenta recuperacion servicios
        tipoInteres = "11212602"
    End If
    sql = " SELECT sum(AvPgDetalle.ValorUnitAvPgDet) as total , ctaingreso FROM AvPgDetalle INNER JOIN AvPgEnc ON AvPgDetalle.NumAvPg = AvPgEnc.NumAvPg"
    sql = sql & " WHERE (AvPgEnc.AvPgEstado = 1) AND (AvPgEnc.Identidad = '" & Identidad & "') AND substring(AvPgDetalle.CtaIngreso,1,6) = ('112126') AND (DATEPART(year, AvPgEnc.FechaVenceAvPg) = " & anio & " )  and ctaingreso = '" & tipoInteres & "'"
    sql = sql & " group by AvPgDetalle.CtaIngreso"
    
    'sql = "SELECT AvPgDetalle.ValorUnitAvPgDet FROM AvPgDetalle INNER JOIN AvPgEnc ON AvPgDetalle.NumAvPg = AvPgEnc.NumAvPg "
    'sql = sql & "WHERE (AvPgEnc.AvPgEstado = 1) AND (AvPgEnc.Identidad = '" & Identidad & "') AND substring(AvPgDetalle.CtaIngreso,1,6) '11212601') AND (DATEPART(year, AvPgEnc.FechaVenceAvPg) = '" & Anio & "')"
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open sql
    Cantidad = DeRia.rsAbonadoSP!Total
    Interes = Cantidad
End Function
Public Function Recargos(ByVal Identidad As String, ByVal anio As String, ByVal cuenta As String) As Double
    Dim sql As String
    Dim Cantidad As String
    Dim cuenta2 As String
    Dim tipoRecargo As String
    cuenta2 = Left(cuenta, 6)
    If cuenta2 = "112122" Then ' cuenta recuperacion impuestos
        tipoRecargo = "11212101"
    ElseIf cuenta2 = "112123" Then ' cuenta recuperacion servicios
        tipoRecargo = "11212102"
    End If
    sql = " SELECT sum(AvPgDetalle.ValorUnitAvPgDet) as total , ctaingreso FROM AvPgDetalle INNER JOIN AvPgEnc ON AvPgDetalle.NumAvPg = AvPgEnc.NumAvPg"
    sql = sql & " WHERE (AvPgEnc.AvPgEstado = 1) AND (AvPgEnc.Identidad = '" & Identidad & "') AND substring(AvPgDetalle.CtaIngreso,1,6) = ('112121') AND (DATEPART(year, AvPgEnc.FechaVenceAvPg) = " & anio & " )  and ctaingreso = '" & tipoRecargo & "'"
    sql = sql & " group by AvPgDetalle.CtaIngreso"
    
    'sql = "SELECT AvPgDetalle.ValorUnitAvPgDet FROM AvPgDetalle INNER JOIN AvPgEnc ON AvPgDetalle.NumAvPg = AvPgEnc.NumAvPg "
    'sql = sql & "WHERE (AvPgEnc.AvPgEstado = 1) AND (AvPgEnc.Identidad = '" & Identidad & "') AND substring(AvPgDetalle.CtaIngreso,1,6) '11212601') AND (DATEPART(year, AvPgEnc.FechaVenceAvPg) = '" & Anio & "')"
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open sql
    Cantidad = DeRia.rsAbonadoSP!Total
    Recargos = Cantidad
End Function

Public Function Impuesto(ByVal Identidad As String)
    Dim sql As String
    Dim cuenta As String
    Dim anio As String
    sql = " SELECT AvPgDescripcion, DATEPART(year, AvPgEnc.FechaVenceAvPg) AS anio, ValorUnitAvPgDet, AvPgDetalle.CtaIngreso "
    sql = sql & " FROM AvPgDetalle INNER JOIN AvPgEnc ON AvPgDetalle.NumAvPg = AvPgEnc.NumAvPg "
    sql = sql & " WHERE (AvPgEnc.AvPgEstado = 1) AND(AvPgEnc.Identidad = '" & Identidad & "') AND (SUBSTRING(AvPgDetalle.CtaIngreso, 1, 8) IN ('11212201')) "
    sql = sql & " ORDER BY AvPgDetalle.NumAvPg "
    
    sql = " SELECT AvPgEnc.AvPgDescripcion, DATEPART(year, AvPgEnc.FechaVenceAvPg) AS aniox, SUM(AvPgDetalle.ValorUnitAvPgDet) AS total, AvPgDetalle.CtaIngreso"
    sql = sql & " FROM AvPgDetalle INNER JOIN"
    sql = sql & " AvPgEnc ON AvPgDetalle.NumAvPg = AvPgEnc.NumAvPg INNER JOIN"
    sql = sql & " CuentaIngreso_A ON AvPgDetalle.CtaIngreso = CuentaIngreso_A.CtaIngreso AND DATEPART(year, AvPgEnc.FechaVenceAvPg) = CuentaIngreso_A.Anio"
    sql = sql & " WHERE (AvPgEnc.AvPgEstado = 1) AND (AvPgEnc.Identidad = '" & Identidad & "') AND (AvPgEnc.AvPgTipoImpuesto IN (1, 5)) AND (CuentaIngreso_A.Tipo <> 2)"
    sql = sql & " GROUP BY AvPgEnc.AvPgDescripcion, DATEPART(year, AvPgEnc.FechaVenceAvPg), AvPgDetalle.CtaIngreso"
    sql = sql & " ORDER BY aniox"

    
    
    
    
    
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open sql
    MsgBox (DeRia.rsAbonadoSP.EOF)
    Do While Not DeRia.rsAbonadoSP.EOF
        cuenta = DeRia.rsAbonadoSP!CtaIngreso
        anio = DeRia.rsAbonadoSP!aniox
        Me.txtMontAdeuConcepto = DeRia.rsAbonadoSP!AvPgDescripcion
        Me.txtMontAdeuAnioImpositivo = DeRia.rsAbonadoSP!aniox
        Me.txtMontAdeuMonto = DeRia.rsAbonadoSP!Total
        'txtMontAdeuMulta
        Me.txtMontAdeuIntereses = Interes(Identidad, anio, cuenta)
        Me.txtMontAdeuRecargos = Recargos(Identidad, anio, cuenta)
        Me.txtMontAdeuValTotal = CDbl(txtMontAdeuMonto.Text) + CDbl(txtMontAdeuIntereses.Text) + CDbl(txtMontAdeuRecargos.Text)
        DeRia.rsAbonadoSP.MoveNext
    Loop
    DeRia.rsAbonadoSP.Close
    
End Function


Public Function Impuesto1(ByVal Identidad As String)
    Dim sql As String
    Dim cuenta As String
    Dim anio As String
    sql = " SELECT AvPgEnc.AvPgDescripcion, DATEPART(year, AvPgEnc.FechaVenceAvPg) AS aniox, SUM(AvPgDetalle.ValorUnitAvPgDet) AS total, AvPgDetalle.CtaIngreso"
    sql = sql & " FROM AvPgDetalle INNER JOIN"
    sql = sql & " AvPgEnc ON AvPgDetalle.NumAvPg = AvPgEnc.NumAvPg INNER JOIN"
    sql = sql & " CuentaIngreso_A ON AvPgDetalle.CtaIngreso = CuentaIngreso_A.CtaIngreso AND DATEPART(year, AvPgEnc.FechaVenceAvPg) = CuentaIngreso_A.Anio"
    sql = sql & " WHERE (AvPgEnc.AvPgEstado = 1) AND (AvPgEnc.Identidad = '" & Identidad & "') AND (AvPgEnc.AvPgTipoImpuesto IN (1, 5)) AND (CuentaIngreso_A.Tipo <> 2)"
    sql = sql & " GROUP BY AvPgEnc.AvPgDescripcion, DATEPART(year, AvPgEnc.FechaVenceAvPg), AvPgDetalle.CtaIngreso"
    sql = sql & " ORDER BY aniox"
    
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open sql
    
    If Not DeRia.rsAbonadoSP.EOF Then
    cuenta = DeRia.rsAbonadoSP!CtaIngreso
    anio = DeRia.rsAbonadoSP!aniox
    
    Me.Detail.Height = Me.Detail.Height
    Set Field1 = Me.Detail.Controls.Add("VB.TextBox")
    With Field1
      .Left = 0
      .Top = Me.Detail.Height + 500
      .Width = 2551
      .Height = 425
      .Text = rsAbonadoSP!AvPgDescripcion
    End With
    Set Field2 = Me.Detail.Controls.Add("VB.TextBox")
    With Field2
      .Left = 2551
      .Top = Me.Detail.Height + 500
      .Width = 1276
      .Height = 425
      .Text = rsAbonadoSP!anio
    End With
    Set Field3 = Me.Detail.Controls.Add("VB.TextBox")
    With Field3
      .Left = 3827
      .Top = Me.Detail.Height + 500
      .Width = 1275
      .Height = 425
      .Text = rsAbonadoSP!Total
    End With
    Set Field4 = Me.Detail.Controls.Add("VB.TextBox")
    With Field4
      .Left = 5100
      .Top = Me.Detail.Height + 500
      .Width = 1275
      .Height = 425
      .Text = ""
    End With
    Set Field5 = Me.Detail.Controls.Add("VB.TextBox")
    With Field5
      .Left = 6378
      .Top = Me.Detail.Height + 500
      .Width = 1276
      .Height = 425
      .Text = Interes(Identidad, anio, cuenta)
    End With
    Set Field6 = Me.Detail.Controls.Add("VB.TextBox")
    With Field6
      .Left = 6661
      .Top = Me.Detail.Height + 500
      .Width = 1276
      .Height = 425
      .Text = Recargos(Identidad, anio, cuenta)
    End With
    Set Field7 = Me.Detail.Controls.Add("VB.TextBox")
    With Field7
      .Left = 8929
      .Top = Me.Detail.Height + 500
      .Width = 1276
      .Height = 425
      .Text = CDbl(Field3.Text) + CDbl(Field5.Text) + CDbl(Field6.Text)
    End With
    RSTDATA.MoveNext
  End If
End Function






Private Sub ActiveReport_ReportStart()
    DatosGenerales (frmAvisoCobro.txtIdentidad)
    Impuesto1 (frmAvisoCobro.txtIdentidad)
End Sub
 

