VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rpAvisoDeCobro 
   Caption         =   "SAFT - rpAvisoDeCobro (ActiveReport)"
   ClientHeight    =   15135
   ClientLeft      =   -28800
   ClientTop       =   -2445
   ClientWidth     =   28800
   _ExtentX        =   50800
   _ExtentY        =   26696
   SectionData     =   "rpAvisoDeCobro.dsx":0000
End
Attribute VB_Name = "rpAvisoDeCobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rReportRs As New ADODB.Recordset
Public Function DatosGenerales(identidad As String)
    txtMontAdeuCalHastMes = Format(Now, "mmmm yyyy")
    strSql = "SELECT Pnombre, SNombre, PApellido, SApellido, direccion FROM Contribuyente WHERE (Identidad = '" & identidad & "')"
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open strSql
    txtMontAdeuNombreCont = Trim(DeRia.rsAbonadoSP!Pnombre) & " " & Trim(DeRia.rsAbonadoSP!sNombre) & " " & Trim(DeRia.rsAbonadoSP!PApellido) & " " & Trim(DeRia.rsAbonadoSP!sApellido)
    txtMontAdeuDireccion = Trim(DeRia.rsAbonadoSP!Direccion)
    txtFechaAviso = Format(Now, "dddd, dd  mmmm  yyyy")
End Function

Public Function Interes(ByVal identidad As String, ByVal anio As String) As Double
    Dim Sql As String
    Dim cantidad As String
    Sql = "SELECT AvPgDetalle.ValorUnitAvPgDet FROM AvPgDetalle INNER JOIN AvPgEnc ON AvPgDetalle.NumAvPg = AvPgEnc.NumAvPg "
    Sql = Sql & "WHERE (AvPgEnc.AvPgEstado = 1) AND (AvPgEnc.Identidad = '" & identidad & "') AND (AvPgDetalle.CtaIngreso = '11212601') AND (DATEPART(year, AvPgEnc.FechaVenceAvPg) = '" & anio & "')"
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open Sql
    cantidad = DeRia.rsAbonadoSP!ValorUnitAvPgDet
    Interes = cantidad
End Function
Public Function Recargos(ByVal identidad As String, ByVal anio As String) As Double
    Dim Sql As String
    Dim cantidad As String
    Sql = "SELECT AvPgDetalle.ValorUnitAvPgDet FROM AvPgDetalle INNER JOIN AvPgEnc ON AvPgDetalle.NumAvPg = AvPgEnc.NumAvPg "
    Sql = Sql & " WHERE (AvPgEnc.AvPgEstado = 1) AND (AvPgEnc.Identidad = '" & identidad & "') AND (AvPgDetalle.CtaIngreso = '11212101') AND (DATEPART(year, AvPgEnc.FechaVenceAvPg) = " & anio & ")"
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open Sql
    cantidad = DeRia.rsAbonadoSP!ValorUnitAvPgDet
    Recargos = cantidad
End Function

Public Function Impuesto(ByVal identidad As String)
    Dim Sql As String
    Sql = " SELECT AvPgDescripcion, DATEPART(year, AvPgEnc.FechaVenceAvPg) AS anio, ValorUnitAvPgDet, AvPgDetalle.CtaIngreso "
    Sql = Sql & " FROM AvPgDetalle INNER JOIN AvPgEnc ON AvPgDetalle.NumAvPg = AvPgEnc.NumAvPg "
    Sql = Sql & " WHERE (AvPgEnc.AvPgEstado = 1) AND(AvPgEnc.Identidad = '" & identidad & "') AND (SUBSTRING(AvPgDetalle.CtaIngreso, 1, 8) IN ('11212201')) "
    Sql = Sql & " ORDER BY AvPgDetalle.NumAvPg "
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open Sql
    Do While Not DeRia.rsAbonadoSP.EOF
        Me.txtMontAdeuConcepto = DeRia.rsAbonadoSP!AvPgDescripcion
        Me.txtMontAdeuAnioImpositivo = DeRia.rsAbonadoSP!anio
        Me.txtMontAdeuMonto = DeRia.rsAbonadoSP!ValorUnitAvPgDet
        'txtMontAdeuMulta
        Me.txtMontAdeuIntereses = Interes(identidad, DeRia.rsAbonadoSP!anio)
        Me.txtMontAdeuRecargos = Recargos(identidad, txtMontAdeuAnioImpositivo.Text)
        Me.txtMontAdeuValTotal = CDbl(txtMontAdeuMonto.Text) + CDbl(txtMontAdeuIntereses.Text) + CDbl(txtMontAdeuRecargos.Text)
        DeRia.rsAbonadoSP.MoveNext
    Loop
    DeRia.rsAbonadoSP.Close
    
End Function


Public Function Impuesto1(ByVal identidad As String)
    Dim Sql As String
    Sql = " SELECT AvPgDescripcion, DATEPART(year, AvPgEnc.FechaVenceAvPg) AS anio, ValorUnitAvPgDet, AvPgDetalle.CtaIngreso "
    Sql = Sql & " FROM AvPgDetalle INNER JOIN AvPgEnc ON AvPgDetalle.NumAvPg = AvPgEnc.NumAvPg "
    Sql = Sql & " WHERE (AvPgEnc.AvPgEstado = 1) AND(AvPgEnc.Identidad = '" & identidad & "') AND (SUBSTRING(AvPgDetalle.CtaIngreso, 1, 8) IN ('11212201')) "
    Sql = Sql & " ORDER BY AvPgDetalle.NumAvPg "
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open Sql
  If Not DeRia.rsAbonadoSP.EOF Then
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
      .Text = rsAbonadoSP!ValorUnitAvPgDet
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
      .Text = Interes(identidad, rsAbonadoSP!anio)
    End With
    Set Field6 = Me.Detail.Controls.Add("VB.TextBox")
    With Field6
      .Left = 6661
      .Top = Me.Detail.Height + 500
      .Width = 1276
      .Height = 425
      .Text = Recargos(identidad, rsAbonadoSP!anio)
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
    DatosGenerales ("1309194700200")
    Impuesto1 ("1309194700200")
End Sub
 

