VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptPlanesPag 
   Caption         =   "SAFT - RptPlanesPag (ActiveReport)"
   ClientHeight    =   8775
   ClientLeft      =   1290
   ClientTop       =   3690
   ClientWidth     =   18765
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   33099
   _ExtentY        =   15478
   SectionData     =   "RptPlanesPag.dsx":0000
End
Attribute VB_Name = "RptPlanesPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsParam As New ADODB.Recordset
Dim rsCont As New ADODB.Recordset
Dim rsAldea As New ADODB.Recordset
Dim VtotalToT, VTotalPag, VTotalPen As Currency

Private Sub ActiveReport_ReportStart()
Dim Str As String
Dim Vval As Currency
VtotalToT = 0
VTotalPag = 0
VTotalPen = 0
TxtFecha.Text = Now
Set RsParam = DeRia.CoRia.Execute("Select * From ParametroCont ")
Me.TxtAlcaldia.Text = Trim(RsParam!NombreEmpresa)

Set rsCont = DeRia.CoRia.Execute("Update PlanPago set EstadoPP = 0 where EstadoPP is null ")
Set rsCont = DeRia.CoRia.Execute("Update PlanPago set TotalPagadoPP = 0 where TotalPagadoPP is null ")

Str = "SELECT PlanPago.SeqPP, SUM(AvPgDetalle.ValorUnitAvPgDet) AS Total, PlanPago.MontoPP FROM PlanPagoFactura INNER JOIN PlanPago ON PlanPagoFactura.SeqPP = PlanPago.SeqPP INNER JOIN "
Str = Str & "AvPgEnc ON PlanPagoFactura.NumAvPg = AvPgEnc.NumAvPg INNER JOIN AvPgDetalle ON AvPgEnc.NumAvPg = AvPgDetalle.NumAvPg "
Str = Str & " Where (PlanPago.EstadoPP = 0) And (AvPgEnc.AvPgEstado <> 3) GROUP BY PlanPago.SeqPP, PlanPago.MontoPP  "

Set rsCont = DeRia.CoRia.Execute(Str)

If Not rsCont.EOF Then

Do While Not rsCont.EOF
Vval = rsCont!MontoPP - rsCont!Total '
If Vval < 1 Then

DeRia.CoRia.Execute ("Update PlanPago set EstadoPP = 1 where SeqPP = " & rsCont!SeqPP & " ")
End If
rsCont.MoveNext
Loop

End If


Str = " SELECT COUNT(AvPgEnc.NumAvPg) AS Cuenta, PlanPago.NumCuotasPP, PlanPago.SeqPP FROM PlanPago INNER JOIN "
Str = Str & " PlanPagoFactura ON PlanPago.SeqPP = PlanPagoFactura.SeqPP INNER JOIN AvPgEnc ON PlanPagoFactura.NumAvPg = AvPgEnc.NumAvPg "
Str = Str & " Where (PlanPago.EstadoPP = 0) And (AvPgEnc.AvPgEstado = 2) GROUP BY PlanPago.NumCuotasPP, PlanPago.SeqPP "
Set rsCont = DeRia.CoRia.Execute(Str)

If Not rsCont.EOF Then

Do While Not rsCont.EOF
Vval = rsCont!NumCuotasPP - rsCont!Cuenta
If Vval = 0 Then

DeRia.CoRia.Execute ("Update PlanPago set EstadoPP = 1 where SeqPP = " & rsCont!SeqPP & " ")
End If
rsCont.MoveNext
Loop

End If

Str = "SELECT PlanPago.SeqPP, SUM(AvPgDetalle.ValorUnitAvPgDet) AS Total, PlanPago.MontoPP FROM PlanPagoFactura INNER JOIN PlanPago ON PlanPagoFactura.SeqPP = PlanPago.SeqPP INNER JOIN "
Str = Str & "AvPgEnc ON PlanPagoFactura.NumAvPg = AvPgEnc.NumAvPg INNER JOIN AvPgDetalle ON AvPgEnc.NumAvPg = AvPgDetalle.NumAvPg "
Str = Str & " Where (PlanPago.EstadoPP = 0) And (AvPgEnc.AvPgEstado = 3) GROUP BY PlanPago.SeqPP, PlanPago.MontoPP  "

Set rsCont = DeRia.CoRia.Execute(Str)

If Not rsCont.EOF Then

Do While Not rsCont.EOF

DeRia.CoRia.Execute ("Update PlanPago set EstadoPP = 1 where SeqPP = " & rsCont!SeqPP & " ")

rsCont.MoveNext
Loop

End If

If VxFechaPP = 0 Then

Str = " SELECT COUNT(AvPgEnc.NumAvPg) AS Cuenta, PlanPago.NumCuotasPP, PlanPago.SeqPP, PlanPago.Identidad, PlanPago.FechaInicioPP, "
Str = Str & " PlanPago.ValorCuotaPP , PlanPago.TotalPagadoPP, PlanPago.MontoPP, PlanPago.EstadoPP FROM PlanPago INNER JOIN "
Str = Str & " PlanPagoFactura ON PlanPago.SeqPP = PlanPagoFactura.SeqPP INNER JOIN AvPgEnc ON PlanPagoFactura.NumAvPg = AvPgEnc.NumAvPg "
Str = Str & " Where (AvPgEnc.AvPgEstado = 1) GROUP BY PlanPago.NumCuotasPP, PlanPago.SeqPP, PlanPago.Identidad, PlanPago.FechaInicioPP, PlanPago.ValorCuotaPP, PlanPago.TotalPagadoPP, "
Str = Str & " PlanPago.MontoPP , PlanPago.EstadoPP "

Else

Str = "  SELECT COUNT(PlanPagoFactura.NumAvPg) AS Cuenta, PlanPago.NumCuotasPP, PlanPago.SeqPP, PlanPago.Identidad, PlanPago.FechaInicioPP, "
Str = Str & " PlanPago.ValorCuotaPP, SUM(ReciboDet.ValorUnitReciboDet) AS TotalPagadoPP, PlanPago.MontoPP, PlanPago.EstadoPP "
Str = Str & " FROM Recibo INNER JOIN ReciboDet ON Recibo.NumRecibo = ReciboDet.NumRecibo INNER JOIN "
Str = Str & " PlanPago INNER JOIN PlanPagoFactura ON PlanPago.SeqPP = PlanPagoFactura.SeqPP ON ReciboDet.NumFactura = PlanPagoFactura.NumAvPg "
Str = Str & " WHERE (Recibo.FechaRecibo BETWEEN '" & VFecPP1 & "' AND '" & VFecPP2 & "') AND (Recibo.ReciboAnulado = 0)"
Str = Str & " GROUP BY PlanPago.NumCuotasPP, PlanPago.SeqPP, PlanPago.Identidad, PlanPago.FechaInicioPP, PlanPago.ValorCuotaPP, PlanPago.TotalPagadoPP, PlanPago.MontoPP , PlanPago.EstadoPP "

End If
'Set rsCont = DeRia.CoRia.Execute("SELECT * from PlanPago where EstadoPP = 0")
 Set rsCont = DeRia.CoRia.Execute(Str)
 
 If rsCont.EOF Then
 MsgBox "Ningun Plan de pago encontrado"
 End If
End Sub

Private Sub Detail_Format()
Dim RsDeclara As New ADODB.Recordset
Dim Str As String
Dim VTotal As Currency
Dim VtotalPo As Currency

Dim VFec As Date
Dim Vfec2 As Date

Static X As Integer
X = X + 1

If X > rsCont.RecordCount Then Exit Sub
TxtId.Text = rsCont!Identidad
TxtNo.Text = rsCont!SeqPP

Set RsDeclara = DeRia.CoRia.Execute("Select PNombre, PApellido from COntribuyente where  Identidad = '" & rsCont!Identidad & "'")

If Not RsDeclara.EOF Then
TxtNegocio.Text = Trim(RsDeclara!Pnombre) & " " & Trim(RsDeclara!PApellido)
Else
TxtNegocio.Text = "---------- No Existe Verifique ---------"

End If

Set RsDeclara = DeRia.CoRia.Execute("Select PNombre, PApellido from COntribuyente where IdRepresentante = '" & rsCont!Identidad & "'")
If Not RsDeclara.EOF Then

TxtPropietario.Text = Trim(RsDeclara!Pnombre) & " " & Trim(RsDeclara!PApellido)
Else
TxtPropietario.Text = ""
End If
   txtTotal.Text = Format(rsCont!MontoPP, "#,###,##0.00")
   VtotalToT = VtotalToT + rsCont!MontoPP

   If VxFechaPP = 0 Then

   Str = "SELECT SUM(AvPgDetalle.ValorUnitAvPgDet) AS Total FROM PlanPagoFactura INNER JOIN PlanPago ON PlanPagoFactura.SeqPP = PlanPago.SeqPP INNER JOIN AvPgEnc ON "
   Str = Str & "PlanPagoFactura.NumAvPg = AvPgEnc.NumAvPg INNER JOIN AvPgDetalle ON AvPgEnc.NumAvPg = AvPgDetalle.NumAvPg Where (PlanPago.SeqPP = " & rsCont!SeqPP & ") And (AvPgEnc.AvPgEstado = 2) "
   Else
   
   Str = "  SELECT SUM(ReciboDet.ValorUnitReciboDet) AS Total FROM PlanPagoFactura INNER JOIN "
   Str = Str & " PlanPago ON PlanPagoFactura.SeqPP = PlanPago.SeqPP INNER JOIN AvPgEnc ON PlanPagoFactura.NumAvPg = AvPgEnc.NumAvPg INNER JOIN "
   Str = Str & " ReciboDet ON AvPgEnc.NumAvPg = ReciboDet.NumFactura INNER JOIN Recibo ON ReciboDet.NumRecibo = Recibo.NumRecibo "
   Str = Str & " WHERE (PlanPago.SeqPP = " & rsCont!SeqPP & ") AND (Recibo.FechaRecibo BETWEEN '" & VFecPP1 & "' AND '" & VFecPP2 & "') AND (Recibo.ReciboAnulado = 0) "
  
   End If
  Set RsDeclara = DeRia.CoRia.Execute(Str)
   
   TxtPagado.Text = Format(IIf(IsNull(RsDeclara!Total), 0, RsDeclara!Total), "#,###,##0.00")
   VTotalPag = VTotalPag + IIf(IsNull(RsDeclara!Total), 0, RsDeclara!Total)

   If rsCont!MontoPP - IIf(IsNull(RsDeclara!Total), 0, RsDeclara!Total) < 1 Then
      TxtPendiente.Text = "0.00"
   Else
      TxtPendiente.Text = Format(rsCont!MontoPP - IIf(IsNull(RsDeclara!Total), 0, RsDeclara!Total), "#,###,##0.00")
      VTotalPen = VTotalPen + rsCont!MontoPP - IIf(IsNull(RsDeclara!Total), 0, RsDeclara!Total)
   End If

rsCont.MoveNext
Detail.PrintSection
End Sub

Private Sub ReportFooter_Format()
TxtTt.Text = Format(VtotalToT, "#,###,##0.00")
TxtTPa.Text = Format(VTotalPag, "#,###,##0.00")
TxtPp.Text = Format(VTotalPen, "#,###,##0.00")
End Sub
