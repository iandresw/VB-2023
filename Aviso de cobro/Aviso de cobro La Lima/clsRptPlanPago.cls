VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRptPlanPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private rReportRs As New ADODB.Recordset
Dim RsParam As New ADODB.Recordset
Dim rsCont As New ADODB.Recordset
Dim rsAldea As New ADODB.Recordset
Dim VtotalToT, VTotalPag, VTotalPen As Currency

Public Sub CrearReporte()
    Dim Str As String
    Dim Vval As Currency
    Dim fecha1 As String
    Dim fecha2 As String
    VtotalToT = 0
    VTotalPag = 0
    VTotalPen = 0
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open ("Select * from ParametroCOnt")
    Set rsCont = DeRia.CoRia.Execute("Update PlanPago set EstadoPP = 0 where EstadoPP is null ")
    Set rsCont = DeRia.CoRia.Execute("Update PlanPago set TotalPagadoPP = 0 where TotalPagadoPP is null ")
    
    'llama las facturas pagadas junto con el monto del plan de pago y verifica cuanto hace falta del pago si esmenor que uno le asiga estado 1 pagado
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

    'actualiza a estado cero el estado del plan de pago a 1 cuando el numero de facturas pagadas corresponde al numero de cuotas
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
    'verifica que si esl plan de pago esta en estado cero (vigente) y tiene facturas anuladas le pone estado 1 que es pagado
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
    
    
    
    
    
    If FrmReportePP.optEntreFecha.Value = True Then
        fecha1 = FormatearFecha(FrmReportePP.txtFecha1)
        fecha2 = FormatearFecha(FrmReportePP.txtFecha2)
        Str = "  SELECT COUNT(PlanPagoFactura.NumAvPg) AS Cuenta, PlanPago.NumCuotasPP, PlanPago.SeqPP, PlanPago.Identidad, PlanPago.FechaInicioPP, "
        Str = Str & " PlanPago.ValorCuotaPP, SUM(ReciboDet.ValorUnitReciboDet) AS TotalPagadoPP, PlanPago.MontoPP, PlanPago.EstadoPP "
        Str = Str & " FROM Recibo INNER JOIN ReciboDet ON Recibo.NumRecibo = ReciboDet.NumRecibo INNER JOIN "
        Str = Str & " PlanPago INNER JOIN PlanPagoFactura ON PlanPago.SeqPP = PlanPagoFactura.SeqPP ON ReciboDet.NumFactura = PlanPagoFactura.NumAvPg "
        Str = Str & " WHERE (Recibo.FechaRecibo BETWEEN '" & fecha1 & "' AND '" & fecha2 & "') AND (Recibo.ReciboAnulado = 0)"
        Str = Str & " GROUP BY PlanPago.NumCuotasPP, PlanPago.SeqPP, PlanPago.Identidad, PlanPago.FechaInicioPP, PlanPago.ValorCuotaPP, PlanPago.TotalPagadoPP, PlanPago.MontoPP , PlanPago.EstadoPP "
        
    ElseIf FrmReportePP.optTodos.Value = True Then
        Str = " SELECT COUNT(AvPgEnc.NumAvPg) AS Cuenta, PlanPago.NumCuotasPP, PlanPago.SeqPP, PlanPago.Identidad, PlanPago.FechaInicioPP, "
        Str = Str & " PlanPago.ValorCuotaPP , PlanPago.TotalPagadoPP, PlanPago.MontoPP, PlanPago.EstadoPP FROM PlanPago INNER JOIN "
        Str = Str & " PlanPagoFactura ON PlanPago.SeqPP = PlanPagoFactura.SeqPP INNER JOIN AvPgEnc ON PlanPagoFactura.NumAvPg = AvPgEnc.NumAvPg "
        Str = Str & " Where (AvPgEnc.AvPgEstado = 1) GROUP BY PlanPago.NumCuotasPP, PlanPago.SeqPP, PlanPago.Identidad, PlanPago.FechaInicioPP, PlanPago.ValorCuotaPP, PlanPago.TotalPagadoPP, "
        Str = Str & " PlanPago.MontoPP , PlanPago.EstadoPP "
        
    ElseIf FrmReportePP.optMeseVencidos.Value = True Then
        Str = " SELECT COUNT(AvPgEnc.NumAvPg) AS Cuenta, PlanPago.NumCuotasPP, PlanPago.SeqPP, PlanPago.Identidad, PlanPago.FechaInicioPP, "
        Str = Str & " PlanPago.ValorCuotaPP , PlanPago.TotalPagadoPP, PlanPago.MontoPP, PlanPago.EstadoPP FROM PlanPago INNER JOIN "
        Str = Str & " PlanPagoFactura ON PlanPago.SeqPP = PlanPagoFactura.SeqPP INNER JOIN AvPgEnc ON PlanPagoFactura.NumAvPg = AvPgEnc.NumAvPg "
        Str = Str & " Where (AvPgEnc.AvPgEstado = 1) AND (AvPgTipoImpuesto = 7) AND (DATEDIFF(month, FechaVenceAvPg, GETDATE()) >= " & FrmReportePP.txtMeses & ")"
        Str = Str & " GROUP BY PlanPago.NumCuotasPP, PlanPago.SeqPP, PlanPago.Identidad, PlanPago.FechaInicioPP, PlanPago.ValorCuotaPP, PlanPago.TotalPagadoPP, "
        Str = Str & " PlanPago.MontoPP , PlanPago.EstadoPP "
    End If
    MsgBox (Str)
    Set rsCont = DeRia.CoRia.Execute(Str)
 
    If rsCont.EOF Then
        MsgBox "Ningun Plan de pago encontrado"
    End If
End Sub

Private Sub CrearRecordSet()
    rReportRs.Fields.Append "Total", adBSTR
    rReportRs.Fields.Append "MontoPP", adBSTR
    'rReportRs.Fields.Append "Nombre", adBSTR
    'rReportRs.Fields.Append "Negocio", adBSTR
    'rReportRs.Fields.Append "ValorTotal", adCurrency
    'rReportRs.Fields.Append "ValorPagado", adCurrency
    'rReportRs.Fields.Append "SaldoPendiente", adCurrency
    rReportRs.Open
End Sub


Public Sub SendToExcel()
    If rsCont.RecordCount = 0 Then MsgBox ("No Hay datos para mostrar")
    Dim RsDeclara As New ADODB.Recordset
    Dim Str As String
    Dim VTotal As Currency
    Dim VtotalPo As Currency
    Dim VFec As Date
    Dim Vfec2 As Date
    Dim oXLApp As Excel.Application         'Declare the object variables
    Dim oXLBook As Excel.Workbook
    Dim oXLSheet As Excel.Worksheet
    Dim nRow, nCol As Long
    Dim oAlcaldia As New Alcaldia
    Dim exRng As Excel.Range
    Dim exRngData As Excel.Range
    
    Set oXLApp = New Excel.Application    'Create a new instance of Excel
    Set oXLBook = oXLApp.Workbooks.Add    'Add a new workbook
    Set oXLSheet = oXLBook.Worksheets(1)  'Work with the first
    
    oXLSheet.Cells(1, 1) = "MUNICIPIO: " & oAlcaldia.Nombre
    oXLSheet.Cells(2, 1) = "REPORTE DE PLANES DE PAGO"
    oXLSheet.Cells(6, 1) = "Departamento: " & oAlcaldia.Departamento
    If FrmReportePP.optEntreFecha.Value = True Then
        oXLSheet.Cells(7, 1) = "De: " & FrmReportePP.txtFecha1.Text & " Hasta: " & FrmReportePP.txtFecha2.Text
    End If
    oXLSheet.Cells(8, 1) = "Fecha Elaboración: " & DateToWord(Date)
    
    'Center Titles
    Set exRng = oXLSheet.Range("A1:G1"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A2:G2"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A3:G3"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A4:G4"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
   
    
    nRow = 12: nCol = 1
    'Pongamos los titulos de las columnas aqui, para que las agarre el autofit siguiente.
    oXLSheet.Cells(nRow, nCol) = "Identidad"                                   'A
    oXLSheet.Cells(nRow, nCol + 1) = "No."                                     'B
    oXLSheet.Cells(nRow, nCol + 2) = "Nombre Contribuyente"                    'C
    oXLSheet.Cells(nRow, nCol + 3) = "Nombre del Establecimiento Comercial"    'D
    oXLSheet.Cells(nRow, nCol + 4) = "Valor Total"                             'E
    oXLSheet.Cells(nRow, nCol + 5) = "Valor Pagado"                            'F
    oXLSheet.Cells(nRow, nCol + 6) = "Saldo Pendiente"                         'G
    
    '------------------------------------------------------------------------------------
  
    'Format titles Bold
    oXLSheet.Range("A1:G12").Font.Bold = True 'Los titulos del reporte
    
    'set a standard width
    Set exRng = oXLSheet.Columns("A")
    exRng.ColumnWidth = 15
    Set exRng = oXLSheet.Columns("B")
    exRng.ColumnWidth = 5
    Set exRng = oXLSheet.Columns("C")
    exRng.ColumnWidth = 50
    Set exRng = oXLSheet.Columns("D")
    exRng.ColumnWidth = 50
    Set exRng = oXLSheet.Columns("E")
    exRng.ColumnWidth = 11
    Set exRng = oXLSheet.Columns("F")
    exRng.ColumnWidth = 12
    Set exRng = oXLSheet.Columns("G")
    exRng.ColumnWidth = 15
    
    
    'Dibuje lineas en los titulos
    Set exRng = oXLSheet.Range("A12:A12")
    ApplyOutsideBorderToExcelCell exRng
    Set exRng = oXLSheet.Range("B12:B12")
    ApplyOutsideBorderToExcelCell exRng
    Set exRng = oXLSheet.Range("C12:C12")
    ApplyOutsideBorderToExcelCell exRng


    
    'Centrar los titulos del cuerpo
    Set exRng = oXLSheet.Range("A9:G12")
    exRng.HorizontalAlignment = xlCenter
    
    'Center between columns
    Set exRng = oXLSheet.Range("A6:G6"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
   ' Set exRng = oXLSheet.Range("Q10:R11"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
   ' Set exRng = oXLSheet.Range("S10:T11"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    
    'Define Starting Cells
    nRow = 13: nCol = 1:
    'Potencial Recaudación
    
    With rsCont
    .MoveFirst
    Do While .EOF = False
        oXLSheet.Cells(nRow, nCol) = rsCont!Identidad
        oXLSheet.Cells(nRow, nCol + 1) = rsCont!SeqPP
        
        Set RsDeclara = DeRia.CoRia.Execute("Select PNombre, PApellido from COntribuyente where  Identidad = '" & rsCont!Identidad & "'")
        If Not RsDeclara.EOF Then
            oXLSheet.Cells(nRow, nCol + 2) = Trim(RsDeclara!Pnombre) & " " & Trim(RsDeclara!PApellido)
        Else
            oXLSheet.Cells(nRow, nCol + 2) = "---------- No Existe Verifique ---------"
        End If
        
        Set RsDeclara = DeRia.CoRia.Execute("Select PNombre, PApellido from COntribuyente where IdRepresentante = '" & rsCont!Identidad & "'")
        If Not RsDeclara.EOF Then
            oXLSheet.Cells(nRow, nCol + 3) = Trim(RsDeclara!Pnombre) & " " & Trim(RsDeclara!PApellido)
        Else
            oXLSheet.Cells(nRow, nCol + 3) = ""
        End If
        
        oXLSheet.Cells(nRow, nCol + 4) = Format(rsCont!MontoPP, "#,###,##0.00")
        VtotalToT = VtotalToT + rsCont!MontoPP
        
        If dFecha1 = 0 Then
            Str = "SELECT SUM(AvPgDetalle.ValorUnitAvPgDet) AS Total FROM PlanPagoFactura INNER JOIN PlanPago ON PlanPagoFactura.SeqPP = PlanPago.SeqPP INNER JOIN AvPgEnc ON "
            Str = Str & "PlanPagoFactura.NumAvPg = AvPgEnc.NumAvPg INNER JOIN AvPgDetalle ON AvPgEnc.NumAvPg = AvPgDetalle.NumAvPg Where (PlanPago.SeqPP = " & rsCont!SeqPP & ") And (AvPgEnc.AvPgEstado = 2) "
        Else
            Str = "  SELECT SUM(ReciboDet.ValorUnitReciboDet) AS Total FROM PlanPagoFactura INNER JOIN "
            Str = Str & " PlanPago ON PlanPagoFactura.SeqPP = PlanPago.SeqPP INNER JOIN AvPgEnc ON PlanPagoFactura.NumAvPg = AvPgEnc.NumAvPg INNER JOIN "
            Str = Str & " ReciboDet ON AvPgEnc.NumAvPg = ReciboDet.NumFactura INNER JOIN Recibo ON ReciboDet.NumRecibo = Recibo.NumRecibo "
            Str = Str & " WHERE (PlanPago.SeqPP = " & rsCont!SeqPP & ") AND (Recibo.FechaRecibo BETWEEN '" & dFecha1 & "' AND '" & dFecha1 & "') AND (Recibo.ReciboAnulado = 0) "
        End If
            Set RsDeclara = DeRia.CoRia.Execute(Str)
            oXLSheet.Cells(nRow, nCol + 5) = Format(IIf(IsNull(RsDeclara!Total), 0, RsDeclara!Total), "#,###,##0.00")
            VTotalPag = VTotalPag + IIf(IsNull(RsDeclara!Total), 0, RsDeclara!Total)

        If rsCont!MontoPP - IIf(IsNull(RsDeclara!Total), 0, RsDeclara!Total) < 1 Then
           oXLSheet.Cells(nRow, nCol + 6) = "0.00"
        Else
           oXLSheet.Cells(nRow, nCol + 6) = Format(rsCont!MontoPP - IIf(IsNull(RsDeclara!Total), 0, RsDeclara!Total), "#,###,##0.00")
           VTotalPen = VTotalPen + rsCont!MontoPP - IIf(IsNull(RsDeclara!Total), 0, RsDeclara!Total)
        End If
        rsCont.MoveNext
        nRow = nRow + 1
     Loop
 
    End With
    
    
    'Calcule los totales
    nCol = 1
    oXLSheet.Cells(nRow + 1, nCol + 3) = "T o t a l"
    oXLSheet.Cells(nRow + 1, nCol + 4) = "=sum(E13:E" & nRow - 1 & ")"
    oXLSheet.Cells(nRow + 1, nCol + 5) = "=sum(F13:F" & nRow - 1 & ")"
    oXLSheet.Cells(nRow + 1, nCol + 6) = "=sum(G13:G" & nRow - 1 & ")"

     
     oXLSheet.Range("A" & nRow + 1 & ":" & "G" & nRow + 1).Font.Bold = True  'NEGRITA
     'Set exRng = oXLSheet.Range("A" & nRow + 1 & ":" & "G" & nRow + 1): exRng.HorizontalAlignment = xlCenter: exRng.Merge ' MARGEN
     
    nRow = nRow + 1
    'Imprime el periodo del reporte
    nRow = nRow + 1
    'oXLSheet.Cells(nRow, nCol) = strPeriodo
    Set exRng = oXLSheet.Range("A" & nRow, "G" & nRow): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    '---------------------------------------------------

    'Formateamos numeros
    Set exRng = oXLSheet.Columns("E:G")
    exRng.NumberFormat = "#,###,###,##0.00"
    'Formateamos texto
    'Set exRng = oXLSheet.Columns("M:N")
    'Set exRng = oXLSheet.Range("E11:P11")
    'exRng.NumberFormat = "@"
    
    'Autofit Columns
    'Set exRng = oXLSheet.Range("A11:B12")
    'exRng.EntireColumn.AutoFit
        
    'Dibuje lineas en los titulos0
    'Set exRng = oXLSheet.Range("A10:A11")
    'ApplyOutsideBorderToExcelCell exRng
    
    'Poner lineas en los datos
    Set exRng = oXLSheet.Range("A12:G" & nRow)
    ApplyBorderToExcelCell exRng
    


    'Cambiar el fondo de el titulo
    Set exRng = oXLSheet.Range("A12:G12")
    ApplyBackColorToExcelCell exRng
    
    '---------------------------------------------------
    oXLApp.Visible = True
    '---------------------------------------------------
    
    'Disconnect from Excel (let the user take over)
    Set oXLApp = Nothing
    '----------------------------------------------

End Sub



Private Function FormatearFecha(fecha As Date) As String
    Dim fechaFormateada As String
    fechaFormateada = Format(fecha, "yyyyMMdd")
    FormatearFecha = fechaFormateada
End Function

