VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRptJuridicosDeclara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private rReportRs As New ADODB.Recordset
Private rsAbonadoSP As New ADODB.Recordset
Private rscmdContrib As New ADODB.Recordset
Public Sub CrearReporte(anio As String)
    Dim consultaSQL As String
    Dim Id As String
   
    
    crearRecorset
    consultaSQL = " SELECT Contribuyente.Identidad, Contribuyente.Pnombre, Contribuyente.IdRepresentante, c.Pnombre AS rPNombre, c.SNombre AS rSNombre, c.PApellido AS rPApellido, c.SApellido AS rSapellido, CuentaIngreso_A.NombreCtaIngreso, "
    consultaSQL = consultaSQL & " DeclaraContJurid.ProdNoReg, DeclaraContJurid.ProdReg, DeclaraContJurid.ImpVolVenta, SUM(AvPgDetalle.ValorUnitAvPgDet) AS TotalMe, DeclaraContJurid.Periodo "
    consultaSQL = consultaSQL & " FROM Contribuyente INNER JOIN AvPgEnc ON Contribuyente.Identidad = AvPgEnc.Identidad INNER JOIN DeclaraContJurid ON Contribuyente.Identidad = DeclaraContJurid.Identidad INNER JOIN "
    consultaSQL = consultaSQL & " Contribuyente AS c ON Contribuyente.IdRepresentante = c.Identidad INNER JOIN CuentaIngreso_A ON Contribuyente.CodProfesion = CuentaIngreso_A.CtaIngreso INNER JOIN AvPgDetalle ON AvPgEnc.NumAvPg = AvPgDetalle.NumAvPg "
    consultaSQL = consultaSQL & " WHERE(Contribuyente.Tipo = 'TRUE') AND (AvPgEnc.AvPgTipoImpuesto = 2) AND (DeclaraContJurid.Periodo = " & anio & " ) AND (CuentaIngreso_A.Anio = " & anio & ") AND (AvPgEnc.AvPgEstado = 2) AND (AvPgEnc.AvPgTipoImpuesto = 2) AND "
    consultaSQL = consultaSQL & " (DATEPART(YEAR, AvPgEnc.FechaVenceAvPg) = '" & anio & "')  AND (DeclaraContJurid.EstadoDeclaraIC = 0)"
    consultaSQL = consultaSQL & " GROUP BY Contribuyente.Identidad, Contribuyente.Pnombre, Contribuyente.IdRepresentante, c.Pnombre, c.SNombre, c.PApellido, c.SApellido, CuentaIngreso_A.NombreCtaIngreso, DeclaraContJurid.ProdNoReg, DeclaraContJurid.ProdReg,  "
    consultaSQL = consultaSQL & " DeclaraContJurid.ImpVolVenta, DeclaraContJurid.Periodo ORDER BY Contribuyente.Identidad "
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open (consultaSQL)
    Do While Not DeRia.rsAbonadoSP.EOF
    rReportRs.AddNew
    rReportRs!identidad = DeRia.rsAbonadoSP!identidad
    Id = DeRia.rsAbonadoSP!identidad
    rReportRs!Pnombre = DeRia.rsAbonadoSP!Pnombre
    rReportRs!IdRepresentante = DeRia.rsAbonadoSP!IdRepresentante
    rReportRs!NombreContribuyente = Trim(DeRia.rsAbonadoSP!rPNombre) & " " & Trim(DeRia.rsAbonadoSP!rSNombre) & " " & Trim(DeRia.rsAbonadoSP!rPApellido) & " " & Trim(DeRia.rsAbonadoSP!rsApellido)
    rReportRs!rubro = DeRia.rsAbonadoSP!NombreCtaIngreso
    rReportRs!valorDeclarado = CDbl(DeRia.rsAbonadoSP!ProdNoReg) + CDbl(DeRia.rsAbonadoSP!ProdReg)
    rReportRs!ImpVolVenta = Impuestos(Id, DeRia.rsAbonadoSP!periodo)
    rReportRs!PerOP = PerOP(Id, DeRia.rsAbonadoSP!periodo)
    rReportRs!Tasas = Tasas(Id, DeRia.rsAbonadoSP!periodo)
    rReportRs!Servicios = Servicios(Id, DeRia.rsAbonadoSP!periodo)
    rReportRs!Descuentos = Descuentos(Id, DeRia.rsAbonadoSP!periodo)
    rReportRs!Intereses = Intereses(DeRia.rsAbonadoSP!identidad, anio)
    rReportRs!Recargos = Recargos(Id, DeRia.rsAbonadoSP!periodo)
   ' rReportRs!Total = Impuestos(Id, DeRia.rsAbonadoSP!periodo) + Intereses(DeRia.rsAbonadoSP!identidad, anio) + Recargos(Id, DeRia.rsAbonadoSP!periodo) + Tasas(Id, DeRia.rsAbonadoSP!periodo) + Descuentos(Id, DeRia.rsAbonadoSP!periodo) + Servicios(Id, DeRia.rsAbonadoSP!periodo)
    DeRia.rsAbonadoSP.MoveNext
    Loop
End Sub

Private Sub crearRecorset()
    rReportRs.Fields.Append "identidad", adBSTR
    rReportRs.Fields.Append "Pnombre", adBSTR
    rReportRs.Fields.Append "IdRepresentante", adBSTR
    rReportRs.Fields.Append "NombreContribuyente", adBSTR
    rReportRs.Fields.Append "Rubro", adBSTR
    rReportRs.Fields.Append "valorDeclarado", adDouble
    rReportRs.Fields.Append "ImpVolVenta", adDouble
    rReportRs.Fields.Append "PerOP", adDouble
    rReportRs.Fields.Append "Tasas", adDouble
    rReportRs.Fields.Append "Servicios", adDouble
    rReportRs.Fields.Append "Descuentos", adDouble
    rReportRs.Fields.Append "Intereses", adDouble
    rReportRs.Fields.Append "Recargos", adDouble
    rReportRs.Fields.Append "Total", adDouble
    rReportRs.Open
End Sub

Public Function Intereses(ByVal identidad As String, ByVal periodo As String) As Double
    Dim dato As Double
    strCodigo = "SELECT Sum(AvPgDetalle.ValorUnitAvPgDet) as Valor FROM AvPgEnc INNER JOIN  AvPgDetalle ON AvPgEnc.NumAvPg = AvPgDetalle.NumAvPg"
    strCodigo = strCodigo & " WHERE AvPgEnc.Identidad = '" & identidad & "' and (AvPgEnc.AvPgEstado = 2) AND (AvPgEnc.AvPgTipoImpuesto = 2) AND (DATEPART(YEAR, AvPgEnc.FechaVenceAvPg) = '" & periodo & "') And  AvPgDetalle.ctaIngreso = '11212601'"
    If DeRia.rscmdContrib.State = 1 Then DeRia.rscmdContrib.Close
    DeRia.rscmdContrib.Open strCodigo
    dato = IIf(IsNull(DeRia.rscmdContrib!valor), 0#, DeRia.rscmdContrib!valor)
    Intereses = dato
End Function

Public Function Recargos(ByVal identidad As String, ByVal periodo As String) As Double
    Dim dato As Double
    strCodigo = "SELECT Sum(AvPgDetalle.ValorUnitAvPgDet) as Valor FROM AvPgEnc INNER JOIN  AvPgDetalle ON AvPgEnc.NumAvPg = AvPgDetalle.NumAvPg"
    strCodigo = strCodigo & " WHERE AvPgEnc.Identidad = '" & identidad & "' and (AvPgEnc.AvPgEstado = 2) AND (AvPgEnc.AvPgTipoImpuesto = 2) AND (DATEPART(YEAR, AvPgEnc.FechaVenceAvPg) = '" & periodo & "') And  AvPgDetalle.ctaIngreso = '11212101'"
    If DeRia.rscmdContrib.State = 1 Then DeRia.rscmdContrib.Close
    DeRia.rscmdContrib.Open strCodigo
    dato = IIf(IsNull(DeRia.rscmdContrib!valor), 0#, DeRia.rscmdContrib!valor)
    Recargos = dato
End Function
Public Function Tasas(ByVal identidad As String, ByVal periodo As String) As Double
    Dim dato As Double
    strCodigo = "SELECT Sum(AvPgDetalle.ValorUnitAvPgDet) as Valor FROM AvPgEnc INNER JOIN  AvPgDetalle ON AvPgEnc.NumAvPg = AvPgDetalle.NumAvPg"
    strCodigo = strCodigo & " WHERE AvPgEnc.Identidad = '" & identidad & "' and (AvPgEnc.AvPgEstado = 2) AND (AvPgEnc.AvPgTipoImpuesto = 2) AND (DATEPART(YEAR, AvPgEnc.FechaVenceAvPg) = '" & periodo & "') And  (SUBSTRING(AvPgDetalle.CtaIngreso, 4, 3) IN ('119')) "
    If DeRia.rscmdContrib.State = 1 Then DeRia.rscmdContrib.Close
    DeRia.rscmdContrib.Open strCodigo
    dato = IIf(IsNull(DeRia.rscmdContrib!valor), 0#, DeRia.rscmdContrib!valor)
    Tasas = dato
End Function
Public Function Servicios(ByVal identidad As String, ByVal periodo As String) As Double
    Dim dato As Double
    strCodigo = "SELECT Sum(AvPgDetalle.ValorUnitAvPgDet) as Valor FROM AvPgEnc INNER JOIN  AvPgDetalle ON AvPgEnc.NumAvPg = AvPgDetalle.NumAvPg"
    strCodigo = strCodigo & " WHERE AvPgEnc.Identidad = '" & identidad & "' and (AvPgEnc.AvPgEstado = 2) AND (AvPgEnc.AvPgTipoImpuesto = 2) AND (DATEPART(YEAR, AvPgEnc.FechaVenceAvPg) = '" & periodo & "') And  (SUBSTRING(AvPgDetalle.CtaIngreso, 4, 3) IN ('118')) "
    If DeRia.rscmdContrib.State = 1 Then DeRia.rscmdContrib.Close
    DeRia.rscmdContrib.Open strCodigo
    dato = IIf(IsNull(DeRia.rscmdContrib!valor), 0#, DeRia.rscmdContrib!valor)
    Servicios = dato
End Function
Public Function PerOP(ByVal identidad As String, ByVal periodo As String) As Double
    Dim dato As Double
    strCodigo = "SELECT Sum(AvPgDetalle.ValorUnitAvPgDet) as Valor FROM AvPgEnc INNER JOIN  AvPgDetalle ON AvPgEnc.NumAvPg = AvPgDetalle.NumAvPg"
    strCodigo = strCodigo & " WHERE AvPgEnc.Identidad = '" & identidad & "' and (AvPgEnc.AvPgEstado = 2) AND (AvPgEnc.AvPgTipoImpuesto = 3) AND (DATEPART(YEAR, AvPgEnc.FechaVenceAvPg) = '" & periodo & "') And  (SUBSTRING(AvPgDetalle.CtaIngreso, 4, 3) IN ('119')) "
    If DeRia.rscmdContrib.State = 1 Then DeRia.rscmdContrib.Close
    DeRia.rscmdContrib.Open strCodigo
    dato = IIf(IsNull(DeRia.rscmdContrib!valor), 0#, DeRia.rscmdContrib!valor)
    PerOP = dato
End Function
Public Function Descuentos(ByVal identidad As String, ByVal periodo As String) As Double
    Dim dato As Double
    strCodigo = "SELECT Sum(AvPgDetalle.ValorUnitAvPgDet) as Valor FROM AvPgEnc INNER JOIN  AvPgDetalle ON AvPgEnc.NumAvPg = AvPgDetalle.NumAvPg"
    strCodigo = strCodigo & " WHERE AvPgEnc.Identidad = '" & identidad & "' and (AvPgEnc.AvPgEstado = 2) AND (AvPgEnc.AvPgTipoImpuesto = 2) AND (DATEPART(YEAR, AvPgEnc.FechaVenceAvPg) = '" & periodo & "') And  AvPgDetalle.ctaIngreso = '11212701'"
    If DeRia.rscmdContrib.State = 1 Then DeRia.rscmdContrib.Close
    DeRia.rscmdContrib.Open strCodigo
    dato = IIf(IsNull(DeRia.rscmdContrib!valor), 0#, DeRia.rscmdContrib!valor)
    Descuentos = dato
End Function
Public Function Impuestos(ByVal identidad As String, ByVal periodo As String) As Double
    Dim dato As Double
    strCodigo = "SELECT Sum(AvPgDetalle.ValorUnitAvPgDet) as Valor FROM AvPgEnc INNER JOIN  AvPgDetalle ON AvPgEnc.NumAvPg = AvPgDetalle.NumAvPg"
    strCodigo = strCodigo & " WHERE AvPgEnc.Identidad = '" & identidad & "' and (AvPgEnc.AvPgEstado = 2) AND (AvPgEnc.AvPgTipoImpuesto = 2) AND (DATEPART(YEAR, AvPgEnc.FechaVenceAvPg) = '" & periodo & "') AND (SUBSTRING(AvPgDetalle.CtaIngreso, 4, 3) IN ('112', '113', '114'))"
    If DeRia.rscmdContrib.State = 1 Then DeRia.rscmdContrib.Close
    DeRia.rscmdContrib.Open strCodigo
    dato = IIf(IsNull(DeRia.rscmdContrib!valor), 0#, DeRia.rscmdContrib!valor)
    Impuestos = dato
End Function


Public Sub SendToExcel()
    
    If rReportRs.RecordCount = 0 Then: MsgBox ("No Hay datos para mostrar"): Exit Sub
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
    Dim strCodigo As String
    
    oXLSheet.Cells(1, 1) = "MUNICIPIO: " & oAlcaldia.Nombre
    oXLSheet.Cells(2, 1) = "LISTA DE ESTABLECIMIENTOS COMERCIALES CON DECLARACION JURADA"
    oXLSheet.Cells(5, 1) = "Departamento: " & oAlcaldia.Departamento
    oXLSheet.Cells(6, 1) = "Municipio: " & oAlcaldia.Nombre
    'oXLSheet.Cells(7, 1) = "Fecha Elaboración: " & DateToWord(Date)
    oXLSheet.Cells(8, 1) = "Estado de Facturas: Pagadas "
    'Center Titles
    Set exRng = oXLSheet.Range("A1:K1"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A2:K2"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A3:K3"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A5:K5"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
    Set exRng = oXLSheet.Range("A6:K6"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
    Set exRng = oXLSheet.Range("A7:K7"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
    
  
   
   
    
    nRow = 12: nCol = 1
    'Pongamos los titulos de las columnas aqui, para que las agarre el autofit siguiente.
    oXLSheet.Cells(nRow, nCol) = "RTM"                               'A
    oXLSheet.Cells(nRow, nCol + 1) = "Nombre Establecimeiento"       'B
    oXLSheet.Cells(nRow, nCol + 2) = "Identidad Propietario"         'C
    oXLSheet.Cells(nRow, nCol + 3) = "Nombre Propietario"            'D
    oXLSheet.Cells(nRow, nCol + 4) = "Actividad Economica"           'E
    oXLSheet.Cells(nRow, nCol + 5) = "Valor Declarado"               'F
    oXLSheet.Cells(nRow, nCol + 6) = "Impuesto"                       'G
    oXLSheet.Cells(nRow, nCol + 7) = "Permiso de Operacion"           'H
    oXLSheet.Cells(nRow, nCol + 8) = "Tasas"                      'I
    oXLSheet.Cells(nRow, nCol + 9) = "Servicios"                    'J
    oXLSheet.Cells(nRow, nCol + 10) = "Descuentos"                       'K
    oXLSheet.Cells(nRow, nCol + 11) = "Intereses"                      'L
    oXLSheet.Cells(nRow, nCol + 12) = "Recargos"                     'M
    oXLSheet.Cells(nRow, nCol + 13) = "Total"                      'N
    '------------------------------------------------------------------------------------
  
    'Format titles Bold
    oXLSheet.Range("A1:N12").Font.Bold = True 'Los titulos del reporte
    'set a standard width
    Set exRng = oXLSheet.Columns("A")
    exRng.ColumnWidth = 15
    Set exRng = oXLSheet.Columns("B")
    exRng.ColumnWidth = 54
    Set exRng = oXLSheet.Columns("C")
    exRng.ColumnWidth = 24
    Set exRng = oXLSheet.Columns("D")
    exRng.ColumnWidth = 35
    Set exRng = oXLSheet.Columns("E")
    exRng.ColumnWidth = 54
    Set exRng = oXLSheet.Columns("F")
    exRng.ColumnWidth = 20
    Set exRng = oXLSheet.Columns("G")
    exRng.ColumnWidth = 15
    Set exRng = oXLSheet.Columns("H")
    exRng.ColumnWidth = 21
    Set exRng = oXLSheet.Columns("I")
    exRng.ColumnWidth = 15
    Set exRng = oXLSheet.Columns("J")
    exRng.ColumnWidth = 15
    Set exRng = oXLSheet.Columns("K")
    exRng.ColumnWidth = 15
    Set exRng = oXLSheet.Columns("L")
    exRng.ColumnWidth = 15
    Set exRng = oXLSheet.Columns("M")
    exRng.ColumnWidth = 15
    Set exRng = oXLSheet.Columns("N")
    exRng.ColumnWidth = 15
    
    
    
    'Dibuje lineas en los titulos
    Set exRng = oXLSheet.Range("A12:A12")
    ApplyOutsideBorderToExcelCell exRng
    Set exRng = oXLSheet.Range("B12:B12")
    ApplyOutsideBorderToExcelCell exRng
    Set exRng = oXLSheet.Range("C12:C12")
    ApplyOutsideBorderToExcelCell exRng
    Set exRng = oXLSheet.Range("D12:D12")
    ApplyOutsideBorderToExcelCell exRng
    Set exRng = oXLSheet.Range("E12:E12")
    ApplyOutsideBorderToExcelCell exRng

    'Centrar los titulos del cuerpo
    Set exRng = oXLSheet.Range("A12:N12")
    exRng.HorizontalAlignment = xlCenter
    

    
    'Define Starting Cells
    nRow = 13: nCol = 1:
    'Potencial Recaudación
    With rReportRs
        .MoveFirst
        Do While .EOF = False
            oXLSheet.Cells(nRow, nCol) = !identidad
            oXLSheet.Cells(nRow, nCol + 1) = !Pnombre
            oXLSheet.Cells(nRow, nCol + 2) = !IdRepresentante
            oXLSheet.Cells(nRow, nCol + 3) = !NombreContribuyente
            oXLSheet.Cells(nRow, nCol + 4) = !rubro
            oXLSheet.Cells(nRow, nCol + 5) = Format(!valorDeclarado, "###,###,####0.00")
            oXLSheet.Cells(nRow, nCol + 6) = Format(!ImpVolVenta, "###,###,####0.00")
            oXLSheet.Cells(nRow, nCol + 7) = Format(!PerOP, "###,###,####0.00")
            oXLSheet.Cells(nRow, nCol + 8) = Format(!Tasas, "###,###,####0.00")
            oXLSheet.Cells(nRow, nCol + 9) = Format(!Servicios, "###,###,####0.00")
            oXLSheet.Cells(nRow, nCol + 10) = Format(!Descuentos, "###,###,####0.00")
            oXLSheet.Cells(nRow, nCol + 11) = Format(!Intereses, "###,###,####0.00")
            oXLSheet.Cells(nRow, nCol + 12) = Format(!Recargos, "###,###,####0.00")
            oXLSheet.Cells(nRow, nCol + 13).Formula = "=sum(G" & nRow & ":M" & nRow & ")"
            oXLSheet.Cells(nRow, nCol + 13).NumberFormat = " #,###,###,##0.00"
            .MoveNext
            nRow = nRow + 1
        Loop
    End With
    
    'Calcule los totales
    nCol = 1
    oXLSheet.Cells(nRow + 1, nCol + 4) = "T o t a l e s"
    oXLSheet.Cells(nRow + 1, nCol + 5).Formula = "=sum(F10:F" & nRow - 1 & ")"
    oXLSheet.Cells(nRow + 1, nCol + 5).NumberFormat = "L #,###,###,##0.00"
    oXLSheet.Cells(nRow + 1, nCol + 6).Formula = "=sum(G10:G" & nRow - 1 & ")"
    oXLSheet.Cells(nRow + 1, nCol + 6).NumberFormat = "L #,###,###,##0.00"
    oXLSheet.Cells(nRow + 1, nCol + 7).Formula = "=sum(H10:H" & nRow - 1 & ")"
    oXLSheet.Cells(nRow + 1, nCol + 7).NumberFormat = "L #,###,###,##0.00"
    oXLSheet.Cells(nRow + 1, nCol + 8).Formula = "=sum(I10:I" & nRow - 1 & ")"
    oXLSheet.Cells(nRow + 1, nCol + 8).NumberFormat = "L #,###,###,##0.00"
    oXLSheet.Cells(nRow + 1, nCol + 9).Formula = "=sum(J10:J" & nRow - 1 & ")"
    oXLSheet.Cells(nRow + 1, nCol + 9).NumberFormat = "L #,###,###,##0.00"
    
    oXLSheet.Cells(nRow + 1, nCol + 10).Formula = "=sum(K10:K" & nRow - 1 & ")"
    oXLSheet.Cells(nRow + 1, nCol + 10).NumberFormat = "L #,###,###,##0.00"
    
    oXLSheet.Cells(nRow + 1, nCol + 11).Formula = "=sum(L10:L" & nRow - 1 & ")"
    oXLSheet.Cells(nRow + 1, nCol + 11).NumberFormat = "L #,###,###,##0.00"
    
    oXLSheet.Cells(nRow + 1, nCol + 12).Formula = "=sum(M10:M" & nRow - 1 & ")"
    oXLSheet.Cells(nRow + 1, nCol + 12).NumberFormat = "L #,###,###,##0.00"
    oXLSheet.Cells(nRow + 1, nCol + 13).Formula = "=sum(N10:N" & nRow - 1 & ")"
    oXLSheet.Cells(nRow + 1, nCol + 13).NumberFormat = "L #,###,###,##0.00"
    
    
    
    
    
    oXLSheet.Range("E" & nRow + 1 & ":" & "N" & nRow + 1).Font.Bold = True: ApplyBackColorToExcelCell exRng   'NEGRITA
    Set exRng = oXLSheet.Range("E" & nRow + 1 & ":" & "N" & nRow + 1):  ApplyBackColorToExcelCell exRng: oXLSheet.Range("B" & nRow & ":" & "I" & nRow).Font.Bold = True   ' MARGEN
    nRow = nRow + 1
    'Imprime el periodo del reporte
    nRow = nRow + 1
    'oXLSheet.Cells(nRow, nCol) = strPeriodo
    Set exRng = oXLSheet.Range("A" & nRow, "N" & nRow): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    
    'Poner lineas en los datos
    Set exRng = oXLSheet.Range("A12:N" & nRow)
    ApplyBorderToExcelCell exRng
    


    'Cambiar el fondo de el titulo
    Set exRng = oXLSheet.Range("A12:N12")
    ApplyBackColorToExcelCell exRng
    
    '---------------------------------------------------
    oXLApp.Visible = True
    '---------------------------------------------------
    
    'Disconnect from Excel (let the user take over)
    Set oXLApp = Nothing
    '----------------------------------------------

End Sub

