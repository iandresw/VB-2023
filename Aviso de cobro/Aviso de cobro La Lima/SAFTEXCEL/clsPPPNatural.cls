VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPPPNatural"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private rReportRs As New ADODB.Recordset
Private rsAbonadoSP As New ADODB.Recordset
Public Sub CrearReporte()
    Dim sql As String
    crearRecorset
    
    sql = " SELECT PlanPago.SeqPP, Contribuyente.Identidad, Contribuyente.Pnombre, Contribuyente_1.Identidad AS Expr1, Contribuyente_1.Pnombre AS Expr2, Contribuyente_1.SNombre, Contribuyente_1.PApellido, Contribuyente_1.SApellido,"
    sql = sql & " PlanPago.FechaInicioPP , PlanPago.NumCuotasPP, PlanPago.ValorCuotaPP, PlanPago.MontoPP"
    sql = sql & " FROM PlanPago INNER JOIN Contribuyente ON PlanPago.Identidad = Contribuyente.Identidad INNER JOIN"
    sql = sql & " Contribuyente AS Contribuyente_1 ON Contribuyente.IdRepresentante = Contribuyente_1.Identidad"
    sql = sql & " WHERE (Contribuyente.Tipo = 'FALSE') AND (PlanPago.EstadoPP = 0)"
    sql = sql & " GROUP BY PlanPago.SeqPP, Contribuyente.Identidad, Contribuyente.Pnombre, Contribuyente_1.Identidad, Contribuyente_1.Pnombre, Contribuyente_1.SNombre, Contribuyente_1.PApellido, Contribuyente_1.SApellido,"
    sql = sql & " PlanPago.FechaInicioPP , PlanPago.NumCuotasPP, PlanPago.ValorCuotaPP, PlanPago.MontoPP"
    sql = sql & " ORDER BY PlanPago.SeqPP"
    
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open (sql)
    
    Do While Not DeRia.rsAbonadoSP.EOF
        rReportRs.AddNew
        rReportRs!SeqPP = DeRia.rsAbonadoSP!SeqPP
        rReportRs!Identidad = DeRia.rsAbonadoSP!Identidad
        rReportRs!Pnombre = DeRia.rsAbonadoSP!Pnombre
        rReportRs!IdRepresentante = DeRia.rsAbonadoSP!Expr1
        rReportRs!NombrePropietario = Trim(DeRia.rsAbonadoSP!Expr2) & " " & Trim(DeRia.rsAbonadoSP!sNombre) & " " & Trim(DeRia.rsAbonadoSP!PApellido) & " " & Trim(DeRia.rsAbonadoSP!sApellido)
        rReportRs!FechaInicioPP = DeRia.rsAbonadoSP!FechaInicioPP
        rReportRs!NumCuotasPP = DeRia.rsAbonadoSP!NumCuotasPP
        rReportRs!ValorCuotaPP = DeRia.rsAbonadoSP!ValorCuotaPP
        rReportRs!MontoPP = DeRia.rsAbonadoSP!MontoPP
        DeRia.rsAbonadoSP.MoveNext
    Loop
End Sub

Private Sub crearRecorset()
    rReportRs.Fields.Append "SeqPP", adBSTR
    rReportRs.Fields.Append "Identidad", adBSTR
    rReportRs.Fields.Append "Pnombre", adBSTR
    rReportRs.Fields.Append "IdRepresentante", adBSTR
    rReportRs.Fields.Append "NombrePropietario", adBSTR
    rReportRs.Fields.Append "FechaInicioPP", adBSTR
    rReportRs.Fields.Append "NumCuotasPP", adBSTR
    rReportRs.Fields.Append "ValorCuotaPP", adBSTR
    rReportRs.Fields.Append "MontoPP", adBSTR
    rReportRs.Open
End Sub

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
    oXLSheet.Cells(2, 1) = "PLANES DE PAGO INDUSTRIA Y COMERCIO VIGENTES"
    oXLSheet.Cells(5, 1) = "Departamento: " & oAlcaldia.Departamento
    oXLSheet.Cells(6, 1) = "Municipio: " & oAlcaldia.Nombre
    oXLSheet.Cells(7, 1) = "Fecha Elaboración: " & DateToWord(Date)
    
    
    'Center Titles
    Set exRng = oXLSheet.Range("A1:I1"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A2:I2"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A3:I3"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A5:I5"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
    Set exRng = oXLSheet.Range("A6:I6"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
    Set exRng = oXLSheet.Range("A7:I7"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
 
    
    
    
    nRow = 9: nCol = 1
    'Pongamos los titulos de las columnas aqui, para que las agarre el autofit siguiente.
    oXLSheet.Cells(nRow, nCol) = "Numero PP"                         'A
    oXLSheet.Cells(nRow, nCol + 1) = "R.T.M."                        'B
    oXLSheet.Cells(nRow, nCol + 2) = "Nombre Establecimiento"        'C
    oXLSheet.Cells(nRow, nCol + 3) = "Identidad"                     'D
    oXLSheet.Cells(nRow, nCol + 4) = "Nombre Propietario"            'E
    oXLSheet.Cells(nRow, nCol + 5) = "Fecha de Inicio"               'F
    oXLSheet.Cells(nRow, nCol + 6) = "Num Cuotas"                    'G
    oXLSheet.Cells(nRow, nCol + 7) = "Valor Cuota PP"                'H
    oXLSheet.Cells(nRow, nCol + 8) = "Monto PP"                      'I
    '------------------------------------------------------------------------------------
  
    'Format titles Bold
    oXLSheet.Range("A1:I9").Font.Bold = True 'Los titulos del reporte
    'set a standard width
    Set exRng = oXLSheet.Columns("A")
    exRng.ColumnWidth = 11
    Set exRng = oXLSheet.Columns("B")
    exRng.ColumnWidth = 11
    Set exRng = oXLSheet.Columns("C")
    exRng.ColumnWidth = 60
    Set exRng = oXLSheet.Columns("D")
    exRng.ColumnWidth = 14
    Set exRng = oXLSheet.Columns("E")
    exRng.ColumnWidth = 40
    Set exRng = oXLSheet.Columns("F")
    exRng.ColumnWidth = 11
    Set exRng = oXLSheet.Columns("G")
    exRng.ColumnWidth = 11
    Set exRng = oXLSheet.Columns("H")
    exRng.ColumnWidth = 15
    Set exRng = oXLSheet.Columns("I")
    exRng.ColumnWidth = 15

    
    
    'Dibuje lineas en los titulos
    Set exRng = oXLSheet.Range("A9:A9")
    ApplyOutsideBorderToExcelCell exRng
    Set exRng = oXLSheet.Range("B9:B9")
    ApplyOutsideBorderToExcelCell exRng
    Set exRng = oXLSheet.Range("C9:C9")
    ApplyOutsideBorderToExcelCell exRng
    Set exRng = oXLSheet.Range("D12:D12")
    ApplyOutsideBorderToExcelCell exRng
    Set exRng = oXLSheet.Range("E12:E12")
    ApplyOutsideBorderToExcelCell exRng
    Set exRng = oXLSheet.Range("F12:F12")
    ApplyOutsideBorderToExcelCell exRng
    Set exRng = oXLSheet.Range("G12:G12")
    ApplyOutsideBorderToExcelCell exRng
    Set exRng = oXLSheet.Range("H12:H12")
    ApplyOutsideBorderToExcelCell exRng
    Set exRng = oXLSheet.Range("I12:I12")
    ApplyOutsideBorderToExcelCell exRng
    Set exRng = oXLSheet.Range("J12:J12")
    ApplyOutsideBorderToExcelCell exRng



    'Centrar los titulos del cuerpo
    Set exRng = oXLSheet.Range("A9:I9")
    exRng.HorizontalAlignment = xlCenter
        
    'Define Starting Cells
    nRow = 10: nCol = 1:
    'Potencial Recaudación
    With rReportRs
        .MoveFirst
        Do While .EOF = False
            oXLSheet.Cells(nRow, nCol) = !SeqPP
            oXLSheet.Cells(nRow, nCol + 1) = !Identidad
            oXLSheet.Cells(nRow, nCol + 2) = !Pnombre
            oXLSheet.Cells(nRow, nCol + 3) = !IdRepresentante
            oXLSheet.Cells(nRow, nCol + 4) = !NombrePropietario
            oXLSheet.Cells(nRow, nCol + 5) = !FechaInicioPP
            oXLSheet.Cells(nRow, nCol + 6) = !NumCuotasPP
            oXLSheet.Cells(nRow, nCol + 7) = Format(!ValorCuotaPP, "#,###,##0.00")
            oXLSheet.Cells(nRow, nCol + 8) = Format(!MontoPP, "#,###,##0.00")
            .MoveNext
            nRow = nRow + 1
        Loop
    End With
    oXLSheet.Range("A" & nRow + 1 & ":" & "I" & nRow + 1).Font.Bold = True  'NEGRITA
    Set exRng = oXLSheet.Range("A" & nRow + 1 & ":" & "B" & nRow + 1): exRng.HorizontalAlignment = xlCenter: exRng.Merge ' MARGEN
    nRow = nRow + 1
    'Imprime el periodo del reporte
    nRow = nRow + 1
    Set exRng = oXLSheet.Range("A" & nRow, "I" & nRow): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    
    'Poner lineas en los datos
    Set exRng = oXLSheet.Range("A9:I" & nRow)
    ApplyBorderToExcelCell exRng
    


    'Cambiar el fondo de el titulo
    Set exRng = oXLSheet.Range("A9:I9")
    ApplyBackColorToExcelCell exRng
    
    '---------------------------------------------------
    oXLApp.Visible = True
    '---------------------------------------------------
    
    'Disconnect from Excel (let the user take over)
    Set oXLApp = Nothing
    '----------------------------------------------

End Sub

