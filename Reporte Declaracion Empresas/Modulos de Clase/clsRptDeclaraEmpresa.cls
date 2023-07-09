VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRptDeclaraEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private rReportRs As New ADODB.Recordset
Private rsAbonadoSP As New ADODB.Recordset
Public Sub CrearReporte(empresa As String, anio As String)
    Dim codSQL As String
    crearRecorset
    
    codSQL = "  SELECT     DeclaraNat.Empresa, DeclaraNat.Identidad, Contribuyente.Pnombre, Contribuyente.SNombre, Contribuyente.PApellido, Contribuyente.SApellido, DeclaraNat.Ingreso, DeclaraNat.Impuesto,  DeclaraNat.Multa, "
    codSQL = codSQL & " DeclaraNat.Interes, DeclaraNat.Recargo, DeclaraNat.Descuento, DeclaraNat.Total, DeclaraNat.CodDeclaraIP, DeclaraNat.NumAvPg, DeclaraNat.NumRecibo, "
    codSQL = codSQL & " DeclaraNat.AnioDeclara, DeclaraNat.FechaPresenta "
    codSQL = codSQL & " FROM    DeclaraNat INNER JOIN Contribuyente ON DeclaraNat.Identidad = Contribuyente.Identidad "
    codSQL = codSQL & " WHERE (DeclaraNat.Empresa = '" & empresa & "') AND (DeclaraNat.AnioDeclara = '" & anio & "') "
    
    
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open (codSQL)
    
    Do While Not DeRia.rsAbonadoSP.EOF
    rReportRs.AddNew
    rReportRs!empresa = DeRia.rsAbonadoSP!empresa
    rReportRs!identidad = DeRia.rsAbonadoSP!identidad
    rReportRs!NombreCompleto = Trim(DeRia.rsAbonadoSP!PNombre) & " " & Trim(DeRia.rsAbonadoSP!SNombre) & " " & Trim(DeRia.rsAbonadoSP!PApellido) & " " & Trim(DeRia.rsAbonadoSP!SApellido)
    rReportRs!Ingreso = IIf(IsNull(DeRia.rsAbonadoSP!Ingreso), 0, DeRia.rsAbonadoSP!Ingreso)
    rReportRs!Impuesto = IIf(IsNull(DeRia.rsAbonadoSP!Impuesto), 0, DeRia.rsAbonadoSP!Impuesto)
    rReportRs!Multa = IIf(IsNull(DeRia.rsAbonadoSP!Multa), 0, DeRia.rsAbonadoSP!Multa)
    rReportRs!Interes = IIf(IsNull(DeRia.rsAbonadoSP!Interes), 0, DeRia.rsAbonadoSP!Interes)
    rReportRs!Recargo = IIf(IsNull(DeRia.rsAbonadoSP!Recargo), 0, DeRia.rsAbonadoSP!Recargo)
    rReportRs!Descuento = IIf(IsNull(DeRia.rsAbonadoSP!Descuento), 0, DeRia.rsAbonadoSP!Descuento)
    rReportRs!Total = IIf(IsNull(DeRia.rsAbonadoSP!Total), 0, DeRia.rsAbonadoSP!Total)
    rReportRs!CodDeclaraIP = DeRia.rsAbonadoSP!CodDeclaraIP
    rReportRs!NumAvPg = DeRia.rsAbonadoSP!NumAvPg
    rReportRs!NumRecibo = DeRia.rsAbonadoSP!NumRecibo
    rReportRs!AnioDeclara = DeRia.rsAbonadoSP!AnioDeclara
    rReportRs!FechaPresenta = DeRia.rsAbonadoSP!FechaPresenta
    DeRia.rsAbonadoSP.MoveNext
    Loop
End Sub

Private Sub crearRecorset()
    
    rReportRs.Fields.Append "Empresa", adBSTR
    rReportRs.Fields.Append "Identidad", adBSTR
    rReportRs.Fields.Append "NombreCompleto", adBSTR
    rReportRs.Fields.Append "Ingreso", adBSTR
    rReportRs.Fields.Append "Impuesto", adBSTR
    rReportRs.Fields.Append "Multa", adBSTR
    rReportRs.Fields.Append "Interes", adBSTR
    rReportRs.Fields.Append "Recargo", adBSTR
    rReportRs.Fields.Append "Descuento", adBSTR
    rReportRs.Fields.Append "Total", adBSTR
    rReportRs.Fields.Append "CodDeclaraIP", adBSTR
    rReportRs.Fields.Append "NumAvPg", adBSTR
    rReportRs.Fields.Append "NumRecibo", adBSTR
    rReportRs.Fields.Append "AnioDeclara", adBSTR
    rReportRs.Fields.Append "FechaPresenta", adBSTR
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
    oXLSheet.Cells(2, 1) = "DECLARACION DE IMPUESTO PERSONAL PARA EMPRESAS"
    oXLSheet.Cells(5, 1) = "Departamento: " & oAlcaldia.Departamento
    oXLSheet.Cells(6, 1) = "Municipio: " & oAlcaldia.Nombre
    oXLSheet.Cells(7, 1) = "Fecha Elaboración: " & Now()
    
    'Center Titles
    Set exRng = oXLSheet.Range("A1:O1"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A2:O2"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A3:O3"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A5:O5"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
    Set exRng = oXLSheet.Range("A6:O6"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
    Set exRng = oXLSheet.Range("A7:O7"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
    
     
    
    nRow = 9: nCol = 1
    'Pongamos los titulos de las columnas aqui, para que las agarre el autofit siguiente.
    oXLSheet.Cells(nRow, nCol) = "Empresa"                     'A
    oXLSheet.Cells(nRow, nCol + 1) = "Identidad"               'B
    oXLSheet.Cells(nRow, nCol + 2) = "Nombre Contribuyente"    'C
    oXLSheet.Cells(nRow, nCol + 3) = "Ingreso"                 'D
    oXLSheet.Cells(nRow, nCol + 4) = "Impuesto"                'E
    oXLSheet.Cells(nRow, nCol + 5) = "Multa"                   'F
    oXLSheet.Cells(nRow, nCol + 6) = "Interes"                 'G
    oXLSheet.Cells(nRow, nCol + 7) = "Recargo"                 'H
    oXLSheet.Cells(nRow, nCol + 8) = "Descuento"               'I
    oXLSheet.Cells(nRow, nCol + 9) = "Total"                   'J
    oXLSheet.Cells(nRow, nCol + 10) = "Declaracion"            'K
    oXLSheet.Cells(nRow, nCol + 11) = "Factura"                'L
    oXLSheet.Cells(nRow, nCol + 12) = "Recibo"                 'M
    oXLSheet.Cells(nRow, nCol + 13) = "Año Declaracion"        'N
    oXLSheet.Cells(nRow, nCol + 14) = "Fecha Presenta"         'O
                          
    
    
    '------------------------------------------------------------------------------------
  
    'Format titles Bold
    oXLSheet.Range("A1:P9").Font.Bold = True 'Los titulos del reporte
    'set a standard width
    Set exRng = oXLSheet.Columns("A")
    exRng.ColumnWidth = 45
    Set exRng = oXLSheet.Columns("B")
    exRng.ColumnWidth = 15
    Set exRng = oXLSheet.Columns("C")
    exRng.ColumnWidth = 45
    Set exRng = oXLSheet.Columns("D")
    exRng.ColumnWidth = 12
    Set exRng = oXLSheet.Columns("E")
    exRng.ColumnWidth = 12
    Set exRng = oXLSheet.Columns("F")
    exRng.ColumnWidth = 12
    Set exRng = oXLSheet.Columns("G")
    exRng.ColumnWidth = 12
    Set exRng = oXLSheet.Columns("H")
    exRng.ColumnWidth = 12
    Set exRng = oXLSheet.Columns("I")
    exRng.ColumnWidth = 12
    Set exRng = oXLSheet.Columns("J")
    exRng.ColumnWidth = 12
    Set exRng = oXLSheet.Columns("K")
    exRng.ColumnWidth = 12
    Set exRng = oXLSheet.Columns("L")
    exRng.ColumnWidth = 12
    Set exRng = oXLSheet.Columns("M")
    exRng.ColumnWidth = 12
    Set exRng = oXLSheet.Columns("N")
    exRng.ColumnWidth = 12
    Set exRng = oXLSheet.Columns("O")
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
    Set exRng = oXLSheet.Range("A9:P9")
    exRng.HorizontalAlignment = xlCenter

    'Format(!Ingreso, "#,###,##0.00")
    'Define Starting Cells
    nRow = 10: nCol = 1:
    'Potencial Recaudación
    With rReportRs
        .MoveFirst
        Do While .EOF = False
            oXLSheet.Cells(nRow, nCol) = !empresa
            oXLSheet.Cells(nRow, nCol + 1) = !identidad
            oXLSheet.Cells(nRow, nCol + 2) = !NombreCompleto
            oXLSheet.Cells(nRow, nCol + 3) = Format(!Ingreso, "#,###,##0.00")
            oXLSheet.Cells(nRow, nCol + 4) = Format(!Impuesto, "#,###,##0.00")
            oXLSheet.Cells(nRow, nCol + 5) = Format(!Multa, "#,###,##0.00")
            oXLSheet.Cells(nRow, nCol + 6) = Format(!Interes, "#,###,##0.00")
            oXLSheet.Cells(nRow, nCol + 7) = Format(!Recargo, "#,###,##0.00")
            oXLSheet.Cells(nRow, nCol + 8) = Format(!Descuento, "#,###,##0.00")
            oXLSheet.Cells(nRow, nCol + 9) = Format(!Total, "#,###,##0.00")
            oXLSheet.Cells(nRow, nCol + 10) = !CodDeclaraIP
            oXLSheet.Cells(nRow, nCol + 11) = !NumAvPg
            oXLSheet.Cells(nRow, nCol + 12) = !NumRecibo
            oXLSheet.Cells(nRow, nCol + 13) = !AnioDeclara
            oXLSheet.Cells(nRow, nCol + 14) = !FechaPresenta
            .MoveNext
            nRow = nRow + 1
        Loop
    End With
 
    
    'Calcule los totales
    nCol = 1
    oXLSheet.Cells(nRow + 1, nCol + 2) = "T o t a l e s"
    oXLSheet.Cells(nRow + 1, nCol + 3) = "=sum(D10:D" & nRow - 1 & ")"
    oXLSheet.Cells(nRow + 1, nCol + 4) = "=sum(E10:E" & nRow - 1 & ")"
    oXLSheet.Cells(nRow + 1, nCol + 5) = "=sum(F10:F" & nRow - 1 & ")"
    oXLSheet.Cells(nRow + 1, nCol + 6) = "=sum(G10:G" & nRow - 1 & ")"
    oXLSheet.Cells(nRow + 1, nCol + 7) = "=sum(H10:H" & nRow - 1 & ")"
    oXLSheet.Cells(nRow + 1, nCol + 8) = "=sum(I10:I" & nRow - 1 & ")"
    oXLSheet.Cells(nRow + 1, nCol + 9) = "=sum(J10:J" & nRow - 1 & ")"
    
    
    nRow = nRow + 1
    
    oXLSheet.Range("A" & nRow + 1 & ":" & "O" & nRow + 1).Font.Bold = True: ApplyBackColorToExcelCell exRng   'NEGRITA
    Set exRng = oXLSheet.Range("C" & nRow & ":" & "J" & nRow): ApplyBackColorToExcelCell exRng: oXLSheet.Range("C" & nRow & ":" & "J" & nRow).Font.Bold = True   ' MARGEN
    nRow = nRow + 1
    'Imprime el periodo del reporte
    nRow = nRow + 1
    oXLSheet.Cells(nRow, nCol) = strPeriodo
    Set exRng = oXLSheet.Range("A" & nRow, "O" & nRow): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    
    
    
    'Poner lineas en los datos
    Set exRng = oXLSheet.Range("A9:O" & nRow)
    ApplyBorderToExcelCell exRng
    


    'Cambiar el fondo de el titulo
    Set exRng = oXLSheet.Range("A9:O9")
    ApplyBackColorToExcelCell exRng
    
    '---------------------------------------------------
    oXLApp.Visible = True
    '---------------------------------------------------
    
    'Disconnect from Excel (let the user take over)
    Set oXLApp = Nothing
    '----------------------------------------------
    End Sub
