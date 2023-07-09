VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPropiedadesPorDeclaracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private rReportRs As New ADODB.Recordset
Private rsAbonadoSP As New ADODB.Recordset
Public Sub CrearReporte(txtFechaInicio As Date, txtFechaFinal As Date)
    Dim StrPropPorDeclara As String
    crearRecorset
    
    StrPropPorDeclara = " SELECT DeclaraBI.FechaDeclaraBI, DeclaraBI.CodDeclaraBI, catastro.ClaveCatastro, Contribuyente.Identidad, Contribuyente.Pnombre, Contribuyente.SNombre, Contribuyente.PApellido, Contribuyente.SApellido, "
    StrPropPorDeclara = StrPropPorDeclara & " DeclaraBI.DeclaraAreaTerreno, DeclaraBI.DeclaraValorTerreno, DeclaraBI.DeclaraAreaEdificacion, DeclaraBI.DeclaraValorEdificacion, DeclaraBI.ImpuestoBI "
    StrPropPorDeclara = StrPropPorDeclara & " FROM   Contribuyente INNER JOIN  DeclaraBI  ON  Contribuyente.Identidad = DeclaraBI.Identidad inner join catastro on DeclaraBI.clavecatastro = catastro.clavecatastro "
    StrPropPorDeclara = StrPropPorDeclara & " WHERE (DeclaraBI.FechaDeclaraBI BETWEEN '" & txtFechaInicio & "' AND '" & txtFechaFinal & "')and EstadoDeclaraBI = 0"
    StrPropPorDeclara = StrPropPorDeclara & " GROUP BY  DeclaraBI.FechaDeclaraBI,  DeclaraBI.CodDeclaraBI, catastro.ClaveCatastro, Contribuyente.Identidad,  Contribuyente.Pnombre,  Contribuyente.SNombre, Contribuyente.PApellido, Contribuyente.SApellido, "
    StrPropPorDeclara = StrPropPorDeclara & " DeclaraBI.DeclaraAreaTerreno, DeclaraBI.DeclaraValorTerreno, DeclaraBI.DeclaraAreaEdificacion, DeclaraBI.DeclaraValorEdificacion, DeclaraBI.ImpuestoBI "
    StrPropPorDeclara = StrPropPorDeclara & " ORDER BY DeclaraBI.FechaDeclaraBI, catastro.ClaveCatastro "
    
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open (StrPropPorDeclara)
    
    Do While Not DeRia.rsAbonadoSP.EOF
        rReportRs.AddNew
        rReportRs!FechaDeclaraBI = DeRia.rsAbonadoSP!FechaDeclaraBI
        rReportRs!CodDeclaraBI = DeRia.rsAbonadoSP!CodDeclaraBI
        rReportRs!ClaveCatastro = DeRia.rsAbonadoSP!ClaveCatastro
        rReportRs!Identidad = DeRia.rsAbonadoSP!Identidad
        rReportRs!NombreCompleto = Trim(DeRia.rsAbonadoSP!PNombre) & " " & Trim(DeRia.rsAbonadoSP!sNombre) & " " & Trim(DeRia.rsAbonadoSP!PApellido) & " " & Trim(DeRia.rsAbonadoSP!sApellido)
        rReportRs!DeclaraAreaTerreno = DeRia.rsAbonadoSP!DeclaraAreaTerreno
        rReportRs!DeclaraValorTerreno = DeRia.rsAbonadoSP!DeclaraValorTerreno
        rReportRs!DeclaraAreaEdificacion = DeRia.rsAbonadoSP!DeclaraAreaEdificacion
        rReportRs!DeclaraValorEdificacion = DeRia.rsAbonadoSP!DeclaraValorEdificacion
        rReportRs!ImpuestoBI = DeRia.rsAbonadoSP!ImpuestoBI
        DeRia.rsAbonadoSP.MoveNext
    Loop
End Sub

Private Sub crearRecorset()
    rReportRs.Fields.Append "FechaDeclaraBI", adBSTR
    rReportRs.Fields.Append "CodDeclaraBI", adBSTR
    rReportRs.Fields.Append "ClaveCatastro", adBSTR
    rReportRs.Fields.Append "Identidad", adBSTR
    rReportRs.Fields.Append "NombreCompleto", adBSTR
    rReportRs.Fields.Append "DeclaraAreaTerreno", adBSTR
    rReportRs.Fields.Append "DeclaraValorTerreno", adBSTR
    rReportRs.Fields.Append "DeclaraAreaEdificacion", adBSTR
    rReportRs.Fields.Append "DeclaraValorEdificacion", adBSTR
    rReportRs.Fields.Append "ImpuestoBI", adBSTR
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
    oXLSheet.Cells(2, 1) = "DECLARACIONES JURADAS DE BIENES INMUEBLES REALIZADAS POR RANGO DE FECHAS"
    oXLSheet.Cells(5, 1) = "Departamento: " & oAlcaldia.Departamento
    oXLSheet.Cells(6, 1) = "Municipio: " & oAlcaldia.Nombre
    oXLSheet.Cells(7, 1) = "Fecha Elaboración: " & DateToWord(Date)
    oXLSheet.Cells(8, 1) = "Declaraciones emitidas De: " & frmPropiedadesConDeclaracion.txtFechaInicio.Text & " Hasta: " & frmPropiedadesConDeclaracion.txtFechaFinal.Text
    
    'Center Titles
    Set exRng = oXLSheet.Range("A1:J1"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A2:J2"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A3:J3"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A5:J5"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
    Set exRng = oXLSheet.Range("A6:J6"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
    Set exRng = oXLSheet.Range("A7:J7"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
 
    
    
    
    nRow = 12: nCol = 1
    'Pongamos los titulos de las columnas aqui, para que las agarre el autofit siguiente.
    oXLSheet.Cells(nRow, nCol) = "Fecha Declaracion BI"                         'A
    oXLSheet.Cells(nRow, nCol + 1) = "Codigo Declaracion"        'B
    oXLSheet.Cells(nRow, nCol + 2) = "Clave Catastro"            'C
    oXLSheet.Cells(nRow, nCol + 3) = "Identidad"           'D
    oXLSheet.Cells(nRow, nCol + 4) = "Nombre Completo"      'E
    oXLSheet.Cells(nRow, nCol + 5) = "Declara Area Terreno"     'F
    oXLSheet.Cells(nRow, nCol + 6) = "Declara Valor Terreno"     'F
    oXLSheet.Cells(nRow, nCol + 7) = "Declara Area Edificacion"        'G
    oXLSheet.Cells(nRow, nCol + 8) = "Declara Valor Edificacion"  'H
    oXLSheet.Cells(nRow, nCol + 9) = "Impuesto BI"               'J
    '------------------------------------------------------------------------------------
  
    'Format titles Bold
    oXLSheet.Range("A1:J12").Font.Bold = True 'Los titulos del reporte
    'set a standard width
    Set exRng = oXLSheet.Columns("A")
    exRng.ColumnWidth = 20
    Set exRng = oXLSheet.Columns("B")
    exRng.ColumnWidth = 20
    Set exRng = oXLSheet.Columns("C")
    exRng.ColumnWidth = 20
    Set exRng = oXLSheet.Columns("D")
    exRng.ColumnWidth = 20
    Set exRng = oXLSheet.Columns("E")
    exRng.ColumnWidth = 50
    Set exRng = oXLSheet.Columns("F")
    exRng.ColumnWidth = 20
    Set exRng = oXLSheet.Columns("G")
    exRng.ColumnWidth = 20
    Set exRng = oXLSheet.Columns("H")
    exRng.ColumnWidth = 20
    Set exRng = oXLSheet.Columns("I")
    exRng.ColumnWidth = 20
    Set exRng = oXLSheet.Columns("J")
    exRng.ColumnWidth = 20
    
    
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
    Set exRng = oXLSheet.Range("A12:J12")
    exRng.HorizontalAlignment = xlCenter
    


    
    'Define Starting Cells
    nRow = 13: nCol = 1:
    'Potencial Recaudación
    With rReportRs
        .MoveFirst
        Do While .EOF = False
            oXLSheet.Cells(nRow, nCol) = !FechaDeclaraBI
            oXLSheet.Cells(nRow, nCol + 1) = !CodDeclaraBI
            oXLSheet.Cells(nRow, nCol + 2) = !ClaveCatastro
            oXLSheet.Cells(nRow, nCol + 3) = !Identidad
            oXLSheet.Cells(nRow, nCol + 4) = !NombreCompleto
            oXLSheet.Cells(nRow, nCol + 5) = Format(!DeclaraAreaTerreno, "#,###,##0.00")
            oXLSheet.Cells(nRow, nCol + 6) = Format(!DeclaraValorTerreno, "#,###,##0.00")
            oXLSheet.Cells(nRow, nCol + 7) = Format(!DeclaraAreaEdificacion, "#,###,##0.00")
            oXLSheet.Cells(nRow, nCol + 8) = Format(!DeclaraValorEdificacion, "#,###,##0.00")
            oXLSheet.Cells(nRow, nCol + 9) = Format(!ImpuestoBI, "#,###,##0.00")
            .MoveNext
            nRow = nRow + 1
        Loop
    End With
    oXLSheet.Range("A" & nRow + 1 & ":" & "J" & nRow + 1).Font.Bold = True  'NEGRITA
    Set exRng = oXLSheet.Range("A" & nRow + 1 & ":" & "B" & nRow + 1): exRng.HorizontalAlignment = xlCenter: exRng.Merge ' MARGEN
    nRow = nRow + 1
    'Imprime el periodo del reporte
    nRow = nRow + 1
    'oXLSheet.Cells(nRow, nCol) = strPeriodo
    Set exRng = oXLSheet.Range("A" & nRow, "J" & nRow): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    
    'Poner lineas en los datos
    Set exRng = oXLSheet.Range("A12:J" & nRow)
    ApplyBorderToExcelCell exRng
    


    'Cambiar el fondo de el titulo
    Set exRng = oXLSheet.Range("A12:J12")
    ApplyBackColorToExcelCell exRng
    
    '---------------------------------------------------
    oXLApp.Visible = True
    '---------------------------------------------------
    
    'Disconnect from Excel (let the user take over)
    Set oXLApp = Nothing
    '----------------------------------------------

End Sub
