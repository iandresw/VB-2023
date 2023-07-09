VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CrearReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private rReportRs As New ADODB.Recordset
Dim excelApp As Object ' Objeto para la aplicación Excel
Dim excelWorkbook As Object ' Objeto para el libro de Excel
Dim excelWorksheet As Object

Public Sub CrearReporte(fechaInicial As String, FechaFinal As String)
    Dim sql As String
    crearRecorset
    sql = "SELECT Identidad, Negocio, Propietario, Ubicacion, Actividad, TipoCuenta, Observacion, Fecha, NumRecibo, NoPermiso "
    sql = sql + "From Tra_PermOP  WHERE (Fecha BETWEEN '" + fechaInicial + "' AND '" + FechaFinal + "')"
    
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open (sql)
    
    Do While Not DeRia.rsAbonadoSP.EOF
    rReportRs.AddNew
    rReportRs!identidad = DeRia.rsAbonadoSP!identidad
    rReportRs!Negocio = DeRia.rsAbonadoSP!Negocio
    rReportRs!Propietario = DeRia.rsAbonadoSP!Propietario
    rReportRs!Ubicacion = DeRia.rsAbonadoSP!Ubicacion
    rReportRs!Actividad = DeRia.rsAbonadoSP!Actividad
    rReportRs!TipoCuenta = DeRia.rsAbonadoSP!TipoCuenta
    rReportRs!observacion = DeRia.rsAbonadoSP!observacion
    rReportRs!Fecha = DeRia.rsAbonadoSP!Fecha
    rReportRs!numRecibo = DeRia.rsAbonadoSP!numRecibo
    rReportRs!NoPermiso = DeRia.rsAbonadoSP!NoPermiso
    DeRia.rsAbonadoSP.MoveNext
    Loop
End Sub
Private Sub crearRecorset()
    rReportRs.Fields.Append "Identidad", adBSTR
    rReportRs.Fields.Append "Negocio", adBSTR
    rReportRs.Fields.Append "Propietario", adBSTR
    rReportRs.Fields.Append "Ubicacion", adBSTR
    rReportRs.Fields.Append "Actividad", adBSTR
    rReportRs.Fields.Append "TipoCuenta", adBSTR
    rReportRs.Fields.Append "observacion", adBSTR
    rReportRs.Fields.Append "Fecha", adBSTR
    rReportRs.Fields.Append "numRecibo", adBSTR
    rReportRs.Fields.Append "NoPermiso", adBSTR
    rReportRs.Open
End Sub

Public Sub SendToExcel()
    
    If rReportRs.RecordCount = 0 Then: MsgBox ("No Hay datos para mostrar"): Exit Sub
    Dim oXLApp As Excel.Application         'Declare the object variables
    Dim oXLBook As Excel.Workbook
    Dim oXLSheet As Excel.Worksheet
    Dim nRow, nCol As Long
    'Dim oAlcaldia As New Alcaldia
    Dim exRng As Excel.Range
    Dim exRngData As Excel.Range
    Set oXLApp = New Excel.Application    'Create a new instance of Excel
    Set oXLBook = oXLApp.Workbooks.Add    'Add a new workbook
    Set oXLSheet = oXLBook.Worksheets(1)  'Work with the first
    Dim strCodigo As String
    
    oXLSheet.Cells(1, 1) = "MUNICIPIO: LA LABOR"
    oXLSheet.Cells(2, 1) = "PERMISOS DE OPERACION EMITIDOS"
    oXLSheet.Cells(5, 1) = "DEPARTAMENTO: OCOTEPEQUE "
    oXLSheet.Cells(6, 1) = "Municipio: La Labor"
    'oXLSheet.Cells(, 1) = "Fecha Elaboración: " & DateToWord(Date)
    
    'Center Titles
    Set exRng = oXLSheet.Range("A1:F1"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A2:F2"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A3:F3"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A5:F5"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
    Set exRng = oXLSheet.Range("A6:F6"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
    Set exRng = oXLSheet.Range("A7:F7"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
    
    
   
    
    nRow = 12: nCol = 1
    'Pongamos los titulos de las columnas aqui, para que las agarre el autofit siguiente.
    oXLSheet.Cells(nRow, nCol) = "Identidad"                     'A
    oXLSheet.Cells(nRow, nCol + 1) = "Negocio"                   'B
    oXLSheet.Cells(nRow, nCol + 2) = "Propietario"               'C
    oXLSheet.Cells(nRow, nCol + 3) = "Ubicacion"                 'D
    oXLSheet.Cells(nRow, nCol + 4) = "Actividad"                 'E
    oXLSheet.Cells(nRow, nCol + 5) = "TipoCuenta"                'F
    oXLSheet.Cells(nRow, nCol + 6) = "observacion"               'G
    oXLSheet.Cells(nRow, nCol + 7) = "Fecha"                     'H
    oXLSheet.Cells(nRow, nCol + 8) = "numRecibo"                 'I
    oXLSheet.Cells(nRow, nCol + 9) = "NoPermiso"                 'J
    
    '------------------------------------------------------------------------------------
  
    'Format titles Bold
    oXLSheet.Range("A1:J12").Font.Bold = True 'Los titulos del reporte
    'set a standard width
    Set exRng = oXLSheet.Columns("A")
    exRng.ColumnWidth = 15
    Set exRng = oXLSheet.Columns("B")
    exRng.ColumnWidth = 15
    Set exRng = oXLSheet.Columns("C")
    exRng.ColumnWidth = 17
    Set exRng = oXLSheet.Columns("D")
    exRng.ColumnWidth = 35
    Set exRng = oXLSheet.Columns("E")
    exRng.ColumnWidth = 15
    Set exRng = oXLSheet.Columns("F")
    exRng.ColumnWidth = 15
    Set exRng = oXLSheet.Columns("G")
    exRng.ColumnWidth = 15
    Set exRng = oXLSheet.Columns("H")
    exRng.ColumnWidth = 17
    Set exRng = oXLSheet.Columns("I")
    exRng.ColumnWidth = 35
    Set exRng = oXLSheet.Columns("J")
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
    Set exRng = oXLSheet.Range("A12:J12")
    exRng.HorizontalAlignment = xlCenter
    
    
    
    'Define Starting Cells
    nRow = 13: nCol = 1:
    'Potencial Recaudación
    With rReportRs
        .MoveFirst
        Do While .EOF = False
            oXLSheet.Cells(nRow, nCol) = !identidad
            oXLSheet.Cells(nRow, nCol + 1) = !Negocio
            oXLSheet.Cells(nRow, nCol + 2) = !Propietario
            oXLSheet.Cells(nRow, nCol + 3) = !Ubicacion
            oXLSheet.Cells(nRow, nCol + 4) = !Actividad
            oXLSheet.Cells(nRow, nCol + 5) = !TipoCuenta
            oXLSheet.Cells(nRow, nCol + 6) = !observacion
            oXLSheet.Cells(nRow, nCol + 7) = !Fecha
            oXLSheet.Cells(nRow, nCol + 8) = !numRecibo
            oXLSheet.Cells(nRow, nCol + 9) = !NoPermiso
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
