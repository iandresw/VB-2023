VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEnviarExel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private rReportRs As New ADODB.Recordset
Private rsAbonadoSP As New ADODB.Recordset
Public Sub CrearReporte()
    Dim StrFactBancoNoPag As String
    crearRecorset
    StrFactBancoNoPag = "SELECT MIN(FactXTes.NoFacts) AS PrimeraFac, MAX(FactXTes.NoFacts) AS UltimmaFact, Contribuyente.Identidad, { fn CONCAT({ fn CONCAT(Contribuyente.Pnombre, Contribuyente.SNombre) }, { fn CONCAT(Contribuyente.PApellido, "
    StrFactBancoNoPag = StrFactBancoNoPag & "Contribuyente.SApellido) }) } AS NombreCompleto, SUM(AvPgDetalle.ValorUnitAvPgDet) AS Valor "
    StrFactBancoNoPag = StrFactBancoNoPag & "FROM AvPgDetalle INNER JOIN AvPgEnc ON AvPgEnc.NumAvPg = AvPgDetalle.NumAvPg INNER JOIN Contribuyente ON Contribuyente.Identidad "
    StrFactBancoNoPag = StrFactBancoNoPag & "= AvPgEnc.Identidad INNER JOIN FactXTes ON FactXTes.NoFacts = AvPgEnc.NumAvPg "
    StrFactBancoNoPag = StrFactBancoNoPag & "GROUP BY Contribuyente.Identidad, Contribuyente.Pnombre, Contribuyente.SNombre, Contribuyente.PApellido, Contribuyente.SApellido "
    
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open (StrFactBancoNoPag)
    
    Do While Not DeRia.rsAbonadoSP.EOF
    rReportRs.AddNew
    rReportRs!PrimeraFac = DeRia.rsAbonadoSP!PrimeraFac
    rReportRs!UltimaFact = DeRia.rsAbonadoSP!UltimmaFact
    rReportRs!Identidad = DeRia.rsAbonadoSP!Identidad
    rReportRs!NombreCompleto = DeRia.rsAbonadoSP!NombreCompleto
    rReportRs!Valor = DeRia.rsAbonadoSP!Valor
    DeRia.rsAbonadoSP.MoveNext
    Loop
End Sub

Private Sub crearRecorset()
    
    rReportRs.Fields.Append "PrimeraFac", adBSTR
    rReportRs.Fields.Append "UltimaFact", adBSTR
    rReportRs.Fields.Append "Identidad", adBSTR
    rReportRs.Fields.Append "NombreCompleto", adBSTR
    rReportRs.Fields.Append "Valor", adBSTR
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
    oXLSheet.Cells(2, 1) = "PAGOS ENVIDOS AL BANCO DE DIA"
    oXLSheet.Cells(5, 1) = "Departamento: " & oAlcaldia.Departamento
    oXLSheet.Cells(6, 1) = "Municipio: " & oAlcaldia.Nombre
    oXLSheet.Cells(, 1) = "Fecha Elaboración: " & DateToWord(Date)
    
    'Center Titles
    Set exRng = oXLSheet.Range("A1:F1"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A2:F2"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A3:F3"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A5:F5"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
    Set exRng = oXLSheet.Range("A6:F6"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
    Set exRng = oXLSheet.Range("A7:F7"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
    
   
    
    nRow = 12: nCol = 1
    'Pongamos los titulos de las columnas aqui, para que las agarre el autofit siguiente.
    oXLSheet.Cells(nRow, nCol) = "Primera Factura"                     'A
    oXLSheet.Cells(nRow, nCol + 1) = "Ultima Factura"                  'B
    oXLSheet.Cells(nRow, nCol + 2) = "Identidad"                       'C
    oXLSheet.Cells(nRow, nCol + 3) = "Nombre Contibuyente"             'D
    oXLSheet.Cells(nRow, nCol + 4) = "Valor"                           'E
    
    '------------------------------------------------------------------------------------
  
    'Format titles Bold
    oXLSheet.Range("A1:E12").Font.Bold = True 'Los titulos del reporte
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
    Set exRng = oXLSheet.Range("A12:E12")
    exRng.HorizontalAlignment = xlCenter
    

    
    
    
    'Define Starting Cells
    nRow = 13: nCol = 1:
    'Potencial Recaudación
    With rReportRs
        .MoveFirst
        Do While .EOF = False
            oXLSheet.Cells(nRow, nCol) = !PrimeraFac
            oXLSheet.Cells(nRow, nCol + 1) = !UltimaFact
            oXLSheet.Cells(nRow, nCol + 2) = !Identidad
            oXLSheet.Cells(nRow, nCol + 3) = !NombreCompleto
            oXLSheet.Cells(nRow, nCol + 4) = !Valor
            .MoveNext
            nRow = nRow + 1
        Loop
    End With
    oXLSheet.Range("A" & nRow + 1 & ":" & "E" & nRow + 1).Font.Bold = True  'NEGRITA
    Set exRng = oXLSheet.Range("A" & nRow + 1 & ":" & "B" & nRow + 1): exRng.HorizontalAlignment = xlCenter: exRng.Merge ' MARGEN
    nRow = nRow + 1
    'Imprime el periodo del reporte
    nRow = nRow + 1
    'oXLSheet.Cells(nRow, nCol) = strPeriodo
    Set exRng = oXLSheet.Range("A" & nRow, "E" & nRow): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    
    'Poner lineas en los datos
    Set exRng = oXLSheet.Range("A12:E" & nRow)
    ApplyBorderToExcelCell exRng
    


    'Cambiar el fondo de el titulo
    Set exRng = oXLSheet.Range("A12:E12")
    ApplyBackColorToExcelCell exRng
    
    '---------------------------------------------------
    oXLApp.Visible = True
    '---------------------------------------------------
    
    'Disconnect from Excel (let the user take over)
    Set oXLApp = Nothing
    '----------------------------------------------

End Sub
