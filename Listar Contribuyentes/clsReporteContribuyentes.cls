VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReporteContribuyentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private rReportRs As New ADODB.Recordset

Public Sub CrearReporte()
    Dim StrRptContribuyente As String
    crearRecorset
    StrRptContribuyente = "SELECT Identidad, { fn CONCAT({ fn CONCAT(Pnombre, SNombre) }, { fn CONCAT(PApellido, SApellido) }) } AS NombreCompleto, Direccion, FechaNac, DATEDIFF(year, FechaNac, GETDATE()) AS Edad FROM Contribuyente ORDER BY NombreCompleto"
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    
    DeRia.rsAbonadoSP.Open (StrRptContribuyente)
    
    Do While Not DeRia.rsAbonadoSP.EOF
        rReportRs.AddNew
        rReportRs!Identidad = DeRia.rsAbonadoSP!Identidad
        rReportRs!Nombre = DeRia.rsAbonadoSP!NombreCompleto
        rReportRs!Direccion = Trim(DeRia.rsAbonadoSP!Direccion)
        rReportRs!FechaNac = DeRia.rsAbonadoSP!FechaNac
        rReportRs!Edad = DeRia.rsAbonadoSP!Edad
    DeRia.rsAbonadoSP.MoveNext
    Loop
End Sub

Private Sub crearRecorset()
    rReportRs.Fields.Append "Identidad", adBSTR
    rReportRs.Fields.Append "Nombre", adBSTR
    rReportRs.Fields.Append "Direccion", adBSTR
    rReportRs.Fields.Append "FechaNac", adBSTR
    rReportRs.Fields.Append "Edad", adBSTR
    rReportRs.Open
End Sub

Public Sub SendToExcel()

    If rReportRs.RecordCount = 0 Then Exit Sub
    
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
    
    'create and fill a recordset here, called oRecordset
    'oXLSheet.Range("B15").CopyFromRecordset rs
    'Set report title
    oXLSheet.Cells(1, 1) = "MUNICIPIO: " & oAlcaldia.Nombre
    oXLSheet.Cells(2, 1) = "LISTA DE CONTIBUYENTES REGISTADOS"
    
'    oXLSheet.Cells(4, 1) = "(Cantidad en Lempiras)"
    
    oXLSheet.Cells(6, 1) = "Departamento: " & oAlcaldia.Departamento
    oXLSheet.Cells(8, 1) = "Fecha Elaboración: " & DateToWord(Date)
    
    'Center Titles
    Set exRng = oXLSheet.Range("A1:E1"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A2:E2"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A3:E3"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A4:E4"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
   
    
    nRow = 12: nCol = 1
    'Pongamos los titulos de las columnas aqui, para que las agarre el autofit siguiente.
    oXLSheet.Cells(nRow, nCol) = "Identidad"                            'A
    oXLSheet.Cells(nRow, nCol + 1) = "Nombre Contribuyente"             'B 1
    oXLSheet.Cells(nRow, nCol + 2) = "Direccion "                       'C 2
    oXLSheet.Cells(nRow, nCol + 3) = "Fecha de Nacimiento"              'D
    oXLSheet.Cells(nRow, nCol + 4) = "Edad"                             'E

    '------------------------------------------------------------------------------------
  
    'Format titles Bold
    oXLSheet.Range("A1:D12").Font.Bold = True 'Los titulos del reporte
    
    'set a standard width
    Set exRng = oXLSheet.Columns("A")
    exRng.ColumnWidth = 17
    Set exRng = oXLSheet.Columns("B")
    exRng.ColumnWidth = 54
    Set exRng = oXLSheet.Columns("C")
    exRng.ColumnWidth = 54
    Set exRng = oXLSheet.Columns("D")
    exRng.ColumnWidth = 19
    Set exRng = oXLSheet.Columns("E")
    exRng.ColumnWidth = 13
    
    
    
    'Dibuje lineas en los titulos
    Set exRng = oXLSheet.Range("A12:A12")
    ApplyOutsideBorderToExcelCell exRng
    Set exRng = oXLSheet.Range("B12:B12")
    ApplyOutsideBorderToExcelCell exRng
    Set exRng = oXLSheet.Range("C12:C12")
    ApplyOutsideBorderToExcelCell exRng


    
    'Centrar los titulos del cuerpo
    Set exRng = oXLSheet.Range("A10:E12")
    exRng.HorizontalAlignment = xlCenter
    
    'Center between columns
    Set exRng = oXLSheet.Range("C6:E6"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
   ' Set exRng = oXLSheet.Range("Q10:R11"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
   ' Set exRng = oXLSheet.Range("S10:T11"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    
    'Define Starting Cells
    nRow = 13: nCol = 1:
    'Potencial Recaudación
    
    With rReportRs
        .MoveFirst
        Do While .EOF = False
            oXLSheet.Cells(nRow, nCol) = !Identidad
            oXLSheet.Cells(nRow, nCol + 1) = !Nombre
            oXLSheet.Cells(nRow, nCol + 2) = !Direccion
            oXLSheet.Cells(nRow, nCol + 3) = !FechaNac
            oXLSheet.Cells(nRow, nCol + 4) = !Edad
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
    '---------------------------------------------------

    'Formateamos numeros
    'Set exRng = oXLSheet.Columns("E:F")
    'exRng.NumberFormat = "#,###,###,##0.00"
    'Formateamos texto
    'Set exRng = oXLSheet.Columns("M:N")
    'Set exRng = oXLSheet.Range("E11:P11")
    'exRng.NumberFormat = "@"
    
    'Autofit Columns
    'Set exRng = oXLSheet.Range("A11:B12")
    'exRng.EntireColumn.AutoFit
        
    'Dibuje lineas en los titulos
    'Set exRng = oXLSheet.Range("A10:A11")
    'ApplyOutsideBorderToExcelCell exRng
    
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


