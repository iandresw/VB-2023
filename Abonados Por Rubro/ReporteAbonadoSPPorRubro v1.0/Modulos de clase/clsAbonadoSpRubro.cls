VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAbonadoSpRubro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private rReportRs As New ADODB.Recordset
Private rsAbonadoSPEnc As New ADODB.Recordset
Private AbonadoSPDet As New ADODB.Recordset
Public Sub CrearReporte(txtRubroCtaIngreso As String, codBarrio As String)
    Dim StrRptAbonadoSPRubro As String
    crearRecorset
    StrRptAbonadoSPRubro = "SELECT AbonadoSPEnc.ASPE_Seq, { fn CONCAT({ fn CONCAT(Contribuyente.Pnombre, Contribuyente.SNombre) }, { fn CONCAT(Contribuyente.PApellido, Contribuyente.SApellido) }) } AS NombreCompleto, AbonadoSPEnc.Identidad,  Catastro.ClaveCatastro, AbonadoSPDet.ASPD_Valor, TablaBarrio.NombreBarrio "
    StrRptAbonadoSPRubro = StrRptAbonadoSPRubro & "FROM AbonadoSPDet INNER JOIN AbonadoSPEnc ON AbonadoSPEnc.ASPE_Seq = AbonadoSPDet.ASPE_Seq INNER JOIN Catastro  ON AbonadoSPEnc.ClaveCatastro = Catastro.ClaveCatastro  INNER JOIN Contribuyente  ON Catastro.Identidad = Contribuyente.Identidad  INNER JOIN TablaBarrio  ON Catastro.CodBarrio = TablaBarrio.CodBarrio  INNER JOIN AvPgEnc ON AvPgEnc.Identidad = Contribuyente.Identidad "
    StrRptAbonadoSPRubro = StrRptAbonadoSPRubro & "WHERE (AvPgEnc.AvPgEstado = 1) AND (AbonadoSPDet.CtaIngreso = '" & txtRubroCtaIngreso & "') AND (TablaBarrio.CodBarrio = '" & codBarrio & "') "
    StrRptAbonadoSPRubro = StrRptAbonadoSPRubro & "GROUP BY AbonadoSPEnc.ASPE_Seq, AbonadoSPEnc.Identidad, Contribuyente.Pnombre, Contribuyente.SNombre, Contribuyente.PApellido, Contribuyente.SApellido, Catastro.ClaveCatastro, AbonadoSPDet.ASPD_Valor, TablaBarrio.NombreBarrio "
    StrRptAbonadoSPRubro = StrRptAbonadoSPRubro & "ORDER BY AbonadoSPEnc.Identidad "
    
    If DeRia.rsRpPorRubroBarrio.State = 1 Then DeRia.rsRpPorRubroBarrio.Close
    DeRia.rsRpPorRubroBarrio.Open (StrRptAbonadoSPRubro)
    MsgBox (StrRptAbonadoSPRubro)
    Do While Not DeRia.rsRpPorRubroBarrio.EOF
    rReportRs.AddNew
    rReportRs!ASPE_Seq = DeRia.rsRpPorRubroBarrio!ASPE_Seq
    rReportRs!Nombre = DeRia.rsRpPorRubroBarrio!NombreCompleto
    rReportRs!Identidad = DeRia.rsRpPorRubroBarrio!Identidad
    rReportRs!ClaveCatastro = DeRia.rsRpPorRubroBarrio!ClaveCatastro
    rReportRs!ASPD_Valor = DeRia.rsRpPorRubroBarrio!ASPD_Valor
    rReportRs!NombreBarrio = DeRia.rsRpPorRubroBarrio!NombreBarrio
    DeRia.rsRpPorRubroBarrio.MoveNext
    Loop
End Sub

Private Sub crearRecorset()
    rReportRs.Fields.Append "ASPE_Seq", adBSTR
    rReportRs.Fields.Append "Nombre", adBSTR
    rReportRs.Fields.Append "Identidad", adBSTR
    rReportRs.Fields.Append "ClaveCatastro", adBSTR
    rReportRs.Fields.Append "ASPD_Valor", adBSTR
    rReportRs.Fields.Append "NombreBarrio", adBSTR
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
    
    oXLSheet.Cells(1, 1) = "MUNICIPIO: " & oAlcaldia.Nombre
    oXLSheet.Cells(2, 1) = "ABONADOS DE SERVICIOS PUBLICOS POR RUBRO"
    oXLSheet.Cells(6, 1) = "Departamento: " & oAlcaldia.Departamento
    oXLSheet.Cells(7, 1) = "Aldea: " & oAlcaldia.Departamento
    oXLSheet.Cells(8, 1) = "Departamento: " & oAlcaldia.Departamento
    'oXLSheet.Cells(9, 1) = "Fecha Elaboración: " & DateToWord(Date)
    
    'Center Titles
    Set exRng = oXLSheet.Range("A1:F1"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A2:F2"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A3:F3"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A4:F4"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
   
    
    nRow = 12: nCol = 1
    'Pongamos los titulos de las columnas aqui, para que las agarre el autofit siguiente.
    oXLSheet.Cells(nRow, nCol) = "Codigo Abonado"                            'A
    oXLSheet.Cells(nRow, nCol + 1) = "Nombre Contribuyente"                  'B
    oXLSheet.Cells(nRow, nCol + 2) = "Identidad "                            'C
    oXLSheet.Cells(nRow, nCol + 3) = "Clave Catastral"                       'D
    oXLSheet.Cells(nRow, nCol + 4) = "Cuota"                                 'E
    oXLSheet.Cells(nRow, nCol + 5) = "Barrio"                                'F

    '------------------------------------------------------------------------------------
  
    'Format titles Bold
    oXLSheet.Range("A1:D12").Font.Bold = True 'Los titulos del reporte
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
    Set exRng = oXLSheet.Columns("F")
    exRng.ColumnWidth = 20
    
    
    
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
    Set exRng = oXLSheet.Range("C6:F6"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    'Define Starting Cells
    nRow = 13: nCol = 1:
    'Potencial Recaudación
    
    With rReportRs
        .MoveFirst
        Do While .EOF = False
            oXLSheet.Cells(nRow, nCol) = !ASPE_Seq
            oXLSheet.Cells(nRow, nCol + 1) = !Nombre
            oXLSheet.Cells(nRow, nCol + 2) = !Identidad
            oXLSheet.Cells(nRow, nCol + 3) = !ClaveCatastro
            oXLSheet.Cells(nRow, nCol + 4) = !ASPD_Valor
            oXLSheet.Cells(nRow, nCol + 5) = !NombreBarrio
            .MoveNext
            nRow = nRow + 1
        Loop
    End With
    oXLSheet.Range("A" & nRow + 1 & ":" & "F" & nRow + 1).Font.Bold = True  'NEGRITA
    Set exRng = oXLSheet.Range("A" & nRow + 1 & ":" & "B" & nRow + 1): exRng.HorizontalAlignment = xlCenter: exRng.Merge ' MARGEN
    nRow = nRow + 1
    'Imprime el periodo del reporte
    nRow = nRow + 1
    Set exRng = oXLSheet.Range("A" & nRow, "F" & nRow): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A12:F" & nRow)
    ApplyBorderToExcelCell exRng
    'Cambiar el fondo de el titulo
    Set exRng = oXLSheet.Range("A12:F12")
    ApplyBackColorToExcelCell exRng
    
    '---------------------------------------------------
    oXLApp.Visible = True
    '---------------------------------------------------
    
    'Disconnect from Excel (let the user take over)
    Set oXLApp = Nothing
    '----------------------------------------------

End Sub




