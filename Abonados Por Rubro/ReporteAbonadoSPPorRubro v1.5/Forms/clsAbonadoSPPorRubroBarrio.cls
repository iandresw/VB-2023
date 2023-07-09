VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAbonadoSPPorRubroBarrio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private rReportRs As New ADODB.Recordset
Private rsAbonadoSPEnc As New ADODB.Recordset
Private AbonadoSPDet As New ADODB.Recordset




Public Sub CrearReporte(txtRubroCtaIngreso As String, txtEstado As Integer, codBarrio As String)
    Dim StrRptAbonadoSPRubro As String
    crearRecorset
    StrRptAbonadoSPRubro = "SELECT AbonadoSPEnc.ASPE_Seq, Contribuyente.Pnombre, Contribuyente.SNombre, Contribuyente.PApellido, Contribuyente.SApellido, AbonadoSPEnc.Identidad,  Catastro.ClaveCatastro, AbonadoSPDet.ASPD_Valor, TablaBarrio.NombreBarrio "
    StrRptAbonadoSPRubro = StrRptAbonadoSPRubro & "FROM AbonadoSPDet INNER JOIN AbonadoSPEnc ON AbonadoSPEnc.ASPE_Seq = AbonadoSPDet.ASPE_Seq INNER JOIN Catastro  ON AbonadoSPEnc.ClaveCatastro = Catastro.ClaveCatastro  INNER JOIN Contribuyente  ON Catastro.Identidad = Contribuyente.Identidad  INNER JOIN TablaBarrio  ON Catastro.CodBarrio = TablaBarrio.CodBarrio  INNER JOIN AvPgEnc ON AvPgEnc.Identidad = Contribuyente.Identidad "
    StrRptAbonadoSPRubro = StrRptAbonadoSPRubro & "WHERE (AvPgEnc.AvPgEstado = 1) AND (AbonadoSPDet.CtaIngreso = '" & txtRubroCtaIngreso & "') AND (AbonadoSPEnc.ASPE_Estado  = " & txtEstado & ")" & codBarrio
    StrRptAbonadoSPRubro = StrRptAbonadoSPRubro & "GROUP BY AbonadoSPEnc.ASPE_Seq, AbonadoSPEnc.Identidad, Contribuyente.Pnombre, Contribuyente.SNombre, Contribuyente.PApellido, Contribuyente.SApellido, Catastro.ClaveCatastro, AbonadoSPDet.ASPD_Valor, TablaBarrio.NombreBarrio "
    StrRptAbonadoSPRubro = StrRptAbonadoSPRubro & "ORDER BY AbonadoSPEnc.Identidad "
    
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open (StrRptAbonadoSPRubro)
    
    Do While Not DeRia.rsAbonadoSP.EOF
    rReportRs.AddNew
    rReportRs!ASPE_Seq = DeRia.rsAbonadoSP!ASPE_Seq
    rReportRs!Nombre = Trim(DeRia.rsAbonadoSP!PNombre) & " " & Trim(DeRia.rsAbonadoSP!sNombre) & " " & Trim(DeRia.rsAbonadoSP!PApellido) & " " & Trim(DeRia.rsAbonadoSP!sApellido)
    rReportRs!Identidad = DeRia.rsAbonadoSP!Identidad
    rReportRs!ClaveCatastro = DeRia.rsAbonadoSP!ClaveCatastro
    rReportRs!ASPD_Valor = IIf(IsNull(DeRia.rsAbonadoSP!ASPD_Valor), 0, DeRia.rsAbonadoSP!ASPD_Valor)
    rReportRs!NombreBarrio = DeRia.rsAbonadoSP!NombreBarrio
    DeRia.rsAbonadoSP.MoveNext
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

Public Sub SendToExcel(txtRubroCtaIngreso As String, codBarrio As String, CodAldea As String)
    
    If rReportRs.RecordCount = 0 Then: MsgBox ("No Hay datos para mostrar"): Exit Sub
    Dim oXLApp As Excel.Application         'Declare the object variables
    Dim oXLBook As Excel.Workbook
    Dim oXLSheet As Excel.Worksheet
    Dim nRow, nCol As Long
    Dim oAlcaldia As New Alcaldia
    Dim oBarrio As String
    Dim oRubro As String
    Dim oAldea As String
    Dim exRng As Excel.Range
    Dim exRngData As Excel.Range
    Set oXLApp = New Excel.Application    'Create a new instance of Excel
    Set oXLBook = oXLApp.Workbooks.Add    'Add a new workbook
    Set oXLSheet = oXLBook.Worksheets(1)  'Work with the first
    Dim strCodigo As String
    
    If codBarrio = "Todos" = True Then
        oBarrio = "Todos los Barrios"
    ElseIf optInactivo.Value = True Then
        'Busca el barrio
        strCodigo = "SELECT NombreBarrio FROM TablaBarrio WHERE CodBarrio = '" & codBarrio & "'"
        If DeRia.rscmdTablaBarrio.State = 1 Then DeRia.rscmdTablaBarrio.Close
        DeRia.rscmdTablaBarrio.Open strCodigo
        oBarrio = DeRia.rscmdTablaBarrio!NombreBarrio
    End If
    

    'Busca el nombre de la cuenta de ingreso
    strCodigo = "SELECT NombreCtaIngreso FROM CuentaIngreso_A  WHERE (CtaIngreso = '" & txtRubroCtaIngreso & "')"
    If DeRia.rsCatalogoCuentasSP.State = 1 Then DeRia.rsCatalogoCuentasSP.Close
    DeRia.rsCatalogoCuentasSP.Open strCodigo
    oRubro = DeRia.rsCatalogoCuentasSP!NombreCtaIngreso
    
   'Busca el nombre de la aldea
   
    strCodigo = "SELECT NombreAldea FROM Aldea WHERE (CodAldea = '" & CodAldea & "')"
    If DeRia.rscmdAldeas.State = 1 Then DeRia.rscmdAldeas.Close
    DeRia.rscmdAldeas.Open strCodigo
    oAldea = DeRia.rscmdAldeas!NombreAldea
    
    
    
    oXLSheet.Cells(1, 1) = "MUNICIPIO: " & oAlcaldia.Nombre
    oXLSheet.Cells(2, 1) = "ABONADOS DE SERVICIOS PUBLICOS POR RUBRO"
    oXLSheet.Cells(5, 1) = "Departamento: " & oAlcaldia.Departamento
    oXLSheet.Cells(6, 1) = "Municipio: " & oAlcaldia.Nombre
    oXLSheet.Cells(7, 1) = "Aldea: " & oAldea
    oXLSheet.Cells(8, 1) = "Barrio: " & oBarrio
    oXLSheet.Cells(9, 1) = "Tipo de Servicio: " & oRubro
    oXLSheet.Cells(10, 1) = "Fecha Elaboración: " & DateToWord(Date)
    
    'Center Titles
    Set exRng = oXLSheet.Range("A1:F1"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A2:F2"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A3:F3"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A5:F5"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
    Set exRng = oXLSheet.Range("A6:F6"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
    Set exRng = oXLSheet.Range("A7:F7"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
    Set exRng = oXLSheet.Range("A8:F8"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
    Set exRng = oXLSheet.Range("A9:F9"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
    Set exRng = oXLSheet.Range("A10:F10"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
    
   
    
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
    oXLSheet.Range("A1:F12").Font.Bold = True 'Los titulos del reporte
    'set a standard width
    Set exRng = oXLSheet.Columns("A")
    exRng.ColumnWidth = 11
    Set exRng = oXLSheet.Columns("B")
    exRng.ColumnWidth = 54
    Set exRng = oXLSheet.Columns("C")
    exRng.ColumnWidth = 17
    Set exRng = oXLSheet.Columns("D")
    exRng.ColumnWidth = 17
    Set exRng = oXLSheet.Columns("E")
    exRng.ColumnWidth = 10
    Set exRng = oXLSheet.Columns("F")
    exRng.ColumnWidth = 17
    
    
    
    'Dibuje lineas en los titulos
    Set exRng = oXLSheet.Range("A12:A12")
    ApplyOutsideBorderToExcelCell exRng
    Set exRng = oXLSheet.Range("B12:B12")
    ApplyOutsideBorderToExcelCell exRng
    Set exRng = oXLSheet.Range("C12:C12")
    ApplyOutsideBorderToExcelCell exRng
    Set exRng = oXLSheet.Range("E12:E12")
    ApplyOutsideBorderToExcelCell exRng
    Set exRng = oXLSheet.Range("F12:F12")
    ApplyOutsideBorderToExcelCell exRng

    'Centrar los titulos del cuerpo
    Set exRng = oXLSheet.Range("A12:F12")
    exRng.HorizontalAlignment = xlCenter
    
    
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
            oXLSheet.Cells(nRow, nCol + 4) = Format(!ASPD_Valor, "######0")
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
    'oXLSheet.Cells(nRow, nCol) = strPeriodo
    Set exRng = oXLSheet.Range("A" & nRow, "F" & nRow): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    
    'Poner lineas en los datos
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




