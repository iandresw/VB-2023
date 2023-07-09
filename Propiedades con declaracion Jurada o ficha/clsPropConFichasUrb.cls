VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPropConFichasUrb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private rReportRs As New ADODB.Recordset
Private rsAbonadoSP As New ADODB.Recordset
Public Sub CrearReporte()
    Dim Str, str1, str2 As String
    
    crearRecorset
    
    Str = " SELECT AvPgDetalle.ClaveCatastro, Contribuyente.Identidad, Contribuyente.Pnombre, Contribuyente.SNombre, Contribuyente.PApellido, Contribuyente.SApellido, AvPgDetalle.ValorUnitAvPgDet, SUM(AvPgDetalle.ValorUnitAvPgDet)  AS impuesto "
    Str = Str & "FROM AvPgEnc INNER JOIN "
    Str = Str & "AvPgDetalle ON AvPgEnc.NumAvPg = AvPgDetalle.NumAvPg INNER JOIN "
    Str = Str & "CuentaIngreso_A ON AvPgDetalle.CtaIngreso = CuentaIngreso_A.CtaIngreso AND DATEPART(year, AvPgEnc.FechaVenceAvPg) = CuentaIngreso_A.Anio INNER JOIN "
    Str = Str & "Contribuyente ON AvPgEnc.Identidad = Contribuyente.Identidad "
    Str = Str & "Where (AvPgEnc.AvPgTipoImpuesto = 1) And (AvPgEnc.AvPgEstado = 1) And (CuentaIngreso_A.Tipo = 1) "
    Str = Str & "GROUP BY AvPgDetalle.ClaveCatastro, Contribuyente.Identidad, Contribuyente.Pnombre, Contribuyente.SNombre, Contribuyente.PApellido, AvPgDetalle.ValorUnitAvPgDet, Contribuyente.SApellido "
    
            
    str1 = " SELECT        SUM(AvPgDetalle.ValorUnitAvPgDet) AS Intereses"
    str1 = str1 & " FROM            AvPgEnc INNER JOIN"
    str1 = str1 & " AvPgDetalle ON AvPgEnc.NumAvPg = AvPgDetalle.NumAvPg INNER JOIN"
    str1 = str1 & " CuentaIngreso_A ON AvPgDetalle.CtaIngreso = CuentaIngreso_A.CtaIngreso AND DATEPART(year, AvPgEnc.FechaVenceAvPg) = CuentaIngreso_A.Anio"
    str1 = str1 & " WHERE        (AvPgEnc.AvPgTipoImpuesto = 1) AND (AvPgEnc.AvPgEstado = 1) AND (AvPgEnc.ClaveCatastro = '161723001018') AND (SUBSTRING(AvPgDetalle.CtaIngreso, 1, 6) = '112126')"
    str1 = str1 & " GROUP BY AvPgDetalle.ClaveCatastro"
  
            
    str2 = "SELECT SUM(AvPgDetalle.ValorUnitAvPgDet) AS Recargos"
    str2 = str2 & "FROM            AvPgEnc INNER JOIN"
    str2 = str2 & "AvPgDetalle ON AvPgEnc.NumAvPg = AvPgDetalle.NumAvPg INNER JOIN"
    str2 = str2 & "CuentaIngreso_A ON AvPgDetalle.CtaIngreso = CuentaIngreso_A.CtaIngreso AND DATEPART(year, AvPgEnc.FechaVenceAvPg) = CuentaIngreso_A.Anio"
    str2 = str2 & "WHERE        (AvPgEnc.AvPgTipoImpuesto = 1) AND (AvPgEnc.AvPgEstado = 1) AND (AvPgEnc.ClaveCatastro = '161723001018') AND (SUBSTRING(AvPgDetalle.CtaIngreso, 1, 6) = '112121')"
    str2 = str2 & "GROUP BY AvPgDetalle.ClaveCatastro"
    
    
    
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open (Str)
    
    Do While Not DeRia.rsAbonadoSP.EOF
        rReportRs.AddNew
        rReportRs!ClaveCatastro = DeRia.rsAbonadoSP!ClaveCatastro
        rReportRs!Identidad = DeRia.rsAbonadoSP!Identidad
        rReportRs!NombreCompleto = Trim(DeRia.rsAbonadoSP!PNombre) & " " & Trim(DeRia.rsAbonadoSP!sNombre) & " " & Trim(DeRia.rsAbonadoSP!PApellido) & " " & Trim(DeRia.rsAbonadoSP!sApellido)
        rReportRs!ValorUnitAvPgDet = IIf(IsNull(DeRia.rsAbonadoSP!ValorUnitAvPgDet), 0, DeRia.rsAbonadoSP!ValorUnitAvPgDet)
        rReportRs!impuesto = IIf(IsNull(DeRia.rsAbonadoSP!impuesto), 0, DeRia.rsAbonadoSP!impuesto)
        'rReportRs!Recargos = IIf(IsNull(DeRia.rsAbonadoSP!Recargos), 0, DeRia.rsAbonadoSP!Recargos)
        'rReportRs!Intereses = IIf(IsNull(DeRia.rsAbonadoSP!Intereses), 0, DeRia.rsAbonadoSP!Intereses)
        
        
        DeRia.rsAbonadoSP.MoveNext
    Loop
End Sub

Private Sub crearRecorset()
    rReportRs.Fields.Append "ClaveCatastro", adBSTR
    rReportRs.Fields.Append "Identidad", adBSTR
    rReportRs.Fields.Append "NombreCompleto", adBSTR
    rReportRs.Fields.Append "ValorUnitAvPgDet", adBSTR
    rReportRs.Fields.Append "impuesto", adBSTR
    'rReportRs.Fields.Append "Recargos", adBSTR
    'rReportRs.Fields.Append "Intereses", adBSTR
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
    oXLSheet.Cells(2, 1) = "REPORTE DE BIENES INMUEBLES CON FICHA CASTRAL URBANA"
    oXLSheet.Cells(5, 1) = "Departamento: " & oAlcaldia.Departamento
    oXLSheet.Cells(6, 1) = "Municipio: " & oAlcaldia.Nombre
    oXLSheet.Cells(7, 1) = "Fecha Elaboración: " & DateToWord(Date)
    oXLSheet.Cells(8, 1) = "Declaraciones emitidas De: " & frmPropConFichasUrb.txtFechaInicio.Text & " Hasta: " & frmPropConFichasUrb.txtFechaFinal.Text
    
    'Center Titles
    Set exRng = oXLSheet.Range("A1:I1"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A2:I2"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A3:I3"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A5:I5"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
    Set exRng = oXLSheet.Range("A6:I6"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
    Set exRng = oXLSheet.Range("A7:I7"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
 

    nRow = 12: nCol = 1
    'Pongamos los titulos de las columnas aqui, para que las agarre el autofit siguiente.
    oXLSheet.Cells(nRow, nCol) = "Clave Catastro"                         'A
    oXLSheet.Cells(nRow, nCol + 1) = "Identidad"        'B
    oXLSheet.Cells(nRow, nCol + 2) = "Nombre Completo"            'C
    oXLSheet.Cells(nRow, nCol + 3) = "Valor por año"      'D
    oXLSheet.Cells(nRow, nCol + 4) = "Impuesto Total"     'E
    oXLSheet.Cells(nRow, nCol + 5) = "Intereses"     'F
    oXLSheet.Cells(nRow, nCol + 6) = "Recargos"        'G
    '------------------------------------------------------------------------------------
  
    'Format titles Bold
    oXLSheet.Range("A1:I12").Font.Bold = True 'Los titulos del reporte
    'set a standard width
    Set exRng = oXLSheet.Columns("A")
    exRng.ColumnWidth = 20
    Set exRng = oXLSheet.Columns("B")
    exRng.ColumnWidth = 20
    Set exRng = oXLSheet.Columns("C")
    exRng.ColumnWidth = 20
    Set exRng = oXLSheet.Columns("D")
    exRng.ColumnWidth = 50
    Set exRng = oXLSheet.Columns("E")
    exRng.ColumnWidth = 15
    Set exRng = oXLSheet.Columns("F")
    exRng.ColumnWidth = 15
    Set exRng = oXLSheet.Columns("G")
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
    Set exRng = oXLSheet.Range("F12:F12")
    ApplyOutsideBorderToExcelCell exRng
    Set exRng = oXLSheet.Range("G12:G12")
    ApplyOutsideBorderToExcelCell exRng

 


    'Centrar los titulos del cuerpo
    Set exRng = oXLSheet.Range("A12:G12")
    exRng.HorizontalAlignment = xlCenter
    
    Dim RsXXX As New ADODB.Recordset
    
    'Define Starting Cells
    nRow = 13: nCol = 1:
    'Potencial Recaudación
    With rReportRs
        .MoveFirst
        Do While .EOF = False

            oXLSheet.Cells(nRow, nCol) = !ClaveCatastro
            oXLSheet.Cells(nRow, nCol + 1) = !Identidad
            oXLSheet.Cells(nRow, nCol + 2) = !NombreCompleto
            oXLSheet.Cells(nRow, nCol + 3) = Format(!ValorUnitAvPgDet, "#,###,##0.00")
            oXLSheet.Cells(nRow, nCol + 4) = Format(!impuesto, "#,###,##0.00")
            
            str2 = " SELECT SUM(AvPgDetalle.ValorUnitAvPgDet) AS Recargos"
            str2 = str2 & " FROM            AvPgEnc INNER JOIN"
            str2 = str2 & " AvPgDetalle ON AvPgEnc.NumAvPg = AvPgDetalle.NumAvPg INNER JOIN"
            str2 = str2 & " CuentaIngreso_A ON AvPgDetalle.CtaIngreso = CuentaIngreso_A.CtaIngreso AND DATEPART(year, AvPgEnc.FechaVenceAvPg) = CuentaIngreso_A.Anio"
            str2 = str2 & " WHERE        (AvPgEnc.AvPgTipoImpuesto = 1) AND (AvPgEnc.AvPgEstado = 1) AND (AvPgEnc.ClaveCatastro = '" & !ClaveCatastro & "') AND (SUBSTRING(AvPgDetalle.CtaIngreso, 1, 6) = '112121')"
            str2 = str2 & " GROUP BY AvPgDetalle.ClaveCatastro"
            
            If DeRia.rsAbonadosSpRubro.State = 1 Then DeRia.rsAbonadosSpRubro.Close
            DeRia.rsAbonadosSpRubro.Open str2
           
            oXLSheet.Cells(nRow, nCol + 5) = Format(DeRia.rsAbonadoSP!Recargos, "#,###,##0.00")
            
            
            str1 = " SELECT        SUM(AvPgDetalle.ValorUnitAvPgDet) AS Intereses"
            str1 = str1 & " FROM            AvPgEnc INNER JOIN"
            str1 = str1 & " AvPgDetalle ON AvPgEnc.NumAvPg = AvPgDetalle.NumAvPg INNER JOIN"
            str1 = str1 & " CuentaIngreso_A ON AvPgDetalle.CtaIngreso = CuentaIngreso_A.CtaIngreso AND DATEPART(year, AvPgEnc.FechaVenceAvPg) = CuentaIngreso_A.Anio"
            str1 = str1 & " WHERE        (AvPgEnc.AvPgTipoImpuesto = 1) AND (AvPgEnc.AvPgEstado = 1) AND (AvPgEnc.ClaveCatastro = '" & !ClaveCatastro & "') AND (SUBSTRING(AvPgDetalle.CtaIngreso, 1, 6) = '112126')"
            str1 = str1 & " GROUP BY AvPgDetalle.ClaveCatastro"
            
            If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
            DeRia.rsAbonadoSP.Open str1
            oXLSheet.Cells(nRow, nCol + 6) = Format(DeRia.rsAbonadoSP!Intereses, "#,###,##0.00")
            MsgBox ("AQUI 11")
            .MoveNext
            nRow = nRow + 1
        Loop
    End With
    oXLSheet.Range("A" & nRow + 1 & ":" & "G" & nRow + 1).Font.Bold = True  'NEGRITA
    Set exRng = oXLSheet.Range("A" & nRow + 1 & ":" & "B" & nRow + 1): exRng.HorizontalAlignment = xlCenter: exRng.Merge ' MARGEN
    nRow = nRow + 1
    'Imprime el periodo del reporte
    nRow = nRow + 1
    'oXLSheet.Cells(nRow, nCol) = strPeriodo
    Set exRng = oXLSheet.Range("A" & nRow, "G" & nRow): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    
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
