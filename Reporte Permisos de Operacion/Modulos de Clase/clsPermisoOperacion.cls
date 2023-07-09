VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPermisoOperacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private rReportRs As New ADODB.Recordset
Private rsAbonadoSP As New ADODB.Recordset
Public Sub CrearReporte(txtFecha1 As String, txtFecha2 As String, AptoRen As String, sqlAldea As String)
    Dim StrPermisoOp As String
    crearRecorset
    
    StrPermisoOp = " SELECT  CONVERT(VARCHAR(20), Tra_PermOP.Fecha,103) As Fecha, Tra_PermOP.Propietario, Tra_PermOP.Negocio, Tra_PermOP.Actividad,   Tra_PermOP.Ubicacion, Tra_PermOP.NoPermiso, Tra_PermOP.Observacion, Tra_PermOP.NumRecibo "
    'StrPermisoOp = StrPermisoOp & " sum(AvPgDetalle.ValorUnitAvPgDet ) as Total"
    StrPermisoOp = StrPermisoOp & " FROM AvPgEnc INNER JOIN "
    StrPermisoOp = StrPermisoOp & " AvPgDetalle ON AvPgEnc.NumAvPg = AvPgDetalle.NumAvPg INNER JOIN  "
    StrPermisoOp = StrPermisoOp & " Tra_PermOP ON AvPgEnc.Identidad = Tra_PermOP.Identidad "
    StrPermisoOp = StrPermisoOp & " WHERE (Tra_PermOP.Fecha BETWEEN '" & txtFecha1 & "'AND '" & txtFecha2 & "') AND (AvPgEnc.AvPgTipoImpuesto = 3) AND (SUBSTRING(AvPgDetalle.CtaIngreso, 4, 5) IN ('11921')) " & AptoRen & sqlAldea
    StrPermisoOp = StrPermisoOp & " GROUP BY Tra_PermOP.Fecha, Tra_PermOP.Propietario, Tra_PermOP.Actividad, Tra_PermOP.Negocio, Tra_PermOP.Ubicacion, Tra_PermOP.NoPermiso, Tra_PermOP.Observacion, Tra_PermOP.NumRecibo "
    StrPermisoOp = StrPermisoOp & " ORDER BY Tra_PermOP.Fecha "
    
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open (StrPermisoOp)
    
    Do While Not DeRia.rsAbonadoSP.EOF
        rReportRs.AddNew
        rReportRs!Fecha = DeRia.rsAbonadoSP!Fecha
        rReportRs!Propietario = DeRia.rsAbonadoSP!Propietario
        rReportRs!Negocio = DeRia.rsAbonadoSP!Negocio
        rReportRs!Actividad = DeRia.rsAbonadoSP!Actividad
        rReportRs!Ubicacion = DeRia.rsAbonadoSP!Ubicacion
        rReportRs!NoPermiso = DeRia.rsAbonadoSP!NoPermiso
        rReportRs!NumRecibo = DeRia.rsAbonadoSP!NumRecibo
        rReportRs!Observacion = DeRia.rsAbonadoSP!Observacion
        'rReportRs!Total = IIf(IsNull(DeRia.rsAbonadoSP!Total), 0, DeRia.rsAbonadoSP!Total)
        DeRia.rsAbonadoSP.MoveNext
    Loop
End Sub

Private Sub crearRecorset()
    rReportRs.Fields.Append "Fecha", adBSTR
    rReportRs.Fields.Append "Propietario", adBSTR
    rReportRs.Fields.Append "Negocio", adBSTR
    rReportRs.Fields.Append "Actividad", adBSTR
    rReportRs.Fields.Append "Ubicacion", adBSTR
    rReportRs.Fields.Append "NoPermiso", adBSTR
    rReportRs.Fields.Append "NumRecibo", adBSTR
    rReportRs.Fields.Append "Observacion", adBSTR
    'rReportRs.Fields.Append "Total", adBSTR
    rReportRs.Open
End Sub

Public Sub SendToExcel()
    If rReportRs.RecordCount = 0 Then: MsgBox ("No Hay datos para mostrar"): Exit Sub
    Dim oXLApp As Excel.Application         'Declare the object variables
    Dim oXLBook As Excel.Workbook
    Dim oXLSheet As Excel.Worksheet
    Dim oAlcaldia As New Alcaldia
    Dim nRow, nCol As Long
    'Dim oAlcaldia As New Alcaldia
    Dim exRng As Excel.Range
    Dim exRngData As Excel.Range
    Set oXLApp = New Excel.Application    'Create a new instance of Excel
    Set oXLBook = oXLApp.Workbooks.Add    'Add a new workbook
    Set oXLSheet = oXLBook.Worksheets(1)  'Work with the first
    Dim strCodigo As String
    
    oXLSheet.Cells(1, 1) = oAlcaldia.Nombre
    oXLSheet.Cells(2, 1) = "PERMISOS DE OPERACION EMITIDOS POR FECHA"
    oXLSheet.Cells(5, 1) = "Departamento: " & oAlcaldia.Departamento
    oXLSheet.Cells(6, 1) = "Municipio: " & oAlcaldia.Nombre
    oXLSheet.Cells(7, 1) = "Fecha Elaboración: " & Now()
    oXLSheet.Cells(8, 1) = "Permisos de operacion emitidos desde: " & frmRpPermisosOp.txtFecha1.Text & " hasta " & frmRpPermisosOp.txtFecha2.Text
    
    'Center Titles
    Set exRng = oXLSheet.Range("A1:H1"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A2:H2"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A3:H3"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A5:H5"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
    Set exRng = oXLSheet.Range("A6:H6"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
    Set exRng = oXLSheet.Range("A7:H7"): exRng.HorizontalAlignment = xlLeft: exRng.Merge
 
    
    
    
    nRow = 12: nCol = 1
    'Pongamos los titulos de las columnas aqui, para que las agarre el autofit siguiente.
    oXLSheet.Cells(nRow, nCol) = "Fecha de Emicion"            'A
    oXLSheet.Cells(nRow, nCol + 1) = "Propietario"             'B
    oXLSheet.Cells(nRow, nCol + 2) = "Negocio"                 'C
    oXLSheet.Cells(nRow, nCol + 3) = "Actividad"               'D
    oXLSheet.Cells(nRow, nCol + 4) = "Ubicacion"               'E
    oXLSheet.Cells(nRow, nCol + 5) = "No. Permiso"               'F
    oXLSheet.Cells(nRow, nCol + 6) = "No. Recibo"               'F
    oXLSheet.Cells(nRow, nCol + 7) = "Observacion"             'G
   'oXLSheet.Cells(nRow, nCol + 8) = "Valor"                   'H
    '------------------------------------------------------------------------------------
  
   'Format titles Bold
    oXLSheet.Range("A1:H12").Font.Bold = True 'Los titulos del reporte
    
    'set a standard width
    Set exRng = oXLSheet.Columns("A")
    exRng.ColumnWidth = 17
    Set exRng = oXLSheet.Columns("B")
    exRng.ColumnWidth = 50
    Set exRng = oXLSheet.Columns("C")
    exRng.ColumnWidth = 65
    Set exRng = oXLSheet.Columns("D")
    exRng.ColumnWidth = 65
    Set exRng = oXLSheet.Columns("E")
    exRng.ColumnWidth = 30
    Set exRng = oXLSheet.Columns("F")
    exRng.ColumnWidth = 20
    Set exRng = oXLSheet.Columns("G")
    exRng.ColumnWidth = 13
    Set exRng = oXLSheet.Columns("H")
    exRng.ColumnWidth = 13
    'Set exRng = oXLSheet.Columns("I")
    'exRng.ColumnWidth = 13
    
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
    'Set exRng = oXLSheet.Range("I12:I12")
    'ApplyOutsideBorderToExcelCell exRng



    'Centrar los titulos del cuerpo
    Set exRng = oXLSheet.Range("A12:H12")
    exRng.HorizontalAlignment = xlCenter
    


    
    'Define Starting Cells
    nRow = 13: nCol = 1:
    'Potencial Recaudación
  
    With rReportRs
        .MoveFirst
        Do While .EOF = False
            oXLSheet.Cells(nRow, nCol) = !Fecha
            oXLSheet.Cells(nRow, nCol + 1) = !Propietario
            oXLSheet.Cells(nRow, nCol + 2) = !Negocio
            oXLSheet.Cells(nRow, nCol + 3) = !Actividad
            oXLSheet.Cells(nRow, nCol + 4) = !Ubicacion
            oXLSheet.Cells(nRow, nCol + 5) = Format(!NoPermiso, "######0")
            oXLSheet.Cells(nRow, nCol + 6) = Format(!NumRecibo, "######0")
            oXLSheet.Cells(nRow, nCol + 7) = !Observacion
            'oXLSheet.Cells(nRow, nCol + 8) = Format(!Total, "#,###,##0.00")
            .MoveNext
            nRow = nRow + 1
        Loop
    End With
    oXLSheet.Range("A" & nRow + 1 & ":" & "H" & nRow + 1).Font.Bold = True  'NEGRITA
    Set exRng = oXLSheet.Range("A" & nRow + 1 & ":" & "B" & nRow + 1): exRng.HorizontalAlignment = xlCenter: exRng.Merge ' MARGEN
    nRow = nRow + 1
    'Imprime el periodo del reporte
    nRow = nRow + 1
    'oXLSheet.Cells(nRow, nCol) = strPeriodo
    Set exRng = oXLSheet.Range("A" & nRow, "H" & nRow): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    
    'Poner lineas en los datos
    Set exRng = oXLSheet.Range("A12:H" & nRow)
    ApplyBorderToExcelCell exRng
    


    'Cambiar el fondo de el titulo
    Set exRng = oXLSheet.Range("A12:H12")
    ApplyBackColorToExcelCell exRng
    
    '---------------------------------------------------
    oXLApp.Visible = True
    '---------------------------------------------------
    
    'Disconnect from Excel (let the user take over)
    Set oXLApp = Nothing
    '----------------------------------------------

End Sub
    
    
