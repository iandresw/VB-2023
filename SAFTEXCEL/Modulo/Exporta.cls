VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Exporta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private strPeriodo As String
Private rReportRs As New ADODB.Recordset
Private rReportRs2 As New ADODB.Recordset
Private Anio1, Anio2, Anio3, Anio4, Anio5 As Integer
'Public Sub CrearReporte(strPrograma As String, dFecha1 As Date, dFecha2 As Date)
Public Sub CrearReporte(dFecha1 As Date, dFecha2 As Date)


    Dim sSql, sSql2 As String
    Dim rsPar, rsRep As New Recordset
    Dim rsAnios As New ADODB.Recordset
    Dim VLen As Integer
    CrearRecordSet
    'Calcule el periodo
    strPeriodo = GetMonth(dFecha1) & " a " & GetMonth(dFecha2) & " del " & Year(dFecha1)
    '------------------
    'Verifiquemos los parametros
    Set rsPar = DeRia.CoRia.Execute("select * from Parametro")
    If rsPar.RecordCount = 0 Then Exit Sub
    
     
  

      sSql = " SELECT  *  From Contribuyente "

  
  
    Set rsRep = New Recordset
    Set rsRep = DeRia.CoRia.Execute(sSql)
    
    Do While rsRep.EOF = False
    
      
            rReportRs.AddNew


                rReportRs!Identidad = rsRep!Identidad
                rReportRs!Nombre = Trim(rsRep!PNombre) & " " & Trim(rsRep!sNombre) & " " & Trim(rsRep!pApellido) & " " & Trim(rsRep!sApellido)
                rReportRs!UltPeriodo = rsRep!UltPeriodoFact
        
        rsRep.MoveNext
       Loop
     

End Sub
Private Sub CrearRecordSet()
    rReportRs.Fields.Append "Identidad", adBSTR
    rReportRs.Fields.Append "Nombre", adBSTR
    rReportRs.Fields.Append "UltPeriodo", adInteger
    'rReportRs.Fields.Append "BI", adCurrency
    'rReportRs.Fields.Append "IC", adCurrency
    'rReportRs.Fields.Append "ATL", adCurrency
    
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
    oXLSheet.Cells(2, 1) = "MORA DE IMPUESTOS Y SERVICIOS PUBLICOS POR FECHAS"
    
'    oXLSheet.Cells(4, 1) = "(Cantidad en Lempiras)"
    
    oXLSheet.Cells(6, 1) = "Departamento: " & oAlcaldia.Departamento
    oXLSheet.Cells(6, 3) = "De: " & FrmReporteMo.txtFecha1.Text & " Hasta: " & FrmReporteMo.txtFecha2.Text
    oXLSheet.Cells(8, 1) = "Fecha Elaboración: " & DateToWord(Date)
    
    'Center Titles
    Set exRng = oXLSheet.Range("A1:P1"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A2:P2"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A3:P3"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    Set exRng = oXLSheet.Range("A4:P4"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
   
    
    nRow = 12: nCol = 1
    'Pongamos los titulos de las columnas aqui, para que las agarre el autofit siguiente.
    oXLSheet.Cells(nRow, nCol) = "Identidad"                            'A
    oXLSheet.Cells(nRow, nCol + 1) = "Nombre Contribuyente"                    'B 1
    oXLSheet.Cells(nRow, nCol + 2) = "Ultimo Periodo Facturado"                            'C 2
  '  oXLSheet.Cells(nRow, nCol + 3) = "IMPUESTO SOBRE INDUSTRIA Y COMERCIO"
 '   oXLSheet.Cells(nRow, nCol + 4) = "Agua, Tren de Aseo y Limpieza de Cementerio"
  '  oXLSheet.Cells(nRow, nCol + 5) = "TOTAL IMPUESTO A PAGAR"
    

    
    
    '------------------------------------------------------------------------------------
  
    'Format titles Bold
    oXLSheet.Range("A1:F12").Font.Bold = True 'Los titulos del reporte
    
    'set a standard width
    Set exRng = oXLSheet.Columns("A")
    exRng.ColumnWidth = 17
    Set exRng = oXLSheet.Columns("B")
    exRng.ColumnWidth = 44
    Set exRng = oXLSheet.Columns("C")
    exRng.ColumnWidth = 14
    Set exRng = oXLSheet.Columns("D")
    exRng.ColumnWidth = 14
    
    Set exRng = oXLSheet.Columns("E")
    exRng.ColumnWidth = 14
    
    Set exRng = oXLSheet.Columns("F")
    exRng.ColumnWidth = 14
    
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
   ' Set exRng = oXLSheet.Range("Q10:R11"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
   ' Set exRng = oXLSheet.Range("S10:T11"): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    
    'Define Starting Cells
    nRow = 14: nCol = 1:
    'Potencial Recaudación
    
    With rReportRs
        .MoveFirst
        Do While .EOF = False
           oXLSheet.Cells(nRow, nCol) = !Identidad
           oXLSheet.Cells(nRow, nCol + 1) = !Nombre
           
            oXLSheet.Cells(nRow, nCol + 2) = !UltPeriodo
            
          '  oXLSheet.Cells(nRow, nCol + 3) = !IC
           ' oXLSheet.Cells(nRow, nCol + 4) = !ATL
            
            oXLSheet.Cells(nRow, nCol + 5) = "=C" & nRow & "+D" & nRow & "+E" & nRow & ""

            .MoveNext
            nRow = nRow + 1
        Loop
    End With
    

    
    'Calcule los totales
    nCol = 1
    oXLSheet.Cells(nRow + 1, nCol) = "T o t a l"
    
    oXLSheet.Cells(nRow + 1, nCol + 2) = "=sum(C13:C" & nRow - 1 & ")"
    oXLSheet.Cells(nRow + 1, nCol + 3) = "=sum(D13:D" & nRow - 1 & ")"
    oXLSheet.Cells(nRow + 1, nCol + 4) = "=sum(E13:E" & nRow - 1 & ")"
    oXLSheet.Cells(nRow + 1, nCol + 5) = "=sum(F13:F" & nRow - 1 & ")"


     
     oXLSheet.Range("A" & nRow + 1 & ":" & "E" & nRow + 1).Font.Bold = True  'NEGRITA
     Set exRng = oXLSheet.Range("A" & nRow + 1 & ":" & "B" & nRow + 1): exRng.HorizontalAlignment = xlCenter: exRng.Merge ' MARGEN
     
    nRow = nRow + 1
    'Imprime el periodo del reporte
    nRow = nRow + 1
    'oXLSheet.Cells(nRow, nCol) = strPeriodo
    Set exRng = oXLSheet.Range("A" & nRow, "E" & nRow): exRng.HorizontalAlignment = xlCenter: exRng.Merge
    '---------------------------------------------------

    'Formateamos numeros
    Set exRng = oXLSheet.Columns("C:F")
    exRng.NumberFormat = "#,###,###,##0.00"
    'Formateamos texto
    'Set exRng = oXLSheet.Columns("M:N")
    Set exRng = oXLSheet.Range("E11:P11")
    exRng.NumberFormat = "@"
    
    'Autofit Columns
    'Set exRng = oXLSheet.Range("A11:B12")
    'exRng.EntireColumn.AutoFit
        
    'Dibuje lineas en los titulos
    'Set exRng = oXLSheet.Range("A10:A11")
    'ApplyOutsideBorderToExcelCell exRng
    
    'Poner lineas en los datos
    Set exRng = oXLSheet.Range("A11:F" & nRow)
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




'  sSql = "SELECT AvPgEnc.Identidad FROM ParamRptMoraSPTipo INNER JOIN ParamRptMoraSP ON ParamRptMoraSPTipo.IdTipo = ParamRptMoraSP.CodTipoServicio INNER JOIN "
'  sSql = sSql & " AvPgEnc INNER JOIN AvPgDetalle ON AvPgEnc.NumAvPg = AvPgDetalle.NumAvPg ON ParamRptMoraSP.CtaIngreso = AvPgDetalle.CtaIngreso "
'  sSql = sSql & " WHERE (AvPgEnc.AvPgEstado = 1) "
'  sSql = sSql & " AND ( AvPgEnc.FechaVenceAvPg BETWEEN '" & dFecha1 & "' AND '" & dFecha2 & "') AND (ParamRptMoraSPTipo.CodModulo = 2)"
'  sSql = sSql & " GROUP BY AvPgEnc.Identidad "
  

