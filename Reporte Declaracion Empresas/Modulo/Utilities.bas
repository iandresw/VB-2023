Attribute VB_Name = "Utilities"
Public Sub ApplyBorderToExcelCell(oRange As Excel.Range)

    oRange.Borders(xlDiagonalDown).LineStyle = xlNone
    oRange.Borders(xlDiagonalUp).LineStyle = xlNone
    With oRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With oRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With oRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With oRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With oRange.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With oRange.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

End Sub
Public Sub ApplyOutsideBorderToExcelCell(oRange As Excel.Range)

    oRange.Borders(xlDiagonalDown).LineStyle = xlNone
    oRange.Borders(xlDiagonalUp).LineStyle = xlNone
    With oRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With oRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With oRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With oRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    oRange.Borders(xlInsideVertical).LineStyle = xlNone
    oRange.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub
Public Sub ApplyBackColorToExcelCell(oRange As Excel.Range)
    With oRange.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
End Sub

Public Sub ExportRsToExcel(rs As ADODB.Recordset)
    'Dim oXLApp As Excel.Application         'Declare the object variables
    'Dim oXLBook As Excel.Workbook
    'Dim oXLSheet As Excel.Worksheet
    
    'Set oXLApp = New Excel.Application    'Create a new instance of Excel
    'Set oXLBook = oXLApp.Workbooks.Add    'Add a new workbook
    'Set oXLSheet = oXLBook.Worksheets(1)  'Work with the first
    
    'create and fill a recordset here, called oRecordset
    'oXLSheet.Range("B15").CopyFromRecordset rs
    'oXLApp.Visible = True
End Sub
Public Function ObtenerAlcaldia() As ADODB.Recordset
    Set ObtenerAlcaldia = DeRia.CoRia.Execute("select * from Parametro")
    
End Function
