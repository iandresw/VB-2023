VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rpAvisoDeCobro 
   Caption         =   "Proyecto1 - rpAvisoDeCobro (ActiveReport)"
   ClientHeight    =   9825
   ClientLeft      =   840
   ClientTop       =   1035
   ClientWidth     =   16845
   _ExtentX        =   29713
   _ExtentY        =   17330
   SectionData     =   "rpAvisoDeCobro.dsx":0000
End
Attribute VB_Name = "rpAvisoDeCobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rReportRs As New ADODB.Recordset
Dim rsCont As New ADODB.Recordset
Dim GranTotal As Currency

Public Function DatosGenerales(Identidad As String)
    lbMunicipio = ""
    Dim sql As String
    sql = "SELECT NombreMuni FROM Parametro"
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open sql
    lbMunicipio = Trim(DeRia.rsAbonadoSP!NombreMuni)
    

    txtMontAdeuCalHastMes = Format(Now, "mmmm yyyy")
    strSql = "SELECT identidad, Pnombre, SNombre, PApellido, SApellido, direccion FROM Contribuyente WHERE (Identidad = '" & Identidad & "')"
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open strSql
    Me.txtIdentidad = Trim(DeRia.rsAbonadoSP!Identidad)
    txtMontAdeuNombreCont = Trim(DeRia.rsAbonadoSP!Pnombre) & " " & Trim(DeRia.rsAbonadoSP!sNombre) & " " & Trim(DeRia.rsAbonadoSP!PApellido) & " " & Trim(DeRia.rsAbonadoSP!sApellido)
    txtMontAdeuDireccion = Trim(DeRia.rsAbonadoSP!Direccion)
    txtFechaAviso = Format(Now, "dddd, dd  mmmm  yyyy")
    
    Dim resultado As String
    ' Realizar la consulta de SQL
    sql = "SELECT  ASPE_Seq FROM AbonadoSPEnc WHERE  (Identidad = '" & Me.txtIdentidad & "')AND (ASPE_Estado = 0)"
    Set rsCont = DeRia.CoRia.Execute(sql)
    ' Verificar si hay más de dos registros en la columna
    If Not rsCont.EOF Then
        resultado = rsCont.Fields("ASPE_Seq").Value
        rsCont.MoveNext
        Do While Not rsCont.EOF
            resultado = resultado & ", " & rsCont.Fields("ASPE_Seq").Value
            rsCont.MoveNext
        Loop
    End If
    rsCont.Close
    ' Asignar el resultado al TextBox
    Me.txtNoAbonado.Text = resultado
End Function

Public Function Interes(ByVal Identidad As String, ByVal anio As String, ByVal TipoCuenta As String, ByVal ClaveCatastro As String) As Double
    Dim sql As String
    Dim Cantidad As String
    Dim cuenta2 As String
    Dim tipoInteres As String
    cuenta2 = Left(TipoCuenta, 6)
    If TipoCuenta = "1" Then ' cuenta recuperacion impuestos
        tipoInteres = "11212601"
    ElseIf TipoCuenta = "5" Then ' cuenta recuperacion servicios
        tipoInteres = "11212602"
    'ElseIf cuenta2 = "111110" Then ' cuenta bienes inmuebles
        'tipoInteres = "11212601"
    'ElseIf cuenta2 = "111118" Then ' cuenta servicios Publicos
        'tipoInteres = "11212602"
    End If
    sql = " SELECT sum(AvPgDetalle.ValorUnitAvPgDet) as total , ctaingreso FROM AvPgDetalle INNER JOIN AvPgEnc ON AvPgDetalle.NumAvPg = AvPgEnc.NumAvPg"
    sql = sql & " WHERE (AvPgEnc.AvPgEstado = 1) AND (AvPgEnc.Identidad = '" & Identidad & "') AND substring(AvPgDetalle.CtaIngreso,1,6) = ('112126') AND (DATEPART(year, AvPgEnc.FechaVenceAvPg) = " & anio & " )  and ctaingreso = '" & tipoInteres & "' AND (AvPgEnc.ClaveCatastro = '" & ClaveCatastro & "')"
    sql = sql & " group by AvPgDetalle.CtaIngreso"
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open sql
    If Not DeRia.rsAbonadoSP.EOF Then
    Cantidad = IIf(IsNull(DeRia.rsAbonadoSP!Total), 0#, DeRia.rsAbonadoSP!Total)
    Interes = Cantidad
    Else
    Cantidad = 0
    Interes = Cantidad
    End If
End Function
Public Function Recargos(ByVal Identidad As String, ByVal anio As String, ByVal TipoCuenta As String, ByVal ClaveCatastro As String) As Double
    Dim sql As String
    Dim Cantidad As String
    Dim cuenta2 As String
    Dim tipoRecargo As String
    cuenta2 = Left(TipoCuenta, 6)
    If TipoCuenta = "1" Then  ' cuenta recuperacion impuestos
        tipoRecargo = "11212101"
    ElseIf TipoCuenta = "5" Then ' cuenta recuperacion servicios
        tipoRecargo = "11212102"
    'ElseIf cuenta = "111110" Then ' cuenta bienes inmuebles
        'tipoRecargo = "11212101"
    'ElseIf cuenta = "111118" Then ' cuenta servicios Publicos
        'tipoRecargo = "11212102"
    End If
    sql = " SELECT sum(AvPgDetalle.ValorUnitAvPgDet) as total , ctaingreso FROM AvPgDetalle INNER JOIN AvPgEnc ON AvPgDetalle.NumAvPg = AvPgEnc.NumAvPg"
    sql = sql & " WHERE (AvPgEnc.AvPgEstado = 1) AND (AvPgEnc.Identidad = '" & Identidad & "') AND substring(AvPgDetalle.CtaIngreso,1,6) = ('112121') AND (DATEPART(year, AvPgEnc.FechaVenceAvPg) = " & anio & " ) and ctaingreso = '" & tipoRecargo & "'AND (AvPgEnc.ClaveCatastro = '" & ClaveCatastro & "')"
    sql = sql & " group by AvPgDetalle.CtaIngreso"
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open sql
    If Not DeRia.rsAbonadoSP.EOF Then
    Cantidad = IIf(IsNull(DeRia.rsAbonadoSP!Total), 0#, DeRia.rsAbonadoSP!Total)
    Recargos = Cantidad
    Else
    Cantidad = 0
    Recargos = Cantidad
    End If
End Function
Public Function Multa(ByVal Identidad As String, ByVal anio As String) As Double
    Dim sql As String
    Dim Cantidad As Double
    sql = " SELECT sum(AvPgDetalle.ValorUnitAvPgDet) as total , ctaingreso FROM AvPgDetalle INNER JOIN AvPgEnc ON AvPgDetalle.NumAvPg = AvPgEnc.NumAvPg"
    sql = sql & " WHERE (AvPgEnc.AvPgEstado = 1) AND (AvPgEnc.Identidad = '" & Identidad & "') AND substring(AvPgDetalle.CtaIngreso,1,6) = ('112120') AND (DATEPART(year, AvPgEnc.FechaVenceAvPg) = " & anio & " ) "
    sql = sql & " group by AvPgDetalle.CtaIngreso"
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open sql
    If Not DeRia.rsAbonadoSP.EOF Then
    Cantidad = IIf(IsNull(DeRia.rsAbonadoSP!Total), 0#, DeRia.rsAbonadoSP!Total)
    Multa = Cantidad
    Else
    Cantidad = 0
    End If
End Function

Public Function Descuentos(ByVal Identidad As String, ByVal anio As String, ByVal ClaveCatastro As String, ByVal TipoCuenta As String) As Double
    Dim sql As String
    Dim Cantidad As Double
    sql = " SELECT sum(AvPgDetalle.ValorUnitAvPgDet) as total FROM AvPgDetalle INNER JOIN AvPgEnc ON AvPgDetalle.NumAvPg = AvPgEnc.NumAvPg"
    sql = sql & " WHERE (AvPgEnc.AvPgEstado = 1) AND (AvPgEnc.Identidad = '" & Identidad & "') AND substring(AvPgDetalle.CtaIngreso,1,6) = ('112127') "
    sql = sql & " AND (DATEPART(year, AvPgEnc.FechaVenceAvPg) = " & anio & " ) AND (AvPgEnc.ClaveCatastro = '" & ClaveCatastro & "') AND  AvPgEnc.AvPgTipoImpuesto = " & TipoCuenta
   ' sql = sql & " group by "
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open sql
    If Not DeRia.rsAbonadoSP.EOF Then
    Cantidad = IIf(IsNull(DeRia.rsAbonadoSP!Total), 0#, DeRia.rsAbonadoSP!Total)
    Descuentos = Cantidad
    Else
    Cantidad = 0
    Descuentos = Cantidad
    End If
End Function

Private Sub Detail_Format()
    Dim TipoCuenta As String
    Dim anio As String
    Dim ClaveCatastro As String
    Static X As Integer
    X = X + 1
    If X > rsCont.RecordCount Then Exit Sub
        TipoCuenta = rsCont!AvPgTipoImpuesto
        anio = rsCont!aniox
        ClaveCatastro = rsCont!ClaveCatastro
        Me.txtMontAdeuConcepto = rsCont!AvPgDescripcion
        Me.txtMontAdeuAnioImpositivo = rsCont!aniox
        Me.txtMontAdeuMonto = Format(rsCont!Total, "#,###,##0.00")
        Me.txtMontDescu = Format(Descuentos(frmAvisoCobro.txtIdentidad, anio, ClaveCatastro, TipoCuenta), "#,###,##0.00")
        Me.txtMontAdeuMulta = Format(Multa(frmAvisoCobro.txtIdentidad, anio), "#,###,##0.00")
        Me.txtMontAdeuIntereses = Format(Interes(frmAvisoCobro.txtIdentidad, anio, TipoCuenta, ClaveCatastro), "#,###,##0.00")
        Me.txtMontAdeuRecargos = Format(Recargos(frmAvisoCobro.txtIdentidad, anio, TipoCuenta, ClaveCatastro), "#,###,##0.00")
        Me.txtMontAdeuValTotal = Format(CDbl(txtMontDescu.Text) + (CDbl(txtMontAdeuMonto.Text) + CDbl(txtMontAdeuIntereses.Text) + CDbl(txtMontAdeuRecargos.Text) + CDbl(txtMontAdeuMulta.Text)), "#,###,##0.00")
        GranTotal = GranTotal + Me.txtMontAdeuValTotal.Text
    rsCont.MoveNext
    txtMontoAdeudadoTotal = Format(GranTotal, "#,###,##0.00")
    Detail.PrintSection
End Sub

Private Sub ActiveReport_ReportStart()
    Dim sql As String
    DatosGenerales (frmAvisoCobro.txtIdentidad)
    sql = " SELECT AvPgEnc.AvPgDescripcion, DATEPART(year, AvPgEnc.FechaVenceAvPg) AS aniox, SUM(AvPgDetalle.ValorUnitAvPgDet) AS total, AvPgEnc.AvPgTipoImpuesto, AvPgEnc.ClaveCatastro"
    sql = sql & " FROM AvPgDetalle INNER JOIN"
    sql = sql & " AvPgEnc ON AvPgDetalle.NumAvPg = AvPgEnc.NumAvPg INNER JOIN"
    sql = sql & " CuentaIngreso_A ON AvPgDetalle.CtaIngreso = CuentaIngreso_A.CtaIngreso AND DATEPART(year, AvPgEnc.FechaVenceAvPg) = CuentaIngreso_A.Anio"
    sql = sql & " WHERE (AvPgEnc.AvPgEstado = 1) AND (AvPgEnc.Identidad = '" & frmAvisoCobro.txtIdentidad & "') AND (AvPgEnc.AvPgTipoImpuesto IN (5,1)) AND (CuentaIngreso_A.Tipo <> 2)"
    sql = sql & " GROUP BY AvPgEnc.AvPgDescripcion, DATEPART(year, AvPgEnc.FechaVenceAvPg), AvPgEnc.AvPgTipoImpuesto, AvPgEnc.ClaveCatastro"
    sql = sql & " ORDER BY  ANIOX"
    Set rsCont = DeRia.CoRia.Execute(sql)
    If rsCont.EOF Then
        MsgBox "No tiene facturas pendientes"
    End If
End Sub
 

