VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rpAvisoDeCobro 
   Caption         =   "Proyecto1 - rpAvisoDeCobro (ActiveReport)"
   ClientHeight    =   10470
   ClientLeft      =   750
   ClientTop       =   1065
   ClientWidth     =   16545
   _ExtentX        =   29184
   _ExtentY        =   18468
   SectionData     =   "rpAvisoDeCobro44.dsx":0000
End
Attribute VB_Name = "rpAvisoDeCobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rReportRs As New ADODB.Recordset
Dim rsCont As New ADODB.Recordset
Dim GranTotal As Currency
Dim RsDatos As New ADODB.Recordset

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
    sql = "SELECT  ClaveCatastro FROM catastro WHERE  (Identidad = '" & Me.txtIdentidad & "')"
    Set RsDatos = DeRia.CoRia.Execute(sql)
    ' Verificar si hay más de dos registros en la columna
    If Not RsDatos.EOF Then
        resultado = RsDatos.Fields("ClaveCatastro").Value
        RsDatos.MoveNext
        Do While Not RsDatos.EOF
            resultado = resultado & ", " & RsDatos.Fields("ClaveCatastro").Value
            RsDatos.MoveNext
        Loop
    End If
    RsDatos.Close
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
    Dim rsDetalle As New ADODB.Recordset
    X = X + 1
    GranTotal = 0
        If X > rsCont.RecordCount Then Exit Sub
            DatosGenerales (rsCont!Identidad)
        
            sql = " SELECT AvPgEnc.AvPgDescripcion, DATEPART(year, AvPgEnc.FechaVenceAvPg) AS aniox, SUM(AvPgDetalle.ValorUnitAvPgDet) AS total, AvPgEnc.AvPgTipoImpuesto, AvPgEnc.ClaveCatastro"
            sql = sql & " FROM AvPgDetalle INNER JOIN"
            sql = sql & " AvPgEnc ON AvPgDetalle.NumAvPg = AvPgEnc.NumAvPg INNER JOIN"
            sql = sql & " CuentaIngreso_A ON AvPgDetalle.CtaIngreso = CuentaIngreso_A.CtaIngreso AND DATEPART(year, AvPgEnc.FechaVenceAvPg) = CuentaIngreso_A.Anio"
            sql = sql & " WHERE (AvPgEnc.AvPgEstado = 1) AND (AvPgEnc.Identidad = '" & rsCont!Identidad & "') AND (AvPgEnc.AvPgTipoImpuesto IN (5,1)) AND (CuentaIngreso_A.Tipo <> 2)"
            sql = sql & " GROUP BY AvPgEnc.AvPgDescripcion, DATEPART(year, AvPgEnc.FechaVenceAvPg), AvPgEnc.AvPgTipoImpuesto, AvPgEnc.ClaveCatastro"
            sql = sql & " ORDER BY  ANIOX"
            
            
            sql = "  SELECT AvPgEnc.AvPgDescripcion, DATEPART(year, MIN(AvPgEnc.FechaVenceAvPg)) AS aniox, DATEPART(year, MAX(AvPgEnc.FechaVenceAvPg)) AS anioMAx, SUM(AvPgDetalle.ValorUnitAvPgDet) AS total, AvPgEnc.AvPgTipoImpuesto, "
            sql = sql & " AvPgEnc.ClaveCatastro FROM AvPgDetalle INNER JOIN AvPgEnc ON AvPgDetalle.NumAvPg = AvPgEnc.NumAvPg INNER JOIN "
            sql = sql & " CuentaIngreso_A ON AvPgDetalle.CtaIngreso = CuentaIngreso_A.CtaIngreso AND DATEPART(year, AvPgEnc.FechaVenceAvPg) = CuentaIngreso_A.Anio "
            sql = sql & " WHERE (AvPgEnc.AvPgEstado = 1) AND AvPgEnc.AvPgTipoImpuesto = 1 AND (AvPgEnc.ClaveCatastro = '" & rsCont!ClaveCatastro & "') "
            sql = sql & " GROUP BY AvPgEnc.AvPgDescripcion, AvPgEnc.AvPgTipoImpuesto, AvPgEnc.ClaveCatastro ORDER BY ANIOX "
            'AND (CuentaIngreso_A.Tipo <> 2)
            
            Set rsDetalle = DeRia.CoRia.Execute(sql)
         '  Do While Not rsDetalle.EOF
            TipoCuenta = rsDetalle!AvPgTipoImpuesto
            anio = rsDetalle!ANIOX
            ClaveCatastro = rsDetalle!ClaveCatastro
            Me.txtMontAdeuConcepto = rsDetalle!AvPgDescripcion
            If rsDetalle!ANIOX = rsDetalle!anioMAx Then
               Me.txtMontAdeuAnioImpositivo = rsDetalle!anioMAx
            Else
               Me.txtMontAdeuAnioImpositivo = "Del: " & rsDetalle!ANIOX & " Al: " & rsDetalle!anioMAx
            End If
            Me.txtMontAdeuValTotal = Format(rsDetalle!Total, "#,###,##0.00")
            'Me.txtMontDescu = Format(Descuentos(rsCont!Identidad, anio, ClaveCatastro, TipoCuenta), "#,###,##0.00")
            'Me.txtMontAdeuMulta = Format(Multa(rsCont!Identidad, anio), "#,###,##0.00")
            'Me.txtMontAdeuIntereses = Format(Interes(rsCont!Identidad, anio, TipoCuenta, ClaveCatastro), "#,###,##0.00")
            'Me.txtMontAdeuRecargos = Format(Recargos(rsCont!Identidad, anio, TipoCuenta, ClaveCatastro), "#,###,##0.00")
            'Me.txtMontAdeuValTotal = Format(CDbl(Val(txtMontDescu.Text)) + (CDbl(txtMontAdeuMonto.Text) + CDbl(Val(txtMontAdeuIntereses.Text)) + CDbl(Val(txtMontAdeuRecargos.Text)) + CDbl(Val(txtMontAdeuMulta.Text))), "#,###,##0.00")
            GranTotal = GranTotal + Me.txtMontAdeuValTotal.Text
            txtMontoAdeudadoTotal = Format(GranTotal, "#,###,##0.00")
            
            rsDetalle.MoveNext
           
         '   Loop
        
    rsCont.MoveNext
         Detail.PrintSection
    
End Sub
Private Sub ActiveReport_ReportStart()
    Dim sql As String
    Dim id As String
    'id = frmAvisoCobro.txtIdentidad
    

    
    sql = "SELECT Contribuyente.Identidad, Contribuyente.Pnombre, Contribuyente.SNombre, Contribuyente.PApellido, Contribuyente.SApellido"
    sql = sql & " FROM Contribuyente INNER JOIN Catastro ON Contribuyente.Identidad = Catastro.Identidad INNER JOIN"
    sql = sql & " AbonadoSPEnc ON Contribuyente.Identidad = AbonadoSPEnc.Identidad AND Catastro.ClaveCatastro = AbonadoSPEnc.ClaveCatastro AND Catastro.ClaveCatastro = AbonadoSPEnc.ClaveCatastro"
    sql = sql & " GROUP BY Contribuyente.Identidad, Contribuyente.Pnombre, Contribuyente.SNombre, Contribuyente.PApellido, Contribuyente.SApellido"
    
    
    sql = " SELECT Contribuyente.Identidad, Contribuyente.Pnombre, Contribuyente.SNombre, Contribuyente.PApellido, Contribuyente.SApellido, AvPgEnc.ClaveCatastro "
    sql = sql & " FROM Contribuyente INNER JOIN Catastro ON Contribuyente.Identidad = Catastro.Identidad INNER JOIN "
    sql = sql & " AvPgEnc ON Contribuyente.Identidad = AvPgEnc.Identidad "
    sql = sql & " WHERE (AvPgEnc.AvPgEstado = 1) AND AvPgEnc.AvPgTipoImpuesto = 1 and catastro.codbarrio = '161810001' and fechavenceavpg  < '" & Format(Now, "dd/mm/yyyy") & "' "
    sql = sql & " GROUP BY Contribuyente.Identidad, Contribuyente.Pnombre, Contribuyente.SNombre, Contribuyente.PApellido, Contribuyente.SApellido, AvPgEnc.ClaveCatastro "
    
    
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open ("Select * from ParametroCont")
    Set rsCont = DeRia.CoRia.Execute(sql)
    If rsCont.EOF Then
        MsgBox "No tiene facturas pendientes"
    End If
End Sub

Private Function cargaID(ByVal Identidad As String) As String
    Dim id As String
    id = Identidad
    cargaID = id
End Function



 

