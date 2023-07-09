VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rpAvisoDeCobro 
   Caption         =   "AvisodeCobroQuimistan - rpAvisoDeCobro (ActiveReport)"
   ClientHeight    =   10350
   ClientLeft      =   4815
   ClientTop       =   1380
   ClientWidth     =   16980
   _ExtentX        =   29951
   _ExtentY        =   18256
   SectionData     =   "rpAvisoDeCobro33.dsx":0000
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
    Me.txtMontAdeuDiaspPago = diasParaPago
    strSql = "SELECT identidad, Pnombre, SNombre, PApellido, SApellido, direccion FROM Contribuyente WHERE (Identidad = '" & Identidad & "')"
    
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open strSql
    Me.txtidentidad = Trim(DeRia.rsAbonadoSP!Identidad)
    Me.txtMontAdeuNombreCont = Trim(DeRia.rsAbonadoSP!Pnombre) & " " & Trim(DeRia.rsAbonadoSP!SNombre) & " " & Trim(DeRia.rsAbonadoSP!PApellido) & " " & Trim(DeRia.rsAbonadoSP!SApellido)
    Me.txtMontAdeuDireccion = Trim(DeRia.rsAbonadoSP!Direccion)
    Me.txtFechaAviso = Format(Now, "dddd, dd  mmmm  yyyy")
End Function

Private Sub Detail_Format()
    Dim TipoCuenta As String
    Dim anio As String
    Dim conceptoPago As String
    Static X As Integer
    Dim rsDetalle As New ADODB.Recordset
    X = X + 1
        If X > rsCont.RecordCount Then Exit Sub
            DatosGenerales (rsCont!Identidad)
            If tipoImpuesto = 1 Then
                sql = "  SELECT AvPgEnc.AvPgDescripcion, DATEPART(year, MIN(AvPgEnc.FechaVenceAvPg)) AS aniox, DATEPART(year, MAX(AvPgEnc.FechaVenceAvPg)) AS anioMAx, SUM(AvPgDetalle.ValorUnitAvPgDet) AS total, AvPgEnc.AvPgTipoImpuesto, "
                sql = sql & " AvPgEnc.ClaveCatastro FROM AvPgDetalle INNER JOIN AvPgEnc ON AvPgDetalle.NumAvPg = AvPgEnc.NumAvPg INNER JOIN "
                sql = sql & " CuentaIngreso_A ON AvPgDetalle.CtaIngreso = CuentaIngreso_A.CtaIngreso AND DATEPART(year, AvPgEnc.FechaVenceAvPg) = CuentaIngreso_A.Anio "
                sql = sql & " WHERE (AvPgEnc.AvPgEstado = 1) AND AvPgEnc.AvPgTipoImpuesto = 1 AND (AvPgEnc.ClaveCatastro = '" & rsCont!ClaveCatastro & "') "
                sql = sql & " GROUP BY AvPgEnc.AvPgDescripcion, AvPgEnc.AvPgTipoImpuesto, AvPgEnc.ClaveCatastro ORDER BY ANIOX "
                Me.txtClavecatastral.Text = rsCont!ClaveCatastro
                Me.Line46.Visible = True
                Me.lbcodigoOclave.Visible = True
                Me.lbcodigoOclave = "Clave Catastral:"
            ElseIf tipoImpuesto = 2 Then
                sql = " SELECT DATEPART(year, MIN(AvPgEnc.FechaVenceAvPg)) AS aniox, DATEPART(year, MAX(AvPgEnc.FechaVenceAvPg)) AS anioMAx, SUM(AvPgDetalle.ValorUnitAvPgDet) AS total,  AvPgEnc.Identidad"
                sql = sql & " FROM AvPgDetalle INNER JOIN"
                sql = sql & " AvPgEnc ON AvPgDetalle.NumAvPg = AvPgEnc.NumAvPg"
                sql = sql & " WHERE (AvPgEnc.AvPgEstado = 1) AND (AvPgEnc.AvPgTipoImpuesto IN (2, 3)) AND (AvPgEnc.Identidad = '" & rsCont!Identidad & "')"
                sql = sql & " GROUP BY  AvPgEnc.Identidad"
                getDatosComer (rsCont!Identidad)
            ElseIf tipoImpuesto = 5 Then
                sql = " SELECT DATEPART(year, MIN(AvPgEnc.FechaVenceAvPg)) AS aniox, DATEPART(year, MAX(AvPgEnc.FechaVenceAvPg)) AS anioMAx, SUM(AvPgDetalle.ValorUnitAvPgDet) AS total, AvPgEnc.AvPgTipoImpuesto, AvPgEnc.Identidad,"
                sql = sql & " AvPgEnc.AvPgDescripcion FROM AvPgDetalle INNER JOIN"
                sql = sql & " AvPgEnc ON AvPgDetalle.NumAvPg = AvPgEnc.NumAvPg"
                sql = sql & " WHERE (AvPgEnc.AvPgEstado = 1) AND (AvPgEnc.AvPgTipoImpuesto IN (5)) AND (AvPgEnc.Identidad = '" & rsCont!Identidad & "')"
                sql = sql & " GROUP BY AvPgEnc.AvPgTipoImpuesto, AvPgEnc.Identidad, AvPgEnc.AvPgDescripcion"
                Me.txtNoAbonado.Text = rsCont!ASPE_Seq
                Me.txtClavecatastral.Text = rsCont!ClaveCatastro
                Me.Line46.Visible = True
                Me.Line50.Visible = True
                Me.lcCodAbonado.Visible = True
                Me.lbcodigoOclave.Visible = True
                Me.lbcodigoOclave = "Clave Catastral:"
            End If
            
            Set rsDetalle = DeRia.CoRia.Execute(sql)
                'TipoCuenta = rsDetalle!AvPgTipoImpuesto
                anio = rsDetalle!ANIOX
                If tipoImpuesto = 2 Then
                    Me.txtMontAdeuConcepto = "Facturas de Industria Comercio y Servicio"
                Else
                   Me.txtMontAdeuConcepto = rsDetalle!AvPgDescripcion
                End If
                If rsDetalle!ANIOX = rsDetalle!anioMAx Then
                   Me.txtMontAdeuAnioImpositivo = rsDetalle!anioMAx
                Else
                   Me.txtMontAdeuAnioImpositivo = "Del: " & rsDetalle!ANIOX & " Al: " & rsDetalle!anioMAx
                End If
                Me.txtMontAdeuValTotal = Format(rsDetalle!Total, "#,###,##0.00")
            rsDetalle.MoveNext
        rsCont.MoveNext
    Detail.PrintSection
End Sub
Private Sub ActiveReport_ReportStart()
    Dim sql As String
    Dim fecha As Date
    fecha = Format(Now, "dd/mm/yyyy")
    If tipoImpuesto = 1 Then
        sql = " SELECT Contribuyente.Identidad, Contribuyente.Pnombre, Contribuyente.SNombre, Contribuyente.PApellido, Contribuyente.SApellido, AvPgEnc.ClaveCatastro"
        sql = sql & "  FROM Contribuyente INNER JOIN Catastro ON Contribuyente.Identidad = Catastro.Identidad INNER JOIN"
        sql = sql & "  AvPgEnc ON Contribuyente.Identidad = AvPgEnc.Identidad"
        sql = sql & "  WHERE (AvPgEnc.AvPgEstado = 1) AND AvPgEnc.AvPgTipoImpuesto = 1 and catastro.codbarrio = '" & codigoBarrio & "' "
        sql = sql & "  AND DATEDIFF(month, avpgenc.fechavenceavpg," & fecha & ")<" & mesesAdeudados & " "
        sql = sql & "  GROUP BY Contribuyente.Identidad, Contribuyente.Pnombre, Contribuyente.SNombre, Contribuyente.PApellido, Contribuyente.SApellido, AvPgEnc.ClaveCatastro"
    ElseIf tipoImpuesto = 2 Then
        sql = " SELECT Contribuyente.Identidad"
        sql = sql & " FROM Contribuyente INNER JOIN"
        sql = sql & " AvPgEnc ON Contribuyente.Identidad = AvPgEnc.Identidad"
        sql = sql & " WHERE (AvPgEnc.AvPgEstado = 1) AND AvPgEnc.AvPgTipoImpuesto in (2,3) and Contribuyente.codbarrio = '" & codigoBarrio & "' "
        sql = sql & " AND DATEDIFF(month, avpgenc.fechavenceavpg," & fecha & ")<" & mesesAdeudados & " "
        sql = sql & " GROUP BY Contribuyente.Identidad"
    ElseIf tipoImpuesto = 5 Then
        sql = " SELECT Contribuyente.Identidad, Contribuyente.Pnombre, Contribuyente.SNombre, Contribuyente.PApellido, Contribuyente.SApellido, AbonadoSPEnc.ASPE_Seq, AbonadoSPEnc.ClaveCatastro "
        sql = sql & " FROM Contribuyente INNER JOIN"
        sql = sql & " AvPgEnc ON Contribuyente.Identidad = AvPgEnc.Identidad INNER JOIN "
        sql = sql & " AbonadoSPEnc ON Contribuyente.Identidad = AbonadoSPEnc.Identidad "
        sql = sql & " WHERE AvPgEnc.AvPgEstado = 1 AND (AvPgEnc.AvPgTipoImpuesto = 5) AND (Contribuyente.CodBarrio = '" & codigoBarrio & "') "
        sql = sql & " AND DATEDIFF(month, avpgenc.fechavenceavpg," & fecha & ")<" & mesesAdeudados & " "
        sql = sql & " GROUP BY Contribuyente.Identidad, Contribuyente.Pnombre, Contribuyente.SNombre, Contribuyente.PApellido, Contribuyente.SApellido, AbonadoSPEnc.ASPE_Seq, AbonadoSPEnc.ClaveCatastro "
    End If
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open ("Select * from ParametroCont")
    Set rsCont = DeRia.CoRia.Execute(sql)
    If rsCont.EOF Then
       MsgBox "No tiene facturas pendientes"
    End If
    
End Sub
Public Function getDatosComer(Identidad As String)
Dim rsDatosCom As New ADODB.Recordset
Dim consulta As String
Dim id As String
id = Identidad
consulta = " SELECT Contribuyente.Identidad, Contribuyente.Pnombre, Contribuyente.Direccion, Contribuyente_1.Identidad AS Expr1, Contribuyente_1.Pnombre AS Expr2, Contribuyente_1.SNombre, Contribuyente_1.PApellido, "
consulta = consulta & " Contribuyente_1.SApellido, CuentaIngreso_A.NombreCtaIngreso"
consulta = consulta & " FROM Contribuyente INNER JOIN"
consulta = consulta & " CuentaIngreso_A ON Contribuyente.CodProfesion = CuentaIngreso_A.CtaIngreso INNER JOIN"
consulta = consulta & " Contribuyente AS Contribuyente_1 ON Contribuyente.IdRepresentante = Contribuyente_1.Identidad"
consulta = consulta & " WHERE Contribuyente.identidad = '" & id & "' and (Contribuyente.Tipo = 'true') AND (CuentaIngreso_A.Anio = DATEPART(year, GETDATE()))"
Set rsDatosCom = DeRia.CoRia.Execute(consulta)
Me.Label37 = "R.T.M"
Me.txtidentidad = id
Me.txtNoAbonado.Text = IIf(IsNull(rsDatosCom!NombreCtaIngreso), "Verifique actividad", rsDatosCom!NombreCtaIngreso)
Me.txtClavecatastral.Text = Trim(rsDatosCom!Expr2) & " " & Trim(rsDatosCom!SNombre) & " " & Trim(rsDatosCom!PApellido) & " " & Trim(rsDatosCom!SApellido)
    Me.Line46.Visible = True
    Me.Line50.Visible = True
    Me.lcCodAbonado.Visible = True
    Me.lbcodigoOclave.Visible = True
    Me.lbcodigoOclave = "Nombre Propietario:"
    Me.lcCodAbonado = "Actividad Economica:"
    Me.lbEstimado = "Establecimiento:"
    
    
End Function


 

