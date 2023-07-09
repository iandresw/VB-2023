Attribute VB_Name = "Contrib"
Option Explicit

Sub AddEncFacturaTemp(sIdentidad As String, dFecha As Date, sDescrip As String)
    'Se agregan los datos a un recordset global, para ser mostrados al momento de llamar
    'la ventana de facturacion
    
    'Si ya existen datos, salga
    If gRsEnc.RecordCount = 1 Then Exit Sub
    gRsEnc.AddNew
    gRsEnc!Identidad = sIdentidad
    gRsEnc!FechaEmAvPg = dFecha
    gRsEnc!AvPgDescripcion = sDescrip
    gRsEnc.Update
End Sub
Sub AddDetFacturaTemp(sCtaIng As String, sCodCat As String, cSubTotal As Currency, cDescuento As Currency, cRecargo As Currency, sCodDeclara As String)
    gRsDet.AddNew
    gRsDet!CtaIngreso = sCtaIng
    gRsDet!ClaveCatastro = sCodCat
    gRsDet!CantAvPgDet = 1
    gRsDet!ValorUnitAvPgDet = cSubTotal
    gRsDet!DescuentoAvPgDet = cDescuento
    gRsDet!RecargoAvPgDet = cRecargo
    gRsDet!RefAvPgDet = sCodDeclara
    gRsDet.Update
End Sub
Public Function ComputeMora(cValor As Currency, ByVal cMulta1 As Currency, ByVal dFechaVence As Date, dFechaActual As Date) As Currency
    'Esta rutina, calcula el Recargo por Mora y el Recargo Sobre Saldo
    
    'Rutina para calcular la mora de una factura pendiente de pago y vencida.
    'Aplica las tasas de cada año, segun definidas en la tabla ParamRia.
    
    'Vamos a hacer una rutina tomando los estandares de aplicacion de multas
    'Se le aplica un porcentaje anual (se divide entre 12, para hacerlo mensual)
    'y tambien se le aplica un recargo mensual sobre saldo.
    'rParam!RecargoAtrasopagoSL = Mantiene el interes bancario anual segun la banca
    'rParam!Recargosobresaldo = mantiene el porcentaje anual de recargo sobre saldo
    
    'Actualizacion: esta rutina la partimos en otras 2 rutinas porque en esta hacemos el
    'calculo de dos recargos, tal como lo requiere una factura vencida; pero me tope
    'con un problema en las declaraciones de Industria y Comercio donde hacemos
    'el calculo de un recargo mes por mes y el segundo recargo de aqui,
    'lo hacemos de un solo para todos los meses
    'Esta rutina esta siendo llamada desde la declaracion de
    'la ventana de atencion al cliente.
    
    Dim rParam As New Recordset
    Dim iYear As Integer
    Dim cMulta, cRecMensual As Currency
    Dim cSaldo As Currency
    Dim cinteres As Currency
    
    ComputeMora = 0
    If dFechaVence >= dFechaActual Then Exit Function
    iYear = Year(dFechaVence)
    Do While iYear <= Year(dFechaActual)
        Set rParam = DeRia.CoRia.Execute("select * from ParamRia where PeriodoFact=" & iYear & " ")
        If rParam.RecordCount = 0 Then
            MsgBox "No estan definidos los parametros para el año " & iYear
            Exit Function
        End If
        If IsNull(rParam!RecargoAtrasoPagoSL) Or IsNull(rParam!RecargoSobreSaldo) Then
            MsgBox "No estan definidos los parametros de multas para el año " & iYear
            Exit Function
        End If
        Do While True
            'Calculemos el recargo por mora
            'si dFechaVence es mayor que la fecha actual o el año no es igual al que esta en proceso, salga del loop
            If dFechaVence > dFechaActual Or Year(dFechaVence) <> iYear Then Exit Do
            cRecMensual = cRecMensual + (rParam!RecargoAtrasoPagoSL * cValor)
            dFechaVence = DateAdd("m", 1, dFechaVence)
        Loop
        'Ahora calculemos el recargo sobre saldo
        'cMulta = Round(rParam!Recargosobresaldo / 12, 5)
        cSaldo = cRecMensual + cMulta1 + cValor
        cinteres = Round((rParam!RecargoSobreSaldo / 12), 5)
        cMulta = cinteres * cSaldo
        'cMulta = cMulta + (Round((rParam!RecargoSobreSaldo / 12), 5) * (cRecMensual + cMulta1 + cValor))
        iYear = Year(dFechaVence) 'Aqui dFechaVence debe estar en Enero del siguiente año
        If dFechaVence > dFechaActual Then Exit Do
    Loop
    'El recargo por mora, lo retornamos como valor de la funcion y el
    'recargo sobre saldo lo retornamos en cRecSobreSaldo, porque se pasa como referencia.
    ComputeMora = cRecMensual + cMulta
    'cRecSobreSaldo = cMulta 'cRecSobreSaldo estaba como parametro
End Function
Public Function ComputeMora2(cValor As Currency, cMulta1 As Currency, ByVal dFechaVence As Date, dFechaActual As Date) As Currency
    'Rutina para calcular la mora de una factura pendiente de pago y vencida.
    'Aplica las tasas de cada año, segun definidas en la tabla ParamRia.
    
    'Vamos a hacer una rutina tomando los estandares de aplicacion de multas
    'Se le aplica un porcentaje anual (se divide entre 12, para hacerlo mensual)
    'y tambien se le aplica un recargo mensual sobre saldo.
    'rParam!RecargoAtrasopagoSL = Mantiene el interes bancario anual segun la banca
    'rParam!Recargosobresaldo = mantiene el porcentaje anual de recargo sobre saldo
    
    'Actualizacion: esta rutina la partimos en otras 2 rutinas porque en esta hacemos el
    'calculo de dos recargos, tal como lo requiere una factura vencida; pero me tope
    'con un problema en las declaraciones de Industria y Comercio donde hacemos
    'el calculo de un recargo mes por mes y el segundo recargo de aqui,
    'lo hacemos de un solo para todos los meses
    'Esta rutina esta siendo llamada desde la declaracion de
    'la ventana de atencion al cliente.
    
    'Esta es la rutina Numero 1
    
    Dim rParam As New Recordset
    Dim iYear As Integer
    Dim cMulta, cRecMensual As Currency
    
    ComputeMora2 = 0
    If dFechaVence >= dFechaActual Then Exit Function
    iYear = Year(dFechaVence)
    Set rParam = DeRia.CoRia.Execute("select * from ParamRia where PeriodoFact=" & iYear & " ")
    If rParam.RecordCount = 0 Then
        MsgBox "No estan definidos los parametros para el año a procesar..!"
        Exit Function
    End If
    If IsNull(rParam!RecargoAtrasoPagoSL) Or IsNull(rParam!RecargoSobreSaldo) Then
        MsgBox "No estan definidos los parametros de multas..!"
        Exit Function
    End If
    Do While True
    'si dFechaVence es mayor que la fecha actual o el año no es igual al que esta en proceso, salga del loop
    If dFechaVence > dFechaActual Or Year(dFechaVence) <> iYear Then Exit Do
        cRecMensual = cRecMensual + (rParam!RecargoAtrasoPagoSL * cValor)
        dFechaVence = DateAdd("m", 1, dFechaVence)
    Loop
    'sumamos los recargos y retornamos ese valor
    ComputeMora2 = cRecMensual
End Function
'Public Function MultaPorMesesNoPagadosIC(cValor As Currency, ByVal dFechaVence As Date, dFechaActual As Date, cValorCuotasVencidas As Currency) As Currency
'    'Esta rutina se llama desde Declaracion de Impuesto a Industria y Comercio.
'    'Calcula la multa por no pagar el impuesto mensual a tiempo. Tambien retorna
'    'el Monto en Deuda
'    '--------------------------------------------------------------------------
'    'La cuota de Enero vence el 31 de Enero, las demas
'    'vencen el dia de cada de mes que dice parametros
'    'OJO=Si se pasa un dia se le aplica el mes.
'    'No estamos usando esta rutina
'    Dim rParam As New Recordset
'    Dim iYear As Integer
'    Dim cMulta As Currency
'    MultaPorMesesNoPagadosIC = 0
'    If dFechaVence >= dFechaActual Then Exit Function
'    iYear = Year(dFechaVence)
'    Set rParam = DeRia.CoRia.Execute("select * from ParamRia where PeriodoFact=" & iYear & " ")
'    If rParam.RecordCount = 0 Then
'        MsgBox "No estan definidos los parametros para el año a procesar..!"
'        Exit Function
'    End If
'    If IsNull(rParam!RecargoAtrasoPagoSL) Then
'        MsgBox "No estan definidos los parametros de multas..!"
'        Exit Function
'    End If
'    'Enero es caso especial
'    If Month(dFechaVence) = 1 Then
'        If dFechaActual > "31/01/" & Year(dFechaVence) Then
'            cMulta = rParam!RecargoAtrasoPagoSL * cValor
'            cValorCuotasVencidas = cValor
'        End If
'        dFechaVence = DateAdd("m", 1, dFechaVence)
'    End If
'    Do While True
'        'si dFechaVence es mayor que la fecha actual o el año no es igual al que esta en proceso, salga del loop
'        If dFechaVence > dFechaActual Or Year(dFechaVence) <> iYear Then Exit Do
'        cMulta = cMulta + (rParam!RecargoAtrasoPagoSL * cValor)
'        dFechaVence = DateAdd("m", 1, dFechaVence)
'        cValorCuotasVencidas = cValorCuotasVencidas + cValor
'    Loop
'    MultaPorMesesNoPagadosIC = cMulta
'End Function
Public Function ComputeMora3(cValor As Currency, cMulta1 As Currency, ByVal dFechaVence As Date, dFechaActual As Date) As Currency
    'Rutina para calcular la mora de una factura pendiente de pago y vencida.
    'Aplica las tasas de cada año, segun definidas en la tabla ParamRia.
    
    'Vamos a hacer una rutina tomando los estandares de aplicacion de multas
    'Se le aplica un porcentaje anual (se divide entre 12, para hacerlo mensual)
    'y tambien se le aplica un recargo mensual sobre saldo.
    'rParam!RecargoAtrasopagoSL = Mantiene el interes bancario anual segun la banca
    'rParam!Recargosobresaldo = mantiene el porcentaje anual de recargo sobre saldo
    
    'Actualizacion: esta rutina la partimos en otras 2 rutinas porque en esta hacemos el
    'calculo de dos recargos, tal como lo requiere una factura vencida; pero me tope
    'con un problema en las declaraciones de Industria y Comercio donde hacemos
    'el calculo de un recargo mes por mes y el segundo recargo de aqui,
    'lo hacemos de un solo para todos los meses
    'Esta rutina esta siendo llamada desde la declaracion de
    'la ventana de atencion al cliente.
    
    'Esta es la rutina Numero 2

    Dim rParam As New Recordset
    Dim iYear As Integer
    Dim nTasa As Single
    Dim cMulta, cRecMensual As Currency
    
    ComputeMora3 = 0
    If dFechaVence >= dFechaActual Then Exit Function
    iYear = Year(dFechaVence)
    Set rParam = DeRia.CoRia.Execute("select * from ParamRia where PeriodoFact=" & iYear & " ")
    If rParam.RecordCount = 0 Then
        MsgBox "No estan definidos los parametros para el año a procesar..!"
        Exit Function
    End If
    If IsNull(rParam!RecargoAtrasoPagoSL) Or IsNull(rParam!RecargoSobreSaldo) Then
        MsgBox "No estan definidos los parametros de multas..!"
        Exit Function
    End If
    'ahora aplicamos un recargo anual sobre el saldo
    nTasa = Round((rParam!RecargoSobreSaldo / 12), 4)
    cMulta = cMulta + (nTasa * (cMulta1 + cValor))
    'sumamos los recargos y retornamos ese valor
    ComputeMora3 = cMulta
End Function
Public Sub InsertAvPgEnc(cConn As Connection, lNum As Long, sId As String, dFeEm As Date, dFeVence As Variant, bEstado As Byte, bTipo As Byte, bTipoImp As Byte, sDescrip As String, sCodDeclara As String)
    Dim sSql As String
    If IsDate(dFeVence) Then
        sSql = "insert into AvPgEnc (NumAvPg,Identidad,FechaEmAvPg,FechaVenceAvPg,AvPgEstado,TipoAvPg,AvPgTipoImpuesto,AvPgDescripcion,CodDeclara) "
        sSql = sSql + " values(" & lNum & ",'" & sId & "','" & Format(dFeEm, "dd/mm/yyyy") & "','" & Format(dFeVence, "dd/mm/yyyy") & "'," & bEstado & "," & bTipo & "," & bTipoImp & ",'" & sDescrip & "','" & sCodDeclara & "') "
    Else
        sSql = "insert into AvPgEnc (NumAvPg,Identidad,FechaEmAvPg,AvPgEstado,TipoAvPg,AvPgTipoImpuesto,AvPgDescripcion,CodDeclara) "
        sSql = sSql + " values(" & lNum & ",'" & sId & "','" & Format(dFeEm, "dd/mm/yyyy") & "'," & bEstado & "," & bTipo & "," & bTipoImp & ",'" & sDescrip & "','" & sCodDeclara & "') "
    End If
    
    'MsgBox sSql
    cConn.Execute (sSql)
End Sub
Public Sub InsertAvPgDetalle(cConn As Connection, lNum As Long, cValorUnit As Currency, sCatastro As String, sCtaIng As String, cCant As Currency, sRef As String, cDescto As Currency, cMulta As Currency, cMulta2 As Currency)
    Dim sSql As String
    
    sSql = "insert into AvPgDetalle (NumAvPg,ValorUnitAvPgDet,ClaveCatastro,CtaIngreso,CantAvPgDet,RefAvPgDet,DescuentoAvPgDet,RecargoAvPgDet,ValorXAvPgDet) "
    sSql = sSql + "values(" & lNum & "," & cValorUnit & ",'" & sCatastro & "','" & sCtaIng & "'," & cCant & ",'" & sRef & "'," & cDescto & "," & cMulta & "," & cMulta2 & ")"
    'MsgBox sSql
    cConn.Execute (sSql)
End Sub
Public Function ComputeDescuento(cValor As Currency, dFechaVence As Date, dFechaActual As Date) As Currency
    'Vamos a hacer una rutina tomando los estandares de aplicacion de descuentos
    'Para que una factura tenga descuento, tiene que pagarse 4 meses antes de vencer.
    'rParam!TiempoParaDescuento=indica el tiempo aplicable para el descuento
    'rParam!DescuentoPagoAnticipado=indica el descuento por pago anticipado
    
    Dim rParam As New Recordset
    Dim iMeses As Integer
    Dim iYear As Integer
    
    ComputeDescuento = 0
    iMeses = DateDiff("M", dFechaActual, dFechaVence)
    iYear = Year(dFechaVence)
    Set rParam = DeRia.CoRia.Execute("select * from ParamRia where PeriodoFact=" & iYear & " ")
    If rParam.RecordCount = 1 Then
        If IsNull(rParam!TiempoParaDescuento) Or IsNull(rParam!DescuentoPagoAnticipado) Then
            MsgBox "No estan definidos los parametros de valores de descuento...!"
            Exit Function
        End If
        If iMeses >= rParam!TiempoParaDescuento Then
            ComputeDescuento = cValor * rParam!DescuentoPagoAnticipado
        End If
    End If
End Function
Public Function CalculeDescuentoBI(cValor As Currency, dFechaVence As Date, dFechaActual As Date) As Currency
    'calcula el descuento para un pago anticipado.
    'Se extraen los parametros de la tabla ParamRia
    'Se extraen los campos DescuentoPagoAnticipado,TiempoParaDescuento
    'Se aplican a los parametros y calculamos el descuento, si aplica
    
    Dim rParam As New Recordset
    Dim iYear As Integer
    Dim dFechaParaDescuento As Date
    
    CalculeDescuentoBI = 0
    iYear = Year(dFechaVence)
    Set rParam = DeRia.CoRia.Execute("select * from ParamRia where PeriodoFact=" & iYear & " ")
    If rParam.RecordCount = 0 Then
        MsgBox "No estan definidos los parametros para el año a procesar..!"
        Exit Function
    End If
    If IsNull(rParam!DescuentoPagoAnticipado) Or IsNull(rParam!TiempoParaDescuento) Then
        MsgBox "No estan definidos los parametros de descuentos..!"
        Exit Function
    End If
    If dFechaVence <= dFechaActual Then Exit Function 'ya vencio o se vence el dia
    'Aqui pasa solo si la fecha de vencimiento es mayor que la fecha actual
    dFechaParaDescuento = DateAdd("m", -rParam!TiempoParaDescuento, dFechaVence)
    If dFechaActual <= dFechaParaDescuento Then
        'Tiene Descuento
        CalculeDescuentoBI = cValor * rParam!DescuentoPagoAnticipado
    End If
End Function
Public Function CheckParametros() As Boolean
    Dim rsPar As New Recordset
    Dim iFlag As Integer
    
    On Error GoTo CheckParametros_Error
    iFlag = 1
    CheckParametros = False
    Set rsPar = DeRia.CoRia.Execute("select * from Parametro")
    If rsPar.RecordCount = 0 Then GoTo CheckParametros_Exit
    If IsNull(rsPar!CodMuni) Or IsNull(rsPar!NombreMuni) Or IsNull(rsPar!NombreDepto) Or IsNull(rsPar!DiaProcesoCT) Then GoTo CheckParametros_Exit
    If Trim(rsPar!CodMuni) = "" Or Trim(rsPar!NombreMuni) = "" Or Trim(rsPar!NombreDepto) = "" Then GoTo CheckParametros_Exit
    If Not IsDate(rsPar!DiaProcesoCT) Then GoTo CheckParametros_Exit
    rsPar.Close
    iFlag = 2
    Set rsPar = DeRia.CoRia.Execute("select * from ParamRia")
    If rsPar.RecordCount = 0 Then GoTo CheckParametros_Exit
    With rsPar
    Do While rsPar.EOF = False
        If IsNull(!periodofact) Or !periodofact = 0 Then GoTo CheckParametros_Exit
        'If IsNull(!BiMultaDeclaraTarde) Or !BiMultaDeclaraTarde = 0 Then GoTo CheckParametros_Exit 'MSx3
        'If IsNull(!BIMultaDeclaraTarde2) Or !BIMultaDeclaraTarde2 = 0 Then GoTo CheckParametros_Exit 'MSx3
        If IsNull(!BiFechaMaxPago) Or Not IsDate(!BiFechaMaxPago) Then GoTo CheckParametros_Exit
        'If IsNull(!BIFechaMaxDeclara) Or Not IsDate(!BIFechaMaxDeclara) Then GoTo CheckParametros_Exit
        If IsNull(!ICFechaMaxDeclara) Or Not IsDate(!ICFechaMaxDeclara) Then GoTo CheckParametros_Exit
        'If IsNull(!ICMultaDeclaraTarde) Or !BiMultaDeclaraTarde = 0 Then GoTo CheckParametros_Exit 'MSx3
        If IsNull(!ICDiaMaxPago) Or !ICDiaMaxPago = 0 Then GoTo CheckParametros_Exit
        If IsNull(!IPFechaMaxDeclara) Or Not IsDate(!IPFechaMaxDeclara) Then GoTo CheckParametros_Exit
        'If IsNull(!IPMultaDeclaraTarde) Or !IPMultaDeclaraTarde = 0 Then GoTo CheckParametros_Exit 'MSx3
        If IsNull(!IPFechaMaxPago) Or Not IsDate(!IPFechaMaxPago) Then GoTo CheckParametros_Exit
        'If IsNull(!RecargoSobreSaldo) Or !RecargoSobreSaldo = 0 Then GoTo CheckParametros_Exit 'MSx3
        'If IsNull(!RecargoAtrasoPagoSL) Or !RecargoAtrasoPagoSL = 0 Then GoTo CheckParametros_Exit 'MSx3
        'If IsNull(!DescuentoPagoAnticipado) Or !DescuentoPagoAnticipado = 0 Then GoTo CheckParametros_Exit 'MSx3
        If IsNull(!TiempoParaDescuento) Or !TiempoParaDescuento = 0 Then GoTo CheckParametros_Exit
        .MoveNext
    Loop
    End With
    rsPar.Close
    iFlag = 3
    Set rsPar = DeRia.CoRia.Execute("select * from systemparam")
    If rsPar.RecordCount = 0 Then GoTo CheckParametros_Exit
    With rsPar
        If IsNull(!CtaIngresoIP) Or !CtaIngresoIP = "" Then GoTo CheckParametros_Exit
        'If IsNull(!CtaIngresoPermOp) Or !CtaIngresoPermOp = "" Then GoTo CheckParametros_Exit
        If IsNull(!CtaIngresoMultaOPSinPermiso) Or !CtaIngresoMultaOPSinPermiso = "" Then GoTo CheckParametros_Exit
        If IsNull(!CtaIngresoDescuento) Or !CtaIngresoDescuento = "" Then GoTo CheckParametros_Exit
        If IsNull(!CtaEgresoFondoPos) Or !CtaEgresoFondoPos = "" Then GoTo CheckParametros_Exit
        If IsNull(!CtaIngresoBiRural) Or !CtaIngresoBiRural = "" Then GoTo CheckParametros_Exit
        If IsNull(!CtaIngresoBiUrb) Or !CtaIngresoBiUrb = "" Then GoTo CheckParametros_Exit
        If IsNull(!CtaIngresoRecargoImp) Or !CtaIngresoRecargoImp = "" Then GoTo CheckParametros_Exit
        If IsNull(!CtaIngresoRecargoServ) Or !CtaIngresoRecargoServ = "" Then GoTo CheckParametros_Exit
        If IsNull(!CtaIngresoIntImp) Or !CtaIngresoIntImp = "" Then GoTo CheckParametros_Exit
        If IsNull(!CtaIngresoIntServ) Or !CtaIngresoIntServ = "" Then GoTo CheckParametros_Exit
    End With
    iFlag = 0
    CheckParametros = True
    
CheckParametros_Exit:

'    Select Case iFlag
'        Case 1
'            frmParamGen.Show
'            frmParamGen.txtDiaProcesoCT.Enabled = True
'        Case 2
'            frmParamRia2.Show
'        Case 3
'            frmParamGen.Show
'            frmParamGen.SSTab1.Tab = 1
'    End Select
    
    Exit Function

CheckParametros_Error:
    MsgBox Err.Description
    Resume CheckParametros_Exit
    
End Function
Public Function MultaDeclaraTardeIP(dFechaPresenta As Date, cImpuesto As Currency, iPeriodo As Integer) As Currency
    Dim rsPar As New Recordset
    
    MultaDeclaraTardeIP = 0
    Set rsPar = DeRia.CoRia.Execute("select * from ParamRia where PeriodoFact=" & iPeriodo & " ")
    'al periodo de la declaracion se le suma 1, porque este ano se hace la declaracion del ano pasado
    'entonces la tabla que buscamos es la de este ano.
    If rsPar.RecordCount = 0 Then
        MsgBox "No estan definidos los datos de parametros para el período solicitado. Factura no puede crearse."
        Exit Function
    End If
    'calcula multa por presentar la declaracion tarde
    MultaDeclaraTardeIP = cImpuesto * IIf(IsNull(rsPar!IPMultaDeclaraTarde), 0, rsPar!IPMultaDeclaraTarde)
End Function
Public Sub MoraPorFactura(RsFact As Recordset)
    'Se calcula y graba el recargo por cada factura que se envia en el recordset como parametro
    '---------------------------------------------------------------------------------------
    'Proceso:
    '1. Ubicarse en la factura deseada.
    '2. Calcula el interes
    '3. Calcule el recargo sobre saldo
    'Las cuentas que son de MULTA no se les calcula el recargo por mora
    '------------------------------------------------------------------
    'Cambio hecho el 23 de Noviembre 2006.
    'Se cambia la fecha base de calculo de los recargos e intereses
    'antes se tomaba como base la fecha del windows, ahora se toma la fecha
    'definida en parametros DiaProcesoCT, para esto se cambio
    'todas las instrucciones Date por DiaProcesoCT.
    
    Dim strCtaIngreso As String
    Dim strFormulaRecargo As String
    Dim curInteres As Currency
    Dim curValor As Currency
    Dim strFormulaInteres As String
    Dim rsDet As New Recordset
    Dim rsDescuenta As New Recordset
    Dim cinteres As Currency
    Dim rsMora As Recordset
    Dim cMulta As Currency
    Dim rsSys As New Recordset
    Dim rsPar As New Recordset
    Dim rsCuenta As Recordset
    Dim cDescuento As Currency
    Dim cRecargoSobreSaldo As Currency
    Dim cRecargoPorMora As Currency
    Dim nMesesVencidos As Single 'puede tener negativos
    Dim cDummy As Currency
    Dim strSql As String
    Dim dDiaProcesoCT As Date
    Dim VMesVen As Date
    Dim RsFecVen As New Recordset
    Dim RsFecVen2 As New Recordset
   ' Dim RsSP As New ADODB.Recordset
    Dim VSp As Currency
    vFactVen = 0
    VFactNum = 0
    Set rsPar = DeRia.CoRia.Execute("select * from Parametro")
    If rsPar.RecordCount = 0 Then
        MsgBox "No estan definidos los parametros del modulo."
        Exit Sub
    End If
    If Not IsDate(rsPar!DiaProcesoCT) Then
        MsgBox "La fecha en proceso es invalida."
        Exit Sub
    End If
    Set rsSys = DeRia.CoRia.Execute("select * from SystemParam")
    Do While RsFact.EOF = False
      If IsDate(RsFact!FechaVenceAvPg) Then
        
        
            If VQuita = 1 Then
    Qdia = "09" 'Format(dDiaProcesoCT, "DD")
    QMes = Format(DiaEnProcesoCT, "MM")
    QAno = Format(DiaEnProcesoCT, "YYYY")
    If QMes = "01" Then
        QMes = 12
        QAno = QAno - 1
    Else
    QMes = Val(QMes) - 1
    If Len(QMes) = 1 Then
    QMes = 0 & QMes
    End If
    End If
    dDiaProcesoCT = Qdia & "/" & QMes & "/" & QAno
   ' nMesesVencidos = MesesVencidos(rsFact!FechaVenceAvPg, dDiaProcesoCT)
    VFactNum = RsFact!NumAvPg
        Set RsFecVen = DeRia.CoRia.Execute("Select * from FactAbono where numavpg = " & RsFact!NumAvPg & "")
        If Not RsFecVen.EOF Then
                     Set RsFecVen2 = DeRia.CoRia.Execute("Select * From AvpgEnc where numavpg = " & RsFecVen!NumFactAbono & "")
           nMesesVencidos = MesesVencidos(RsFecVen2!FechaVenceAvPg, rsPar!DiaProcesoCT)
        Else
           nMesesVencidos = MesesVencidos(RsFact!FechaVenceAvPg, rsPar!DiaProcesoCT)
        End If
    
    Else
    VFactNum = RsFact!NumAvPg
    Set RsFecVen = DeRia.CoRia.Execute("Select * from FactAbono where numavpg = " & RsFact!NumAvPg & "")
        If Not RsFecVen.EOF Then
          Set RsFecVen2 = DeRia.CoRia.Execute("Select * From AvpgEnc where numavpg = " & RsFecVen!NumFactAbono & "")
           If Not RsFecVen2.EOF Then
              nMesesVencidos = MesesVencidos(RsFecVen2!FechaVenceAvPg, rsPar!DiaProcesoCT)
              Else
              nMesesVencidos = MesesVencidos(RsFecVen!FechaVenceAvPg, rsPar!DiaProcesoCT)
           End If
        Else
           nMesesVencidos = MesesVencidos(RsFact!FechaVenceAvPg, rsPar!DiaProcesoCT)
        End If
    End If
        
        
      '  nMesesVencidos = MesesVencidos(rsFact!FechaVenceAvPg, rsPar!DiaProcesoCT) ' Original
         vFactVen = RsFact!NumAvPg
        
        VFactNum = RsFact!NumAvPg
        Dim VANioFactS As Integer
        VANioFactS = Format(RsFact!FechaVenceAvPg, "YYYY")
        
        'Seleccione los detalles de cada factura que no sean interes/Multas/Recargos, porque
        'a esos no se les calcula recargo ni interes.
        
       ' strSql = "select A.* from AvPgDetalle A, CuentaIngreso B where NumAvPg=" & RsFact!NumAvPg & " and "
       ' strSql = strSql & " B.Tipo <> 2 and B.CtaIngreso=A.CtaIngreso order by A.SeqAvPgDet asc"
        
        strSql = "SELECT A.NumAvPg, A.CtaIngreso, A.ValorUnitAvPgDet, A.CantAvPgDet, B.Tipo FROM AvPgDetalle AS A INNER JOIN CuentaIngreso_A AS B ON A.CtaIngreso = B.CtaIngreso Where (b.Tipo <> 2)"
        strSql = strSql & "GROUP BY A.NumAvPg, A.CtaIngreso, A.SeqAvPgDet, A.ValorUnitAvPgDet, A.CantAvPgDet, B.Tipo Having (A.NumAvPg = " & RsFact!NumAvPg & ") ORDER BY A.SeqAvPgDet"
        
        Set rsDet = DeRia.CoRia.Execute(strSql)
        Dim MesVencido As Integer
        MesVencido = nMesesVencidos
        
        
        
        
        'Modificar la 118 por parametro.
        
        
        Do While rsDet.EOF = False
        
        If Not Mid(rsDet!CtaIngreso, 1, 6) = "111112" Or Not Mid(rsDet!CtaIngreso, 1, 6) = "111113" Or Not Mid(rsDet!CtaIngreso, 1, 6) = "111114" Then
        nMesesVencidos = 0
        End If
        
        If rsDet!CtaIngreso = "11111101" Or rsDet!CtaIngreso = "11111001" Or rsDet!CtaIngreso = "11111002" Then
        nMesesVencidos = MesVencido
        End If
        
        If Mid(rsDet!CtaIngreso, 1, 6) = "111112" Or Mid(rsDet!CtaIngreso, 1, 6) = "111113" Or Mid(rsDet!CtaIngreso, 1, 6) = "111114" Then
        nMesesVencidos = MesVencido
        End If
        
        If Mid(rsDet!CtaIngreso, 1, 6) = "112122" Then 'Recuperacio tasas o impuestos
        nMesesVencidos = MesVencido
        End If
        If VANioFactS < 2013 Then ' ===============================Cargue SP para años antes del 2013
        
        If Mid(rsDet!CtaIngreso, 1, 6) = "111117" Then ' 08 de Julio Que cargue intereses a cuentas de Servicio Publico
        nMesesVencidos = MesVencido
        End If
        Else
        
        If Mid(rsDet!CtaIngreso, 1, 6) = "111118" Then
        nMesesVencidos = MesVencido
        
        End If ' ====================================================================================
        
        End If
        If Mid(rsDet!CtaIngreso, 1, 6) = "112123" Then ' 08 de Julio Que cargue intereses a cuentas Recuperación de Servicio Publico
        nMesesVencidos = MesVencido
        'nMesesVencidos = 0
        End If
        
        
        
     If VANioFactS < 2013 Then ' ===============================Cargue 'Hace que no calcule intereses a Permiso de Operación
        If Mid(rsDet!CtaIngreso, 1, 8) = "11111821" Or Mid(rsDet!CtaIngreso, 1, 8) = "11212208" Then
        nMesesVencidos = MesVencido
        End If
     Else
        If Mid(rsDet!CtaIngreso, 1, 8) = "11111921" Or Mid(rsDet!CtaIngreso, 1, 8) = "11212209" Then
        nMesesVencidos = MesVencido
        End If
     
     End If '================================================================================================
        'Para Amnistia
        If DTE.rsAmni.State = 1 Then DTE.rsAmni.Close
        DTE.rsAmni.Open ("Select * from systemparam ")
        
        If Not DTE.rsAmni.EOF Then
        VAmniIp = DTE.rsAmni!IIP
        VAmniIC = DTE.rsAmni!Iic
        VAmniBi = DTE.rsAmni!ibi
        VAmniSP = DTE.rsAmni!SP
        End If
        Dim rsAmni As New ADODB.Recordset
        
        Set rsAmni = DeRia.CoRia.Execute("Select FechaAmni FROM SystemParam")
        'Impuesto Personal
        If VAmniIp = 1 And RsFact!AvPgTipoImpuesto = 4 And rsAmni!FechaAmni > RsFact!FechaVenceAvPg Then  ' 1. No Aplica interes 0. Si Aplica
            nMesesVencidos = 0
        End If
        
        If VAmniIC = 1 And RsFact!AvPgTipoImpuesto = 2 And rsAmni!FechaAmni > RsFact!FechaVenceAvPg Then
            nMesesVencidos = 0
            DeRia.CoRia.Execute ("Delete FROM AvPgDetalle where NumAvPg = " & VFactNum & " and SUBSTRING(AvPgDetalle.CtaIngreso, 1, 8) = '11212002'")
            'DeRia.CoRia.Execute ("Delete FROM AvPgDetalle where NumAvPg = " & VFactNum & " and SUBSTRING(AvPgDetalle.CtaIngreso, 1, 8) = '11212101'")
            'DeRia.CoRia.Execute ("Delete FROM AvPgDetalle where NumAvPg = " & VFactNum & " and SUBSTRING(AvPgDetalle.CtaIngreso, 1, 8) = '11212601'")
        End If
        
        If VAmniIC = 1 And RsFact!AvPgTipoImpuesto = 3 And rsAmni!FechaAmni > RsFact!FechaVenceAvPg Then
            nMesesVencidos = 0
            DeRia.CoRia.Execute ("Delete FROM AvPgDetalle where NumAvPg = " & VFactNum & " and SUBSTRING(AvPgDetalle.CtaIngreso, 1, 8) = '11212002'")
            'DeRia.CoRia.Execute ("Delete FROM AvPgDetalle where NumAvPg = " & VFactNum & " and SUBSTRING(AvPgDetalle.CtaIngreso, 1, 8) = '11212101'")
            'DeRia.CoRia.Execute ("Delete FROM AvPgDetalle where NumAvPg = " & VFactNum & " and SUBSTRING(AvPgDetalle.CtaIngreso, 1, 8) = '11212601'")
        End If
 
        
        VCuentaServicios = ""
  
        
        If VAmniBi = 1 And RsFact!AvPgTipoImpuesto = 1 And rsAmni!FechaAmni > RsFact!FechaVenceAvPg Then
            nMesesVencidos = 0
        End If
            
            If MesVencido < 0 Then
            nMesesVencidos = MesVencido
            End If
            
           Dim rsFactS230 As Recordset
     
         Set rsFactS230 = DeRia.CoRia.Execute("Select * From FactXTes where NoFacts = " & RsFact!NumAvPg & " ")
        If Not rsFactS230.EOF Then
                
              If nMesesVencidos > 0 Then
                 nMesesVencidos = 0
              End If
                
                
        End If
     
            Dim RsIntAb As New Recordset
            Dim RsAbono As New Recordset
            Dim RsAbono2 As New Recordset
            Dim VFVence As Date
          
            
    Set RsFecVen = DeRia.CoRia.Execute("Select * from FactAbono where numavpg = " & RsFact!NumAvPg & "")
        If Not RsFecVen.EOF Then
                     Set RsFecVen2 = DeRia.CoRia.Execute("Select * From AvpgEnc where numavpg = " & RsFecVen!NumFactAbono & "")
                     
           If Not RsFecVen2.EOF Then
            VFVence = RsFecVen2!FechaVenceAvPg
           Else
            VFVence = RsFact!FechaVenceAvPg
           End If
        Else
           VFVence = RsFact!FechaVenceAvPg
   End If
            
            
      Dim RsDetUp As New ADODB.Recordset
        Dim Str As String
        Dim VFactSp As Long
        Dim vCuenta As String
        Dim RsTpo As New ADODB.Recordset
            
            
            
            
 
       Set RsDetUp = DeRia.CoRia.Execute(" SELECT AvPgDetalle.CtaIngreso, AvPgEnc.FechaVenceAvPg, AvPgDetalle.NumAvPg  FROM AvPgDetalle INNER JOIN AvPgEnc ON AvPgDetalle.NumAvPg = AvPgEnc.NumAvPg WHERE (AvPgDetalle.NumAvPg = " & RsFact!NumAvPg & ")")
        
        
        Do While Not RsDetUp.EOF
     '---------------------------------------------------------------------------------------------------------------------------------
     Set RsTpo = DeRia.CoRia.Execute("Select Tipo from CuentaIngreso_A where Anio = " & Year(RsDetUp!FechaVenceAvPg) & " and CtaIngreso = '" & RsDetUp!CtaIngreso & "' ")
     
     If RsTpo.EOF Then
     Set RsTpo = DeRia.CoRia.Execute("Select Tipo from CuentaIngreso_A where CtaIngreso = '" & RsDetUp!CtaIngreso & "' ")
     End If
     
     strCtaIngreso = RsDetUp!CtaIngreso
     vCuenta = RsDetUp!CtaIngreso
     VFactSp = RsDetUp!NumAvPg
     
           If Format(RsDetUp!FechaVenceAvPg, "YYYY") < 2013 Then

                      If Mid(strCtaIngreso, 4, 3) = "118" And RsTpo!Tipo = 3 Then
                     strCtaIngreso = Mid(strCtaIngreso, 1, 3) & "117" & Mid(strCtaIngreso, 7, 5)
                     vCuenta = Mid(strCtaIngreso, 1, 3) & "117" & Mid(strCtaIngreso, 7, 5)
                     If vCuenta <> RsDetUp!CtaIngreso Then
                       DeRia.CoRia.Execute ("Update Avpgdetalle set CtaIngreso = '" & vCuenta & "' where CtaIngreso = '" & RsDetUp!CtaIngreso & "' and NumAvpg = " & RsDetUp!NumAvPg & " ")
                     End If
                   ' rsDet!CtaIngreso = strCtaIngreso
                  End If

                 '------------------------------------------------------------------------------------------------------------
            If Mid(strCtaIngreso, 4, 3) = "119" And RsTpo!Tipo = 1 Then
                     strCtaIngreso = Mid(strCtaIngreso, 1, 3) & "118" & Mid(strCtaIngreso, 7, 5)
                     vCuenta = Mid(strCtaIngreso, 1, 3) & "118" & Mid(strCtaIngreso, 7, 5)
                     If vCuenta <> RsDetUp!CtaIngreso Then
                      DeRia.CoRia.Execute ("Update Avpgdetalle set CtaIngreso = '" & vCuenta & "' where CtaIngreso = '" & RsDetUp!CtaIngreso & "' and NumAvpg = " & RsDetUp!NumAvPg & " ")
                     End If
                   ' rsDet!CtaIngreso = strCtaIngreso
                  End If



             Else
                  '------------------------------------------------------------------------------------------------------------
                  If Mid(strCtaIngreso, 4, 3) = "117" And RsTpo!Tipo = 3 Then
                     strCtaIngreso = Mid(strCtaIngreso, 1, 3) & "118" & Mid(strCtaIngreso, 7, 5)
                     vCuenta = Mid(strCtaIngreso, 1, 3) & "118" & Mid(strCtaIngreso, 7, 5)
                   If vCuenta <> RsDetUp!CtaIngreso Then
                      DeRia.CoRia.Execute ("Update Avpgdetalle set CtaIngreso = '" & vCuenta & "' where CtaIngreso = '" & RsDetUp!CtaIngreso & "' and NumAvpg = " & RsDetUp!NumAvPg & " ")
                   End If
                   ' rsDet!CtaIngreso = strCtaIngreso
                  End If
 
                  If Mid(strCtaIngreso, 4, 3) = "118" And RsTpo!Tipo = 1 Then
                     strCtaIngreso = Mid(strCtaIngreso, 1, 3) & "119" & Mid(strCtaIngreso, 7, 5)
                     vCuenta = Mid(strCtaIngreso, 1, 3) & "119" & Mid(strCtaIngreso, 7, 5)
                   If vCuenta <> RsDetUp!CtaIngreso Then
                      DeRia.CoRia.Execute ("Update Avpgdetalle set CtaIngreso = '" & vCuenta & "' where CtaIngreso = '" & RsDetUp!CtaIngreso & "' and NumAvpg = " & RsDetUp!NumAvPg & " ")
                   End If
                   ' rsDet!CtaIngreso = strCtaIngreso
                  End If
            End If
     
     
     '---------------------------------------------------------------------------------------------------------------------------------
        RsDetUp.MoveNext
        Loop
                                
'00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
 
            
            
            
            Select Case nMesesVencidos
                    
                Case Is >= 1
                 
                    Select Case RsFact!AvPgTipoImpuesto
                        Case 1
                            'Bienes Inmuebles Interes
                            
                            
                                Dim rsFacSp As New ADODB.Recordset
                                
        strSql = "SELECT A.NumAvPg, A.CtaIngreso, A.ValorUnitAvPgDet, A.CantAvPgDet, B.Tipo FROM AvPgDetalle AS A INNER JOIN CuentaIngreso_A AS B ON A.CtaIngreso = B.CtaIngreso Where (b.Tipo = 1)"
        strSql = strSql & "GROUP BY A.NumAvPg, A.CtaIngreso, A.SeqAvPgDet, A.ValorUnitAvPgDet, A.CantAvPgDet, B.Tipo Having (A.NumAvPg = " & RsFact!NumAvPg & ") ORDER BY A.SeqAvPgDet"
        
                                
                                Set rsFacSp = DeRia.CoRia.Execute(strSql)

                            
                            CalculeRecargoBI rsFacSp!ValorUnitAvPgDet, rsPar!DiaProcesoCT, VFVence, cinteres, cRecargoPorMora, cRecargoSobreSaldo, "", ""
                            cDescuento = 0
                       'Interes Abono
                        Set RsAbono = DeRia.CoRia.Execute("Select * from FactAbono where  numavpg = " & RsFact!NumAvPg & " ")
                         If Not RsAbono.EOF Then
                         Set RsAbono2 = DeRia.CoRia.Execute("Select * from FactAbono where substring(CtaIngreso,1, 6) = '112126' and numavpg = " & RsFact!NumAvPg & "")
                           ' Set RsIntAb = DeRia.CoRia.Execute("Select * from avpgdetalle where substring(CtaIngreso,1, 6) = '112126' and numavpg = " & RsAbono!numavpg & " ")
                                   
                                    If Not RsAbono2.EOF Then
                                    
                                        cinteres = cinteres + IIf(IsNull(RsAbono2!ValorUnitAvPgDet), 0, RsAbono2!ValorUnitAvPgDet) - IIf(IsNull(RsAbono2!ValorAbonado), 0, RsAbono2!ValorAbonado)
                                    End If
                                                                    
                                    
                                    
                         End If
                         
                       'Recargo Abono
                       
                       
                       Set RsAbono = DeRia.CoRia.Execute("Select * from FactAbono where  numavpg = " & RsFact!NumAvPg & " ")
                         If Not RsAbono.EOF Then
                         Set RsAbono2 = DeRia.CoRia.Execute("Select * from FactAbono where substring(CtaIngreso,1, 6) = '112121' and numavpg = " & RsFact!NumAvPg & "")
                           ' Set RsIntAb = DeRia.CoRia.Execute("Select * from avpgdetalle where substring(CtaIngreso,1, 6) = '112121' and numavpg = " & RsAbono!numavpg & " ")
                                   
                                    If Not RsAbono2.EOF Then
                                    
                                        cRecargoSobreSaldo = cRecargoSobreSaldo + IIf(IsNull(RsAbono2!ValorUnitAvPgDet), 0, RsAbono2!ValorUnitAvPgDet) - IIf(IsNull(RsAbono2!ValorAbonado), 0, RsAbono2!ValorAbonado)
                                    End If
                         End If
                         

                         
                            UpdateIntRecDesc cDescuento, cinteres, cRecargoSobreSaldo, RsFact!AvPgTipoImpuesto, RsFact!NumAvPg
                            cinteres = 0: cRecargoPorMora = 0: cRecargoSobreSaldo = 0
                        Case 2
                            'Facturas generadas en Declaracion de Industria y Comercio
                            'No se le calcula recargo e Interes a una cuenta que sirve de permiso de operacion.
                            'Ahora veamos si no es de permiso de operacion. Esta tampoco tiene Recargo o Interes.
                            'Esto se remueve en validacion de CEIBA
                            'Set rsCuenta = DeRia.CoRia.Execute("select count(*) as Num from CuentaIngreso where CtaPermOP='" & rsDet!CtaIngreso & "' ")
                            'If rsCuenta!Num = 0 Then
                                ClearArrayCuotas
                                'Calculemos interes y recargo
                                ' VFVence = rsFact!FechaVenceAvPg
                                 If Not RsFecVen.EOF Then
                                  VFVence = "10/" & Format(VFVence, "mm") & "/" & Format(VFVence, "YYYY")
                                 End If
                                aCuotasMensuales(1, 0) = VFVence 'rsFact!FechaVenceAvPg
                                aCuotasMensuales(1, 1) = rsDet!ValorUnitAvPgDet
                                aCuotasMensuales(1, 4) = MesesVencidosIC(VFVence, rsPar!DiaProcesoCT)
                                MultaPorMesesNoPagadosIC (rsPar!DiaProcesoCT)
                                RecargoPorMesesNoPagadosIC (rsPar!DiaProcesoCT)
                                cDescuento = 0
                                
  '00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
                                        
        
        
 Str = " SELECT AvPgDetalle.CtaIngreso, AvPgEnc.FechaVenceAvPg, AvPgDetalle.NumAvPg  FROM AvPgDetalle INNER JOIN AvPgEnc ON AvPgDetalle.NumAvPg = AvPgEnc.NumAvPg "
Str = Str & "WHERE     (AvPgDetalle.NumAvPg = " & rsDet!NumAvPg & ")"
 
       Set RsDetUp = DeRia.CoRia.Execute(Str)
        
        
        Do While Not RsDetUp.EOF
     '---------------------------------------------------------------------------------------------------------------------------------
     Set RsTpo = DeRia.CoRia.Execute("Select Tipo from CuentaIngreso_A where Anio = " & Year(RsDetUp!FechaVenceAvPg) & " and CtaIngreso = '" & RsDetUp!CtaIngreso & "' ")
     
     If RsTpo.EOF Then
     Set RsTpo = DeRia.CoRia.Execute("Select Tipo from CuentaIngreso_A where CtaIngreso = '" & RsDetUp!CtaIngreso & "' ")
     End If
     
     strCtaIngreso = RsDetUp!CtaIngreso
     vCuenta = RsDetUp!CtaIngreso
     VFactSp = RsDetUp!NumAvPg
     
           If Format(RsDetUp!FechaVenceAvPg, "YYYY") < 2013 Then

                      If Mid(strCtaIngreso, 4, 3) = "118" And RsTpo!Tipo = 3 Then
                     strCtaIngreso = Mid(strCtaIngreso, 1, 3) & "117" & Mid(strCtaIngreso, 7, 5)
                     vCuenta = Mid(strCtaIngreso, 1, 3) & "117" & Mid(strCtaIngreso, 7, 5)
                     If vCuenta <> RsDetUp!CtaIngreso Then
                       DeRia.CoRia.Execute ("Update Avpgdetalle set CtaIngreso = '" & vCuenta & "' where CtaIngreso = '" & RsDetUp!CtaIngreso & "' and NumAvpg = " & RsDetUp!NumAvPg & " ")
                     End If
                   ' rsDet!CtaIngreso = strCtaIngreso
                  End If

                 '------------------------------------------------------------------------------------------------------------
            If Mid(strCtaIngreso, 4, 3) = "119" And RsTpo!Tipo = 1 Then
                     strCtaIngreso = Mid(strCtaIngreso, 1, 3) & "118" & Mid(strCtaIngreso, 7, 5)
                     vCuenta = Mid(strCtaIngreso, 1, 3) & "118" & Mid(strCtaIngreso, 7, 5)
                     If vCuenta <> RsDetUp!CtaIngreso Then
                      DeRia.CoRia.Execute ("Update Avpgdetalle set CtaIngreso = '" & vCuenta & "' where CtaIngreso = '" & RsDetUp!CtaIngreso & "' and NumAvpg = " & RsDetUp!NumAvPg & " ")
                     End If
                   ' rsDet!CtaIngreso = strCtaIngreso
                  End If

             Else
                  '------------------------------------------------------------------------------------------------------------
                  If Mid(strCtaIngreso, 4, 3) = "117" And RsTpo!Tipo = 3 Then
                     strCtaIngreso = Mid(strCtaIngreso, 1, 3) & "118" & Mid(strCtaIngreso, 7, 5)
                     vCuenta = Mid(strCtaIngreso, 1, 3) & "118" & Mid(strCtaIngreso, 7, 5)
                   If vCuenta <> RsDetUp!CtaIngreso Then
                      DeRia.CoRia.Execute ("Update Avpgdetalle set CtaIngreso = '" & vCuenta & "' where CtaIngreso = '" & RsDetUp!CtaIngreso & "' and NumAvpg = " & RsDetUp!NumAvPg & " ")
                   End If
                   ' rsDet!CtaIngreso = strCtaIngreso
                  End If
 
                  If Mid(strCtaIngreso, 4, 3) = "118" And RsTpo!Tipo = 1 Then
                     strCtaIngreso = Mid(strCtaIngreso, 1, 3) & "119" & Mid(strCtaIngreso, 7, 5)
                     vCuenta = Mid(strCtaIngreso, 1, 3) & "119" & Mid(strCtaIngreso, 7, 5)
                   If vCuenta <> RsDetUp!CtaIngreso Then
                      DeRia.CoRia.Execute ("Update Avpgdetalle set CtaIngreso = '" & vCuenta & "' where CtaIngreso = '" & RsDetUp!CtaIngreso & "' and NumAvpg = " & RsDetUp!NumAvPg & " ")
                   End If
                   ' rsDet!CtaIngreso = strCtaIngreso
                  End If
            End If
     
     
     '---------------------------------------------------------------------------------------------------------------------------------
        RsDetUp.MoveNext
        Loop
                                
'00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
         'Interes Abono
                        Set RsAbono = DeRia.CoRia.Execute("Select * from FactAbono where  numavpg = " & RsFact!NumAvPg & " ")
                         If Not RsAbono.EOF Then
                         Set RsAbono2 = DeRia.CoRia.Execute("Select * from FactAbono where substring(CtaIngreso,1, 6) = '112126' and numavpg = " & RsFact!NumAvPg & "")
                           ' Set RsIntAb = DeRia.CoRia.Execute("Select * from avpgdetalle where substring(CtaIngreso,1, 6) = '112126' and numavpg = " & RsAbono!numavpg & " ")
                                   
                                    If Not RsAbono2.EOF Then
                                    
                                        cinteres = cinteres + IIf(IsNull(RsAbono2!ValorUnitAvPgDet), 0, RsAbono2!ValorUnitAvPgDet) - IIf(IsNull(RsAbono2!ValorAbonado), 0, RsAbono2!ValorAbonado)
                                    End If
                         End If
                         
                       'Recargo Abono
                       
                       
                       Set RsAbono = DeRia.CoRia.Execute("Select * from FactAbono where  numavpg = " & RsFact!NumAvPg & " ")
                         If Not RsAbono.EOF Then
                         Set RsAbono2 = DeRia.CoRia.Execute("Select * from FactAbono where substring(CtaIngreso,1, 6) = '112121' and numavpg = " & RsFact!NumAvPg & "")
                           ' Set RsIntAb = DeRia.CoRia.Execute("Select * from avpgdetalle where substring(CtaIngreso,1, 6) = '112121' and numavpg = " & RsAbono!numavpg & " ")
                                   
                                    If Not RsAbono2.EOF Then
                                    
                                        cRecargoSobreSaldo = cRecargoSobreSaldo + IIf(IsNull(RsAbono2!ValorUnitAvPgDet), 0, RsAbono2!ValorUnitAvPgDet) - IIf(IsNull(RsAbono2!ValorAbonado), 0, RsAbono2!ValorAbonado)
                                    End If
                         End If
                         
             Dim RsSP As New ADODB.Recordset
             
                                
                                UpdateIntRecDesc cDescuento, Val(aCuotasMensuales(1, 5)), Val(aCuotasMensuales(1, 7)), RsFact!AvPgTipoImpuesto, RsFact!NumAvPg
                                        
                                        
                                    
                   UpdateIntRecDesc cDescuento, Val(aCuotasMensuales(1, 5)), Val(aCuotasMensuales(1, 7)), RsFact!AvPgTipoImpuesto, RsFact!NumAvPg
                                        
                                        
                                        
                        Case 3
                            'Impuesto Personal
                            
                            'VFVence = rsFact!FechaVenceAvPg
                            
                            cDescuento = ComputeDescuento(rsDet!ValorUnitAvPgDet * rsDet!CantAvPgDet, VFVence, rsPar!DiaProcesoCT)
                            cDescuento = cDescuento * -1  'el descuento va negativo en la factura
                            cinteres = CalculeIntMoraIP(rsDet!ValorUnitAvPgDet, rsPar!DiaProcesoCT, VFVence)
                            cRecargoSobreSaldo = CalculeRecargoSobreSaldoIP(rsDet!ValorUnitAvPgDet + cinteres, rsPar!DiaProcesoCT, VFVence)
                            
                            
                                  'Interes Abono
                        Set RsAbono = DeRia.CoRia.Execute("Select * from FactAbono where  numavpg = " & RsFact!NumAvPg & " ")
                         If Not RsAbono.EOF Then
                         Set RsAbono2 = DeRia.CoRia.Execute("Select * from FactAbono where substring(CtaIngreso,1, 6) = '112126' and numavpg = " & RsFact!NumAvPg & "")
                           ' Set RsIntAb = DeRia.CoRia.Execute("Select * from avpgdetalle where substring(CtaIngreso,1, 6) = '112126' and numavpg = " & RsAbono!numavpg & " ")
                                   
                                    If Not RsAbono2.EOF Then
                                    
                                        cinteres = cinteres + IIf(IsNull(RsAbono2!ValorUnitAvPgDet), 0, RsAbono2!ValorUnitAvPgDet) - IIf(IsNull(RsAbono2!ValorAbonado), 0, RsAbono2!ValorAbonado)
                                    End If
                         End If
                         
                       'Recargo Abono
                       
                       
                       Set RsAbono = DeRia.CoRia.Execute("Select * from FactAbono where  numavpg = " & RsFact!NumAvPg & " ")
                         If Not RsAbono.EOF Then
                         Set RsAbono2 = DeRia.CoRia.Execute("Select * from FactAbono where substring(CtaIngreso,1, 6) = '112121' and numavpg = " & RsFact!NumAvPg & "")
                           ' Set RsIntAb = DeRia.CoRia.Execute("Select * from avpgdetalle where substring(CtaIngreso,1, 6) = '112121' and numavpg = " & RsAbono!numavpg & " ")
                                   
                                    If Not RsAbono2.EOF Then
                                    
                                         cRecargoSobreSaldo = cRecargoSobreSaldo + IIf(IsNull(RsAbono2!ValorUnitAvPgDet), 0, RsAbono2!ValorUnitAvPgDet) - IIf(IsNull(RsAbono2!ValorAbonado), 0, RsAbono2!ValorAbonado)
                                    End If
                         End If
                         
                       
                            
                            UpdateIntRecDesc cDescuento, cinteres, cRecargoSobreSaldo, RsFact!AvPgTipoImpuesto, RsFact!NumAvPg
                                                      
                                                      
                        Case 4
                        
                        'VFVence = rsFact!FechaVenceAvPg
                            'Impuesto Personal
                            cDescuento = ComputeDescuento(rsDet!ValorUnitAvPgDet * rsDet!CantAvPgDet, VFVence, rsPar!DiaProcesoCT)
                            cDescuento = cDescuento * -1  'el descuento va negativo en la factura
                            cinteres = CalculeIntMoraIP(rsDet!ValorUnitAvPgDet, rsPar!DiaProcesoCT, VFVence)
                            cRecargoSobreSaldo = CalculeRecargoSobreSaldoIP(rsDet!ValorUnitAvPgDet + cinteres, rsPar!DiaProcesoCT, VFVence)
                            
                         
                         
                                  'Interes Abono
                        Set RsAbono = DeRia.CoRia.Execute("Select * from FactAbono where  numavpg = " & RsFact!NumAvPg & " ")
                         If Not RsAbono.EOF Then
                         Set RsAbono2 = DeRia.CoRia.Execute("Select * from FactAbono where substring(CtaIngreso,1, 6) = '112126' and numavpg = " & RsFact!NumAvPg & "")
                           ' Set RsIntAb = DeRia.CoRia.Execute("Select * from avpgdetalle where substring(CtaIngreso,1, 6) = '112126' and numavpg = " & RsAbono!numavpg & " ")
                                   
                                    If Not RsAbono2.EOF Then
                                    
                                        cinteres = cinteres + IIf(IsNull(RsAbono2!ValorUnitAvPgDet), 0, RsAbono2!ValorUnitAvPgDet) - IIf(IsNull(RsAbono2!ValorAbonado), 0, RsAbono2!ValorAbonado)
                                    End If
                         End If
                         
                       'Recargo Abono
                       
                       
                       Set RsAbono = DeRia.CoRia.Execute("Select * from FactAbono where  numavpg = " & RsFact!NumAvPg & " ")
                         If Not RsAbono.EOF Then
                         Set RsAbono2 = DeRia.CoRia.Execute("Select * from FactAbono where substring(CtaIngreso,1, 6) = '112121' and numavpg = " & RsFact!NumAvPg & "")
                           ' Set RsIntAb = DeRia.CoRia.Execute("Select * from avpgdetalle where substring(CtaIngreso,1, 6) = '112121' and numavpg = " & RsAbono!numavpg & " ")
                                   
                                    If Not RsAbono2.EOF Then
                                    
                                         cRecargoSobreSaldo = cRecargoSobreSaldo + IIf(IsNull(RsAbono2!ValorUnitAvPgDet), 0, RsAbono2!ValorUnitAvPgDet) - IIf(IsNull(RsAbono2!ValorAbonado), 0, RsAbono2!ValorAbonado)
                                    End If
                         End If
                         
                     
                            UpdateIntRecDesc cDescuento, cinteres, cRecargoSobreSaldo, RsFact!AvPgTipoImpuesto, RsFact!NumAvPg
                            
                        Case 7
                                 
                                 
                        Case 5


                    End Select
                Case Is = 0
                    'No hay mora ni descuentos. Los ponemos en cero
                      vFactVen = RsFact!NumAvPg
                    Select Case RsFact!AvPgTipoImpuesto
                        Case 1
                            'Bienes Inmuebles
                            cinteres = 0: cRecargoPorMora = 0: cRecargoSobreSaldo = 0
                            UpdateIntRecDesc cDescuento, cinteres, cRecargoSobreSaldo, RsFact!AvPgTipoImpuesto, RsFact!NumAvPg
                        Case 2
                            'Declaraciones de industria y comercio, ponemos en cero toda la mora
                            cDescuento = 0
                            UpdateIntRecDesc cDescuento, 0, 0, RsFact!AvPgTipoImpuesto, RsFact!NumAvPg
                        Case 4
                            'Impuesto Personal
                            cDescuento = 0
                            cinteres = 0
                            cRecargoSobreSaldo = 0
                            UpdateIntRecDesc cDescuento, cinteres, cRecargoSobreSaldo, RsFact!AvPgTipoImpuesto, RsFact!NumAvPg
                        Case 5


                    End Select
                Case Is < 1
                    'Hay descuentos
                    
                    Dim VFecDesc As Date
                    
                        If VQuita = 2 Then
    VFecDesc = VFechaQuita
    End If
                    
                    Select Case RsFact!AvPgTipoImpuesto
                        Case 1
                        
                        
 '-------------------------------------------------------------------------------------------------------------------------------------------
         'Para Servicios Publicos
         
 
 '-------------------------------------------------------------------------------------------------------------------------------------------
  Set rsDescuenta = DeRia.CoRia.Execute("SELECT SUM(A.ValorUnitAvPgDet * A.CantAvPgDet) AS total FROM  AvPgDetalle AS A INNER JOIN CuentaIngreso_A AS B ON A.CtaIngreso = B.CtaIngreso WHERE (A.NumAvPg = " & RsFact!NumAvPg & ") AND (B.Tipo <> 2) AND ANIO = " & Format(RsFact!FechaVenceAvPg, "YYYY") & "")
                           
 '-------------------------------------------------------------------------------------------------------------------------------------------
 'Bienes Inmuebles
                            cinteres = 0: cRecargoPorMora = 0: cRecargoSobreSaldo = 0
                            
                                                            If VQuita = 2 Then
                                                          VFecDesc = VFechaQuita
                                                         Else
    
                                                            VFecDesc = rsPar!DiaProcesoCT
                                                        End If
                            
                            cDescuento = CalculeDescuentoBI(IIf(IsNull(rsDescuenta!Total), 0, rsDescuenta!Total), RsFact!FechaVenceAvPg, VFecDesc)
                            'cDescuento = CalculeDescuentoBI(rsDet!ValorUnitAvPgDet * rsDet!CantAvPgDet, RsFact!FechaVenceAvPg, rsPar!DiaProcesoCT)
                            cDescuento = cDescuento * -1
                            UpdateIntRecDesc cDescuento, cinteres, cRecargoSobreSaldo, RsFact!AvPgTipoImpuesto, RsFact!NumAvPg
                            cinteres = 0: cRecargoPorMora = 0: cRecargoSobreSaldo = 0: cDescuento = 0
                        Case 2, 3
                            'Facturas generadas en Declaracion de Industria y Comercio
                            'No se le calcula recargo e Interes a una cuenta que sirve de permiso de operacion.
                            'Ahora veamos si no es de permiso de operacion. Esta tampoco tiene Recargo o Interes.
                            Set rsCuenta = DeRia.CoRia.Execute("select count(*) as Num from CuentaIngreso_A where CtaPermOP='" & rsDet!CtaIngreso & "' AND ANIO = " & Format(RsFact!FechaVenceAvPg, "YYYY") & "")
                            If rsCuenta!Num = 0 Then
                            'Suma total cuentas para descuento '07 Enero 2011
                            
                            Set rsDescuenta = DeRia.CoRia.Execute("SELECT SUM(A.ValorUnitAvPgDet * A.CantAvPgDet) AS total FROM  AvPgDetalle AS A INNER JOIN CuentaIngreso_A AS B ON A.CtaIngreso = B.CtaIngreso WHERE (A.NumAvPg = " & RsFact!NumAvPg & ") AND (B.Tipo <> 2) AND ANIO = " & Format(RsFact!FechaVenceAvPg, "YYYY") & "")
                               
                                                       If VQuita = 2 Then
                                                          VFecDesc = VFechaQuita
                                                         Else
    
                                                            VFecDesc = rsPar!DiaProcesoCT
                                                        End If
                               
                               
                                cDescuento = ComputeDescuento(IIf(IsNull(rsDescuenta!Total), 0, rsDescuenta!Total), RsFact!FechaVenceAvPg, VFecDesc) 'rsDet!ValorUnitAvPgDet * rsDet!CantAvPgDet
                                cDescuento = cDescuento * -1
                                UpdateIntRecDesc cDescuento, 0, 0, RsFact!AvPgTipoImpuesto, RsFact!NumAvPg
                                cDescuento = 0
                            End If
                        Case 4
                            'Impuesto Personal
                            cDescuento = ComputeDescuento(rsDet!ValorUnitAvPgDet * rsDet!CantAvPgDet, RsFact!FechaVenceAvPg, rsPar!DiaProcesoCT)
                            cDescuento = cDescuento * -1  'el descuento va negativo en la factura
                            cinteres = 0
                            cRecargoSobreSaldo = 0
                            UpdateIntRecDesc cDescuento, cinteres, cRecargoSobreSaldo, RsFact!AvPgTipoImpuesto, RsFact!NumAvPg
                            cDescuento = 0
                        Case 5
                        
                        Case 7 ' Interes plandes de pago

              
                            
                    End Select

                End Select
                rsDet.MoveNext
            Loop
        End If
        'Cambiamos el codigo de la cuenta de ingreso a recuperacion para facturas
        'de años anteriores
        ApliqueCtaRecuperacion RsFact!NumAvPg, rsPar!DiaProcesoCT, RsFact!FechaVenceAvPg, DeRia.CoRia
        '------------------------------------------------------------------------
        If RsFact.EOF Then Exit Sub
        RsFact.MoveNext
    Loop
End Sub
Public Function MesesVencidos(ByVal dFechaVence As Date, ByVal dFechaActual As Date) As Single
    'calcula los meses de diferencia entre fechavence y fechaactual, tomando el dia
    'que se manda, como referencia para cambiar de mes.
    'Esta rutina es especial para calcular meses vencidos.
    Dim RsVen As New ADODB.Recordset
    
     If VQuita = 1 Then
    Qdia = "09" 'Format(dDiaProcesoCT, "DD")
    QMes = Format(dFechaActual, "MM")
    QAno = Format(dFechaActual, "YYYY")
    If QMes = "01" Then
        QMes = 12
        QAno = QAno - 1
    Else
    QMes = Val(QMes) - 1
    If Len(QMes) = 1 Then
    QMes = 0 & QMes
    End If
    End If
    dFechaActual = Qdia & "/" & QMes & "/" & QAno
    End If
    
    If VQuita = 2 Then
      dFechaActual = VFechaQuita
    End If
    
    Dim dFecha1 As Date 'para formar el ultimo dia del mes
    Dim dDiaProcesoCT As Date 'Para Quitar un mes de interes
    
    MesesVencidos = 0
    If dFechaVence < dFechaActual Then 'ya esta vencido

 'Set RsVen = DeRia.CoRia.Execute("Select DATEDIFF(month, '" & dFechaVence & "', '" & dFechaActual & "') AS Meses from AvPgEnc where NumAvPg = " & VFactNum & "")
  
'''' Set RsVen = DeRia.CoRia.Execute("Select top(1) DATEDIFF(month, '" & dFechaVence & "', '" & dFechaActual & "') AS Meses from AvPgEnc ")
''''  If Not RsVen.EOF Then
''''     MesesVencidos = RsVen!Meses
''''  End If
''''
If Val(Format(dFechaVence, "DD")) > 10 Then
  MesesVencidos = DateDiff("m", dFechaVence, dFechaActual)
Else
  MesesVencidos = DateDiff("m", dFechaVence, dFechaActual) + 1
End If
  
'        Do While dFechaVence < dFechaActual
'            MesesVencidos = MesesVencidos + 1
'            dFechaVence = DateAdd("m", 1, dFechaVence)
'            'calculamos el ultimo dia del mes
'            dFecha1 = DateAdd("d", -Day(dFechaVence) + 1, dFechaVence) 'calculamos el 1er dia del mes
'           dFecha1 = DateAdd("m", 1, dFecha1) 'calculamos el 1er dia del siguiente mes
'            dFecha1 = DateAdd("d", -1, dFecha1) 'restamos un dia, y tenemos el ultimo dia del mes
'            dFechaVence = dFecha1
'        Loop

 
        Exit Function
    End If
    
    If dFechaVence > dFechaActual Then 'significa que no ha vencido
        
         '  Set RsVen = DeRia.CoRia.Execute("Select DATEDIFF(month, '" & dFechaVence & "', '" & dFechaActual & "') AS Meses from AvPgEnc where NumAvPg = " & VFactNum & "")
         
''''         Set RsVen = DeRia.CoRia.Execute("Select top(1) DATEDIFF(month, '" & dFechaVence & "', '" & dFechaActual & "') AS Meses from AvPgEnc ")
''''  If Not RsVen.EOF Then
''''     MesesVencidos = RsVen!Meses
''''  End If
  
  
'        Do While dFechaVence > dFechaActual
'            MesesVencidos = MesesVencidos - 1
'            dFechaVence = DateAdd("m", -1, dFechaVence)
'        Loop
'
 MesesVencidos = DateDiff("m", dFechaVence, dFechaActual)
        
        Exit Function
    End If
End Function
Public Function MesesVencidosIC_reemplazadaporladeabajo(ByVal dFechaVence As Date, ByVal dFechaActual As Date) As Single
    'calcula los meses de diferencia entre fechavence y fechaactual, tomando el dia
    'que se manda, como referencia para cambiar de mes.
    'Esta rutina es especial para calcular meses vencidos.
    'hicimos una rutina especial para Industria y Comercio porque Enero se trata de forma
    'diferente.
    Dim dEnero As Date
    Dim sYearMonthVence As String
    Dim sYearMonthActual As String
    
    '========MesesVencidosIC = 0
    'Enero recibe trato especial
    dEnero = "31/01/" & Year(dFechaVence)
    If dFechaVence < dFechaActual Then
        '======MesesVencidosIC = 1
    Else
        Exit Function
    End If
    '--------------------------
    'Pasemos al siguiente mes
    dFechaVence = DateAdd("m", 1, dFechaVence)
    '--------------------------
    sYearMonthVence = Year(dFechaVence) & Format(Month(dFechaVence), "00") 'Debe ir formato AAAAMM
    sYearMonthActual = Year(dFechaActual) & Format(Month(dFechaActual), "00") 'Debe ir formato AAAAMM
    If Val(sYearMonthVence) <= Val(sYearMonthActual) Then
        'ya esta vencido
        Do While Val(sYearMonthVence) <= Val(sYearMonthActual)
            '=====MesesVencidosIC = MesesVencidosIC + 1
            dFechaVence = DateAdd("m", 1, dFechaVence)
            sYearMonthVence = Year(dFechaVence) & Format(Month(dFechaVence), "00")
        Loop
        Exit Function
    End If
End Function
Public Function MesesVencidosIC(ByVal dFechaVence As Date, ByVal dFechaActual As Date) As Single
    'calcula los meses de diferencia entre fechavence y fechaactual, tomando el dia
    'que se manda, como referencia para cambiar de mes.
    'Esta rutina es especial para calcular meses vencidos.
    'hicimos una rutina especial para Industria y Comercio porque Enero se trata de forma
    'diferente.
    Dim dEnero As Date
    Dim sYearMonthVence As String
    Dim sYearMonthActual As String
    
    MesesVencidosIC = 0
    'Enero recibe trato especial
    dEnero = "31/01/" & Year(dFechaVence)
    If dFechaVence < dFechaActual Then
        MesesVencidosIC = 1
    Else
        Exit Function
    End If
    '--------------------------
    'Pasemos al siguiente mes
    dFechaVence = DateAdd("m", 1, dFechaVence)
    dFechaVence = CDate("10/" & Month(dFechaVence) & "/" & Year(dFechaVence))
    '--------------------------
    'sYearMonthVence = Year(dFechaVence) & Format(Month(dFechaVence), "00") 'Debe ir formato AAAAMM
    'sYearMonthActual = Year(dFechaActual) & Format(Month(dFechaActual), "00") 'Debe ir formato AAAAMM
    If dFechaVence < dFechaActual Then
        'ya esta vencido
        Do While dFechaVence < dFechaActual
            MesesVencidosIC = MesesVencidosIC + 1
            dFechaVence = DateAdd("m", 1, dFechaVence)
            'sYearMonthVence = Year(dFechaVence) & Format(Month(dFechaVence), "00")
        Loop
        
        
        Exit Function
    End If
End Function

Function CuentasMillar(lv_Valor As Currency, VRCtaIng As String, VRAnio As Integer, VAper As Integer)
    Dim Impuesto As Single
    Dim ValorDesde As Single
    Dim ValorHasta As Single
    Dim Miles As Single

    Impuesto = 0
    'VRVal2 = 0
    If DeRia.rsCmdRangoBiSP.State = adStateOpen Then
        DeRia.rsCmdRangoBiSP.Close
    End If
    
  '  strSql = "SELECT De AS ValorMinimo, Hasta AS ValorMaximo, Valor FROM CuentaIngreso_R WHERE (De <= ?) AND (CtaIngreso = ?) AND (Anio = ?) AND Apertura = ? ORDER BY ValorMinimo "
    
    DeRia.CmdRangoBiSP (lv_Valor), (VRCtaIng), (VRAnio), VAper  'Seleccione los rangos de la tabla
    If Not DeRia.rsCmdRangoBiSP.EOF Then
    ' Do While Not DeRia.rsCmdRangoBiSP.EOF()
        ValorDesde = DeRia.rsCmdRangoBiSP!ValorMinimo - 1  'Para que tome correctamente los miles
        If DeRia.rsCmdRangoBiSP!ValorMaximo <= lv_Valor Then
            ValorHasta = DeRia.rsCmdRangoBiSP!ValorMaximo    'Antes o igual al valor de ingresos
        Else
            ValorHasta = lv_Valor   ' Cuando el valor de ingreso esta dentro de un rango especifico
        End If
        'Miles = (ValorHasta - ValorDesde) / 1000
        Miles = (ValorHasta) / 1000
        Impuesto = Impuesto + (Miles * DeRia.rsCmdRangoBiSP!valor)
        DeRia.rsCmdRangoBiSP.MoveNext
    End If
    'Loop
   ' CalculeIP = Impuesto
   VRVal2 = Impuesto
End Function





Function CalculeIP(lv_Valor As Currency)
    Dim Impuesto As Single
    Dim ValorDesde As Single
    Dim ValorHasta As Single
    Dim Miles As Single
    
    Impuesto = 0
    If DeRia.rscmdRangosIP.State = adStateOpen Then
        DeRia.rscmdRangosIP.Close
    End If
    DeRia.cmdRangosIP (lv_Valor) 'Seleccione los rangos de la tabla
    Do While Not DeRia.rscmdRangosIP.EOF()
        ValorDesde = DeRia.rscmdRangosIP!ValorMinimo - 1  'Para que tome correctamente los miles
        If DeRia.rscmdRangosIP!ValorMaximo <= lv_Valor Then
            ValorHasta = DeRia.rscmdRangosIP!ValorMaximo    'Antes o igual al valor de ingresos
        Else
            ValorHasta = lv_Valor   ' Cuando el valor de ingreso esta dentro de un rango especifico
        End If
        'Miles = Int((ValorHasta - ValorDesde) / 1000)
        Miles = (ValorHasta - ValorDesde) / 1000
        Impuesto = Impuesto + (Miles * DeRia.rscmdRangosIP!ValorMillar)
        DeRia.rscmdRangosIP.MoveNext
    Loop
    CalculeIP = Impuesto
End Function
Function CalculeIC(lv_Valor As Currency) As Currency
    Dim Impuesto As Single
    Dim ValorDesde As Single
    Dim ValorHasta As Single
    Dim Miles As Single
    Dim RsRangos As New Recordset
    Dim sSql As String
    
    Impuesto = 0
    'Seleccione los rangos de la tabla de valores
    sSql = "SELECT ValorMinimo, ValorMaximo, ValorxMillar " & _
    "FROM TablaImpCom WHERE ValorMinimo <= " & lv_Valor & " ORDER BY ValorMinimo"
    Set RsRangos = DeRia.CoRia.Execute(sSql)
    '--------------------------------------------
    Do While RsRangos.EOF = False
        ValorDesde = RsRangos!ValorMinimo - 1  'Para que tome correctamente los miles
        If RsRangos!ValorMaximo <= lv_Valor Then
            ValorHasta = RsRangos!ValorMaximo    'Antes o igual al valor de ingresos
        Else
            ValorHasta = lv_Valor   ' Cuando el valor de ingreso esta dentro de un rango especifico
        End If
        Miles = (ValorHasta - ValorDesde) / 1000
        Impuesto = Impuesto + (Miles * RsRangos!ValorxMillar)
        RsRangos.MoveNext
    Loop
    CalculeIC = Impuesto
End Function
Function CalculeICReg(lv_Valor As Currency) As Currency
    Dim Impuesto As Single
    Dim ValorDesde As Single
    Dim ValorHasta As Single
    Dim Miles As Single
    Dim RsRangos As New Recordset
    Dim sSql As String
    
    Impuesto = 0
    'Seleccione los rangos de la tabla de valores
    sSql = "SELECT ValorMinimo, ValorMaximo, ValorxMillar " & _
    "FROM TablaImpProdReg WHERE ValorMinimo <= " & lv_Valor & " ORDER BY ValorMinimo"
    Set RsRangos = DeRia.CoRia.Execute(sSql)
    '--------------------------------------------
    Do While RsRangos.EOF = False
        ValorDesde = RsRangos!ValorMinimo - 1  'Para que tome correctamente los miles
        If RsRangos!ValorMaximo <= lv_Valor Then
            ValorHasta = RsRangos!ValorMaximo    'Antes o igual al valor de ingresos
        Else
            ValorHasta = lv_Valor   ' Cuando el valor de ingreso esta dentro de un rango especifico
        End If
        Miles = Int((ValorHasta - ValorDesde) / 1000)
        Impuesto = Impuesto + (Miles * RsRangos!ValorxMillar)
        RsRangos.MoveNext
    Loop
    CalculeICReg = Impuesto
End Function
Public Sub AgregarFacturas(rsEnc As Recordset, rsDet As Recordset, cnnFact As ADODB.Connection, lResult As Boolean)
'Se recibe un recordset encabezado y un recordset detalle donde temporalmente estan los valores
'a insertar en las tablas de facturas.
'Todo el proceso se hace en un batch, mientras este batch ocurre la tabla de parametros
'donde se guarda el siguiente numero de factura, debe estar bloqueado a lectura/escritura, asi
'se impide que otra factura se genere, por otro proceso o usuario.
'Proceso:
'1.- Lea el siguiente número de factura, y haga un bloqueo de lectura/escritura.
'2.- Para cada factura en el encabezado.
'3.-    inserte el encabezado de la factura
'4.-    seleccione el detalle de la factura temporal
'5.-    para cada detalle
'7.-        inserte el detalle de la factura
'8-     siguiente detalle
'9-     Aumente el numero de factura
'9-  Siguiente factura
'10- actualize el ultimo numero de factura en parametros y cierre el recordset bloqueado

    Dim rsPar As New ADODB.Recordset
    Dim lNumFact As Long
    Dim sSql As String
    Dim sBoolean As String
    Dim VANioFactS As Integer
    
    On Error GoTo AgregarFacturas_Error
    
    rsPar.Open "Select * from ParametroCont", cnnFact, _
    adOpenKeyset, adLockPessimistic, adCmdText
    'tambien se puede bloquear de la siguiente forma
    'rsPar.Open "ParametroCont", cnnFact, adOpenKeyset, adLockPessimistic, adCmdTable
    'See how database is opened too
    
    'Try to lock record
LockRecord:
    On Error GoTo LockError
    rsPar.Requery
    rsPar!UltNumFact = rsPar!UltNumFact + 1
    'Record succesfully locked, sino va a LockError
    rsPar!UltNumFact = rsPar!UltNumFact - 1 'regresamos el numero al estado original
    '------------------
    On Error GoTo AgregarFacturas_Error
    If IsNull(rsPar!UltNumFact) Then
        rsPar!UltNumFact = 1
    End If
    rsEnc.MoveFirst
    Do While rsEnc.EOF = False
        'Aumentamos en 1 el numero de factura
        rsPar!UltNumFact = rsPar!UltNumFact + 1
        '------------------------------------
        'Grabamos la factura MSx3 Fechas modificado Format
        sSql = "insert into AvPgEnc (NumAvPg,Identidad,FechaEmAvPg,FechaVenceAvPg," & _
        "AvPgEstado,TipoAvPg,AvPgTipoImpuesto,AvPgDescripcion,CodDeclara, " & _
        "AvPgTotalPeriodo,ClaveCatastro,CreadoPor,FechaCreado,ModificadoPor,FechaModificado) " & _
        "Values(" & rsPar!UltNumFact & ",'" & rsEnc!Identidad & "','" & Format(rsEnc!FechaEmAvPg, "dd/mm/yyyy") & "', " & _
        "'" & Format(rsEnc!FechaVenceAvPg, "dd/mm/yyyy") & "'," & rsEnc!AvPgEstado & "," & rsEnc!TipoAvPg & ", " & _
        "" & rsEnc!AvPgTipoImpuesto & ",'" & rsEnc!AvPgDescripcion & "'," & _
        "'" & rsEnc!CodDeclara & "'," & rsEnc!AvPgTotalPeriodo & ", '" & rsEnc!ClaveCatastro & "'," & _
        "'" & rsEnc!CreadoPor & "','" & Format(rsEnc!FechaCreado, "dd/mm/yyyy") & "', " & _
        "'" & rsEnc!ModificadoPor & "', '" & Format(rsEnc!FechaModificado, "dd/mm/yyyy") & "') "
        'MsgBox sSql
        cnnFact.Execute (sSql)
        rsDet.MoveFirst
        
        Do While rsDet.EOF = False
            If rsDet!NumAvPg = rsEnc!NumAvPg Then
                If rsDet!VisibleEnTesAvPgDet = False Then sBoolean = "False" Else sBoolean = "True"
                
               ''' MSx3
                If sBoolean = False Then
                sBoolean = 0
                ElseIf sBoolean = True Then
                sBoolean = 1
                End If
                ''''''
                
                sSql = "insert into AvPgDetalle " & _
                "(NumAvPg,ValorUnitAvPgDet,ClaveCatastro," & _
                "CtaIngreso,CantAvPgDet,RefAvPgDet," & _
                "DescuentoAvPgDet,RecargoAvPgDet," & _
                "ValorPagadoAvPgDet,VisibleEnTesAvPgDet," & _
                "ValorXAvPgDet) " & _
                "Values(" & rsPar!UltNumFact & "," & rsDet!ValorUnitAvPgDet & ",'" & rsDet!ClaveCatastro & "'," & _
                "'" & rsDet!CtaIngreso & "'," & rsDet!CantAvPgDet & ",'" & rsDet!RefAvPgDet & "', " & _
                "" & rsDet!DescuentoAvPgDet & "," & rsDet!RecargoAvPgDet & "," & _
                "" & rsDet!ValorPagadoAvPgDet & "," & sBoolean & "," & _
                "" & rsDet!ValorXAvPgDet & ")"
                'MsgBox sSql
                cnnFact.Execute (sSql)
      'Round(rsDet!ValorUnitAvPgDet, 2)
      VANioFactS = Format(rsEnc!FechaVenceAvPg, "yyyy")
If Not rsEnc!AvPgTipoImpuesto = 7 Then
 If VANioFactS < 2013 Then ' ==============================='Cargue 2013
    If rsDet!CtaIngreso Like "11111821*" Or rsDet!CtaIngreso Like "11212208*" Then  'rsEnc!AvPgTipoImpuesto <> 0 Or
        rsDet.MoveNext
    End If
    
 Else
 If rsDet!CtaIngreso Like "11111921*" Or rsDet!CtaIngreso Like "11212209*" Then  'rsEnc!AvPgTipoImpuesto <> 0 Or
        rsDet.MoveNext
    End If
    
 
 End If
    
End If

'    If Not rsDet!CtaIngreso Like "111112*" Or Not rsDet!CtaIngreso Like "111113*" Or Not rsDet!CtaIngreso Like "111114*" Then
'        rsDet.MoveNext
'    End If
    
            End If
            If Not rsDet.EOF Then
            rsDet.MoveNext
            End If
        Loop
        'Ponemos el numero real de la factura, para referencias fuera de este procedure
        'como por ejemplo los planes de pago.
        rsEnc!NumAvPg = rsPar!UltNumFact
        rsEnc.MoveNext
    Loop
    rsPar.Update 'release lock
    rsPar.Close

AgregarFacturas_Exit:
    Exit Sub

AgregarFacturas_Error:
    MsgBox Err.Number & ":" & Err.Description
    lResult = False
    GoTo AgregarFacturas_Exit
    
LockError:
    Select Case cnnFact.Errors(0).SQLState
    Case 3197, 3260, 3218
        'record is locked or has changed since opened.
        If MsgBox(cnnFact.Errors(0).Description & Chr(13) & Chr(10) _
        & "Bloqueo activado por otro usuario. Desea volver a intentar grabar los datos?", vbQuestion + vbYesNo, "Registro Bloqueado") = vbYes Then
            Resume LockRecord
        Else
            Resume AgregarFacturas_Error
        End If
    Case Else
        MsgBox Err.Number & ":" & Err.Description
        Resume AgregarFacturas_Error
    End Select
   
End Sub
Function CalculeRecargoPorMora(cImpuesto As Currency, cMulta As Currency, dFechaVence As Date, dFechaPago As Date, rs As Recordset) As Currency
'Vea manual de pseudocode para descripcion de funcion
'Proceso llamado desde: Declaracion Impuesto Personal

    Dim dMesA As Date
    Dim cRecargoA As Currency
    Dim iMeses, i As Integer
    Dim rsPar As Recordset
    Dim sSql As String
    
    'Defina el recordset en blanco, que mantiene las fechas vencidas
    rs.Fields.Append "Fecha", adDate
    rs.Fields.Append "Impuesto", adCurrency
    rs.Fields.Append "MultaNoDeclara", adCurrency
    rs.Fields.Append "RecMora", adCurrency
    rs.Fields.Append "RecSobreSaldo", adCurrency
    rs.LockType = adLockOptimistic
    rs.CursorLocation = adUseServer
    rs.Open
    '---------------------------------------------------------------
    CalculeRecargoPorMora = 0
    cRecargoA = 0
    iMeses = DateDiff("m", dFechaVence, dFechaPago)
    dMesA = dFechaVence
    For i = 1 To iMeses
        dMesA = DateAdd("m", 1, dMesA)
        Set rsPar = DeRia.CoRia.Execute("select RecargoAtrasopagoSL from ParamRia where PeriodoFact=" & Year(dMesA) & " ")
        If rsPar.RecordCount = 0 Or IsNull(rsPar!RecargoAtrasoPagoSL) Then
            MsgBox "Error al obtener el porcentaje de recargo por mora, de los parametros del año " & Year(dMesA)
            Exit Function
        End If
        cRecargoA = cRecargoA + (rsPar!RecargoAtrasoPagoSL * cImpuesto)
        rs.AddNew
        rs!Fecha = dMesA: rs!Impuesto = cImpuesto: rs!MultaNoDeclara = cMulta: rs!RecMora = cRecargoA: rs!RecSobreSaldo = 0
        rs.Update
    Next i
    CalculeRecargoPorMora = cRecargoA
End Function
Function CalculeRecargoSobreSaldo_(rs As Recordset) As Currency
'Vea manual de pseudocode para descripcion de funcion
'Proceso llamado desde: Declaracion Impuesto Personal
'No se esta usando
    
    Dim cRecargoA, cinteres, cSaldo As Currency
    Dim rsPar As Recordset
    
    cRecargoA = 0
    If rs.RecordCount > 0 Then rs.MoveFirst
    Do While rs.EOF = False
        Set rsPar = DeRia.CoRia.Execute("select Recargosobresaldo from ParamRia where PeriodoFact=" & Year(rs!Fecha) & " ")
        If rsPar.RecordCount = 0 Or IsNull(rsPar!RecargoSobreSaldo) Then
            MsgBox "Error al obtener el porcentaje de recargo sobre saldo, de los parametros del año " & Year(rs!Fecha)
            Exit Function
        End If
        cSaldo = rs!Impuesto + rs!RecMora
        cinteres = Round((rsPar!RecargoSobreSaldo / 12), 5)
        cRecargoA = cRecargoA + (cinteres * cSaldo)
        'MsgBox "Fecha:" & rs!fecha & ", Saldo:" & cSaldo & ", Recargo: " & cRecargoA
        rs.MoveNext
    Loop
    CalculeRecargoSobreSaldo_ = cRecargoA
End Function
Function FacturaFinanciada(lNumFact As Long) As Boolean
'Verifique si esta factura, no es producto de un plan de pago

    Dim rsPP As New Recordset
    Set rsPP = DeRia.CoRia.Execute("select * from PlanPagoFactura where NumAvPg=" & lNumFact & " ")
    If rsPP.RecordCount = 0 Then FacturaFinanciada = False Else FacturaFinanciada = True
End Function
Public Sub CalculeRecargoBI(cImpuesto As Currency, dFechaFact As Date, dFechaVence As Date, cRec1 As Currency, cRec2 As Currency, cRec3 As Currency, lblInteres As String, lblRecargo As String)
    'Esta rutina ya no se usa, actualmente se usa desde actualizar la factura, hacer el cambio
    'ahora interes y recargo son functions separadas
    
    Dim iMeses As Integer
    Dim iPeriodo As Integer
    Dim rsPar As New Recordset
    Dim cValor As Currency
    
    On Error GoTo CalculeRecargoBI_Error
    
    iPeriodo = Year(dFechaVence)
    If dFechaFact <= dFechaVence Then Exit Sub
    'Extrae los parametros del modulo
    Set rsPar = DeRia.CoRia.Execute("select * from ParamRia where PeriodoFact=" & iPeriodo & " ")
    'Se asume que la facturacion va para el periodo que se indica en la fecha de facturacion.
    If rsPar.RecordCount = 0 Then
        MsgBox "No estan definidos los datos de parametros para el período solicitado. Factura no puede crearse."
        Exit Sub
    End If
    
    'Interes, igual a la tasa bancaria. (Art 109 reformado)
    iMeses = MesesVencidos(dFechaVence, dFechaFact)
    cRec1 = Round(cImpuesto * rsPar!RecargoAtrasoPagoSL, 4)
    lblInteres = cImpuesto & " * " & rsPar!RecargoAtrasoPagoSL & " * " & iMeses
    cRec1 = cRec1 * iMeses
    
    'Recargo anual sobre saldos. (Art. 109 reformado)
    cValor = rsPar!RecargoSobreSaldo / 12
    cRec3 = cImpuesto + cRec1
    lblRecargo = cRec3 & " * " & cValor & " * " & iMeses
    cRec3 = Round(cRec3 * cValor * iMeses, 2)
    Exit Sub
    
CalculeRecargoBI_Error:
    MsgBox Err.Description
End Sub
Public Function CalculeIntMoraIP(cImpuesto As Currency, dFechaActual As Date, dFechaVence As Date) As Currency
'Rutina especial para Declaracion de Impuesto Personal
'Calcula el Interes por Mora sobre un impuesto
'-----------------------------------------------------
'Proceso:
'calculamos los meses vencidos
'Extraemos la tasa de interes desde ParamRia
'llamamos la rutina general de calculo de Interes
'-----------------------------------------------------
    Dim nMesesVencidos As Integer
    Dim rsParamRia As New Recordset
    Dim nPeriodo As Integer
    Dim nTasa As Single
    
    On Error GoTo CalculeIntMoraIP_Error
    CalculeIntMoraIP = 0
    If dFechaActual <= dFechaVence Then Exit Function
    
    nMesesVencidos = DateDiff("m", dFechaVence, dFechaActual)
    nPeriodo = Year(dFechaVence)
    Set rsParamRia = DeRia.CoRia.Execute("select * from ParamRia where PeriodoFact=" & nPeriodo & "")
    If rsParamRia.RecordCount = 0 Then
        MsgBox "Error: No estan definidos los parametros de tasas para el año: " & nPeriodo
        Exit Function
    End If
    'Acepta Recargos en 0 Mayo 2010
    'If IsNull(rsParamRia!RecargoAtrasoPagoSL) Or rsParamRia!RecargoAtrasoPagoSL = 0 Then
    '    MsgBox "No estan definidos los parametros de Impuesto Personal para el año:" & nPeriodo
    '    Exit Function
    'End If
    nTasa = rsParamRia!RecargoAtrasoPagoSL
    CalculeIntMoraIP = CalculeInteres(cImpuesto, nTasa, nMesesVencidos)
    Exit Function
    
CalculeIntMoraIP_Error:
    MsgBox Err.Description
End Function
Public Function CalculeRecargoSobreSaldoIP(cSaldo As Currency, dFechaActual As Date, dFechaVence As Date) As Currency
'Rutina especial para Declaracion de Impuesto Personal
'Calcula el Recargo sobre un saldo vencido
'El saldo debe estar compuesto por el Impuesto + Interes
'-----------------------------------------------------
'Proceso:
'calculamos los meses vencidos
'Extraemos la tasa de interes desde ParamRia
'llamamos la rutina general de calculo de recargo
'-----------------------------------------------------

    Dim nMesesVencidos As Integer
    Dim rsParamRia As New Recordset
    Dim nPeriodo As Integer
    Dim nTasa As Single

    On Error GoTo CalculeRecargoSobreSaldoIP_Error
    
    CalculeRecargoSobreSaldoIP = 0
    If dFechaActual <= dFechaVence Then Exit Function
    
    nMesesVencidos = DateDiff("m", dFechaVence, dFechaActual)
    nPeriodo = Year(dFechaVence)
    Set rsParamRia = DeRia.CoRia.Execute("select * from ParamRia where PeriodoFact=" & nPeriodo & "")
    If rsParamRia.RecordCount = 0 Then
        MsgBox "Error: No estan definidos los parametros de tasas para el año: " & nPeriodo
        Exit Function
    End If
    'Acepta Recargos en 0 Mayo 2010
    'If IsNull(rsParamRia!RecargoSobreSaldo) Or rsParamRia!RecargoSobreSaldo = 0 Then
     '   MsgBox "No estan definidos los parametros de Impuesto Personal para el año:" & nPeriodo
     '   Exit Function
    'End If
    nTasa = Round((rsParamRia!RecargoSobreSaldo / 12), 4)
    CalculeRecargoSobreSaldoIP = CalculeRecargoSobreSaldo(cSaldo, nTasa, nMesesVencidos)
    Exit Function
    
CalculeRecargoSobreSaldoIP_Error:
    MsgBox Err.Description
End Function
Public Function CalculeIntMoraSP(cImpuesto As Currency, dFechaActual As Date, dFechaVence As Date) As Currency
'Rutina especial para Servicios Publicos
'Calcula el Interes por Mora sobre un impuesto
'-----------------------------------------------------
'Proceso:
'calculamos los meses vencidos
'Extraemos la tasa de interes desde ParamRia
'llamamos la rutina general de calculo de Interes
'-----------------------------------------------------
    Dim nMesesVencidos As Integer
    Dim rsParamRia As New Recordset
    Dim nPeriodo As Integer
    Dim nTasa As Single

    On Error GoTo CalculeIntMoraSP_Error
    CalculeIntMoraSP = 0
    If dFechaActual <= dFechaVence Then Exit Function
    nMesesVencidos = DateDiff("m", dFechaVence, dFechaActual)
    nPeriodo = Year(dFechaVence)
    Set rsParamRia = DeRia.CoRia.Execute("select * from ParamRia where PeriodoFact=" & nPeriodo & "")
    If rsParamRia.RecordCount = 0 Then
        MsgBox "Error: No estan definidos los parametros de tasas para el año: " & nPeriodo
        Exit Function
    End If
    If IsNull(rsParamRia!RecargoAtrasoPagoSL) Or rsParamRia!RecargoAtrasoPagoSL = 0 Then
        MsgBox "No estan definidos los parametros de Intereses para el año:" & nPeriodo
        Exit Function
    End If
    nTasa = rsParamRia!RecargoAtrasoPagoSL
    CalculeIntMoraSP = CalculeInteres(cImpuesto, nTasa, nMesesVencidos)
    Exit Function
    
CalculeIntMoraSP_Error:
    MsgBox Err.Description
End Function
Public Function CalculeRecargoSobreSaldoSP(cSaldo As Currency, dFechaActual As Date, dFechaVence As Date) As Currency
'Calcula el Recargo sobre un saldo vencido
'El saldo debe estar compuesto por el Impuesto + Interes
'-----------------------------------------------------
'Proceso:
'calculamos los meses vencidos
'Extraemos la tasa de interes desde ParamRia
'llamamos la rutina general de calculo de recargo
'-----------------------------------------------------

    Dim nMesesVencidos As Integer
    Dim rsParamRia As New Recordset
    Dim nPeriodo As Integer
    Dim nTasa As Single

    On Error GoTo CalculeRecargoSobreSaldoSP_Error
    
    CalculeRecargoSobreSaldoSP = 0
    If dFechaActual <= dFechaVence Then Exit Function
    
    nMesesVencidos = DateDiff("m", dFechaVence, dFechaActual)
    nPeriodo = Year(dFechaVence)
    Set rsParamRia = DeRia.CoRia.Execute("select * from ParamRia where PeriodoFact=" & nPeriodo & "")
    If rsParamRia.RecordCount = 0 Then
        MsgBox "Error: No estan definidos los parametros de tasas para el año: " & nPeriodo
        Exit Function
    End If
    If IsNull(rsParamRia!RecargoSobreSaldo) Or rsParamRia!RecargoSobreSaldo = 0 Then
        MsgBox "No estan definidos los parametros de Recargos para el año:" & nPeriodo
        Exit Function
    End If
    nTasa = Round((rsParamRia!RecargoSobreSaldo / 12), 4)
    CalculeRecargoSobreSaldoSP = CalculeRecargoSobreSaldo(cSaldo, nTasa, nMesesVencidos)
    Exit Function
    
CalculeRecargoSobreSaldoSP_Error:
    MsgBox Err.Description
End Function
Public Sub ClearArrayCuotas()
    Dim X As Integer
    For X = 1 To 12
        aCuotasMensuales(X, 0) = "" 'Fecha que vence
        aCuotasMensuales(X, 1) = "" 'Valor Cuota
        aCuotasMensuales(X, 2) = "" 'Meses Vencidos
        aCuotasMensuales(X, 3) = "" 'Intereses = Valor Cuota * Tasa * Meses Vencidos
        aCuotasMensuales(X, 4) = "" 'Saldo Acumulado=Valor Cuota+Interess
        aCuotasMensuales(X, 5) = "" 'Recargo=Saldo Acumulado * Tasa * Meses Vencidos
        aCuotasMensuales(X, 6) = "" 'Saldo Vencido del Impuesto = Valor Cuota * Meses Vencidos
        aCuotasMensuales(X, 7) = ""
        aCuotasMensuales(X, 8) = ""
        aCuotasMensuales(X, 9) = ""
    Next
End Sub
Public Function MultaPorMesesNoPagadosIC(dFechaActual As Date) As Currency
    'Calcula la multa por no pagar el impuesto mensual a tiempo. Toma las cuotas
    'guardadas
    '--------------------------------------------------------------------------
    Dim i As Byte
    Dim cMulta As Currency
    Dim nMesesVencidos As Integer
    Dim cTotalMulta As Currency
    Dim iYear As Integer
    Dim rParam As Recordset
    
    MultaPorMesesNoPagadosIC = 0
    'Extraemos la primera fecha valida en el array aCuotasMensuales
    'recordemos que las fechas no necesariamente empiezan en Enero, como es
    'el caso de los que estan iniciando negocios, donde la facturacion
    'empieza un mes X
    For i = 1 To 12
        If IsDate(aCuotasMensuales(i, 0)) Then
            iYear = Year(CDate(aCuotasMensuales(i, 0))) 'Fecha que vencio la factura
            Exit For
        End If
    Next i
    'Extrae parametros
    Set rParam = DeRia.CoRia.Execute("select * from ParamRia where PeriodoFact=" & iYear & " ")
    If rParam.RecordCount = 0 Then
        MsgBox "No estan definidos los parametros para el año a procesar..!"
        Exit Function
    End If
    '-----------------
    If IsNull(rParam!RecargoAtrasoPagoSL) Then
        MsgBox "No estan definidos los parametros de multas..!"
        Exit Function
    End If
    '-----------------
    'calcule la multa por cada mes.
    cTotalMulta = 0
    VMensualidad = 0
    
    For i = 1 To 12
        If Val(aCuotasMensuales(i, 4)) > 0 Then
        
        If VAmniIC = 1 And Not VCuentaServicios = "" Then
        
        If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close 'aquiva
            DeRia.rsAbonadoSP.Open ("SELECT AvPgDetalle.NumAvPg, AvPgDetalle.ValorUnitAvPgDet, AvPgDetalle.CantAvPgDet FROM AvPgDetalle INNER JOIN CuentaIngreso ON AvPgDetalle.CtaIngreso = CuentaIngreso.CtaIngreso WHERE (AvPgDetalle.NumAvPg = " & VFactNum & ") AND (CuentaIngreso.Tipo <> 2) and SUBSTRING(AvPgDetalle.CtaIngreso, 1, 6) = '" & VCuentaServicios & "' ")
            If Not DeRia.rsAbonadoSP.EOF Then
            Do While Not DeRia.rsAbonadoSP.EOF
            VMensualidad = VMensualidad + DeRia.rsAbonadoSP!ValorUnitAvPgDet
            
            DeRia.rsAbonadoSP.MoveNext
            Loop
            aCuotasMensuales(i, 1) = VMensualidad
            End If
        Else
        If VFactNum = "" Then VFactNum = 0
        If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
                
   DeRia.rsAbonadoSP.Open ("SELECT AvPgDetalle.NumAvPg, AvPgDetalle.ValorUnitAvPgDet, AvPgDetalle.CantAvPgDet FROM AvPgDetalle INNER JOIN CuentaIngreso_A ON AvPgDetalle.CtaIngreso = CuentaIngreso_A.CtaIngreso Where (CuentaIngreso_A.Tipo <> 2) GROUP BY AvPgDetalle.NumAvPg, AvPgDetalle.ValorUnitAvPgDet, AvPgDetalle.CantAvPgDet Having (AvPgDetalle.NumAvPg = " & VFactNum & ") ")
            
            If Not DeRia.rsAbonadoSP.EOF Then
            Do While Not DeRia.rsAbonadoSP.EOF
            VMensualidad = VMensualidad + DeRia.rsAbonadoSP!ValorUnitAvPgDet
            
            DeRia.rsAbonadoSP.MoveNext
            Loop
            aCuotasMensuales(i, 1) = VMensualidad
            End If
            
        End If
            cMulta = Round((rParam!RecargoAtrasoPagoSL * Val(aCuotasMensuales(i, 1))) * Val(aCuotasMensuales(i, 4)), 2)
            aCuotasMensuales(i, 5) = cMulta
            cTotalMulta = cTotalMulta + cMulta
            aCuotasMensuales(i, 6) = Val(aCuotasMensuales(i, 1)) + Val(aCuotasMensuales(i, 2)) + Val(aCuotasMensuales(i, 3)) + cMulta
        End If
    Next
    MultaPorMesesNoPagadosIC = cTotalMulta
      
End Function
Public Function RecargoPorMesesNoPagadosIC(dFechaActual As Date) As Currency
    'Calcula la multa por no pagar el impuesto mensual a tiempo. Toma las cuotas
    'guardadas
    '--------------------------------------------------------------------------
    Dim i As Byte
    Dim cMulta As Currency
    Dim nMesesVencidos As Integer
    Dim cTotalMulta As Currency
    Dim cPorcentajeMensual As Single
    Dim iYear As Integer
    Dim rParam As Recordset
    Dim curSaldo As Currency
    
    'RecargoPorMesesNoPagadosIC = 0 ' Aqui
    'Extraemos la primera fecha valida en el array aCuotasMensuales
    'recordemos que las fechas no necesariamente empiezan en Enero, como es
    'el caso de los que estan iniciando negocios, donde la facturacion
    'empieza un mes X
    For i = 1 To 12
        If IsDate(aCuotasMensuales(i, 0)) Then
            iYear = Year(CDate(aCuotasMensuales(i, 0))) 'Fecha que vencio la factura
            Exit For
        End If
    Next i
    'Extrae parametros
    Set rParam = DeRia.CoRia.Execute("select * from ParamRia where PeriodoFact=" & iYear & " ")
    If rParam.RecordCount = 0 Then
        MsgBox "No estan definidos los parametros para el año a procesar..!"
        Exit Function
    End If
    '-----------------
    If IsNull(rParam!RecargoSobreSaldo) Then
        MsgBox "No estan definidos los parametros de multas..!"
        Exit Function
    End If
    '--------------------------------
    'calcule la multa por cada mes.
    '--------------------------------
    cTotalMulta = 0
    
    VMensualidad = 0
    For i = 1 To 12
        If Val(aCuotasMensuales(i, 4)) > 0 Then
            cPorcentajeMensual = Round(rParam!RecargoSobreSaldo / 12, 4)
        
        
        If VAmniIC = 1 And Not VCuentaServicios = "" Then
            
            If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
            
           DeRia.rsAbonadoSP.Open ("SELECT AvPgDetalle.NumAvPg, AvPgDetalle.ValorUnitAvPgDet, AvPgDetalle.CantAvPgDet FROM AvPgDetalle INNER JOIN CuentaIngreso_A ON AvPgDetalle.CtaIngreso = CuentaIngreso_A.CtaIngreso WHERE (SUBSTRING(AvPgDetalle.CtaIngreso, 1, 6) = '" & VCuentaServicios & "') AND (CuentaIngreso_A.Tipo <> 2) GROUP BY AvPgDetalle.NumAvPg, AvPgDetalle.ValorUnitAvPgDet, AvPgDetalle.CantAvPgDet Having (AvPgDetalle.NumAvPg = " & VFactNum & ")")
            
            
            If Not DeRia.rsAbonadoSP.EOF Then
            Do While Not DeRia.rsAbonadoSP.EOF
            VMensualidad = VMensualidad + DeRia.rsAbonadoSP!ValorUnitAvPgDet
            
            DeRia.rsAbonadoSP.MoveNext
            Loop
            aCuotasMensuales(i, 1) = VMensualidad
            End If
            
        Else
            
            If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
             
             DeRia.rsAbonadoSP.Open ("SELECT AvPgDetalle.NumAvPg, AvPgDetalle.ValorUnitAvPgDet, AvPgDetalle.CantAvPgDet FROM AvPgDetalle INNER JOIN CuentaIngreso_A ON AvPgDetalle.CtaIngreso = CuentaIngreso_A.CtaIngreso WHERE (CuentaIngreso_A.Tipo <> 2) GROUP BY AvPgDetalle.NumAvPg, AvPgDetalle.ValorUnitAvPgDet, AvPgDetalle.CantAvPgDet Having (AvPgDetalle.NumAvPg = " & VFactNum & ")")
           
            
            If Not DeRia.rsAbonadoSP.EOF Then
            Do While Not DeRia.rsAbonadoSP.EOF
            VMensualidad = VMensualidad + DeRia.rsAbonadoSP!ValorUnitAvPgDet
            
            DeRia.rsAbonadoSP.MoveNext
            Loop
            aCuotasMensuales(i, 1) = VMensualidad
            End If
            
        End If
            curSaldo = Val(aCuotasMensuales(i, 1) + Val(aCuotasMensuales(i, 5))) 'Impuesto + Interes
            nMesesVencidos = Val(aCuotasMensuales(i, 4))
            cMulta = Round(curSaldo * nMesesVencidos * cPorcentajeMensual, 2)
            aCuotasMensuales(i, 7) = cMulta
            cTotalMulta = cTotalMulta + cMulta
        End If
    Next
    RecargoPorMesesNoPagadosIC = cTotalMulta
    
End Function
Public Function DescuentoTerceraEdad(sId As String, cValor As Currency) As Currency
'calcula el descuento por tercera edad si procede.
'se esperan como parametros la identidad y el valor sobre el cual se le calcula descuento

    Dim rsPar As New Recordset
    Dim iEdad As Integer
    
    DescuentoTerceraEdad = 0
    Set rsPar = DeRia.CoRia.Execute("select * from Parametro")
    If rsPar.RecordCount = 0 Then
        MsgBox "No estan definidos los parametros generales del modulo. Vaya al menu mantenimiento."
        Exit Function
    End If
    If IsNull(rsPar!AnosTerceraEdad) Or IsNull(rsPar!TasaTerceraEdad) Then
        MsgBox "Debe definir correctamente los parametros de tercera edad, en parametros generales del menu mantenimiento."
        Exit Function
    End If
    iEdad = EdadPersona(sId, DeRia.CoRia)
    If iEdad >= rsPar!AnosTerceraEdad Then
        DescuentoTerceraEdad = cValor * rsPar!TasaTerceraEdad
    End If
End Function
Public Sub InsertEncFacturaTemp(rsEncTemp As Recordset, lNum As Long, sId As String, dFeEm As Date, dFeVence As Date, bTipoImp As Byte, bTipo As Byte, bEstado As Byte, sDescrip As String, sClaveCatastro As String)
'Para cuestiones de multiusuario, cuando se crea una factura,
'primero se crea una imagen, y luego se manda a grabar todo en un solo batch, en ese batch
'se bloquea el numero de factura, para que otro usuario no pueda insertar otra factura en ese
'momento
    rsEncTemp.AddNew
    rsEncTemp!NumAvPg = lNum
    rsEncTemp!Identidad = sId
    rsEncTemp!FechaEmAvPg = dFeEm
    rsEncTemp!FechaVenceAvPg = dFeVence
    rsEncTemp!AvPgTipoImpuesto = bTipoImp '1=Bienes Inmuebles, 2=Volumen Ventas,3=Permiso Oper... Ver tabla
    rsEncTemp!TipoAvPg = bTipo '1=Contado, 2=Pendiente
    rsEncTemp!AvPgEstado = bEstado '1=No Pagada,2=Pagada,3=Anulada,4=En Tesoreria, 5=Pagada Parcial, 6=Plan Pago
    rsEncTemp!AvPgDescripcion = sDescrip
    rsEncTemp!AvPgTotalPeriodo = 0
    rsEncTemp!ClaveCatastro = sClaveCatastro
    rsEncTemp!CreadoPor = gsUsername
    rsEncTemp!FechaCreado = Date
    rsEncTemp!ModificadoPor = ""
    rsEncTemp!FechaModificado = Date
    rsEncTemp.Update
End Sub
Public Sub InsertDetFacturaTemp(rsDetTemp As Recordset, lNum As Long, cValorUnit As Currency, sCatastro As String, sCtaIng As String, sRef As String, cCant As Currency, cDescuento As Currency, cinteres As Currency, cRecargo As Currency, nVisibleEnTes As Integer)
    rsDetTemp.AddNew
    rsDetTemp!NumAvPg = lNum
    rsDetTemp!CtaIngreso = Trim(sCtaIng)
    rsDetTemp!RefAvPgDet = sRef
    rsDetTemp!ClaveCatastro = sCatastro
    rsDetTemp!CantAvPgDet = cCant
    rsDetTemp!ValorUnitAvPgDet = cValorUnit
    rsDetTemp!DescuentoAvPgDet = cDescuento
    rsDetTemp!RecargoAvPgDet = cinteres
    rsDetTemp!ValorXAvPgDet = cRecargo
    rsDetTemp!ValorPagadoAvPgDet = 0
    If nVisibleEnTes = 1 Then rsDetTemp!VisibleEnTesAvPgDet = True Else rsDetTemp!VisibleEnTesAvPgDet = False
    rsDetTemp.Update
End Sub
Private Sub UpdateIntRecDesc(cDesc As Currency, cinteres As Currency, cRec As Currency, intTipo As String, lngNumFact As Long)
'actualiza una factura especifica, para reflejar el descuento, interes y recargo.
'Si el tipo es 1,2,4 es un bloque el 5 otro bloque
'Para actualizar descuento,Interes, Recargo. Se hace una para cada tipo
'aqui muestro de ejemplo para descuento.
'Proceso:
'busque en SystemParam la cuenta de descuento, interesporimpuesto, interesporservicio,recargoporimpuesto,recargoporservicio
'   busque en el detalle de la factura la cuenta de descuento
'   si se encuentra
'       si Descuento es mayor que 0
'           se actualiza.
'       sino
'           se borra item
'   si no se encuentra
'       si Descuento es mayor que 0
'           se agrega
'
'Si el tipo es 5, son servicios.
    Dim rsSysPar, rs As Recordset
    Dim strSql As String
    Dim cValor As Currency
    Dim sCuenta As String
    
    Set rsSysPar = DeRia.CoRia.Execute("select * from SystemParam")
    If rsSysPar.RecordCount = 0 Then Exit Sub
    'Trabajamos primero las facturas por servicio
    If intTipo = 5 Then
        'Intereses
        Set rs = DeRia.CoRia.Execute("select * from AvPgDetalle where NumAvPg=" & lngNumFact & " and CtaIngreso='" & rsSysPar!CtaIngresoIntServ & "'")
        If rs.RecordCount > 0 Then
            If cinteres > 0 Then
                DeRia.CoRia.Execute ("update AvPgDetalle set ValorUnitAvPgDet=" & Round(cinteres, 2) & " where NumAvPg=" & lngNumFact & " and CtaIngreso='" & rsSysPar!CtaIngresoIntServ & "'")
            Else
                DeRia.CoRia.Execute ("delete from AvPgDetalle where NumAvPg=" & lngNumFact & " and CtaIngreso='" & rsSysPar!CtaIngresoIntServ & "'")
            End If
        Else
            If cinteres > 0 Then
                strSql = "insert into AvPgDetalle (NumAvPg,ValorUnitAvPgDet,CtaIngreso,"
                strSql = strSql & "CantAvPgDet,VisibleEnTesAvPgDet) values(" & lngNumFact & ","
                strSql = strSql & "" & Round(cinteres, 2) & ",'" & rsSysPar!CtaIngresoIntServ & "',"
                strSql = strSql & "1,0)"
                'MsgBox strSql
                DeRia.CoRia.Execute (strSql)
            End If
        End If
        'Recargo
        Set rs = DeRia.CoRia.Execute("select * from AvPgDetalle where NumAvPg=" & lngNumFact & " and CtaIngreso='" & rsSysPar!CtaIngresoRecargoServ & "'")
        If rs.RecordCount > 0 Then
            If cRec > 0 Then
                DeRia.CoRia.Execute ("update AvPgDetalle set ValorUnitAvPgDet=" & Round(cRec, 2) & " where NumAvPg=" & lngNumFact & " and CtaIngreso='" & rsSysPar!CtaIngresoRecargoServ & "'")
            Else
                DeRia.CoRia.Execute ("delete from AvPgDetalle where NumAvPg=" & lngNumFact & " and CtaIngreso='" & rsSysPar!CtaIngresoRecargoServ & "'")
            End If
        Else
            If cRec > 0 Then
                strSql = "insert into AvPgDetalle (NumAvPg,ValorUnitAvPgDet,CtaIngreso,"
                strSql = strSql & "CantAvPgDet,VisibleEnTesAvPgDet) values(" & lngNumFact & ","
                strSql = strSql & "" & Round(cRec, 2) & ",'" & rsSysPar!CtaIngresoRecargoServ & "',"
                strSql = strSql & "1,0)"
                DeRia.CoRia.Execute (strSql)
            End If
        End If
        Exit Sub
    End If
    'Ahora trabajamos las otras tipos de facturas, los impuestos: la 1,2,4
    '
    'Descuentos
    cValor = Round(cDesc, 2)
    sCuenta = rsSysPar!CtaIngresoDescuento
    Set rs = DeRia.CoRia.Execute("select * from AvPgDetalle where NumAvPg=" & lngNumFact & " and CtaIngreso='" & sCuenta & "'")
    If rs.RecordCount > 0 Then
        If Abs(cValor) > 0 Then
            DeRia.CoRia.Execute ("update AvPgDetalle set ValorUnitAvPgDet=" & cValor & " where NumAvPg=" & lngNumFact & " and CtaIngreso='" & sCuenta & "'")
        Else
            DeRia.CoRia.Execute ("delete from AvPgDetalle where NumAvPg=" & lngNumFact & " and CtaIngreso='" & sCuenta & "'")
        End If
    Else
        If Abs(cValor) > 0 Then 'el descuento debe venir negative
            strSql = "insert into AvPgDetalle (NumAvPg,ValorUnitAvPgDet,CtaIngreso,"
            strSql = strSql & "CantAvPgDet,VisibleEnTesAvPgDet) values(" & lngNumFact & ","
            strSql = strSql & "" & Round(cValor, 2) & ",'" & sCuenta & "',"
            strSql = strSql & "1,0)"
            DeRia.CoRia.Execute (strSql)
        End If
    End If
    '----------
    'Intereses
    cValor = cinteres
    sCuenta = Trim(rsSysPar!CtaIngresoIntImp)
    Set rs = DeRia.CoRia.Execute("select * from AvPgDetalle where NumAvPg=" & lngNumFact & " and rtrim(ltrim(CtaIngreso))='" & sCuenta & "'")
    If rs.RecordCount > 0 Then
        If cValor > 0 Then
            DeRia.CoRia.Execute ("update AvPgDetalle set ValorUnitAvPgDet=" & Round(cValor, 2) & " where NumAvPg=" & lngNumFact & " and rtrim(ltrim(CtaIngreso))='" & sCuenta & "'")
        Else
           'MSx32AM' DeRia.CoRia.Execute ("delete from AvPgDetalle where NumAvPg=" & lngNumFact & " and CtaIngreso='" & sCuenta & "'")
        End If
    Else
        If cValor > 0 Then
            strSql = "insert into AvPgDetalle (NumAvPg,ValorUnitAvPgDet,CtaIngreso,"
            strSql = strSql & "CantAvPgDet,VisibleEnTesAvPgDet) values(" & lngNumFact & ","
            strSql = strSql & "" & Round(cValor, 2) & ",'" & sCuenta & "',"
            strSql = strSql & "1,0)"
            DeRia.CoRia.Execute (strSql)
        End If
    End If
    '--------
    'Recargos
    cValor = cRec
    sCuenta = Trim(rsSysPar!CtaIngresoRecargoImp)
    Set rs = DeRia.CoRia.Execute("select * from AvPgDetalle where NumAvPg=" & lngNumFact & " and rtrim(ltrim(CtaIngreso))='" & sCuenta & "'")
    If rs.RecordCount > 0 Then
        If cValor > 0 Then
            DeRia.CoRia.Execute ("update AvPgDetalle set ValorUnitAvPgDet=" & Round(cValor, 2) & " where NumAvPg=" & lngNumFact & " and rtrim(ltrim(CtaIngreso))='" & sCuenta & "'")
        Else
          'MSx32AM  DeRia.CoRia.Execute ("delete from AvPgDetalle where NumAvPg=" & lngNumFact & " and CtaIngreso='" & sCuenta & "'")
        End If
    Else
        If cValor > 0 Then
            strSql = "insert into AvPgDetalle (NumAvPg,ValorUnitAvPgDet,CtaIngreso,"
            strSql = strSql & "CantAvPgDet,VisibleEnTesAvPgDet) values(" & lngNumFact & ","
            strSql = strSql & "" & Round(cValor, 2) & ",'" & sCuenta & "',"
            strSql = strSql & "1,0)"
            DeRia.CoRia.Execute (strSql)
        End If
    End If
End Sub
Private Function EsIntRecMulta(sCuenta As String) As Boolean
'Es Interes, Recargo, o Multa
'Verifica sCuenta, si es Interes, Recargo o Multa retorna Verdadero
    Dim rs As Recordset
    
    EsIntRecMulta = False
    Set rs = DeRia.CoRia.Execute("select * from SystemParam")
    If rs.RecordCount = 0 Then Exit Function
    Select Case sCuenta
        Case rs!CtaIngresoIntImp
            EsIntRecMulta = True
        Case rs!CtaIngresoIntServ
            EsIntRecMulta = True
        Case rs!CtaIngresoRecargoImp
            EsIntRecMulta = True
        Case rs!CtaIngresoRecargoServ
            EsIntRecMulta = True
        Case rs!CtaIngresoMultaDeclaraTarde
            EsIntRecMulta = True
    End Select
End Function
Public Function GetClaveCatastro(lngNumAvPg As Long) As String
'Extrae la clave de catastro del detalle, dando el numero de factura.
    Dim rs As Recordset
    
    GetClaveCatastro = ""
    Set rs = DeRia.CoRia.Execute("select ClaveCatastro from AvPgEnc where NumAvPg=" & lngNumAvPg & " ")
    If rs.RecordCount = 0 Then Exit Function
    GetClaveCatastro = rs!ClaveCatastro
End Function
Public Function CalculeSaldoSP(strClaveCatastro As String) As Currency
'Calcula saldo en facturas de servicio publico para una propiedad
    Dim rs As Recordset
    Dim rs2 As Recordset
    Dim strSql As String
    
    CalculeSaldoSP = 0
    'Seleccionamos todas las facturas pendientes de pago de la clave dada
    strSql = "select * from AvPgEnc where AvPgEstado=1 and AvPgTipoImpuesto=5 "
    strSql = strSql & "and ClaveCatastro='" & strClaveCatastro & "'"
    Set rs = DeRia.CoRia.Execute(strSql)
    '--------------------------------------------------------------------
    If rs.RecordCount = 0 Then Exit Function
    'calculamos los recargos e intereses
    MoraPorFactura rs
    '-----------------------------------
    'Ahora extraemos el saldo total pendiente de la propiedad
    rs.MoveFirst
    Do While rs.EOF = False
        strSql = "select sum(ValorUnitAvPgDet) as valor from AvPgDetalle where NumAvPg=" & rs!NumAvPg & " "
        Set rs2 = DeRia.CoRia.Execute(strSql)
        If IIf(IsNull(rs2!valor), 0, rs2!valor) > 0 Then
            CalculeSaldoSP = CalculeSaldoSP + rs2!valor
        End If
        rs.MoveNext
    Loop
    '---------------------------------------------------------
End Function
Public Sub PrintAvisoPago(lngNumAvPg As Long, cnnCon As Connection)
'Imprime un aviso de pago de servicios publicos
'o sea una factura cuyo numero viene de parametro
    Dim rs As New Recordset
    
    'cnnCon.Execute ("delete from Temp1")
    'cnnCon.Execute ("insert into Temp1 (NumPartida) values(" & lngNumAvPg & ")")
    'el siguiente select solo es para que se haga el flush de los datos en el buffer
    'y funcione bien el reporte
    'Set rs = cnnCon.Execute("select NumPartida from Temp1")
    '-------------------------------------------------------------------------------
    'If De.rsAvisosPagoEnc.State = adStateOpen Then De.rsAvisosPagoEnc.Close
    'De.rsAvisosPagoEnc.Open
    'rptFacturaSP.Orientation = rptOrientPortrait
    'rptFacturaSP.Show
End Sub

Public Function AnularPlanPago(lngNumPP As Long) As Boolean
    Dim rsFacturaGenerada As New Recordset
    Dim rsFacturaOriginal As New Recordset
    Dim rs As New Recordset
    Dim strSql As String
    
    'Ver proceso de anulacion de PlanPago en Directorio de Programas fuente
    AnularPlanPago = False
    'Verificamos que ninguna factura del plan de pagos este pagada
    strSql = "select count(*) as Contador from AvPgEnc where AvPgEstado=2 and NumAvPg In " & _
    "(select NumAvPg from PlanPagoFactura where SeqPP=" & lngNumPP & ")"
    Set rs = DeRia.CoRia.Execute(strSql)
    If rs!Contador > 0 Then
        MsgBox "Al menos una factura del plan de pagos " & lngNumPP & " ya ha sido pagada. No se puede anular plan de pagos."
        Exit Function
    End If
    '--------------------------------------------------------------
    'Ahora buscamos las facturas originales
    strSql = "Select NumAvPg from PlanPagoDetalle where SeqPP=" & lngNumPP & " "
    Set rsFacturaOriginal = DeRia.CoRia.Execute(strSql)
    '-------------------------------------
    'Ahora buscamos las facturas que fueron creadas al generar el plan de pago
    strSql = "select NumAvPg from PlanPagoFactura where SeqPP=" & lngNumPP & ""
    Set rsFacturaGenerada = DeRia.CoRia.Execute(strSql)
    '-------------------------------------------------------------------------
    
    'Pongamos el plan de pago en estado 2, anulado
    DeRia.CoRia.Execute ("Update PlanPago set EstadoPP =2 where SeqPP=" & lngNumPP & "")
    'Anulamos las facturas generadas por el plan de pago
    Do While rsFacturaGenerada.EOF = False
        DeRia.CoRia.Execute ("update AvPgEnc set AvPgEstado=3 where NumAvPg=" & rsFacturaGenerada!NumAvPg & " ")
        rsFacturaGenerada.MoveNext
    Loop
    'Ahora activamos las facturas que generaron el plan de pagos
    Do While rsFacturaOriginal.EOF = False
        DeRia.CoRia.Execute ("update AvPgEnc set AvPgEstado=1 where NumAvPg=" & rsFacturaOriginal!NumAvPg & " ")
        rsFacturaOriginal.MoveNext
    Loop
    AnularPlanPago = True
End Function

Public Function EsImpuesto(strCta As String) As Boolean
    Dim rs As New Recordset
    
    EsImpuesto = False
    Set rs = DeRia.CoRia.Execute("select * from CuentaIngreso where CtaIngreso='" & strCta & "'")
    If rs.RecordCount = 0 Then Exit Function
    If rs!Tipo = 1 Then
        EsImpuesto = True
    End If
End Function

Public Function CalculeImpuestoBI(strClaveCatastro, lngPeriodo As Long) As Currency
    Dim rs As New Recordset
    Dim curValorPropiedad As Currency
    Dim rsPar As New Recordset
    Dim sngTasa As Single
    
    CalculeImpuestoBI = 0
    Set rs = DeRia.CoRia.Execute("select * from Catastro where ClaveCatastro='" & strClaveCatastro & "'")
    If rs.RecordCount = 0 Then Exit Function
    
    'Calcule valor de propiedad
    curValorPropiedad = rs!ValorTerreno + rs!ValorEdificacion - rs!ValorExencion
    '--------------------------
    'Extrae la tasa por millar y otros parametros referentes a Bienes Inmuebles
    Set rsPar = DeRia.CoRia.Execute("select * from ParamRia where PeriodoFact=" & lngPeriodo & "")
    If rsPar.RecordCount = 0 Then
        MsgBox "No existen parametros definidos para el año:" & lngPeriodo
        Exit Function
    Else
        If rs!Ubicacion = 0 Then
            sngTasa = IIf(IsNull(rsPar!RecargoAtrasoPago), 0, rsPar!RecargoAtrasoPago)
        Else
            sngTasa = IIf(IsNull(rsPar!TasaBIRural), 0, rsPar!TasaBIRural)
        End If
    End If
    '-----------------------------------
    CalculeImpuestoBI = CalcImpCat(curValorPropiedad, sngTasa)
    '-----------------------------------
End Function

Public Function DiaEnProcesoCT() As Date
    Dim rs As New Recordset
    
    Set rs = DeRia.CoRia.Execute("select DiaProcesoCT from Parametro")
    If rs.RecordCount = 0 Then Exit Function
    DiaEnProcesoCT = rs!DiaProcesoCT
End Function


Public Function PuedeModificarPeriodo(strId As String, lngUltPerFact As String) As Boolean
    Dim rs As New Recordset
    Dim rsFactura As New Recordset
    
    'Verificar por tipo de impuesto, este no esta finalizado.
    PuedeModificarPeriodo = True
    Exit Function
    '------------------------------
    PuedeModificarPeriodo = False
    Set rs = DeRia.CoRia.Execute("select * from Contribuyente where Identidad='" & strId & "' ")
    If rs.RecordCount = 0 Then
        PuedeModificarPeriodo = True
        Exit Function
    End If
    If rs!UltPeriodoFact = lngUltPerFact Then
        PuedeModificarPeriodo = True
        Exit Function
    End If
    Set rsFactura = DeRia.CoRia.Execute("select * from AvPgEnc where Identidad ='" & strId & "' and AvPgEstado<>3")
    If rsFactura.RecordCount > 0 Then
        Exit Function
    End If
    PuedeModificarPeriodo = True
End Function

Public Function Interes(curImpuesto As Currency, dFechaFact As Date, dFechaVence As Date, strFormula As String) As Currency
    Dim intMeses As Integer
    Dim intPeriodo As Integer
    Dim rsPar As New Recordset
    Dim curValor As Currency
    
    Interes = 0
    intPeriodo = Year(dFechaVence)
    If dFechaFact <= dFechaVence Then Exit Function
    'Extrae los parametros del modulo
    Set rsPar = DeRia.CoRia.Execute("select * from ParamRia where PeriodoFact=" & intPeriodo & " ")
    'Se asume que la facturacion va para el periodo que se indica en la fecha de facturacion.
    If rsPar.RecordCount = 0 Then
        MsgBox "No estan definidos los datos de parametros para el período solicitado. Factura no puede crearse."
        Exit Function
    End If
    
    'Interes, igual a la tasa bancaria. (Art 109 reformado)
    intMeses = MesesVencidos(dFechaVence, dFechaFact)
    Interes = Round(curImpuesto * rsPar!RecargoAtrasoPagoSL, 4)
    strFormula = curImpuesto & " * " & rsPar!RecargoAtrasoPagoSL & " * " & intMeses
    
    Interes = Interes * intMeses
    
End Function
Public Function Recargo(curValorDeuda As Currency, dFechaVence As Date, dFechaFact As Date, strFormula As String) As Currency
    Dim intMeses As Integer
    Dim intPeriodo As Integer
    Dim rsPar As New Recordset
    Dim curValor As Currency
    Dim curTasa As Currency
    
    Recargo = 0
    intPeriodo = Year(dFechaVence)
    If dFechaFact <= dFechaVence Then Exit Function
    'Extrae los parametros del modulo
    Set rsPar = DeRia.CoRia.Execute("select * from ParamRia where PeriodoFact=" & intPeriodo & " ")
    'Se asume que la facturacion va para el periodo que se indica en la fecha de facturacion.
    If rsPar.RecordCount = 0 Then
        MsgBox "No estan definidos los datos de parametros para el período solicitado. Factura no puede crearse."
        Exit Function
    End If
    intMeses = MesesVencidos(dFechaVence, dFechaFact)
    'Recargo anual sobre saldos. (Art. 109 reformado)
    curTasa = rsPar!RecargoSobreSaldo / 12
    strFormula = curValorDeuda & " * " & curTasa & " * " & intMeses
    Recargo = Round(curValorDeuda * curTasa * intMeses, 2)
    
End Function

Public Function EsIntMultRecDesc(strCta As String) As Boolean
    Dim rs As New Recordset
    
    EsIntMultRecDesc = False
    Set rs = DeRia.CoRia.Execute("select * from CuentaIngreso where CtaIngreso='" & strCta & "'")
    If rs.RecordCount = 0 Then Exit Function
    If rs!Tipo = 2 Then
        EsIntMultRecDesc = True
    End If
End Function
Public Function RepresentanteNegocio(strIdentidad As String) As ADODB.Recordset
    
    Dim rs As New ADODB.Recordset
    
    Set rs = DeRia.CoRia.Execute("select * from Contribuyente where Identidad='" & strIdentidad & "'")
    Set RepresentanteNegocio = rs
    
End Function

Public Function AnuleFactura(lngNumFact As Long)
    Dim strSql As String
    
    strSql = "Update AvPgEnc set " & _
    "AvPgEstado=3, " & _
    "ModificadoPor= '" & gsUsername & "', " & _
    "FechaModificado= '" & Format(DiaEnProcesoCT, "dd/mm/yyyy") & "' " & _
    "where NumAvPg=" & lngNumFact & ""
    DeRia.CoRia.Execute (strSql)
    
End Function
Public Function ObtenerCuentasContables() As ADODB.Recordset
    Set ObtenerCuentasContables = DeRia.CoRia.Execute("select * from Catalogo where Operable=True order by NombreCta")
End Function
Public Function ObtenerCuentasContablesIngreso() As ADODB.Recordset
    Set ObtenerCuentasContablesIngreso = DeRia.CoRia.Execute("select * from Catalogo where CtaContable like '4%' and Operable=True order by NombreCta")
End Function

Public Function ObtenerCuentaContablePorID(strId As String) As ADODB.Recordset
    Set ObtenerCuentaContablePorID = DeRia.CoRia.Execute("select * from Catalogo where CtaContable='" & strId & "'")
End Function
Public Function ObtenerCuentasIngreso() As ADODB.Recordset
    Set ObtenerCuentasIngreso = DeRia.CoRia.Execute("select * from CuentaIngreso order by NombreCtaIngreso")
End Function
Public Function ObtenerCuentaIngresoPorID(strId As String) As ADODB.Recordset
    Set ObtenerCuentaIngresoPorID = DeRia.CoRia.Execute("select * from CuentaIngreso where CtaIngreso='" & strId & "' order by NombreCtaIngreso")
End Function
Public Function ObtenerCuentasPermisoOperacion(VFec1 As Date) As ADODB.Recordset
    Set ObtenerCuentasPermisoOperacion = DeRia.CoRia.Execute("select * from CuentaIngreso where CtaIngreso like '11111821%' order by NombreCtaIngreso")
End Function
Public Function ObtenerCuentasRecuperacion() As ADODB.Recordset
    Set ObtenerCuentasRecuperacion = DeRia.CoRia.Execute("select * from CuentaIngreso where CtaIngreso like '112122%' order by NombreCtaIngreso")
End Function
Public Sub CalculeSaldoVencido()
    Dim i As Integer
    
    For i = 1 To 12
        If Val(aCuotasMensuales(i, 4)) > 0 Then
            aCuotasMensuales(i, 9) = Val(Format(aCuotasMensuales(i, 1), "####.00")) + Val(Format(aCuotasMensuales(i, 5), "####.00"))
        End If
    Next
    
End Sub

Public Function DeclaracionTieneRecibosPosteadosEnConta(strNumDeclara As String) As Boolean
'Verificar si el recibo esta operado en contabilidad
'Verificar si los recibos de las facturas de las declaraciones estan posteados en contabilidad
'Si es así retorna True, sino retorna Falso
    Dim rsFacturas As New ADODB.Recordset
    Dim rsRecibos As New ADODB.Recordset
    Dim strSql As String
    
    DeclaracionTieneRecibosPosteadosEnConta = False
    strSql = "select NumAvPg from AvPgDetalle where RefAvPgDet='" & strNumDeclara & "' "
    Set rsFacturas = dal.ObtenerRecordset(strSql)
    Do While rsFacturas.EOF = False
        Set rsRecibos = dal.ObtenerRecibosPorFactura(rsFacturas!NumAvPg)
        Do While rsRecibos.EOF = False
            If rsRecibos!SentToCont = True Then
                DeclaracionTieneRecibosPosteadosEnConta = True
                Exit Function
            End If
            rsRecibos.MoveNext
        Loop
        rsFacturas.MoveNext
    Loop
    '-----------------------------------------
End Function
Public Sub ActualizeSaldosPorContribuyente(strIdentidad As String)
    'traiga todas las facturas pendientes del contribuyente, que no sean de servicios publicos
    'Tambien trae los negocios de los cuales es el propietario o representante.
    Dim rsFacturas As New ADODB.Recordset
    Dim sSql As String
    
    Set rsFacturas = New Recordset
    
    sSql = "select * from AvPgEnc " & _
    "where (Identidad='" & strIdentidad & "' " & _
    "or Identidad  in (select Identidad from Contribuyente where IdRepresentante='" & strIdentidad & "'))" & _
    "and AvPgTipoImpuesto<>5 " & _
    "and (AvPgEstado=1 or AvPgEstado=5) order by FechaVenceAvPg "
    rsFacturas.Open sSql, DeRia.CoRia, , adLockBatchOptimistic
    'Calcule la mora por cada factura vencida, y actualize la tabla AvPgDetalle
    If rsFacturas.RecordCount > 0 Then
        MoraPorFactura rsFacturas
        rsFacturas.MoveFirst
    End If

End Sub
Public Sub AnuleDeclaracionJuridica(strNumDeclaracion As String)
'Solo se puede anular la ultima declaración realizada para el contribuyente
'seleccione todas las declaraciones con fecha mayore para ese contribuyente contribuyente
'   si hay mas de una, no puede anular
'busque todos los detalles con distinto numero de factura, donde
'numero de referencia sea igual a la declaracion en pantalla.
'para cada detalle
'   Buscamos el encabezado del detalle
'   Si la factura esta pagada. No puede anularse
'   Cambiamos el estado de la factura a anulada
'Cambiamos el estado de la declaracion a anulada
'Restamos 1 al UltPeriodoFacturado del contribuyente.
'

    Dim strSql As String
    Dim rs As New Recordset
    Dim rsFactura As New Recordset
    Dim blnInTransac As Boolean
    Dim intUltMesFact As Integer
    Dim rsDeclaracion As ADODB.Recordset
   
    
    On Error GoTo CheckError
    'Verifique que el primer campo tiene datos
    blnInTransac = False
    
    Set rsDeclaracion = dal.ObtenerDeclaracionICPorNumero(strNumDeclaracion)
    'Verificar si el recibo esta operado en contabilidad
    If DeclaracionTieneRecibosPosteadosEnConta(strNumDeclaracion) = True Then
        MsgBox "Algunas facturas ya fueron pagadas y los recibos de estas ya fueron registrados en contabilidad. Esta declaración no se puede anular."
        Exit Sub
    End If
    '----------------------------------------------------
    'Verifique si es la ultima declaracion. Solo puede anular en orden de ultima a primera
    strSql = "select FechaFacturado from DeclaraContJurid  " & _
    "where Periodo > " & rsDeclaracion!Periodo & " and " & _
    "EstadoDeclaraIC<>2 and Identidad='" & rsDeclaracion!Identidad & "'"
    Set rs = DeRia.CoRia.Execute(strSql)
    If rs.RecordCount > 0 Then
        MsgBox "Esta no es la ultima declaración para este contribuyente. " & _
        "Debe efectuar las anulaciones en orden. De período mayor a menor."
        Exit Sub
    End If
    '-----------------------------------------
    If MsgBox("Realmente desea anular esta declaración...?", vbQuestion + vbYesNo, "Cuidado..!") = vbYes Then
        'Seleccionamos todas las facturas que generó la declaración
        strSql = "select distinct(NumAvPg) as NumAvPg " & _
        "from AvPgDetalle where RefAvPgDet='" & rsDeclaracion!CodDeclaraCJ & "' order by NumAvPg desc "
        Set rsFactura = DeRia.CoRia.Execute(strSql)
        If rsFactura.RecordCount > 0 Then
            DeRia.CoRia.BeginTrans
            blnInTransac = True
            Do While rsFactura.EOF = False
                'Analizamos cada una de las facturas a anular
                strSql = "select NumAvPg, AvPgEstado,FechaVenceAvPg " & _
                "from AvPgEnc where NumAvPg=" & rsFactura!NumAvPg & " "
                Set rs = DeRia.CoRia.Execute(strSql)
                If rs.RecordCount > 0 Then
                    If rs!AvPgEstado = 6 Then
                        MsgBox "La factura esta distribuida en un plan de pago. No puede anular declaración."
                        Exit Sub
                    End If
                    If rs!AvPgEstado = 2 Then
                        MsgBox ("Factura " & rs!NumAvPg & " esta pagada. No puede anular de esta declaración..?")
                     '   If MsgBox("Factura " & rs!NumAvPg & " esta pagada. Desea detener el proceso de anulación de esta declaración..?", vbQuestion + vbYesNo, "Cuidado..!") = vbYes Then
                            DeRia.CoRia.RollbackTrans
                            blnInTransac = False
                            Exit Sub
                        'Else 'No se pueden eliminar declaraciones ya pagadas. 15/05/2012
                            'Anule el recibo
                            'strSql = "update Recibo set ReciboAnulado=True where NumRecibo in " & _
                            '"(select NumRecibo from ReciboDet where NumFactura=" & rs!NumAvPg & ")"
                            'DeRia.CoRia.Execute (strSql)
                            'Un recibo puede pagar varias facturas
                           
                            'AnularReciboDeFactura rs!NumAvPg
                       ' End If ' anulaba factura
                    End If
                End If
                AnuleFactura rsFactura!NumAvPg
                intUltMesFact = Month(rs!FechaVenceAvPg)
                rsFactura.MoveNext
            Loop
            strSql = "update DeclaraContJurid set EstadoDeclaraIC=2 where CodDeclaraCJ='" & rsDeclaracion!CodDeclaraCJ & "'"
            DeRia.CoRia.Execute (strSql)
            'Al contribuyente le vamos a decir que el ultimo mes factura es el ultimo mes
            'anulado menos 1, solo Enero que es 1 pasa a ser 12
            If intUltMesFact = 1 Then
                intUltMesFact = 12
            Else
                intUltMesFact = intUltMesFact - 1
            End If
            strSql = "update Contribuyente set " & _
            "UltPeriodoFact=" & rsDeclaracion!Periodo & " -1,UltMesFact=" & intUltMesFact & " where " & _
            "Identidad='" & rsDeclaracion!Identidad & "'"
            DeRia.CoRia.Execute (strSql)
            '----------------------------------------------------------------------------
            DeRia.CoRia.CommitTrans
            blnInTransac = False
        End If
    End If
    Exit Sub

CheckError:
    MsgBox Err.Description
    If blnInTransac = True Then DeRia.CoRia.RollbackTrans

End Sub
Public Sub AnuleDeclaracionPersonal(strNumDeclaracion As String)
'Solo se puede anular la ultima declaración realizada para el contribuyente
'seleccione todas las declaraciones con fecha mayore para ese contribuyente contribuyente
'   si hay mas de una, no puede anular
'busque todos los detalles con distinto numero de factura, donde
'numero de referencia sea igual a la declaracion en pantalla.
'para cada detalle
'   Buscamos el encabezado del detalle
'   Si la factura esta pagada. No puede anularse
'   Cambiamos el estado de la factura a anulada
'Cambiamos el estado de la declaracion a anulada
'Restamos 1 al UltPeriodoFacturado del contribuyente.
'

    Dim strSql As String
    Dim rs As New Recordset
    Dim rsFactura As New Recordset
    Dim blnInTransac As Boolean
    Dim intUltMesFact As Integer
    Dim rsDeclaracion As ADODB.Recordset
    
    On Error GoTo CheckError
    'Verifique que el primer campo tiene datos
    blnInTransac = False
    
    Set rsDeclaracion = dal.ObtenerDeclaracionesIPporNumero(strNumDeclaracion)
    'Verificar si el recibo esta operado en contabilidad
    If DeclaracionTieneRecibosPosteadosEnConta(strNumDeclaracion) = True Then
        MsgBox "Algunas facturas ya fueron pagadas y los recibos de estas ya fueron registrados en contabilidad. Esta declaración no se puede anular."
        Exit Sub
    End If
    '----------------------------------------------------
    'Verifique si es la ultima declaracion. Solo puede anular en orden de ultima a primera
    strSql = "select FechaFacturadoIP from DeclaraImpInd  " & _
    "where PeriodoDeclara > " & rsDeclaracion!PeriodoDeclara & " and " & _
    "EstadoDeclaraIP<>2 and Identidad='" & rsDeclaracion!Identidad & "'"
    Set rs = DeRia.CoRia.Execute(strSql)
    If rs.RecordCount > 0 Then
        MsgBox "Esta no es la ultima declaración para este contribuyente. " & _
        "Debe efectuar las anulaciones en orden. De período mayor a menor."
        Exit Sub
    End If
    '-----------------------------------------
    If MsgBox("Realmente desea anular esta declaración...?", vbQuestion + vbYesNo, "Cuidado..!") = vbYes Then
        'Seleccionamos todas las facturas que generó la declaración
        strSql = "select distinct(NumAvPg) as NumAvPg " & _
        "from AvPgDetalle where RefAvPgDet='" & rsDeclaracion!CodDeclaraIP & "' order by NumAvPg desc "
        Set rsFactura = DeRia.CoRia.Execute(strSql)
        If rsFactura.RecordCount > 0 Then
            DeRia.CoRia.BeginTrans
            blnInTransac = True
            Do While rsFactura.EOF = False
                'Analizamos cada una de las facturas a anular
                strSql = "select NumAvPg, AvPgEstado,FechaVenceAvPg " & _
                "from AvPgEnc where NumAvPg=" & rsFactura!NumAvPg & " "
                Set rs = DeRia.CoRia.Execute(strSql)
                If rs.RecordCount > 0 Then
                    If rs!AvPgEstado = 6 Then
                        MsgBox "La factura esta distribuida en un plan de pago. No puede anular declaración."
                        Exit Sub
                    End If
                    If rs!AvPgEstado = 2 Then
                        If MsgBox("Factura " & rs!NumAvPg & " esta pagada. Desea detener el proceso de anulación de esta declaración..?", vbQuestion + vbYesNo, "Cuidado..!") = vbYes Then
                            DeRia.CoRia.RollbackTrans
                            blnInTransac = False
                            Exit Sub
                        Else
                            'Anule el recibo
                            'strSql = "update Recibo set ReciboAnulado=True where NumRecibo in " & _
                            '"(select NumRecibo from ReciboDet where NumFactura=" & rs!NumAvPg & ")"
                            'DeRia.CoRia.Execute (strSql)
                            'Un recibo puede pagar varias facturas
                            AnularReciboDeFactura rs!NumAvPg
                        End If
                    End If
                End If
                AnuleFactura rsFactura!NumAvPg
                intUltMesFact = Month(rs!FechaVenceAvPg)
                rsFactura.MoveNext
            Loop
            strSql = "update DeclaraImpInd set EstadoDeclaraIP=2 where CodDeclaraIP='" & rsDeclaracion!CodDeclaraIP & "'"
            DeRia.CoRia.Execute (strSql)
            'Al contribuyente le vamos a decir que el ultimo mes factura es el ultimo mes
            'anulado menos 1, solo Enero que es 1 pasa a ser 12
            If intUltMesFact = 1 Then
                intUltMesFact = 12
            Else
                intUltMesFact = intUltMesFact - 1
            End If
            strSql = "update Contribuyente set " & _
            "UltPeriodoFact=" & rsDeclaracion!PeriodoDeclara & " -1,UltMesFact=" & intUltMesFact & " where " & _
            "Identidad='" & rsDeclaracion!Identidad & "'"
            DeRia.CoRia.Execute (strSql)
            '----------------------------------------------------------------------------
            DeRia.CoRia.CommitTrans
            blnInTransac = False
        End If
    End If
    Exit Sub
    
CheckError:
    MsgBox Err.Description
    If blnInTransac = True Then DeRia.CoRia.RollbackTrans

End Sub
Public Sub AnularReciboDeFactura(lngNumFactura As Long)
    'anula los recibos que pertenecen a una factura
    'un recibo puede pagar varias facturas, asi que al anular un recibo
    'debemos anular las facturas que paga ese recibo.
    'tambien cada factura puede depender de una declaracion que fija la ultima de pago para un
    'contribuyente.
    
    'Proceso:
    'Obtener el recibo de una factura
    'para cada recibo
    '   seleccione las distintas facturas que paga el recibo de mayor a menor.
    '   Para cada factura
    '       guarde los numeros de declaracion
    '       Si existen varios numeros de declaracion
    '           informe al usuario que debe actualizar manualmente el estado de cuenta del usuario
    '       anule factura
    '   anule el recibo
    
    
    Dim rsRecibos As New ADODB.Recordset
    Dim rsFacturas As New ADODB.Recordset
    Dim rsFactura As New ADODB.Recordset
    Dim strNumDeclara1 As String
    Dim blnVariasDeclaraciones As Boolean
    Dim strSql As String
    
    blnVariasDeclaraciones = False
    'obtenga el numero de recibo con el cual se pago la factura que estamos anulando
    Set rsRecibos = dal.ObtenerRecordset("select distinct(NumRecibo) as NumRecibo from ReciboDet where NumFactura=" & lngNumFactura & "")
    Do While rsRecibos.EOF = False
        'obtenga las facturas que paga el recibo que vamos a anular
        Set rsFacturas = dal.ObtenerRecordset("select distinct(NumFactura) as NumFactura from ReciboDet where NumRecibo=" & rsRecibos!NumRecibo & "")
        If rsFacturas.EOF = False Then
            Set rsFactura = dal.ObtenerDetallesDeFactura(rsFacturas!NumFactura)
            strNumDeclara1 = rsFactura!RefAvPgDet
        End If
        'informe al usuario, si el recibo anula facturas de varias declaraciones
        Do While rsFacturas.EOF = False
            Set rsFactura = dal.ObtenerDetallesDeFactura(rsFacturas!NumFactura)
            If strNumDeclara1 <> rsFactura!RefAvPgDet Then
                MsgBox "El recibo que se esta anulando, anula facturas que vienen de varias declaraciones. Revise el recibo y las facturas que paga y actualize las declaraciones de forma manual, asi como el ultimo periodo facturado para el contribuyente."
            End If
            'anule factura
            AnuleFactura rsFacturas!NumFactura
            rsFacturas.MoveNext
        Loop
        'Anule el recibo
        'strSql = "update Recibo set ReciboAnulado=True where NumRecibo=" & rsRecibos!NumRecibo & " "
        strSql = "update recibo set ReciboAnulado=1, " & _
        "ModificadoPor='" & gsUsername & "', FechaModificado= " & _
        "'" & Format(DiaEnProcesoCT, "dd/mm/yyyy") & "' " & _
        "where " & _
        "NumRecibo=" & rsRecibos!NumRecibo & ""

        DeRia.CoRia.Execute (strSql)
        rsRecibos.MoveNext
    Loop
End Sub
Public Function GeneraFacturaBI(strClaveCatastro As String, dFechaFacturado As Date, strFormula As String) As ADODB.Recordset
    Dim rsFactura As New ADODB.Recordset
    Dim strCtaIngreso As String
    Dim rsSysPar As New Recordset
    Dim RsCat As New Recordset
    Dim RsDeclara As New Recordset
    Dim rsPar As New Recordset
    Dim rs As New Recordset
    Dim curImpuesto As Currency
    Dim curInteres As Currency
    Dim curRecargo As Currency
    Dim curDescuentoTE As Currency
    Dim curDescuentoPE As Currency
    Dim curMulta As Currency
    Dim curValor As Currency
    Dim curTotal As Currency
    Dim lngPeriodoAFacturar As Long
    Dim blnAplicaTE As Boolean
    Dim sngTasaDescTE As Single
    Dim strNombre As String
    Dim strFormulaInteres As String
    Dim strFormulaRecargo As String
    Dim dFechaVence As Date
    
    rsFactura.Fields.Append "CtaIngreso", adBSTR, 50
    rsFactura.Fields.Append "Valor", adCurrency
    rsFactura.Open
        
    Set RsCat = DeRia.CoRia.Execute("select * from Catastro where ClaveCatastro='" & strClaveCatastro & "' ")
    If RsCat.RecordCount = 0 Then
        MsgBox "Clave Catastral no encontrada."
        Exit Function
    End If
    If IsNull(RsCat!UltPeriodoFact) Then
        MsgBox "Debe definir el ultimo periodo que se facturó Bienes Inmuebles para esta propiedad: " & strClaveCatastro
        Exit Function
    End If
    If IsNull(RsCat!Impuesto) Or RsCat!Impuesto <= 0 Then
        MsgBox "La propiedad: " & strClaveCatastro & ", no tiene el avaluo correspondiente."
        Exit Function
    End If
    lngPeriodoAFacturar = RsCat!UltPeriodoFact + 1
    'Extraemos los parametros de impuestos para el periodo a facturar
    Set rsPar = DeRia.CoRia.Execute("select * from ParamRia where PeriodoFact=" & lngPeriodoAFacturar & " ")
    If rsPar.RecordCount = 0 Then
        MsgBox "No estan definidos los datos de parametros para el " & lngPeriodoAFacturar & ". Factura no puede crearse."
        Exit Function
    End If
    dFechaVence = rsPar!BiFechaMaxPago
    'Me.lblFechaVence = dFechaVence
    'Calculamos el impuesto de la propiedad
    curImpuesto = 0: curInteres = 0: curRecargo = 0: curDescuentoTE = 0
    curDescuentoPE = 0: curMulta = 0
    'Si la fecha que vence la factura es de un año anterior al actual va a recuperaciones.
    strCtaIngreso = GetCtaRecuperacion(IIf(RsCat!Ubicacion = 1, rsSysPar!CtaIngresoBiRural, rsSysPar!CtaIngresoBiUrb), DeRia.CoRia, DiaEnProcesoCT(), dFechaVence)
    '------------------------------------------------------------------------------------
    'Esto lo remuevo porque cuando el avaluo viene de catastro la funcion de abajo no esta tomando
    'en cuenta los valores de concertacion.
    '24 Octubre 2008
    'curImpuesto = CalculeImpuestoBI(Me.txtClaveCatastro, lngPeriodoAFacturar)
    curImpuesto = RsCat!Impuesto
    FindCuentaIngreso rs, DeRia.CoRia, strCtaIngreso, Format(DiaEnProcesoCT, "YYYY")
    If rs.RecordCount = 0 Then
        MsgBox "Cuenta " & strCtaIngreso & " no existe en el catalogo de cuentas de ingreso tributarias." & Chr(13) & _
        "Verifique los parametros del sistema SAFT y corrija la cuenta del impuesto de Bienes Inmuebles."
        Exit Function
    End If
    curTotal = curImpuesto
    rsFactura.AddNew
    rsFactura!CtaIngreso = strCtaIngreso
    rsFactura!valor = curImpuesto
    rsFactura.Update
    '--------------------------------------
    'Calculamos descuento por pagos adelantados, este van negativo
    strCtaIngreso = rsSysPar!CtaIngresoDescuento
    curValor = CalculeDescuentoBI(curImpuesto, dFechaVence, dFechaFacturado) * -1
    If Abs(curValor) > 0 Then
        FindCuentaIngreso rs, DeRia.CoRia, strCtaIngreso, Format(DiaEnProcesoCT, "YYYY")
        curTotal = curTotal + curValor
        curDescuentoPE = curValor
        rsFactura.AddNew
        rsFactura!CtaIngreso = strCtaIngreso
        rsFactura!valor = curDescuentoPE
        rsFactura.Update
    End If
    '------------------------------------------
    'Calculamos descuento por tercera edad, este va negativo
    If RsCat!HabitaPropietario = 1 Then
        'strCtaIngreso = rsSysPar!CtaIngresoCJ
        AplicaTerceraEdad RsCat!Identidad, blnAplicaTE, strCtaIngreso, strNombre, sngTasaDescTE, DeRia.CoRia, dFechaVence
        If blnAplicaTE Then
            Set rs = DeRia.CoRia.Execute("select DescMaxBI from Parametro")
            'Se aplica un porcentaje de descuento por los primeros DescMaxBI, si el impuesto
            'es menor se aplica a todo el impuesto
            If curImpuesto <= rs!DescMaxBi Then
                curValor = (curImpuesto * sngTasaDescTE)
            Else
                curValor = (rs!DescMaxBi * sngTasaDescTE)
            End If
            curValor = curValor * -1
            FindCuentaIngreso rs, DeRia.CoRia, strCtaIngreso, Format(DiaEnProcesoCT, "YYYY")
            curTotal = curTotal + curValor
            curDescuentoTE = curValor
            rsFactura.AddNew
            rsFactura!CtaIngreso = strCtaIngreso
            rsFactura!valor = curDescuentoTE
            rsFactura.Update
        End If
    End If
    '------------------------------------------
    'Calculamos la multa por declarar tarde.
    'Extraemos la ultima declaracion, si la hay, para ver si hay multa por declarar tarde
    Set RsDeclara = DeRia.CoRia.Execute("select * from DeclaraBI where ClaveCatastro='" & strClaveCatastro & "' and PeriodoDeclaBI =" & lngPeriodoAFacturar & " ")
    If RsDeclara.RecordCount > 0 Then
        curValor = IIf(IsNull(RsDeclara!RecargoBI), 0, RsDeclara!RecargoBI)
    End If
    If curValor > 0 Then
        strCtaIngreso = rsSysPar!CtaIngresoMultaDeclaraTarde
        FindCuentaIngreso rs, DeRia.CoRia, strCtaIngreso, Format(DiaEnProcesoCT, "YYYY")
        curTotal = curTotal + curValor
        curMulta = curValor
        rsFactura.AddNew
        rsFactura!CtaIngreso = strCtaIngreso
        rsFactura!valor = curMulta
        rsFactura.Update
    End If
    'Calculamos el interes
    curValor = Interes(curImpuesto, dFechaFacturado, dFechaVence, strFormulaInteres)
    If curValor > 0 Then
        strCtaIngreso = rsSysPar!CtaIngresoIntImp
        FindCuentaIngreso rs, DeRia.CoRia, strCtaIngreso, Format(DiaEnProcesoCT, "YYYY")
        curTotal = curTotal + curValor
        curInteres = curValor
        rsFactura.AddNew
        rsFactura!CtaIngreso = strCtaIngreso
        rsFactura!valor = curInteres
        rsFactura.Update
    End If
    'Calculamos el recargo, es sobre el impuesto+Interes sin descuentos
    curValor = Recargo(curImpuesto + curInteres, dFechaVence, dFechaFacturado, strFormulaRecargo)
    If curValor > 0 Then
        strCtaIngreso = rsSysPar!CtaIngresoRecargoImp
        FindCuentaIngreso rs, DeRia.CoRia, strCtaIngreso, Format(DiaEnProcesoCT, "YYYY")
        curTotal = curTotal + curValor
        curRecargo = curValor
        rsFactura.AddNew
        rsFactura!CtaIngreso = strCtaIngreso
        rsFactura!valor = curRecargo
        rsFactura.Update
    End If
    
End Function


