Attribute VB_Name = "Utilities"
Sub SetEditMode(lv_Flag As Integer)
    If lv_Flag = pv_ModoEdicion Then
        For Each Control In Screen.ActiveForm.Controls
            Select Case Control.Tag
                Case "IE" 'Inactive on Data Entry
                    Control.Enabled = False
                Case "AE" 'Active on Data Entry
                    If TypeOf Control Is TextBox Then
                        Control.BackColor = vbWhite
                        Control.Locked = False ' Only textbox can be locked
                    End If
                    If TypeOf Control Is MaskEdBox Then
                        Control.BackColor = vbWhite
                    End If
                    Control.Enabled = True
            End Select
        Next Control
    Else
        For Each Control In Screen.ActiveForm.Controls
            Select Case Control.Tag
                Case "AE"
                    If TypeOf Control Is TextBox Then
                        Control.BackColor = vbWhite
                        Control.Locked = True
                    End If
                    If TypeOf Control Is MaskEdBox Then
                        'Control.BackColor = vbWhite
                    End If
                    Control.Enabled = False
                Case "IE"
                    Control.Enabled = True
            End Select
        Next Control
    End If
End Sub
Function CalcImpCat(lv_ValorPropiedad As Currency, lv_Tasa As Single)
    Dim lv_Miles As Single

    lv_Miles = lv_ValorPropiedad / 1000
    CalcImpCat = lv_Miles * lv_Tasa
End Function
Function SearchContrib()
    Load frmListaContrib
    SearchContrib = pv_Identidad
End Function
Function Discount(lv_Impuesto As Currency, lv_FechaPaga As Date, lv_FechaVence As Date)
    'lv_Impuesto= a impuesto calculado
    'lv_FechaPaga = Fecha en que llega a pagar
    'lv_FechaVence = Fecha que vence el impuesto
    'Este procedimiento calcula un descuento, si paga 4 meses antes del
    'vencimiento de un impuesto.
    Dim Meses As Integer
    Dim rsParamRia As New Recordset
    
    Discount = 0
    Set rsParamRia = DeConta.ConConta.Execute("select * from ParamRia")
    If rsParamRia.RecordCount = 0 Then
        MsgBox "Parametros del modulo no estan definidos...!"
        Exit Function
    End If
    Meses = DateDiff("m", lv_FechaPaga, lv_FechaVence)
    If Meses >= rsParamRia!TiempoParaDescuento Then
        Discount = lv_Impuesto * rsParamRia!DescuentoPagoAnticipado
    End If
End Function
Function CalcImpMensual(lv_Cantidad As Integer, lv_Cuenta As String)
    If lv_Cantidad <= 0 Then
        CalcImpMensual = 0
        Exit Function
    End If
    If DeRia.rscmdGetCtaIngreso.State = adStateOpen Then
        DeRia.rscmdGetCtaIngreso.Close
    End If
    DeRia.cmdGetCtaIngreso (lv_Cuenta)
    If DeRia.rscmdGetCtaIngreso.EOF() Then
        'mensaje que cuenta no se encontro
        CalcImpMensual = 0
    Else
        CalcImpMensual = DeRia.rscmdGetCtaIngreso!ValorMensual * lv_Cantidad
    End If
End Function
Function MultaDeclaraTardeIC(lv_FechaDeclara As Date, lv_ValorMensual As Single)
    ' Si presenta la declaracion, despues de la fecha de vencimiento, tiene un
    ' una multa
    ' lv_FechaDeclara = Fecha que presenta la declaracion
    ' lv_ValorMulta= Es el valor de la multa que pagara
    ' lv_FechaTope = Cargamos la fecha maxima de presentacion de declaracion
    ' si la fecha que presenta la declaracion es mayor que la fecha maxima
    ' entonces le aplicamos el valor que viene en lv_ValorMulta
    Dim rsParamRia As New Recordset
    
    Set rsParamRia = DeConta.ConConta.Execute("select * from ParamRia")
    If rsParamRia.RecordCount = 0 Then
        MsgBox "Parametros del modulo no definidos"
        Exit Function
    End If
    
    MultaDeclaraTardeIC = 0
    If lv_ValorMensual <= 0 Then
        Exit Function
    End If
    If lv_FechaDeclara > rsParamRia!ICFechaMaxDeclara Then
        MultaDeclaraTardeIC = lv_ValorMensual * rsParamRia!ICMultaDeclaraTarde
    End If
End Function
Sub OpenParametros()
    If DeRia.rscmdParametro.State = adStateOpen Then
        DeRia.rscmdParametro.Close
    End If
    DeRia.cmdParametro
End Sub
Public Sub DeleteAllRecords(TargetSet As Recordset)
    'borre los regitros del temporal
    Dim i As Integer
    Dim n As Integer
    
    With TargetSet
        n = .RecordCount
        If n = 0 Then Exit Sub
        .MoveFirst
        For i = 1 To n
            .Delete
            .MoveNext
        Next i
    End With
End Sub

Sub EditCopyProc()
    ' Copia el texto seleccionado al Portapapeles.
    Clipboard.SetText Screen.ActiveControl.SelText
End Sub

Sub EditCutProc()
    ' Copia el texto seleccionado al Portapapeles.
    Clipboard.SetText frmMDI.ActiveForm.ActiveControl.SelText
    ' Elimina el texto seleccionado.
    frmMDI.ActiveForm.ActiveControl.SelText = ""
End Sub
Sub EditPasteProc()
    ' Coloca el texto del Portapapeles en el control activo.
    frmMDI.ActiveForm.ActiveControl.SelText = Clipboard.GetText()
End Sub
Public Function BusqueContrib(Id As String)
    With DeRia.rscmdBuscaContrib
        If .State = adStateOpen Then .Close
        DeRia.cmdBuscaContrib (Id)
        If .RecordCount = 0 Then
            BusqueContrib = ""
            Exit Function
        End If
        BusqueContrib = !ContNombre
    End With
End Function
Function OpenDataPath()
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    Dim fs, f, ts, s
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(App.Path + "\" + Trim(PathFileName))
    Set ts = f.OpenAsTextStream(ForReading, TristateUseDefault)
    s = ts.ReadLine
    'MsgBox s
    ts.Close
    OpenDataPath = s
End Function

Public Function GetMonth(dFecha As Date) As String
    Dim i As Integer
    Dim s As String
    
    i = DatePart("m", dFecha)
    Select Case i
        Case 1
            s = "Enero"
        Case 2
            s = "Febrero"
        Case 3
            s = "Marzo"
        Case 4
            s = "Abril"
        Case 5
            s = "Mayo"
        Case 6
            s = "Junio"
        Case 7
            s = "Julio"
        Case 8
            s = "Agosto"
        Case 9
            s = "Septiembre"
        Case 10
            s = "Octubre"
        Case 11
            s = "Noviembre"
        Case 12
            s = "Diciembre"
    End Select
    GetMonth = s
End Function
Public Function GetMonthYearWord(dFecha As Date) As String
    Dim s As String
    
    s = GetMonth(dFecha)
    s = s + " del " + Str(Year(dFecha))
    GetMonthYearWord = s
End Function
Public Function NettoWorkdays(ByVal dtmStart As Date, ByVal dtmEnd As Date) As Integer

'This function calculates the number of working days (monday to friday) between 2 dates,
'including the first and the last day
    
    Dim intDays As Integer
    Dim intSubtract As Integer
    
    ' if end is smaller then start return -1
    If dtmEnd < dtmStart Then
        NettoWorkdays = -1
    Else
        ' Get the start and end dates to be weekdays.
        While Weekday(dtmStart) = vbSaturday Or Weekday(dtmStart) = vbSunday
            dtmStart = dtmStart + 1
        Wend
        While Weekday(dtmEnd) = vbSaturday Or Weekday(dtmEnd) = vbSunday
            dtmEnd = dtmEnd - 1
        Wend
        If dtmStart > dtmEnd Then
            ' Sorry, no Workdays to be had. Just return 0.
            NettoWorkdays = 0
        Else
            intDays = dtmEnd - dtmStart + 1
            
            ' Subtract off weekend days.  Do this by figuring out how
            ' many calendar weeks there are between the dates, and
            ' multiplying the difference by two (because there are two
            ' weekend days for each week). That is, if the difference
            ' is 0, the two days are in the same week. If the
            ' difference is 1, then we have two weekend days.
            intSubtract = (DateDiff("ww", dtmStart, dtmEnd) * 2)
            
            NettoWorkdays = intDays - intSubtract
        End If
    End If
    
End Function
Public Function GetSaldo(dFecha As Date, sCuenta As String) As Currency
    Dim rsSet As New ADODB.Recordset
    
    'Si la fecha es del mes en proceso, trae el saldo de Catalogo
    'sino la trae de Cierres
    GetSaldo = 0 'mandamos cero, en caso de error
    Set rsSet = DeConta.ConConta.Execute("select MesActivoInicio, MesActivoFin from ParametroCont")
    If rsSet.EOF And rsSet.BOF Then
        MsgBox "Error 207: Error en el archivo de parametros, llame a soporte tecnico."
        Exit Function
    End If
    If dFecha >= rsSet!MesActivoInicio And dFecha <= rsSet!MesActivoFin Then
        'hace la consulta del archivo Catalogo
        Set rsSet = DeConta.ConConta.Execute("select DebitoAcumulado-CreditoAcumulado as Saldo from Catalogo where CtaContable='" & sCuenta & "'")
        If rsSet.EOF And rsSet.BOF Then
            MsgBox "Error 207: No se encontraron registros para la cuenta " + sCuenta
            Exit Function
        End If
        If Not IsNull(rsSet!Saldo) Then GetSaldo = rsSet!Saldo
    Else
        'si la fecha es menor al mes en proceso, haga query desde Cierres
        'sino saldo=0, pide datos de fecha posterior al mes en proceso
        If dFecha < rsSet!MesActivoInicio Then
            Set rsSet = DeConta.ConConta.Execute("select SaldoDebe-SaldoHaber as Saldo from Cierre where MesCerrado=" & dFecha & " and CtaContable='" & sCuenta & "'")
            If rsSet.EOF And rsSet.BOF Then
                'MsgBox "Error 207: No se encontraron registros de cierres, para la cuenta " + sCuenta
                Exit Function
            End If
            GetSaldo = rsSet!Saldo
        End If
    End If
End Function

Public Sub SiguienteRegistro(rs As Recordset)
    On Error GoTo CheckError
    
    rs.MoveNext
    If rs.EOF Then
        rs.MoveLast
        MsgBox "Estamos en el ultimo registro....!", vbQuestion, "Mensaje de Sistema"
    End If
    Exit Sub
    
CheckError:
    MsgBox Err.Description
End Sub

Public Sub AnteriorRegistro(rs As Recordset)
    On Error GoTo CheckError
    
    rs.MovePrevious
    If rs.BOF Then
        rs.MoveFirst
        MsgBox "Estamos en el primer registro....!", vbQuestion, "Mensaje de Sistema"
    End If
    Exit Sub
    
CheckError:
    MsgBox Err.Description
End Sub
Public Sub SizeForm(x As Long, y As Long, FormName As Form)
    Dim TopCorner As Integer
    Dim LeftCorner As Integer
    
    FormName.Width = x
    FormName.Height = y
    TopCorner = (Screen.Height - FormName.Height - 1500) / 2
    'el valor 1500 arriba, representa el status bar, para que centre la vertical bien
    LeftCorner = (Screen.Width - FormName.Width) / 2
    'Me.Move LeftCorner, TopCorner 'center form
End Sub

Public Sub UltimoRegistro(rs As Recordset)
    On Error GoTo CheckError
    
    rs.MoveLast
    Exit Sub
    
CheckError:
    MsgBox Err.Description
End Sub
Public Sub PrimerRegistro(rs As Recordset)
    On Error GoTo CheckError
    
    rs.MoveFirst
    Exit Sub
    
CheckError:
    MsgBox Err.Description
End Sub
Public Sub BorrarRegistro(rs As Recordset)

    On Error GoTo BorrarRegistroError
    
    If MsgBox("Realmente desea borrar este registro...?", vbQuestion + vbYesNo, "Cuidado..!") = vbYes Then
        rs.Delete
        'rs.Requery  cuando se trabaja con bound controls a veces los deja readonly. Shit.
    End If
    Exit Sub
    
BorrarRegistroError:
    MsgBox Err.Description
    Exit Sub
End Sub
Public Function SearchRecordset(vValue As Variant, fld As String, rs As Recordset) As Boolean
    'busca en el recordset, si existe el valor que viene en vValue
    'retorna True, si se encuentra, False si no
    
    SearchRecordset = False
    rs.MoveFirst
    Do While Not rs.EOF
        'If rs! & fld& = vValue Then
            SearchRecordset = True
            Exit Function
        'End If
        rs.MoveNext
    Loop
End Function
Public Sub CalculeEstadoCuenta(dFechaFact As Date, sId As String, cSaldo As Currency, cRecargo As Currency)
    'PROCESO DESCARTADO
    'Esta rutina calcula el estado de cuenta para un determinado contribuyente
    'El Proceso es el siguiente:
    '1= Seleccione el ultimo aviso de pago emitido para el contribuyente
    '2=Si esta vencido
    '   a=Si el Valor Pagado es menor que el total facturado (Total del Mes + Saldoanterior + Recargo)
    '       -calcule el saldo pendiente
    '       -calcule el recargo sobre el saldo
    '3=Busque cargos pendientes no pagados, para el contribuyente
    '4=si existen
    '   a=Sume a saldo pendiente
    
    Dim sRemain As Single
    Dim rsAvPg As New Recordset
    Dim rsParRia As New Recordset
    
    cRecargo = 0
    cSaldo = 0
    'consulte los porcentajes de recargo
    Set rsParRia = DeConta.ConConta.Execute("select * from ParamRia order by PeriodoFact")
    If rsParRia.RecordCount = 0 Then
        MsgBox "Parametros del modulo de contribuyentes no estan definidos..!"
        Exit Sub
    End If
    'Seleccione los avisos de pago
    Set rsAvPg = DeConta.ConConta.Execute("Select * from  AvPgEnc where Identidad='" & sId & "' and AvPgEstado=1") 'Facturas no pa
    If rsAvPg.RecordCount > 0 Then
        rsAvPg.MoveLast 'vaya al ultimo aviso de pago
        'vea si eta vencido
        If rsAvPg!FechaVenceAvPg < dFechaFact Then
            ' ahora consulte si no pago todo
            If rsAvPg!avpgValorPagado < (rsAvPg!AvPgTotalPeriodo + rsAvPg!AvPgSaldoAnterior + rsAvPg!AvPgRecargo) Then
                'No pago el total de la factura, entonces calculamos el recargo para el sobrante
                'el recargo solo es sobre el saldos en mora, no incluye otros recargos
                'cuando se paga una factura, y no se paga en su totalidad, primero se pagan el recargo, luego
                'el saldo en mora, y despues el valor del periodo.
                
                If rsAvPg!avpgValorPagado > rsAvPg!AvPgRecargo Then
                    sRemain = rsAvPg!ValorPagado - rsAvPg!TotalPeri
                End If
                sRemain = rsAvPg!avpgValorPagado - rsAvPg!AvPgRecargo
                If sRemain > 0 Then
                    'al sobrante reste TotalPeriodo-SaldoAnterior
                    sRemain = (rsAvPg!AvPgTotalPeriodo + rsAvPg!AvPgSaldoAnterior) - sRemain
                Else
                    'no queda sobrante, entonces asigne el valor del mes + saldo anterior
                    sRemain = rsAvPg!AvPgSaldoAnterior + rsAvPg!AvPgTotalPeriodo
                End If
                'Aplica el recargo a sRemain
                'Consultar los porcentajes de recargo
                cSaldo = rsAvPg!AvPgTotalPeriodo + rsAvPg!AvPgSaldoAnterior
                cRecargo = sRemain * (rsParRia!RecargoAtrasoPago / 12) 'por estar moroso
                cRecargo = cRecargo + (sRemain * (rsParRia!RecargoSobreSaldo / 12)) 'recargo sobre saldo
            Else
                'pago todo, o pago de mas..!!
                cSaldo = 0
                cRecargo = 0
            End If
        Else
            'No esta vencido el aviso de pago
            cSaldo = rsAvPg!AvPgSaldoAnterior
            cRecargo = rsAvPg!AvPgRecargo
        End If
    End If
End Sub
Function MultaDeclaraTarde(lv_FechaDeclara As Date, lv_ValorMulta As Single, lv_FechaTope)
    ' Si presenta la declaracion, despues de la fecha de vencimiento, tiene un
    ' una multa
    ' lv_FechaDeclara = Fecha que presenta la declaracion
    ' lv_ValorMulta= Es el valor de la multa que pagara
    ' lv_FechaTope = Cargamos la fecha maxima de presentacion de declaracion
    ' si la fecha que presenta la declaracion es mayor que la fecha maxima
    ' entonces le aplicamos el valor que viene en lv_ValorMulta
    
    MultaDeclaraTarde = 0
    If lv_ValorMulta <= 0 Then
        Exit Function
    End If
    If lv_FechaDeclara > lv_FechaTope Then
        MultaDeclaraTarde = lv_ValorMulta
    End If
End Function
Function GetNextFactura() As Long
    'Warning: this is not multiuser.
    Dim rsPar As New Recordset
    
    Set rsPar = DeConta.ConConta.Execute("select UltNumFact as NumFactura from ParametroCont")
    If rsPar.RecordCount = 0 Or IsNull(rsPar!NumFactura) Then
        GetNextFactura = 1
    Else
        GetNextFactura = rsPar!NumFactura + 1
    End If
End Function
Public Sub CalculeRecargoPorAtrasoPago(dFecha1 As Date, dFecha2 As Date, cValorVencido As Currency, cPorAtraso As Currency, cPorSaldo As Currency)
    Dim rsParRia As New Recordset
    
    'consulte los porcentajes de recargo
    Set rsParRia = DeConta.ConConta.Execute("select * from ParamRia")
    If rsParRia.RecordCount = 0 Then
        MsgBox "Parametros del modulo de contribuyentes no estan definidos..!"
        Exit Sub
    End If
    'cPorAtraso = cValorVencido * (rsParRia!RecargoAtrasoPago / 12) 'por estar moroso
    'cPorSaldo = cValorVencido * (rsParRia!RecargoSobreSaldo / 12) 'recargo sobre saldo
    cPorAtraso = CalculeRecargoMensual(dFecha1, dFecha2, cValorVencido, rsParRia!RecargoAtrasoPago / 12)
    cPorSaldo = CalculeRecargoMensual(dFecha1, dFecha2, cValorVencido, rsParRia!RecargoSobreSaldo / 12)
End Sub
Function CalculeRecargoMensual(dFecha1 As Date, dFecha2 As Date, cValorEnMora As Currency, sPorcRecargo As Currency)
    'Calcula el recargo por cada mes que no paga.
    Dim iMeses As Integer

    CalculeRecargoMensual = 0
    iMeses = DateDiff("m", dFecha1, dFecha2)
    If iMeses > 0 Then
        CalculeRecargoMensual = cValorEnMora * sPorcRecargo * iMeses
    End If
End Function
Public Sub CalculeMora(sId As String, dFecha2 As Date, cMora As Currency, cSaldo As Currency)
    Dim RsFact As New Recordset
    Dim rsPar As New Recordset
    Dim iMeses As Integer
    Dim cRecargoxSaldo As Currency
    Dim cRecargoPorMora As Currency
    
    cMora = 0
    cSaldo = 0
    'consulte los porcentajes de recargo
    Set rsPar = DeConta.ConConta.Execute("select * from ParamRia")
    If rsPar.RecordCount = 0 Then
        MsgBox "Parametros del modulo de contribuyentes no estan definidos..!"
        Exit Sub
    End If
    Set RsFact = DeConta.ConConta.Execute("Select * from AvPgEnc where Identidad='" & sId & "' and AvPgEstado=1")
    If RsFact.RecordCount = 0 Then Exit Sub
    RsFact.MoveFirst
    Do While RsFact.EOF = False
        'vea si esta vencida
        iMeses = DateDiff("m", RsFact!FechaVenceAvPg, dFecha2)
        If iMeses > 0 Then
            'calcule el saldo
            cSaldo = cSaldo + RsFact!AvPgTotalPeriodo + RsFact!AvPgRecargo
            'calcule el recargo
            CalculeRecargoPorAtrasoPago RsFact!FechaVenceAvPg, dFecha2, RsFact!AvPgTotalPeriodo, cRecargoxSaldo, cRecargoPorMora
            cMora = cMora + CalculeRecargoMensual(RsFact!FechaVenceAvPg, dFecha2, RsFact!AvPgTotalPeriodo + RsFact!AvPgRecargo, 0.01)
            'si es Bienes Inmuebles, tiene otro impuesto mensual
            If RsFact!AvPgTipoImpuesto = 1 Then
                cMora = cMora + CalculeRecargoMensual(RsFact!FechaVenceAvPg, dFecha2, RsFact!AvPgTotalPeriodo + RsFact!AvPgRecargo, rsPar!BIRecargoAtrasoPago)
            End If
            cMora = cMora + cRecargoPorMora + cRecargoxSaldo
        End If
        RsFact.MoveNext
    Loop
End Sub
Public Sub AgregueDetalle(Renglon As String, Cantidad As Single, Catastro As String, Factura As Long)
    'the recordset must be opened before calling this routine
    With DeRia.rscmdDetFactura
        .AddNew
        !CtaIngreso = Renglon
        !ValorUnitAvPgDet = Cantidad
        !CantAvPgDet = 1
        !ClaveCatastro = Catastro
        !NumAvPg = Factura
        .Update
    End With
End Sub
Function ValideId(sId As String) As String
    If DeRia.rscmdBuscaContrib.State = adStateOpen Then DeRia.rscmdBuscaContrib.Close
    'On Error Resume Next
    DeRia.cmdBuscaContrib sId
    If DeRia.rscmdBuscaContrib.RecordCount = 0 Then
        ValideId = ""
    Else
    ' On Error Resume Next '''Corregir MSx3 iisNull
        ValideId = IIf(IsNull(DeRia.rscmdBuscaContrib!ContNombre), "", DeRia.rscmdBuscaContrib!ContNombre)
    End If
    
    'cmdContrib
    If ValideId = "" Then
    
    If DeRia.rscmdContrib.State = 1 Then DeRia.rscmdContrib.Close
    DeRia.rscmdContrib.Open ("Select * from Contribuyente where identidad = '" & sId & "'")
    If DeRia.rscmdContrib.RecordCount = 0 Then
        ValideId = ""
    Else
    ' On Error Resume Next '''Corregir MSx3 iisNull
        ValideId = IIf(IsNull(Trim(DeRia.rscmdContrib!PNombre)), "", Trim(DeRia.rscmdContrib!PNombre))
    End If
    
    End If
End Function

Function ValideCtaIngreso(sCta As String) As String
    Dim rs As New Recordset
    
    Set rs = DeConta.ConConta.Execute("Select * from CuentaIngreso where CtaIngreso='" & sCta & "'")
    If rs.RecordCount = 0 Then
        ValideCtaIngreso = ""
    Else
        ValideCtaIngreso = rs!NombreCtaIngreso
    End If
End Function

Function ValideClaveCatastro(sClave As String) As String
    Dim rs As New Recordset
    
    Set rs = DeRia.CoRia.Execute("select * from Catastro where ClaveCatastro='" & sClave & "'")
    If rs.RecordCount = 0 Then
        ValideClaveCatastro = ""
    Else
        ValideClaveCatastro = "Existe"
    End If
End Function
Function CenterForm(FormName As Form)
    Dim TopCorner As Integer
    Dim LeftCorner As Integer
    Dim iWinBar As Integer
    
    iWinBar = 1500 'el valor 1500 arriba, representa el status bar de windows, para que centre la vertical bien
    TopCorner = (Screen.Height - FormName.Height - iWinBar) / 2
    LeftCorner = (Screen.Width - FormName.Width) / 2
    FormName.Move LeftCorner, TopCorner
End Function
Public Sub GoLastRecord(rs As Recordset)
    On Error GoTo CheckError
    
    rs.MoveLast
Exit_Sub:
    Exit Sub
    
CheckError:
    MsgBox Err.Description
    GoTo Exit_Sub
End Sub
Public Sub GoFirstRecord(rs As Recordset)
    On Error GoTo CheckError
    
    rs.MoveFirst
Exit_Sub:
    Exit Sub
    
CheckError:
    MsgBox Err.Description
    GoTo Exit_Sub
End Sub
Public Sub GoNextRecord(rs As Recordset)
    On Error GoTo CheckError

    If rs.EOF = False Then
        rs.MoveNext
    End If
    If rs.EOF = True Then
        rs.MovePrevious
        MsgBox "Estamos en el primer registro....!", vbQuestion, "Mensaje de Sistema"
    End If
    
Exit_Sub:
    Exit Sub
    
CheckError:
    MsgBox Err.Description
    GoTo Exit_Sub
End Sub
Public Sub GoPrevRecord(rs As Recordset)
    On Error GoTo CheckError

    If rs.BOF = False Then
        rs.MovePrevious
    End If
    If rs.BOF = True Then
        rs.MoveNext
        MsgBox "Estamos en el primer registro....!", vbQuestion, "Mensaje de Sistema"
    End If
    
Exit_Sub:
    Exit Sub

CheckError:
    MsgBox Err.Description
    GoTo Exit_Sub
End Sub
Public Sub BindFields(rRs As Recordset)
    For Each Control In Screen.ActiveForm.Controls
        If Control.Tag = "AE" Then
            If TypeOf Control Is TextBox Or TypeOf Control Is MaskEdBox Or TypeOf Control Is CheckBox Then
                Set Control.DataSource = rRs
            End If
        End If
    Next Control
End Sub
Public Function MakeDate(sDate As String) As String
    
    If IsDate(sDate) Then 'si la fecha esta bien, regresela
        MakeDate = sDate
        Exit Function
    End If
    If Len(sDate) = 8 Then 'Se espera el string ddmmaaaa
        MakeDate = Mid(sDate, 1, 2) & "/" & Mid(sDate, 3, 2) & "/" & Mid(sDate, 5, 4)
    End If
    If Len(sDate) = 6 Then 'se espera el string ddmmaa
        MakeDate = Mid(sDate, 1, 2) & "/" & Mid(sDate, 3, 2) & "/20" & Mid(sDate, 5, 2)
    End If
    If Len(sDate) = 2 Then 'se le agrega el mes y año actual
        MakeDate = Trim(sDate) & "/" & Month(Date) & "/" & Year(Date)
    End If
End Function
Public Function CanEdit(rs As Recordset) As Boolean
    On Error GoTo CanEditError
    
    CanEdit = False
    If rs.RecordCount = 0 Or rs.EOF = True Or rs.BOF = True Then
        MsgBox "No puede editar. No existen registros, o no existe ninguno seleccionado."
    End If
    CanEdit = True
    Exit Function
    
CanEditError:
    'Error 91 es cuando no esta abierto rs, y no debemos desplegar ese error.
    If Err.Number <> 91 Then MsgBox "Error Número: " & Err.Number & ":" & Err.Description
End Function
Public Sub GotoRecord(bRec As Integer, rs As Recordset)
    '0=Primer registro,1=registro anterior,2=Siguiente registro,3=Ultimo registro
    On Error GoTo GotoRecordError
    
    If Not IsObject(rs) Then Exit Sub
    If rs.RecordCount = 0 Then Exit Sub
    If bRec = 0 Then rs.MoveFirst
    If bRec = 1 Then
        If rs.BOF = False Then rs.MovePrevious
        'Si esta al inicio deja un registro en blanco
        If rs.BOF = True Then rs.MoveNext
    End If
    If bRec = 2 Then
        If rs.EOF = False Then rs.MoveNext
        'Si esta al final deja un registro en blanco
        If rs.EOF = True Then rs.MovePrevious
    End If
    If bRec = 3 Then rs.MoveLast
    Exit Sub
    
GotoRecordError:
    'Error 91 es cuando no esta abierto rs, y no debemos desplegar ese error.
    If Err.Number <> 91 Then MsgBox "Error Número: " & Err.Number & ":" & Err.Description
End Sub
Public Function EstadoFactura(bEstado As Byte) As String
    Select Case bEstado
        Case 1
            EstadoFactura = "No Pagada"
        Case 2
            EstadoFactura = "Pagada"
        Case 3
            EstadoFactura = "Anulada"
        Case 4
            EstadoFactura = "En Tesoreria"
        Case 5
            EstadoFactura = "Pago Parcial"
        Case 6
            EstadoFactura = "Plan Pago"
    End Select
End Function
Public Sub ClearControls()
    Dim sMask As String
    
    For Each Control In Screen.ActiveForm.Controls
        If Control.Tag = "AE" Then
            If TypeOf Control Is TextBox Then Control.Text = ""
            If TypeOf Control Is MaskEdBox Then
                sMask = Control.Mask
                Control.Mask = ""
                Control.Text = ""
                Control.Mask = sMask
            End If
            If TypeOf Control Is DataCombo Then Control.Text = ""
        End If
    Next
    
End Sub

Public Sub BuildSearch(sSql As String, sField As String, sValue As String, sTipo As String)
    If sValue = "" Or sValue = "0" Then Exit Sub
    If Len(sSql) > 0 Then sSql = sSql & " and "
    Select Case sTipo
        Case "Numeric"
            sSql = sSql & " " & sField & "=" & sValue
        Case "String"
            If Mid(sValue, Len(sValue), 1) = "%" Then
                sSql = sSql & " " & sField & " like '" & sValue & "'"
            Else
                sSql = sSql & " " & sField & "='" & sValue & "'"
            End If
        Case "Date"
            sSql = sSql & " " & sField & "='" & sValue & "'"
    End Select
End Sub
Public Function PeriodoEnLetras(dFecha1 As Date, dFecha2 As Date) As String
    Dim s As String
    s = "Período: " & Day(dFecha1) & "-" & GetMonth(dFecha1) & "-" & Year(dFecha1)
    s = s & " al " & Day(dFecha2) & "-" & GetMonth(dFecha2) & "-" & Year(dFecha2)
    PeriodoEnLetras = s
End Function
Public Function DateToWord(dFecha As Date) As String
    Dim sDate, sMes As String
    
    On Error GoTo DateToWordError
    DateToWord = ""
    If Not IsDate(dFecha) Then Exit Function
    sMes = Switch(Month(dFecha) = 1, "Enero", Month(dFecha) = 2, "Febrero", Month(dFecha) = 3, "Marzo", Month(dFecha) = 4, "Abril", Month(dFecha) = 5, "Mayo", Month(dFecha) = 6, "Junio", Month(dFecha) = 7, "Julio", Month(dFecha) = 8, "Agosto", Month(dFecha) = 9, "Septiembre", Month(dFecha) = 10, "Octubre", Month(dFecha) = 11, "Noviembre", Month(dFecha) = 12, "Diciembre")
    sDate = Day(dFecha) & " de " & sMes & " del " & Year(dFecha)
    DateToWord = sDate
    Exit Function
    
DateToWordError:
    MsgBox Err.Description
    Exit Function
End Function
Public Function Nulo(vValue As Variant) As Currency
    'Verifica si un valor es nulo, y retorna cero.
    If IsNull(vValue) Then
        Nulo = 0
    Else
    If vValue = "" Then vValue = 0
        Nulo = vValue
    End If
End Function
'Public Function FindCtaEgreso(cCon As Connection, sCta As String) As String
'    Dim rs As New Recordset
'
'    FindCtaEgreso = ""
'    Set rs = cCon.Execute("select Descripcion From CatalogoEgreso where CtaEgreso='" & sCta & "' ")
'    If rs.RecordCount > 0 Then
'        FindCtaEgreso = rs!Descripcion
'    End If
'End Function
Public Sub RefreshRecordset(rs As Recordset)
    If rs Is Nothing Then Exit Sub
    rs.Requery
End Sub
Public Function FechaDesdeEnero(dFechaFinal As Date) As String
'Retorna una fecha asi: 01 de Enero al 31 de Octubre del 2006
    
    FechaDesdeEnero = ""
    FechaDesdeEnero = "01 de Enero al "
    FechaDesdeEnero = FechaDesdeEnero & DateToWord(dFechaFinal)
End Function
Public Function CalculeInteres(cImpuesto As Currency, nTasa As Single, nMeses As Integer) As Currency
'Calcule el interes sobre un impuesto vencido
'cImpuesto=Impuesto vencido
'nTasa = la tasa mensual a aplicar
'nMeses = Numero de meses vencido el impuesto
    CalculeInteres = cImpuesto * nTasa * nMeses

End Function
Public Function CalculeRecargoSobreSaldo(cSaldo As Currency, nTasa As Single, nMeses As Integer) As Currency
'Calcula el recargo sobre un saldo vencido. El saldo lo componen el Impuesto + Interes
'cSaldo=Saldo Vencido
'nTasa = la tasa mensual a aplicar
'nMeses = Numero de meses vencido el impuesto

    CalculeRecargoSobreSaldo = cSaldo * nTasa * nMeses
    
End Function
Public Function EdadPersona(sIdentidad As String, cCon As Connection) As Integer
    Dim rs As New Recordset
    
    EdadPersona = 0
    Set rs = cCon.Execute("select FechaNac from Contribuyente where Identidad='" & sIdentidad & "'")
    If rs.RecordCount = 0 Then Exit Function
    If Not IsDate(rs!FechaNac) Then
        Exit Function
    End If
    EdadPersona = DateDiff("yyyy", rs!FechaNac, Date)
End Function
Public Function FinDeMes(dFecha1 As Date) As Date
'Esta funcion calcula el fin del mes de una fecha dada
'Ejemplo 10/01/2005 retorna 31/01/2005
'-----------------------------------------------------
'Proceso:
'LLevamos la fecha que nos envian al siguiente mes
'A esa fecha le restamos los dias que tiene, resultado ultimo del mes anterior.
'-------------------------------------------------------------------------------
    Dim iDias As Integer
    
    On Error GoTo FinDeMes_Error
    FinDeMes = DateAdd("m", 1, dFecha1)
    iDias = Day(FinDeMes)
    FinDeMes = DateAdd("d", -iDias, FinDeMes)
    Exit Function
    
FinDeMes_Error:
    MsgBox Err.Description
End Function
Public Sub FindCtaContable(rs As Recordset, cnCon1 As Connection, strCtaContable As String)

    Set rs = cnCon1.Execute("select * from Catalogo where CtaContable='" & strCtaContable & "' ")
    
End Sub

Public Sub FindCtaEgreso(rs As Recordset, cnCon1 As Connection, strCtaEgreso As String)

    Set rs = cnCon1.Execute("select * from CatalogoEgreso where CtaEgreso='" & strCtaEgreso & "' ")
    
End Sub
Public Sub FindCatalogoIngreso(rs As Recordset, cnCon1 As Connection, strCtaIngreso As String)

    Set rs = cnCon1.Execute("select * from CatalogoIngreso where CtaIngreso='" & strCtaIngreso & "' ")
    
End Sub

Public Sub FindCuentaIngreso(rs As Recordset, cnCon1 As Connection, strCtaIngreso As String, strAnio As Integer)

    Set rs = cnCon1.Execute("select * from CuentaIngreso_A where CtaIngreso='" & strCtaIngreso & "' and anio = " & strAnio & "")
    
End Sub

Public Sub FindPropiedad(rs As Recordset, cnCon1 As Connection, strClaveCatastro As String)

    Set rs = cnCon1.Execute("select * from Catastro where ClaveCatastro='" & strClaveCatastro & "' ")

End Sub
Public Sub FindContribuyente(rs As Recordset, cnCon1 As Connection, strId As String)

    Set rs = cnCon1.Execute("select * from Contribuyente where Identidad='" & strId & "' ")

End Sub
Public Sub FindAldea(rs As Recordset, cnCon1 As Connection, strId As String)

    Set rs = cnCon1.Execute("select * from Aldea where CodAldea='" & strId & "' ")

End Sub

Public Sub FindBarrio(rs As Recordset, cnCon1 As Connection, strId As String)

    Set rs = cnCon1.Execute("select * from TablaBarrio where CodBarrio='" & strId & "' ")

End Sub
Public Sub FindProfesion(rs As Recordset, cnCon1 As Connection, strId As String)
    Set rs = cnCon1.Execute("select * from Profesion where CodProfesion='" & strId & "' ")
End Sub
Public Sub FindBanco(rs As Recordset, cnCon1 As Connection, strId As String)
    Set rs = cnCon1.Execute("select * from Banco where CodigoBanco='" & strId & "' ")
End Sub
Public Sub FindCtaContabilidad(rs As Recordset, cnCon1 As Connection, strId As String)
    Set rs = cnCon1.Execute("select * from Catalogo where CtaContable='" & strId & "' ")
End Sub
Public Sub FindProveedor(rs As Recordset, cnCon1 As Connection, strId As String)
    Set rs = cnCon1.Execute("select * from Proveedor where RTNProv='" & strId & "' ")
End Sub

Public Function FormeNombre(rs As Recordset) As String
    If rs.RecordCount = 0 Then
        FormeNombre = ""
    Else
        FormeNombre = Trim(rs!PNombre) & " " & Trim(rs!sNombre) & " " & Trim(rs!pApellido) & " " & Trim(rs!sApellido)
    End If
End Function
Public Function GetCtaRecuperacion(strCuenta As String, cn As Connection, dFechaActual As Date, dFechaVence As Date) As String

'Actualize Cuenta de Recuperacion
'    para las facturas de bienes inmuebles, Impuesto Personal y Servicios Publicos
'    si la factura pertenece a un año anterior, la cuenta de impuesto debe
'    cambiarse a la cuenta de recuperacion.
'    Entonces la regla a aplicar será la siguiente:
'    Si la factura es de años anteriores al actual en proceso
'        para cada item de factura
'            Busque la cuenta en CuentaIngreso
'            Si tiene cuenta de recuperacion
'                busque la cuenta de recuperacion
'                si la encuentra
'                    cambiela por la cuenta de recuperacion
'                sino
'                    mande un mensaje que cuenta de recuperacion no encontrada

    Dim strSql As String
    Dim rs As Recordset
    Dim vAnio2 As Integer
        Dim rsSysPar As New ADODB.Recordset
      Set rsSysPar = DeRia.CoRia.Execute("SELECT * FROM SystemParam")
      If rsSysPar!CtaIngresoCJ = strCuenta Then
      vAnio2 = Format(DiaEnProcesoCT, "YYYY")
      Else
    vAnio2 = Format(dFechaVence, "yyyy")
    End If
    GetCtaRecuperacion = strCuenta
    If Year(dFechaVence) >= Year(dFechaActual) Then
        Exit Function
    End If
    
    strSql = "select CtaRecuperacion from CuentaIngreso_A  " & _
    "where CtaIngreso='" & strCuenta & "' and anio = " & vAnio2 & " "
    Set rs = cn.Execute(strSql)
    If rs.RecordCount = 0 Then
        MsgBox "Cuenta " & strCuenta & " no encontrada. Verifique sus cuentas de ingreso...!"
        Exit Function
    End If
    If IsNull(rs!CtaRecuperacion) Or Trim(rs!CtaRecuperacion) = "" Then
        'no tiene definida cuenta de recuperacion
        Exit Function
    Else
        strSql = "select CtaIngreso from CuentaIngreso_A " & _
        "where CtaIngreso='" & rs!CtaRecuperacion & "' and anio = " & vAnio2 & " "
        Set rs = cn.Execute(strSql)
        If rs.RecordCount = 0 Then
            MsgBox "La cuenta de recuperación para la cuenta " & strCuenta & " no existe en el catalogo. Verifique sus cuentas de ingreso...!"
            Exit Function
        End If
        GetCtaRecuperacion = rs!CtaIngreso
    End If
End Function
Public Function IsNullString(str1 As Variant) As String
    If IsNull(str1) Then
        IsNullString = ""
    Else
        IsNullString = str1
    End If
End Function

Public Sub AplicaTerceraEdad(strId As String, blnSiNO As Boolean, strCtaIngreso As String, strNombre As String, sngTasa As Single, cnConn As Connection, dFechaVenceFactura As Date)
'Verifica si una persona tiene la tercera edad, para efectuar el descuento en los
'impuestos y servicios que paga.
'Retorna una variable que indica si aplica o no, la cuenta de ingreso por descuento y el nombre
'Los descuentos de la tercera edad, son a partir del 2007

    Dim intEdad As Integer
    Dim rs As New Recordset
    Dim RsTe As New ADODB.Recordset
    Dim Str As String
    
    blnSiNO = False
    If Year(dFechaVenceFactura) < 2007 Then
        Exit Sub
    End If
    intEdad = EdadPersona(strId, cnConn)
    Set rs = cnConn.Execute("select AnosTerceraEdad,TasaTerceraEdad from Parametro")
    If rs.RecordCount = 0 Then
        MsgBox "No estan definidos los parametros de Tercera Edad."
        Exit Sub
    End If
    If IsNull(rs!AnosTerceraEdad) Then
        MsgBox "No estan definidos los parametros de Tercera Edad."
        Exit Sub
    End If
    If intEdad < rs!AnosTerceraEdad Then
        Exit Sub
    End If
    If rs!TasaTerceraEdad <= 0 Then
        MsgBox "Debe definir la tasa de descuento por tercera edad."
        Exit Sub
    End If
    sngTasa = rs!TasaTerceraEdad
    Set rs = cnConn.Execute("select CtaIngresoCJ from SystemParam")
    If rs.RecordCount = 0 Then
        MsgBox "No esta definida la cuenta de ingreso por descuento de tercera edad. Verifique los parametros"
        Exit Sub
    End If
    If IsNull(rs!CtaIngresoCJ) Then
        MsgBox "No esta definida la cuenta de ingreso por descuento de tercera edad. Verifique los parametros"
        Exit Sub
    End If
    
        Set RsTe = DeRia.CoRia.Execute("Select * from Cuentaingreso_A where CtaIngreso = '" & rs!CtaIngresoCJ & "' and Anio = " & Year(dFechaVenceFactura) & " ")
       If RsTe.EOF Then
       Set RsTe = DeRia.CoRia.Execute("Select * from Cuentaingreso_A where CtaIngreso = '" & rs!CtaIngresoCJ & "' and Anio = " & Year(DiaEnProcesoCT) & " ")
       Str = "Insert Into CuentaIngreso_A  (CtaIngreso, Anio, Rango, ValorPermOp, ValorMensual, ValorRenovacion, ValorMultaSinPermiso, NombreCtaIngreso, CtaPermOP, CtaRecuperacion , Tipo)"
       Str = Str & "values('" & RsTe!CtaIngreso & "', " & Year(dFechaVenceFactura) & ", " & RsTe!Rango & ", " & RsTe!ValorPermOp & ", " & RsTe!ValorMensual & ", " & RsTe!ValorRenovacion & ", " & RsTe!ValorMultaSinPermiso & ", '" & RsTe!NombreCtaIngreso & "', '" & RsTe!CtaPermOP & "', '" & RsTe!CtaRecuperacion & "' , " & RsTe!Tipo & ")"
       DeRia.CoRia.Execute (Str)
       
       End If
    
    
    FindCuentaIngreso rs, cnConn, rs!CtaIngresoCJ, Format(DiaEnProcesoCT, "YYYY")

    If rs.RecordCount = 0 Then
        MsgBox "La cuenta de descuento de tercera edad definida en parametros, no existe en las cuentas de ingreso."
        Exit Sub
    End If
    If IsNull(rs!CtaIngreso) Then
        MsgBox "La cuenta de descuento de tercera edad definida en parametros, no existe en las cuentas de ingreso."
        Exit Sub
    End If
    strCtaIngreso = rs!CtaIngreso
    strNombre = rs!NombreCtaIngreso
    blnSiNO = True
End Sub
Public Function InicioDeMes(dFecha1 As Date) As Date
'Esta funcion calcula el inicio del mes de una fecha dada
'Ejemplo 10/07/2005 retorna 01/07/2005
'Proceso:
'Calculamos el numero de dias que hay en la fecha que nos envian
'Restamos 1 a ese valor y restamos esos dias de la fecha que nos envian
'----------------------------------------------------------------------
    
    Dim iDias As Integer
    
    On Error GoTo InicioDeMes_Error
    InicioDeMes = dFecha1
    iDias = Day(InicioDeMes) - 1
    InicioDeMes = DateAdd("d", -iDias, InicioDeMes)
    Exit Function
    
InicioDeMes_Error:
    MsgBox Err.Description
End Function
Public Function ValidateNumeric(strText As String) As Boolean
    ValidateNumeric = CBool(strText = "" _
        Or strText = "-" _
        Or strText = "-." _
        Or strText = "." _
        Or IsNumeric(strText))
End Function

Private Function AbrirDB(strDataPath As String, oConn As Connection) As Boolean
    'iniciar sesión en Jet

    Dim strPass As String
    Dim intTries As Integer
    
    On Error GoTo CheckError
    
    strPass = ""
    GoTo OpenDB
    
GetPWD:
    sConnect = Switch(intTries = 0, "lasuiza3025", intTries = 1, "qwerty3025", intTries = 2, "mgia730@lc@3b3s")
    intTries = intTries + 1
    If intTries > 3 Then
        MsgBox "Clave de base de datos desconocida. No se puede continuar."
        Exit Function
    End If
    
OpenDB:
    If Len(strDataPath) > 0 Then
        sProvider = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDataPath & ";Jet OLEDB:Database Password=" & strPass
        oConn.Open sProvider
    End If
    Exit Function

    
CheckError:
    If Err.Number = -2147217843 Then
        'base protegida con contrasena
        Resume GetPWD
    End If
    MsgBox Err.Description
    Exit Function
End Function

Public Function FindTipoImpuesto(lngNumFactura As Long, strReferencia As String, cnConn As ADODB.Connection) As String
    Dim rs As New Recordset
    Dim strSql As String
    
    FindTipoImpuesto = ""
    strReferencia = ""
    strSql = "select * from AvPgEnc where NumAvPg=" & lngNumFactura & ""
    Set rs = cnConn.Execute(strSql)
    Select Case rs!AvPgTipoImpuesto
        Case 0
            FindTipoImpuesto = "OS"
        Case 1
            FindTipoImpuesto = "BI"
            strReferencia = rs!ClaveCatastro
        Case 3
        FindTipoImpuesto = "IC" 'Tributaria
            strSql = "select * from AvPgDetalle where NumAvPg=" & lngNumFactura & ""
            Set rs = cnConn.Execute(strSql)
            strReferencia = rs!RefAvPgDet
        Case 2
            FindTipoImpuesto = "IC"
            strSql = "select * from AvPgDetalle where NumAvPg=" & lngNumFactura & ""
            Set rs = cnConn.Execute(strSql)
            strReferencia = IIf(IsNull(rs!RefAvPgDet), "", rs!RefAvPgDet)
        Case 4
            FindTipoImpuesto = "IP"
            strSql = "select * from AvPgDetalle where NumAvPg=" & lngNumFactura & ""
            Set rs = cnConn.Execute(strSql)
            strReferencia = rs!RefAvPgDet
        Case 5
            FindTipoImpuesto = "SP"
            strSql = "select CodDeclara from AvPgEnc where NumAvPg=" & lngNumFactura & ""
            Set rs = cnConn.Execute(strSql)
            strSql = "select ClaveCatastro from AbonadoSPEnc where ASPE_Seq=" & rs!CodDeclara & ""
            Set rs = cnConn.Execute(strSql)
            strReferencia = rs!ClaveCatastro
        Case 7
            FindTipoImpuesto = "PP"
        Case 8
            FindTipoImpuesto = "AB"
            
    End Select
End Function

Public Sub ApliqueCtaRecuperacion(lngNumFactura As Long, dFechaActual As Date, dFechaVence As Date, cn As Connection)
'Actualize Cuenta de Recuperacion
'   Nos envian un numero de factura con su fecha de vencimiento y la fecha en proceso
'   Determinamos si esa factura es de un año anterior
'   si es asi
'       seleccionamos todos los items de esa factura
'       para cada item
'           determinamos si la cuenta de ingreso del item, tiene cuenta de recuperacion
'           si es asi entonces reemplazamos la cuenta del item, por la cuenta de recuperacion

    Dim strSql As String
    Dim strNuevaCta As String
    Dim rs As Recordset
    Dim rsDetalle As New Recordset
    
    If Year(dFechaVence) >= Year(dFechaActual) Then
        Exit Sub
    End If
    
    Set rsDetalle = cn.Execute("select * from AvPgDetalle where NumAvPg=" & lngNumFactura & "")
    Do While rsDetalle.EOF = False
        strNuevaCta = GetCtaRecuperacion(rsDetalle!CtaIngreso, cn, dFechaActual, dFechaVence)
        If strNuevaCta <> Trim(rsDetalle!CtaIngreso) Then
            cn.Execute ("update AvPgDetalle set CtaIngreso='" & strNuevaCta & "' where SeqAvPgDet=" & rsDetalle!SeqAvPgDet & "")
        End If
        rsDetalle.MoveNext
        DoEvents
    Loop
End Sub

Public Function EnLetras(Numero As String) As String

    Dim b, paso As Integer

    Dim expresion, entero, deci, flag As String

        

    flag = "N"

    For paso = 1 To Len(Numero)

        If Mid(Numero, paso, 1) = "." Then

            flag = "S"

        Else

            If flag = "N" Then

                entero = entero + Mid(Numero, paso, 1) 'Extae la parte entera del numero

            Else

                deci = deci + Mid(Numero, paso, 1) 'Extrae la parte decimal del numero

            End If

        End If

    Next paso

    

    If Len(deci) = 1 Then

        deci = deci & "0"

    End If

    

    flag = "N"

    If Val(Numero) >= -999999999 And Val(Numero) <= 999999999 Then 'si el numero esta dentro de 0 a 999.999.999

        For paso = Len(entero) To 1 Step -1

            b = Len(entero) - (paso - 1)

            Select Case paso

            Case 3, 6, 9

                Select Case Mid(entero, b, 1)

                    Case "1"

                        If Mid(entero, b + 1, 1) = "0" And Mid(entero, b + 2, 1) = "0" Then

                            expresion = expresion & "cien "

                        Else

                            expresion = expresion & "ciento "

                        End If

                    Case "2"

                        expresion = expresion & "doscientos "

                    Case "3"

                        expresion = expresion & "trescientos "

                    Case "4"

                        expresion = expresion & "cuatrocientos "

                    Case "5"

                        expresion = expresion & "quinientos "

                    Case "6"

                        expresion = expresion & "seiscientos "

                    Case "7"

                        expresion = expresion & "setecientos "

                    Case "8"

                        expresion = expresion & "ochocientos "

                    Case "9"

                        expresion = expresion & "novecientos "

                End Select

                

            Case 2, 5, 8

                Select Case Mid(entero, b, 1)

                    Case "1"

                        If Mid(entero, b + 1, 1) = "0" Then

                            flag = "S"

                            expresion = expresion & "diez "

                        End If

                        If Mid(entero, b + 1, 1) = "1" Then

                            flag = "S"

                            expresion = expresion & "once "

                        End If

                        If Mid(entero, b + 1, 1) = "2" Then

                            flag = "S"

                            expresion = expresion & "doce "

                        End If

                        If Mid(entero, b + 1, 1) = "3" Then

                            flag = "S"

                            expresion = expresion & "trece "

                        End If

                        If Mid(entero, b + 1, 1) = "4" Then

                            flag = "S"

                            expresion = expresion & "catorce "

                        End If

                        If Mid(entero, b + 1, 1) = "5" Then

                            flag = "S"

                            expresion = expresion & "quince "

                        End If

                        If Mid(entero, b + 1, 1) > "5" Then

                            flag = "N"

                            expresion = expresion & "dieci"

                        End If

                

                    Case "2"

                        If Mid(entero, b + 1, 1) = "0" Then

                            expresion = expresion & "veinte "

                            flag = "S"

                        Else

                            expresion = expresion & "veinti"

                            flag = "N"

                        End If

                    

                    Case "3"

                        If Mid(entero, b + 1, 1) = "0" Then

                            expresion = expresion & "treinta "

                            flag = "S"

                        Else

                            expresion = expresion & "treinta y "

                            flag = "N"

                        End If

                

                    Case "4"

                        If Mid(entero, b + 1, 1) = "0" Then

                            expresion = expresion & "cuarenta "

                            flag = "S"

                        Else

                            expresion = expresion & "cuarenta y "

                            flag = "N"

                        End If

                

                    Case "5"

                        If Mid(entero, b + 1, 1) = "0" Then

                            expresion = expresion & "cincuenta "

                            flag = "S"

                        Else

                            expresion = expresion & "cincuenta y "

                            flag = "N"

                        End If

                

                    Case "6"

                        If Mid(entero, b + 1, 1) = "0" Then

                            expresion = expresion & "sesenta "

                            flag = "S"

                        Else

                            expresion = expresion & "sesenta y "

                            flag = "N"

                        End If

                

                    Case "7"

                        If Mid(entero, b + 1, 1) = "0" Then

                            expresion = expresion & "setenta "

                            flag = "S"

                        Else

                            expresion = expresion & "setenta y "

                            flag = "N"

                        End If

                

                    Case "8"

                        If Mid(entero, b + 1, 1) = "0" Then

                            expresion = expresion & "ochenta "

                            flag = "S"

                        Else

                            expresion = expresion & "ochenta y "

                            flag = "N"

                        End If

                

                    Case "9"

                        If Mid(entero, b + 1, 1) = "0" Then

                            expresion = expresion & "noventa "

                            flag = "S"

                        Else

                            expresion = expresion & "noventa y "

                            flag = "N"

                        End If

                End Select

                

            Case 1, 4, 7

                Select Case Mid(entero, b, 1)

                    Case "1"

                        If flag = "N" Then

                            If paso = 1 Then

                                expresion = expresion & "uno "

                            Else

                                expresion = expresion & "un "

                            End If

                        End If

                    Case "2"

                        If flag = "N" Then

                            expresion = expresion & "dos "

                        End If

                    Case "3"

                        If flag = "N" Then

                            expresion = expresion & "tres "

                        End If

                    Case "4"

                        If flag = "N" Then

                            expresion = expresion & "cuatro "

                        End If

                    Case "5"

                        If flag = "N" Then

                            expresion = expresion & "cinco "

                        End If

                    Case "6"

                        If flag = "N" Then

                            expresion = expresion & "seis "

                        End If

                    Case "7"

                        If flag = "N" Then

                            expresion = expresion & "siete "

                        End If

                    Case "8"

                        If flag = "N" Then

                            expresion = expresion & "ocho "

                        End If

                    Case "9"

                        If flag = "N" Then

                            expresion = expresion & "nueve "

                        End If

                End Select

            End Select

            If paso = 4 Then

                If Mid(entero, 6, 1) <> "0" Or Mid(entero, 5, 1) <> "0" Or Mid(entero, 4, 1) <> "0" Or _
                  (Mid(entero, 6, 1) = "0" And Mid(entero, 5, 1) = "0" And Mid(entero, 4, 1) = "0" And _
                   Len(entero) <= 6) Then

                    expresion = expresion & "mil "

                End If

            End If

            If paso = 7 Then

                If Len(entero) = 7 And Mid(entero, 1, 1) = "1" Then

                    expresion = expresion & "millón "

                Else

                    expresion = expresion & "millones "

                End If

            End If

        Next paso

        

        If deci <> "" Then

            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo

                EnLetras = "menos " & expresion & "con " & deci ' & "/100"

            Else

                EnLetras = expresion & "con " & deci ' & "/100"

            End If

        Else

            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo

                EnLetras = "menos " & expresion

            Else

                EnLetras = expresion

            End If

        End If

    Else 'si el numero a convertir esta fuera del rango superior e inferior

        EnLetras = ""

    End If

End Function
Public Function NaturalJuridico(bitTipo As Byte) As String
    If bitTipo = 0 Then
        NaturalJuridico = "Natural"
    Else
        NaturalJuridico = "Juridico"
    End If
End Function

Public Function IsFormLoaded(frm As Form) As Boolean
    Dim f As Form
    For Each f In Forms
        If f.Name = frm.Name Then
            IsFormLoaded = True
            Exit Function
        End If
    Next
    IsFormLoaded = False
End Function

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

