VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCrearPermiso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub CrearPermisoOP(numRecibo As Long, formulario As Object)
    If ExistePermisoOP(numRecibo) Then
        MsgBox "El número de recibo ya está registrado.", vbInformation
        CargarDatosPermisoOP numRecibo, formulario
        formulario.cmdImprimir.Enabled = True
        formulario.cmdGuardar.Enabled = False
    Else
        ObtenerDatosRecibo numRecibo, formulario
    End If
End Sub

Private Function ExistePermisoOP(numRecibo As Long) As Boolean
    Dim codSQL As String
    codSQL = "SELECT COUNT(*) FROM Tra_PermOP WHERE NumRecibo = " & numRecibo
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open codSQL
    ExistePermisoOP = (DeRia.rsAbonadoSP.Fields(0).Value > 0)
End Function

Private Sub CargarDatosPermisoOP(numRecibo As Long, formulario As Object)
    Dim codSQL As String
    codSQL = "SELECT NumRecibo, Identidad, NoPermiso, Periodo, Negocio, Propietario, Ubicacion, Actividad, fechaEstablecido, TipoCuenta " & _
             "FROM Tra_PermOP WHERE NumRecibo = " & numRecibo
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open codSQL
    formulario.txtnumrecibo = DeRia.rsAbonadoSP!numRecibo
    formulario.txtRTM = DeRia.rsAbonadoSP!identidad
    formulario.txtCodigo = DeRia.rsAbonadoSP!NoPermiso
    formulario.txtPeriodo = DeRia.rsAbonadoSP!periodo
    formulario.txtNombreEstablecimiento = DeRia.rsAbonadoSP!Negocio
    formulario.txtPropietario = DeRia.rsAbonadoSP!Propietario
    formulario.txtUbicacion = DeRia.rsAbonadoSP!Ubicacion
    formulario.txtActividad = DeRia.rsAbonadoSP!Actividad
    formulario.txtFechaNac = DeRia.rsAbonadoSP!fechaEstablecido
    formulario.txtTipo = DeRia.rsAbonadoSP!TipoCuenta
    formulario.cmdImprimir.Enabled = True
    formulario.cmdGuardar.Enabled = False
End Sub

Private Sub ObtenerDatosRecibo(numRecibo As Long, formulario As Object)
    Dim codSQL As String
    codSQL = "SELECT ReciboDet.NumRecibo, AvPgEnc.Identidad, Contribuyente.Direccion, Contribuyente.IdRepresentante, " & _
             "Contribuyente_1.Pnombre, Contribuyente_1.SNombre, Contribuyente_1.PApellido, Contribuyente_1.SApellido, " & _
             "CuentaIngreso_A.NombreCtaIngreso, Contribuyente.CodProfesion, Contribuyente.FechaNac, Contribuyente.Pnombre AS Expr1 " & _
             "FROM AvPgEnc INNER JOIN ReciboDet ON AvPgEnc.NumAvPg = ReciboDet.NumFactura " & _
             "INNER JOIN Contribuyente ON AvPgEnc.Identidad = Contribuyente.Identidad " & _
             "INNER JOIN Contribuyente AS Contribuyente_1 ON Contribuyente.IdRepresentante = Contribuyente_1.Identidad " & _
             "INNER JOIN CuentaIngreso_A ON Contribuyente.CodProfesion = CuentaIngreso_A.CtaIngreso " & _
             "WHERE (ReciboDet.NumRecibo = " & numRecibo & ") AND ((AvPgEnc.AvPgTipoImpuesto = 2) or (AvPgEnc.AvPgTipoImpuesto = 3)) " & _
             "GROUP BY ReciboDet.NumRecibo, AvPgEnc.Identidad, Contribuyente.Direccion, Contribuyente.IdRepresentante, " & _
             "Contribuyente_1.Pnombre, Contribuyente_1.SNombre, Contribuyente_1.PApellido, Contribuyente_1.SApellido, " & _
             "CuentaIngreso_A.NombreCtaIngreso, Contribuyente.CodProfesion, Contribuyente.FechaNac, Contribuyente.Pnombre"

    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open codSQL
    formulario.txtRTM = DeRia.rsAbonadoSP!identidad
    formulario.txtnumrecibo = numRecibo
    formulario.txtPeriodo = Format(Now, "yyyy")
    formulario.txtNombreEstablecimiento = DeRia.rsAbonadoSP!Expr1
    formulario.txtPropietario = Trim(DeRia.rsAbonadoSP!Pnombre) & " " & Trim(DeRia.rsAbonadoSP!sNombre) & " " & _
                     Trim(DeRia.rsAbonadoSP!PApellido) & " " & Trim(DeRia.rsAbonadoSP!sApellido)
    formulario.txtUbicacion = DeRia.rsAbonadoSP!Direccion
    formulario.txtActividad = DeRia.rsAbonadoSP!NombreCtaIngreso
    formulario.txtFechaNac = DeRia.rsAbonadoSP!FechaNac
    formulario.txtTipo = TipoCuenta(DeRia.rsAbonadoSP!CodProfesion)
    Dim nuevoCodigo As Long
    GenerarNuevoPermisoOP nuevoCodigo
    formulario.txtCodigo = nuevoCodigo
    formulario.cmdImprimir.Enabled = False
    formulario.cmdGuardar.Enabled = True
End Sub

Private Sub GenerarNuevoPermisoOP(ByRef nuevoPermiso As Long)
    Dim codSQL As String
    codSQL = "SELECT UltMunPermioOP FROM ParametroCont"
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open codSQL
    nuevoPermiso = DeRia.rsAbonadoSP!UltMunPermioOP + 1
End Sub

Public Function TipoCuenta(cta As String) As String
    If Left(cta, 6) = "117101" Then
        TipoCuenta = "Establecimiento Industrial"
    ElseIf Left(cta, 6) = "117102" Then
        TipoCuenta = "Establecimiento Comercial"
    ElseIf Left(cta, 6) = "117103" Then
        TipoCuenta = "Establecimiento de Servicio"
    Else
        TipoCuenta = "Tipo de cuenta no reconocido"
    End If
End Function

Public Function Guardar(formulario As Object)
    Dim obser As String
    obser = ObtenerTipoPermiso(formulario)
    If obser = "" Then
        MsgBox ("Por favor, seleccione una opción (Renovación o Apertura).")
        formulario.cmdGuardar.Enabled = True
        formulario.cmdImprimir.Enabled = False
    Else
        Dim sqlInsert As String
        sqlInsert = "INSERT INTO Tra_PermOP (NoPermiso, Periodo, Identidad, Negocio, Propietario, Ubicacion, Actividad, TipoCuenta, Observacion, fechaEstablecido, Fecha, CodAldea, NumRecibo) VALUES ("
        sqlInsert = sqlInsert & "'" & formulario.txtCodigo.Text & "', "
        sqlInsert = sqlInsert & "'" & formulario.txtPeriodo.Text & "', "
        sqlInsert = sqlInsert & "'" & formulario.txtRTM.Text & "', "
        sqlInsert = sqlInsert & "'" & formulario.txtNombreEstablecimiento.Text & "', "
        sqlInsert = sqlInsert & "'" & formulario.txtPropietario.Text & "', "
        sqlInsert = sqlInsert & "'" & formulario.txtUbicacion.Text & "', "
        sqlInsert = sqlInsert & "'" & formulario.txtActividad.Text & "', "
        sqlInsert = sqlInsert & "'" & formulario.txtTipo.Text & "', "
        sqlInsert = sqlInsert & "'" & ObtenerTipoPermiso(formulario) & "', "
        sqlInsert = sqlInsert & "'" & formulario.txtFechaNac.Text & "', "
        sqlInsert = sqlInsert & "'" & Format(Date, "dd/mm/yyyy") & "', "
        sqlInsert = sqlInsert & "'" & CargarBarrio(formulario) & "', "
        sqlInsert = sqlInsert & "'" & formulario.txtnumrecibo.Text & "')"
        If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
        DeRia.rsAbonadoSP.Open (sqlInsert)
        
        sql = "UPDATE ParametroCont SET UltMunPermioOP = " + formulario.txtCodigo.Text
        If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
        DeRia.rsAbonadoSP.Open (sql)
        formulario.cmdGuardar.Enabled = False
        formulario.cmdImprimir.Enabled = True
    End If
End Function

Private Function ObtenerTipoPermiso(formulario As Object) As String
    Dim observacion As String

    If formulario.optRenovacion.Value = True Then
        observacion = "Renovacion"
    ElseIf formulario.optApertura.Value = True Then
        observacion = "Apertura"
    End If

    'If observacion = "" Then
        'MsgBox ("Por favor, seleccione una opción (Renovación o Apertura).")
    'Else
    ObtenerTipoPermiso = observacion
    'End If
End Function


Public Function Imprimir(formulario As Object, reporte As Object)
    
    'Llena los campos del informe con los datos necesarios
    'rpt.txtNumero.Text = Me.txtnumrecibo.Text
    reporte.txtCorrelativo.Text = formulario.txtCodigo.Text
    reporte.txtAnio.Text = formulario.txtPeriodo.Text
    reporte.txtPropietario.Text = formulario.txtPropietario.Text
    reporte.txtUbicacion.Text = formulario.txtUbicacion.Text
    reporte.txtClase.Text = formulario.txtTipo.Text
    reporte.txtActividad.Text = formulario.txtActividad.Text
    reporte.txtNombreEstablecimiento.Text = formulario.txtNombreEstablecimiento.Text
    reporte.txtFechaestablecido.Text = formulario.txtFechaNac.Text
    reporte.txtDia.Text = Format(Now, "dd")
    reporte.txtMes.Text = Format(Now, "mmmm")
    reporte.txtAnio_1.Text = Format(Now, "yyyy")
    'rpt.txtNumSol = (Format(Now, "yyyy")) - Format(Me.txtFechaNac.Text, "yyyy") + 1
    If formulario.optRenovacion.Value = True Then
        reporte.txtNumSol.Text = "Renovacion"
    ElseIf formulario.optApertura.Value = True Then
        reporte.txtNumSol.Text = "Apertura"
    End If
    reporte.Run (False)
    reporte.Show
End Function



Private Function CargarBarrio(formulario As Object) As String
    Dim codSQL As String
    codSQL = "SELECT codaldea FROM contribuyente where identidad = '" + formulario.txtRTM + "'"
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open codSQL
    CargarBarrio = DeRia.rsAbonadoSP!CodAldea
End Function

























