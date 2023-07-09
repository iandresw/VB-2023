VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub CrearPermisoOP(numRecibo As Long)
    Dim codSQL As String
    Dim numpermiso As Long
    codSQL = "SELECT COUNT(*) FROM Tra_PermOP WHERE NumRecibo = " & numRecibo
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open (codSQL)
    If DeRia.rsAbonadoSP.Fields(0).Value > 0 Then
        MsgBox "El número de recibo ya está registrado.", vbInformation
        codSQL = "SELECT NumRecibo,Identidad, NoPermiso, Periodo , Negocio, Propietario, Ubicacion, Actividad, fechaEstablecido, TipoCuenta "
        codSQL = codSQL + "FROM Tra_PermOP WHERE NumRecibo = " & numRecibo
        If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
        DeRia.rsAbonadoSP.Open (codSQL)
        txtnumrecibo = DeRia.rsAbonadoSP!numRecibo
        txtRTM = DeRia.rsAbonadoSP!Identidad
        txtCodigo = DeRia.rsAbonadoSP!NoPermiso
        txtPeriodo = DeRia.rsAbonadoSP!Periodo
        txtNombreEstablecimiento = DeRia.rsAbonadoSP!Negocio
        txtPropietario = DeRia.rsAbonadoSP!Propietario
        txtUbicacion = DeRia.rsAbonadoSP!Ubicacion
        txtActividad = DeRia.rsAbonadoSP!Actividad
        txtFechaNac = DeRia.rsAbonadoSP!fechaEstablecido
        txtTipo = DeRia.rsAbonadoSP!TipoCuenta
    Else
        codSQL = "SELECT ReciboDet.NumRecibo, AvPgEnc.Identidad, Contribuyente.Direccion, Contribuyente.IdRepresentante, Contribuyente_1.Pnombre, Contribuyente_1.SNombre, Contribuyente_1.PApellido, Contribuyente_1.SApellido,"
        codSQL = codSQL + " CuentaIngreso_A.NombreCtaIngreso, Contribuyente.CodProfesion, Contribuyente.FechaNac, Contribuyente.Pnombre AS Expr1 FROM AvPgEnc INNER JOIN "
        codSQL = codSQL + " ReciboDet ON AvPgEnc.NumAvPg = ReciboDet.NumFactura INNER JOIN Contribuyente ON AvPgEnc.Identidad = Contribuyente.Identidad INNER JOIN "
        codSQL = codSQL + " Contribuyente AS Contribuyente_1 ON Contribuyente.IdRepresentante = Contribuyente_1.Identidad INNER JOIN CuentaIngreso_A ON Contribuyente.CodProfesion = CuentaIngreso_A.CtaIngreso "
        codSQL = codSQL + " Where (ReciboDet.numrecibo = " & numRecibo & ") And ( (AvPgEnc.AvPgTipoImpuesto = 2) or (AvPgEnc.AvPgTipoImpuesto = 3) ) GROUP BY ReciboDet.NumRecibo, AvPgEnc.Identidad, Contribuyente.Direccion, Contribuyente.IdRepresentante, Contribuyente_1.Pnombre, Contribuyente_1.SNombre, Contribuyente_1.PApellido, Contribuyente_1.SApellido, "
        codSQL = codSQL + " CuentaIngreso_A.NombreCtaIngreso , Contribuyente.CodProfesion, Contribuyente.FechaNac, Contribuyente.Pnombre "
        
        If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
        DeRia.rsAbonadoSP.Open (codSQL)
        txtRTM = DeRia.rsAbonadoSP!Identidad
        txtnumrecibo = numRecibo
        txtPeriodo = Format(Now, "yyyy")
        txtNombreEstablecimiento = DeRia.rsAbonadoSP!Expr1
        txtPropietario = Trim(DeRia.rsAbonadoSP!Pnombre) + " " + Trim(DeRia.rsAbonadoSP!sNombre) + " " + Trim(DeRia.rsAbonadoSP!PApellido) + " " + Trim(DeRia.rsAbonadoSP!sApellido)
        txtUbicacion = DeRia.rsAbonadoSP!Direccion
        txtActividad = DeRia.rsAbonadoSP!NombreCtaIngreso
        txtFechaNac = DeRia.rsAbonadoSP!FechaNac
        txtTipo = TipoCuenta(DeRia.rsAbonadoSP!CodProfesion)
        
        
        codSQL = "SELECT UltMunPermioOP FROM ParametroCont"
        If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
        DeRia.rsAbonadoSP.Open (codSQL)
        txtCodigo = DeRia.rsAbonadoSP!UltMunPermioOP + 1
     End If
End Sub

