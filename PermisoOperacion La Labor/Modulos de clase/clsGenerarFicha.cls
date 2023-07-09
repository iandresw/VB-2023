VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGenerarFicha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim crearPermiso As New clsCrearPermiso
Public Function crearFicha(rtm As String)
    Dim codSQL As String
    Dim rpt As New FichaEstablecimiento
    Dim rtm1 As String
    
    codSQL = " SELECT Contribuyente_1.Pnombre, Contribuyente_1.SNombre, Contribuyente_1.PApellido, Contribuyente_1.SApellido, Contribuyente_1.Identidad, Contribuyente.Direccion, Contribuyente.Telefono, Contribuyente.RTN,"
    codSQL = codSQL + " Contribuyente.Pnombre AS NombreEstablecimiento, Contribuyente.CodProfesion, CuentaIngreso_A.NombreCtaIngreso, Contribuyente.FechaNac, Contribuyente.Identidad AS Expr1"
    codSQL = codSQL + " FROM Contribuyente INNER JOIN"
    codSQL = codSQL + " Contribuyente AS Contribuyente_1 ON Contribuyente.IdRepresentante = Contribuyente_1.Identidad INNER JOIN"
    codSQL = codSQL + " CuentaIngreso_A ON Contribuyente.CodProfesion = CuentaIngreso_A.CtaIngreso"
    codSQL = codSQL + " WHERE(CuentaIngreso_A.Anio = 2023) AND (Contribuyente.Identidad = '" + rtm + "')"
 
    If DeRia.rsAbonadoSP.State = 1 Then DeRia.rsAbonadoSP.Close
    DeRia.rsAbonadoSP.Open (codSQL)
    rpt.txtPropietario = Trim(DeRia.rsAbonadoSP!Pnombre) + " " + Trim(DeRia.rsAbonadoSP!sNombre) + " " + Trim(DeRia.rsAbonadoSP!PApellido) + " " + Trim(DeRia.rsAbonadoSP!sApellido)
    rpt.txtDNI = Trim(DeRia.rsAbonadoSP!identidad)
    rpt.txtUbicacion = DeRia.rsAbonadoSP!Direccion
    rpt.txtTelefono = DeRia.rsAbonadoSP!Telefono
    rpt.txtRTN = DeRia.rsAbonadoSP!RTN
    rpt.txtNombreEstablecimiento = DeRia.rsAbonadoSP!NombreEstablecimiento
    rpt.txtClaseNegocio = crearPermiso.TipoCuenta(DeRia.rsAbonadoSP!CodProfesion)
    rpt.txtActividadPrincipal = DeRia.rsAbonadoSP!NombreCtaIngreso
    rpt.txtFecha = DeRia.rsAbonadoSP!FechaNac
    rpt.txtRTM = DeRia.rsAbonadoSP!Expr1
    rpt.Run (False)
    rpt.Show
End Function










