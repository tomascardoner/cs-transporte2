------------------------------------------------------------------------------------------
-- REPORT_PLANILLAVIAJE
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'sp_Report_Viaje_Planilla' AND type = 'P')
    DROP PROCEDURE sp_Report_Viaje_Planilla
GO

CREATE PROCEDURE dbo.sp_Report_Viaje_Planilla
	@FechaHora_FILTER smalldatetime,
	@IDRuta_FILTER char(20) AS

	SELECT dbo.udf_GetEntidadApellidoYNombre(Conductor.Apellido, Conductor.Nombre) AS Conductor
			, dbo.udf_GetEntidadApellidoYNombre(Conductor2.Apellido, Conductor2.Nombre) AS Conductor2
			, Vehiculo.Nombre AS Vehiculo, Vehiculo.Asiento - Viaje.AsientoOcupado AS AsientoLibre
			, LugarOrigen.Nombre AS Origen, LugarDestino.Nombre AS Destino
		FROM (((((Viaje INNER JOIN Ruta ON Viaje.IDRuta = Ruta.IDRuta)
			INNER JOIN Lugar AS LugarOrigen ON Ruta.IDOrigen = LugarOrigen.IDLugar)
			INNER JOIN Lugar AS LugarDestino ON Ruta.IDDestino = LugarDestino.IDLugar)
			LEFT JOIN Persona AS Conductor ON Viaje.IDConductor = Conductor.IDPersona)
			LEFT JOIN Persona AS Conductor2 ON Viaje.IDConductor2 = Conductor2.IDPersona)
			LEFT JOIN Vehiculo ON Viaje.IDVehiculo = Vehiculo.IDVehiculo
		WHERE Viaje.FechaHora = @FechaHora_FILTER AND Viaje.IDRuta = @IDRuta_FILTER

GO



------------------------------------------------------------------------------------------
-- REPORT_PLANILLAVIAJE_PASAJERO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'sp_Report_Viaje_Planilla_Pasajero' AND type = 'P')
    DROP PROCEDURE sp_Report_Viaje_Planilla_Pasajero
GO

CREATE PROCEDURE dbo.sp_Report_Viaje_Planilla_Pasajero
	@FechaHora_FILTER smalldatetime,
	@IDRuta_FILTER char(20) AS

	DECLARE @IDRutaOtra char(20)

	SET @IDRutaOtra = (SELECT Texto FROM Parametro WHERE IDParametro = 'Ruta_ID_Otra')

	SELECT dbo.udf_GetEntidadApellidoYNombre(Conductor.Apellido, Conductor.Nombre) AS Conductor
			, dbo.udf_GetEntidadApellidoYNombre(Conductor2.Apellido, Conductor2.Nombre) AS Conductor2
			, Vehiculo.Nombre AS Vehiculo, Vehiculo.Asiento - Viaje.AsientoOcupado AS AsientoLibre
			, ViajeLugarOrigen.Nombre AS ViajeOrigen, ViajeLugarDestino.Nombre AS ViajeDestino
			, dbo.udf_GetEntidadApellidoYNombre(Persona.Apellido, Persona.Nombre) AS Pasajero
			, dbo.udf_GetPasajeroSubeOBaja(ViajeDetalle.IDRuta, ViajeDetalle.Sube, ViajeDetalle.IDOrigen, Ruta.IDOrigen, PasajeroLugarOrigen.Nombre, PasajeroLugarOrigen.NombreCorto) AS PasajeroOrigen
			, dbo.udf_GetPasajeroSubeOBaja(ViajeDetalle.IDRuta, ViajeDetalle.Baja, ViajeDetalle.IDDestino, Ruta.IDDestino, PasajeroLugarDestino.Nombre, PasajeroLugarDestino.NombreCorto) AS PasajeroDestino
			, ViajeDetalle.Importe, ViajeDetalle.ImporteContado + ViajeDetalle.ImporteCuentaCorriente AS ImportePagado, ViajeDetalle.ImprimirSaldo
			, (SELECT Sum(Importe) AS SaldoActual
					FROM CuentaCorriente
					WHERE CuentaCorriente.IDPersona = (CASE ISNULL(ViajeDetalle.IDPersonaCuentaCorriente, 0) WHEN 0 THEN ViajeDetalle.IDPersona ELSE ViajeDetalle.IDPersonaCuentaCorriente END)
					) AS SaldoActual
			, ISNULL(ListaPrecio.Leyenda + ' - ', '') + ISNULL(ViajeDetalle.Notas, '') AS Notas
			, dbo.udf_GetDocumentoTipoYNumero(DocumentoTipo.Nombre, Persona.DocumentoNumero) AS Documento, ViajeDetalle.Realizado
		FROM ((((((((((((Viaje INNER JOIN Ruta ON Viaje.IDRuta = Ruta.IDRuta)
			INNER JOIN Lugar AS ViajeLugarOrigen ON Ruta.IDOrigen = ViajeLugarOrigen.IDLugar)
			INNER JOIN Lugar AS ViajeLugarDestino ON Ruta.IDDestino = ViajeLugarDestino.IDLugar)
			LEFT JOIN Persona AS Conductor ON Viaje.IDConductor = Conductor.IDPersona)
			LEFT JOIN Persona AS Conductor2 ON Viaje.IDConductor2 = Conductor2.IDPersona)
			LEFT JOIN Vehiculo ON Viaje.IDVehiculo = Vehiculo.IDVehiculo)
			LEFT JOIN ViajeDetalle ON Viaje.FechaHora = ViajeDetalle.FechaHora AND Viaje.IDRuta = ViajeDetalle.IDRuta)
			LEFT JOIN Persona ON ViajeDetalle.IDPersona = Persona.IDPersona)
			LEFT JOIN DocumentoTipo ON Persona.IDDocumentoTipo = DocumentoTipo.IDDocumentoTipo)
			LEFT JOIN Lugar AS PasajeroLugarOrigen ON ViajeDetalle.IDOrigen = PasajeroLugarOrigen.IDLugar)
			INNER JOIN RutaDetalle AS PasajeroRutaDetalleOrigen ON Viaje.IDRuta = PasajeroRutaDetalleOrigen.IDRuta AND PasajeroLugarOrigen.IDLugar = PasajeroRutaDetalleOrigen.IDLugar)
			LEFT JOIN Lugar AS PasajeroLugarDestino ON ViajeDetalle.IDDestino = PasajeroLugarDestino.IDLugar)
			LEFT JOIN ListaPrecio ON ViajeDetalle.IDListaPrecio = ListaPrecio.IDListaPrecio
		WHERE Viaje.FechaHora = @FechaHora_FILTER AND Viaje.IDRuta = @IDRuta_FILTER
			AND (ViajeDetalle.OcupanteTipo IS NULL OR ViajeDetalle.OcupanteTipo = 'PA')
			AND (ViajeDetalle.Estado IS NULL OR ViajeDetalle.Estado = '1CO')
		ORDER BY PasajeroRutaDetalleOrigen.Indice, Persona.Apellido, Persona.Nombre

GO



------------------------------------------------------------------------------------------
-- REPORT_PLANILLAVIAJE_COMISION
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'sp_Report_Viaje_Planilla_Comision' AND type = 'P')
    DROP PROCEDURE sp_Report_Viaje_Planilla_Comision
GO

CREATE PROCEDURE dbo.sp_Report_Viaje_Planilla_Comision
	@FechaHora_FILTER smalldatetime,
	@IDRuta_FILTER char(20) AS

	DECLARE @IDRutaOtra char(20)

	SET @IDRutaOtra = (SELECT Texto FROM Parametro WHERE IDParametro = 'Ruta_ID_Otra')

	SELECT dbo.udf_GetEntidadApellidoYNombre(PersonaEnvia.Apellido, PersonaEnvia.Nombre) AS Envia
			, dbo.udf_GetEntidadApellidoYNombre(PersonaRecibe.Apellido, PersonaRecibe.Nombre) + ISNULL(ViajeDetalle.Recibe, '') AS Recibe
			, ViajeDetalle.IDOrigen
			, dbo.udf_GetPasajeroSubeOBaja(ViajeDetalle.IDRuta, ViajeDetalle.Sube, ViajeDetalle.IDOrigen, Ruta.IDOrigen, LugarOrigen.Nombre, LugarOrigen.NombreCorto) AS Origen
			, ViajeDetalle.IDDestino
			, dbo.udf_GetPasajeroSubeOBaja(ViajeDetalle.IDRuta, ViajeDetalle.Baja, ViajeDetalle.IDDestino, Ruta.IDDestino, LugarDestino.Nombre, LugarDestino.NombreCorto) AS Destino
			, ViajeDetalle.Importe, ViajeDetalle.ImporteContado + ViajeDetalle.ImporteCuentaCorriente AS ImportePagado, ViajeDetalle.ImprimirSaldo
			, (SELECT SUM(Importe) AS SaldoActual
					FROM CuentaCorriente
					WHERE CuentaCorriente.IDPersona = (CASE isnull(ViajeDetalle.IDPersonaCuentaCorriente, 0) WHEN 0 THEN ViajeDetalle.IDPersona ELSE ViajeDetalle.IDPersonaCuentaCorriente END)
					) AS SaldoActual
			, ISNULL(ListaPrecio.Leyenda + ' - ', '') + ISNULL(ViajeDetalle.Notas, '') AS Notas
		FROM (((((ViajeDetalle INNER JOIN Persona AS PersonaEnvia ON ViajeDetalle.IDPersona = PersonaEnvia.IDPersona)
			INNER JOIN Ruta ON ViajeDetalle.IDRuta = Ruta.IDRuta)
			INNER JOIN Lugar AS LugarOrigen ON ViajeDetalle.IDOrigen = LugarOrigen.IDLugar)
			INNER JOIN Lugar AS LugarDestino ON ViajeDetalle.IDDestino = LugarDestino.IDLugar)
			LEFT JOIN Persona AS PersonaRecibe ON ViajeDetalle.IDPersonaRecibe = PersonaRecibe.IDPersona)
			LEFT JOIN ListaPrecio ON ViajeDetalle.IDListaPrecio = ListaPrecio.IDListaPrecio
		WHERE ViajeDetalle.FechaHora = @FechaHora_FILTER
			AND ViajeDetalle.IDRuta = @IDRuta_FILTER
			AND ViajeDetalle.OcupanteTipo = 'CO'
			AND ViajeDetalle.Estado = '1CO'
		ORDER BY ViajeDetalle.Orden

GO



------------------------------------------------------------------------------------------
-- REPORT_VIAJE_PLANILLA_DOMICILIO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Report_Viaje_Planilla_Domicilio'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_Viaje_Planilla_Domicilio
GO

CREATE PROCEDURE dbo.sp_Report_Viaje_Planilla_Domicilio
	@FechaHora_FILTER smalldatetime,
	@IDRuta_FILTER char(20) AS

	SELECT Viaje.FechaHora, LugarGrupoOrigen.Nombre AS LugarGrupoOrigen, LugarGrupoDestino.Nombre AS LugarGrupoDestino
			, Vehiculo.Nombre AS Vehiculo, Vehiculo.Dominio
			, dbo.udf_GetEntidadApellidoYNombre(Conductor.Apellido, Conductor.Nombre) AS Conductor
			, dbo.udf_GetEntidadApellidoYNombre(Conductor2.Apellido, Conductor2.Nombre) AS Conductor2
			, ViajeDetalle.Orden
			, dbo.udf_GetEntidadApellidoYNombre(Pasajero.Apellido, Pasajero.Nombre) AS Pasajero
			, DocumentoTipo.Nombre AS DocumentoTipo, Pasajero.DocumentoNumero
			, dbo.udf_Domicilio_GetShort(Pasajero.DomicilioCalle1, Pasajero.DomicilioNumero, Pasajero.DomicilioPiso, Pasajero.DomicilioDepartamento, Pasajero.DomicilioCalle2, Pasajero.DomicilioCalle3) AS Domicilio
		FROM (((((((((((
			Viaje INNER JOIN ViajeDetalle ON Viaje.FechaHora = ViajeDetalle.FechaHora AND Viaje.IDRuta = ViajeDetalle.IDRuta)
			INNER JOIN Ruta ON Viaje.IDRuta = Ruta.IDRuta)
			INNER JOIN RutaDetalle AS RutaDetalleOrigen ON Ruta.IDRuta = RutaDetalleOrigen.IDRuta AND Ruta.IDOrigen = RutaDetalleOrigen.IDLugar) INNER JOIN LugarGrupo AS LugarGrupoOrigen ON RutaDetalleOrigen.IDLugarGrupo = LugarGrupoOrigen.IDLugarGrupo)
			INNER JOIN RutaDetalle AS RutaDetalleDestino ON Ruta.IDRuta = RutaDetalleDestino.IDRuta AND Ruta.IDDestino = RutaDetalleDestino.IDLugar) INNER JOIN LugarGrupo AS LugarGrupoDestino ON RutaDetalleDestino.IDLugarGrupo = LugarGrupoDestino.IDLugarGrupo)
			INNER JOIN RutaDetalle ON ViajeDetalle.IDRuta = RutaDetalle.IDRuta AND ViajeDetalle.IDOrigen = RutaDetalle.IDLugar)
			INNER JOIN Persona AS Pasajero ON ViajeDetalle.IDPersona = Pasajero.IDPersona)
			LEFT JOIN DocumentoTipo ON Pasajero.IDDocumentoTipo = DocumentoTipo.IDDocumentoTipo)
			LEFT JOIN Persona AS Conductor ON Viaje.IDConductor = Conductor.IDPersona)
			LEFT JOIN Persona AS Conductor2 ON Viaje.IDConductor2 = Conductor2.IDPersona)
			LEFT JOIN Vehiculo ON Viaje.IDVehiculo = Vehiculo.IDVehiculo
		WHERE Viaje.FechaHora = @FechaHora_FILTER AND Viaje.IDRuta = @IDRuta_FILTER
			AND ViajeDetalle.OcupanteTipo = 'PA'
			AND ViajeDetalle.Estado = '1CO'
		ORDER BY RutaDetalle.Indice, Pasajero.Apellido, Pasajero.Nombre

GO