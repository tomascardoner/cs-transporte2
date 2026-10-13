------------------------------------------------------------------------------------------
-- REPORTE_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Reporte_Data'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Reporte_Data
GO

CREATE PROCEDURE dbo.sp_Reporte_Data
	@IDReporte_FILTER char(50) AS

	SELECT IDReporte, Tipo, Nombre, Titulo, MostrarEnVisor, Personal
		FROM Reporte
		WHERE IDReporte = @IDReporte_FILTER

GO



------------------------------------------------------------------------------------------
-- REPORTEPARAMETRO_LIST
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ReporteParametro_List' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ReporteParametro_List
GO

CREATE PROCEDURE dbo.sp_ReporteParametro_List
	@IDReporte_FILTER char(50) AS

	SELECT IDParametro, Nombre, Tipo, Requerido, RequeridoLeyenda
		FROM ReporteParametro
		WHERE IDReporte = @IDReporte_FILTER
		ORDER BY Orden

GO



------------------------------------------------------------------------------------------
-- REPORT_VIAJE_LIST
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Report_Viaje_Listado' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_Viaje_Listado
GO

CREATE PROCEDURE dbo.sp_Report_Viaje_Listado
	@DiaSemana_FILTER int,
	@FechaHoraDesde_FILTER smalldatetime,
	@FechaHoraHasta_FILTER smalldatetime,
	@FechaDesde_FILTER smalldatetime,
	@FechaHasta_FILTER smalldatetime,
	@HoraDesde_FILTER smalldatetime,
	@HoraHasta_FILTER smalldatetime,
	@IDRuta_FILTER char(20),
	@IDVehiculo_FILTER int,
	@IDConductor_FILTER int,
	@Estado_FILTER char(2),
	@Charter_FILTER bit,
	@Personal bit AS

	SELECT (CASE datepart(weekday, Viaje.FechaHora) WHEN 1 THEN 'Domingo' WHEN 2 THEN 'Lunes' WHEN 3 THEN 'Martes' WHEN 4 THEN 'Miércoles' WHEN 5 THEN 'Jueves' WHEN 6 THEN 'Viernes' WHEN 7 THEN 'Sábado' END) AS DiaSemana, Viaje.FechaHora, RTRIM(Viaje.IDRuta) + (CASE Viaje.IDRuta WHEN Parametro.Texto THEN ': ' + Viaje.RutaOtra ELSE '' END) AS Ruta, Vehiculo.Nombre AS Vehiculo, (CASE ISNULL(Viaje.IDConductor, 1) WHEN 1 THEN '' ELSE Persona.Apellido + ', ' + Persona.Nombre END) AS Conductor, (CASE Viaje.Estado WHEN 'AC' THEN 'Activo' WHEN 'EP' Then 'En Progreso' WHEN 'FI' THEN 'Finalizado' WHEN 'CA' THEN 'Cancelado' END) AS Estado, (CASE ISNULL(Viaje.IDConductor2, 0) WHEN 0 THEN 'Completo' ELSE (CASE ISNULL(@IDConductor_FILTER, 0) WHEN Viaje.IDConductor THEN 'Tramo 1' WHEN Viaje.IDConductor2 THEN 'Tramo 2' END) END) AS Tramo, (CASE Viaje.Charter WHEN 0 THEN '' WHEN 1 THEN 'Sí' END) AS CharterDisplay
		FROM (Viaje LEFT JOIN Persona ON Viaje.IDConductor = Persona.IDPersona) LEFT JOIN Vehiculo ON Viaje.IDVehiculo = Vehiculo.IDVehiculo, Parametro
		WHERE
			(@DiaSemana_FILTER IS NULL OR datepart(weekday, Viaje.FechaHora) = @DiaSemana_FILTER)
			AND (@FechaHoraDesde_FILTER IS NULL OR Viaje.FechaHora >= @FechaHoraDesde_FILTER)
			AND (@FechaHoraHasta_FILTER IS NULL OR Viaje.FechaHora <= @FechaHoraHasta_FILTER)
			AND (@FechaDesde_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) >= convert(char(10), @FechaDesde_FILTER, 111))
			AND (@FechaHasta_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) <= convert(char(10), @FechaHasta_FILTER, 111))
			AND (@HoraDesde_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) >= convert(char(8), @HoraDesde_FILTER, 108))
			AND (@HoraHasta_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) <= convert(char(8), @HoraHasta_FILTER, 108))
			AND (@IDRuta_FILTER IS NULL OR Viaje.IDRuta = @IDRuta_FILTER)
			AND (@IDVehiculo_FILTER IS NULL OR (Viaje.IDVehiculo = @IDVehiculo_FILTER AND Vehiculo.Activo = 1))
			AND (@IDConductor_FILTER IS NULL OR (Viaje.IDConductor = @IDConductor_FILTER AND Persona.Activo = 1))
			AND (@Estado_FILTER IS NULL OR Viaje.Estado = @Estado_FILTER)
			AND (@Charter_FILTER IS NULL OR Viaje.Charter = @Charter_FILTER)
			AND Parametro.IDParametro = 'Ruta_ID_Otra'
			AND (@Personal = 0 OR Viaje.Personal = 0)
		ORDER BY Viaje.FechaHora, Viaje.IDRuta, Viaje.RutaOtra

GO



------------------------------------------------------------------------------------------
-- REPORT_VIAJE_PASAJEROS_DIASEMANA_PRECIO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects
	   WHERE  name = N'sp_Report_Viaje_Pasajeros_DiaSemana_Precio'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_Viaje_Pasajeros_DiaSemana_Precio
GO

CREATE PROCEDURE dbo.sp_Report_Viaje_Pasajeros_DiaSemana_Precio
	@DiaSemana_FILTER int,
	@FechaHoraDesde_FILTER smalldatetime,
	@FechaHoraHasta_FILTER smalldatetime,
	@FechaDesde_FILTER smalldatetime,
	@FechaHasta_FILTER smalldatetime,
	@HoraDesde_FILTER smalldatetime,
	@HoraHasta_FILTER smalldatetime,
	@IDRuta_FILTER char(20),
	@IDVehiculo_FILTER int,
	@IDConductor_FILTER int,
	@IDOrigen_FILTER int,
	@IDDestino_FILTER int,
	@IDGrupoOrigen_FILTER int,
	@IDGrupoDestino_FILTER int,
	@IDListaPrecio_FILTER int,
	@Personal bit AS

	SELECT DatePart(weekday, Viaje.FechaHora) AS DiaSemana, ListaPrecio.Nombre AS ListaPrecio, ViajeDetalle.Importe AS Precio, Count(ViajeDetalle.Indice) AS CantidadPasajeros
		FROM (((ViajeDetalle INNER JOIN Viaje ON ViajeDetalle.FechaHora = Viaje.FechaHora AND ViajeDetalle.IDRuta = Viaje.IDRuta) INNER JOIN RutaDetalle AS RutaDetalleOrigen ON ViajeDetalle.IDRuta = RutaDetalleOrigen.IDRuta AND ViajeDetalle.IDOrigen = RutaDetalleOrigen.IDLugar) INNER JOIN RutaDetalle AS RutaDetalleDestino ON ViajeDetalle.IDRuta = RutaDetalleDestino.IDRuta AND ViajeDetalle.IDDestino = RutaDetalleDestino.IDLugar) INNER JOIN ListaPrecio ON ViajeDetalle.IDListaPrecio = ListaPrecio.IDListaPrecio 
		WHERE ViajeDetalle.OcupanteTipo = 'PA'
			AND ViajeDetalle.Estado = '1CO'
			AND Viaje.Estado <> 'CA'
			AND (Viaje.Estado <> 'FI' OR ViajeDetalle.Realizado = 1)
			AND (@DiaSemana_FILTER IS NULL OR datepart(weekday, Viaje.FechaHora) = @DiaSemana_FILTER)
			AND (@FechaHoraDesde_FILTER IS NULL OR Viaje.FechaHora >= @FechaHoraDesde_FILTER)
			AND (@FechaHoraHasta_FILTER IS NULL OR Viaje.FechaHora <= @FechaHoraHasta_FILTER)
			AND (@FechaDesde_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) >= convert(char(10), @FechaDesde_FILTER, 111))
			AND (@FechaHasta_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) <= convert(char(10), @FechaHasta_FILTER, 111))
			AND (@HoraDesde_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) >= convert(char(8), @HoraDesde_FILTER, 108))
			AND (@HoraHasta_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) <= convert(char(8), @HoraHasta_FILTER, 108))
			AND (@IDRuta_FILTER IS NULL OR Viaje.IDRuta = @IDRuta_FILTER)
			AND (@IDVehiculo_FILTER IS NULL OR Viaje.IDVehiculo = @IDVehiculo_FILTER)
			AND (@IDConductor_FILTER IS NULL OR Viaje.IDConductor = @IDConductor_FILTER)
			AND (@IDOrigen_FILTER IS NULL OR ViajeDetalle.IDOrigen = @IDOrigen_FILTER)
			AND (@IDDestino_FILTER IS NULL OR ViajeDetalle.IDDestino = @IDDestino_FILTER)
			AND (@IDGrupoOrigen_FILTER IS NULL OR RutaDetalleOrigen.IDLugarGrupo = @IDGrupoOrigen_FILTER)
			AND (@IDGrupoDestino_FILTER IS NULL OR RutaDetalleDestino.IDLugarGrupo = @IDGrupoDestino_FILTER)
			AND (@IDListaPrecio_FILTER IS NULL OR ViajeDetalle.IDListaPrecio = @IDListaPrecio_FILTER)
			AND (@Personal = 0 OR Viaje.Personal = 0)
		GROUP BY DatePart(weekday, Viaje.FechaHora), ListaPrecio.Nombre, ViajeDetalle.Importe
		ORDER BY DiaSemana, ListaPrecio, Precio

GO



------------------------------------------------------------------------------------------
-- REPORT_VIAJE_PASAJEROSVIAJESASIENTOS_DIASEMANA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects
	   WHERE  name = N'sp_Report_Viaje_PasajerosViajesAsientos_DiaSemana'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_Viaje_PasajerosViajesAsientos_DiaSemana
GO

CREATE PROCEDURE dbo.sp_Report_Viaje_PasajerosViajesAsientos_DiaSemana
	@DiaSemana_FILTER int,
	@FechaHoraDesde_FILTER smalldatetime,
	@FechaHoraHasta_FILTER smalldatetime,
	@FechaDesde_FILTER smalldatetime,
	@FechaHasta_FILTER smalldatetime,
	@HoraDesde_FILTER smalldatetime,
	@HoraHasta_FILTER smalldatetime,
	@IDRuta_FILTER char(20),
	@IDVehiculo_FILTER int,
	@IDConductor_FILTER int,
	@IDOrigen_FILTER int,
	@IDDestino_FILTER int,
	@IDGrupoOrigen_FILTER int,
	@IDGrupoDestino_FILTER int,
	@DividirViajes bit,
	@Personal bit AS

	SELECT Viajes.DiaSemana, IsNull(Ocupantes.Ocupantes, 0) AS Ocupantes, (CASE ISNULL(@DividirViajes, 0) WHEN 0 THEN Viajes.Viajes WHEN 1 THEN Viajes.Viajes / 2 END) AS Viajes, IsNull(Viajes.Asientos, 0) AS Asientos
		FROM
			(SELECT DatePart(weekday, Viaje.FechaHora) AS DiaSemana, CONVERT(FLOAT, COUNT(CONVERT(CHAR(19), Viaje.FechaHora, 120) + Viaje.IDRuta)) AS Viajes, SUM(Vehiculo.Asiento) AS Asientos
				FROM Viaje LEFT JOIN Vehiculo ON Viaje.IDVehiculo = Vehiculo.IDVehiculo
				WHERE Viaje.Estado <> 'CA'
					AND (@DiaSemana_FILTER IS NULL OR datepart(weekday, Viaje.FechaHora) = @DiaSemana_FILTER)
					AND (@FechaHoraDesde_FILTER IS NULL OR Viaje.FechaHora >= @FechaHoraDesde_FILTER)
					AND (@FechaHoraHasta_FILTER IS NULL OR Viaje.FechaHora <= @FechaHoraHasta_FILTER)
					AND (@FechaDesde_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) >= convert(char(10), @FechaDesde_FILTER, 111))
					AND (@FechaHasta_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) <= convert(char(10), @FechaHasta_FILTER, 111))
					AND (@HoraDesde_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) >= convert(char(8), @HoraDesde_FILTER, 108))
					AND (@HoraHasta_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) <= convert(char(8), @HoraHasta_FILTER, 108))
					AND (@IDRuta_FILTER IS NULL OR Viaje.IDRuta = @IDRuta_FILTER)
					AND (@IDVehiculo_FILTER IS NULL OR Viaje.IDVehiculo = @IDVehiculo_FILTER)
					AND (@IDConductor_FILTER IS NULL OR Viaje.IDConductor = @IDConductor_FILTER)
					AND (@Personal = 0 OR Viaje.Personal = 0)
				GROUP BY DatePart(weekday, Viaje.FechaHora)) AS Viajes
			LEFT JOIN
			(SELECT DatePart(weekday, Viaje.FechaHora) AS DiaSemana, Count(ViajeDetalle.Indice) AS Ocupantes
				FROM ((ViajeDetalle INNER JOIN Viaje ON ViajeDetalle.FechaHora = Viaje.FechaHora AND ViajeDetalle.IDRuta = Viaje.IDRuta) INNER JOIN RutaDetalle AS RutaDetalleOrigen ON ViajeDetalle.IDRuta = RutaDetalleOrigen.IDRuta AND ViajeDetalle.IDOrigen = RutaDetalleOrigen.IDLugar) INNER JOIN RutaDetalle AS RutaDetalleDestino ON ViajeDetalle.IDRuta = RutaDetalleDestino.IDRuta AND ViajeDetalle.IDDestino = RutaDetalleDestino.IDLugar
				WHERE ViajeDetalle.OcupanteTipo = 'PA'
					AND ViajeDetalle.Estado = '1CO'
					AND Viaje.Estado <> 'CA'
					AND (Viaje.Estado <> 'FI' OR ViajeDetalle.Realizado = 1)
					AND (@DiaSemana_FILTER IS NULL OR datepart(weekday, Viaje.FechaHora) = @DiaSemana_FILTER)
					AND (@FechaHoraDesde_FILTER IS NULL OR Viaje.FechaHora >= @FechaHoraDesde_FILTER)
					AND (@FechaHoraHasta_FILTER IS NULL OR Viaje.FechaHora <= @FechaHoraHasta_FILTER)
					AND (@FechaDesde_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) >= convert(char(10), @FechaDesde_FILTER, 111))
					AND (@FechaHasta_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) <= convert(char(10), @FechaHasta_FILTER, 111))
					AND (@HoraDesde_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) >= convert(char(8), @HoraDesde_FILTER, 108))
					AND (@HoraHasta_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) <= convert(char(8), @HoraHasta_FILTER, 108))
					AND (@IDRuta_FILTER IS NULL OR Viaje.IDRuta = @IDRuta_FILTER)
					AND (@IDVehiculo_FILTER IS NULL OR Viaje.IDVehiculo = @IDVehiculo_FILTER)
					AND (@IDConductor_FILTER IS NULL OR Viaje.IDConductor = @IDConductor_FILTER)
					AND (@IDOrigen_FILTER IS NULL OR ViajeDetalle.IDOrigen = @IDOrigen_FILTER)
					AND (@IDDestino_FILTER IS NULL OR ViajeDetalle.IDDestino = @IDDestino_FILTER)
					AND (@IDGrupoOrigen_FILTER IS NULL OR RutaDetalleOrigen.IDLugarGrupo = @IDGrupoOrigen_FILTER)
					AND (@IDGrupoDestino_FILTER IS NULL OR RutaDetalleDestino.IDLugarGrupo = @IDGrupoDestino_FILTER)
					AND (@Personal = 0 OR Viaje.Personal = 0)
				GROUP BY DatePart(weekday, Viaje.FechaHora)) AS Ocupantes
			ON Viajes.DiaSemana = Ocupantes.DiaSemana
		ORDER BY Viajes.DiaSemana

GO



------------------------------------------------------------------------------------------
-- REPORT_VIAJE_PASAJEROSVIAJESASIENTOS_VEHICULO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Report_Viaje_PasajerosViajesAsientos_Vehiculo' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_Viaje_PasajerosViajesAsientos_Vehiculo
GO

CREATE PROCEDURE dbo.sp_Report_Viaje_PasajerosViajesAsientos_Vehiculo
	@DiaSemana_FILTER int,
	@FechaHoraDesde_FILTER smalldatetime,
	@FechaHoraHasta_FILTER smalldatetime,
	@FechaDesde_FILTER smalldatetime,
	@FechaHasta_FILTER smalldatetime,
	@HoraDesde_FILTER smalldatetime,
	@HoraHasta_FILTER smalldatetime,
	@IDRuta_FILTER char(20),
	@IDVehiculo_FILTER int,
	@IDConductor_FILTER int,
	@IDOrigen_FILTER int,
	@IDDestino_FILTER int,
	@IDGrupoOrigen_FILTER int,
	@IDGrupoDestino_FILTER int,
	@DividirViajes bit,
	@Personal bit AS

	SELECT Viajes.Vehiculo, IsNull(Ocupantes.Ocupantes, 0) AS Ocupantes, (CASE ISNULL(@DividirViajes, 0) WHEN 0 THEN Viajes.Viajes WHEN 1 THEN Viajes.Viajes / 2 END) AS Viajes, IsNull(Viajes.Asientos, 0) AS Asientos
		FROM
			(SELECT Viaje.IDVehiculo, Vehiculo.Nombre AS Vehiculo, CONVERT(FLOAT, COUNT(CONVERT(CHAR(19), Viaje.FechaHora, 120) + Viaje.IDRuta)) AS Viajes, SUM(Vehiculo.Asiento) AS Asientos
				FROM Viaje INNER JOIN Vehiculo ON Viaje.IDVehiculo = Vehiculo.IDVehiculo
				WHERE Viaje.Estado <> 'CA'
					AND (@DiaSemana_FILTER IS NULL OR datepart(weekday, Viaje.FechaHora) = @DiaSemana_FILTER)
					AND (@FechaHoraDesde_FILTER IS NULL OR Viaje.FechaHora >= @FechaHoraDesde_FILTER)
					AND (@FechaHoraHasta_FILTER IS NULL OR Viaje.FechaHora <= @FechaHoraHasta_FILTER)
					AND (@FechaDesde_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) >= convert(char(10), @FechaDesde_FILTER, 111))
					AND (@FechaHasta_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) <= convert(char(10), @FechaHasta_FILTER, 111))
					AND (@HoraDesde_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) >= convert(char(8), @HoraDesde_FILTER, 108))
					AND (@HoraHasta_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) <= convert(char(8), @HoraHasta_FILTER, 108))
					AND (@IDRuta_FILTER IS NULL OR Viaje.IDRuta = @IDRuta_FILTER)
					AND (@IDVehiculo_FILTER IS NULL OR Viaje.IDVehiculo = @IDVehiculo_FILTER)
					AND (@IDConductor_FILTER IS NULL OR Viaje.IDConductor = @IDConductor_FILTER)
					AND (@Personal = 0 OR Viaje.Personal = 0)
				GROUP BY Viaje.IDVehiculo, Vehiculo.Nombre) AS Viajes
			LEFT JOIN
			(SELECT Viaje.IDVehiculo, Count(ViajeDetalle.Indice) AS Ocupantes
				FROM ((ViajeDetalle INNER JOIN Viaje ON ViajeDetalle.FechaHora = Viaje.FechaHora AND ViajeDetalle.IDRuta = Viaje.IDRuta)
					LEFT JOIN RutaDetalle AS RutaDetalleOrigen ON ViajeDetalle.IDRuta = RutaDetalleOrigen.IDRuta AND ViajeDetalle.IDOrigen = RutaDetalleOrigen.IDLugar)
					LEFT JOIN RutaDetalle AS RutaDetalleDestino ON ViajeDetalle.IDRuta = RutaDetalleDestino.IDRuta AND ViajeDetalle.IDDestino = RutaDetalleDestino.IDLugar
				WHERE ViajeDetalle.OcupanteTipo = 'PA'
					AND ViajeDetalle.Estado = '1CO'
					AND Viaje.Estado <> 'CA'
					AND (Viaje.Estado <> 'FI' OR ViajeDetalle.Realizado = 1) 
					AND (@DiaSemana_FILTER IS NULL OR datepart(weekday, Viaje.FechaHora) = @DiaSemana_FILTER)
					AND (@FechaHoraDesde_FILTER IS NULL OR Viaje.FechaHora >= @FechaHoraDesde_FILTER)
					AND (@FechaHoraHasta_FILTER IS NULL OR Viaje.FechaHora <= @FechaHoraHasta_FILTER)
					AND (@FechaDesde_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) >= convert(char(10), @FechaDesde_FILTER, 111))
					AND (@FechaHasta_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) <= convert(char(10), @FechaHasta_FILTER, 111))
					AND (@HoraDesde_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) >= convert(char(8), @HoraDesde_FILTER, 108))
					AND (@HoraHasta_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) <= convert(char(8), @HoraHasta_FILTER, 108))
					AND (@IDRuta_FILTER IS NULL OR Viaje.IDRuta = @IDRuta_FILTER)
					AND (@IDVehiculo_FILTER IS NULL OR Viaje.IDVehiculo = @IDVehiculo_FILTER)
					AND (@IDConductor_FILTER IS NULL OR Viaje.IDConductor = @IDConductor_FILTER)
					AND (@IDOrigen_FILTER IS NULL OR ViajeDetalle.IDOrigen = @IDOrigen_FILTER)
					AND (@IDDestino_FILTER IS NULL OR ViajeDetalle.IDDestino = @IDDestino_FILTER)
					AND (@IDGrupoOrigen_FILTER IS NULL OR RutaDetalleOrigen.IDLugarGrupo = @IDGrupoOrigen_FILTER)
					AND (@IDGrupoDestino_FILTER IS NULL OR RutaDetalleDestino.IDLugarGrupo = @IDGrupoDestino_FILTER)
					AND (@Personal = 0 OR Viaje.Personal = 0)
				GROUP BY Viaje.IDVehiculo) AS Ocupantes
			ON Viajes.IDVehiculo = Ocupantes.IDVehiculo
		ORDER BY Viajes.Vehiculo

GO



------------------------------------------------------------------------------------------
-- REPORT_VIAJE_PASAJEROSVIAJESIMPORTE_DIASEMANA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name
	   FROM   sysobjects
	   WHERE  name = N'sp_Report_Viaje_PasajerosViajesImporte_DiaSemana'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_Viaje_PasajerosViajesImporte_DiaSemana
GO

CREATE PROCEDURE dbo.sp_Report_Viaje_PasajerosViajesImporte_DiaSemana
	@DiaSemana_FILTER int,
	@FechaHoraDesde_FILTER smalldatetime,
	@FechaHoraHasta_FILTER smalldatetime,
	@FechaDesde_FILTER smalldatetime,
	@FechaHasta_FILTER smalldatetime,
	@HoraDesde_FILTER smalldatetime,
	@HoraHasta_FILTER smalldatetime,
	@IDRuta1_FILTER char(20),
	@IDRuta2_FILTER char(20),
	@IDRuta3_FILTER char(20),
	@IDRuta4_FILTER char(20),
	@IDVehiculo_FILTER int,
	@IDConductor_FILTER int,
	@OcupanteTipo_FILTER char(2),
	@Realizado_FILTER bit,
	@IDOrigen_FILTER int,
	@IDDestino_FILTER int,
	@IDGrupoOrigen_FILTER int,
	@IDGrupoDestino_FILTER int,
	@DividirViajes bit,
	@Personal bit AS

	SELECT Viajes.DiaSemana, IsNull(Ocupantes.Ocupantes, 0) AS Ocupantes, IsNull(Ocupantes.ImporteTotal, 0) + IsNull(Viajes.Importe, 0) AS ImporteTotal, (CASE ISNULL(@DividirViajes, 0) WHEN 0 THEN Viajes.Viajes WHEN 1 THEN Viajes.Viajes / 2 END) AS Viajes
		FROM
			(SELECT DatePart(weekday, Viaje.FechaHora) AS DiaSemana, CONVERT(FLOAT, COUNT(CONVERT(CHAR(19), Viaje.FechaHora, 120) + Viaje.IDRuta)) AS Viajes, SUM(CASE ISNULL(@OcupanteTipo_FILTER, 'NU') WHEN 'CO' THEN 0 ELSE Viaje.Importe END) AS Importe
				FROM Viaje
				WHERE Viaje.Estado <> 'CA'
					AND (@DiaSemana_FILTER IS NULL OR datepart(weekday, Viaje.FechaHora) = @DiaSemana_FILTER)
					AND (@FechaHoraDesde_FILTER IS NULL OR Viaje.FechaHora >= @FechaHoraDesde_FILTER)
					AND (@FechaHoraHasta_FILTER IS NULL OR Viaje.FechaHora <= @FechaHoraHasta_FILTER)
					AND (@FechaDesde_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) >= convert(char(10), @FechaDesde_FILTER, 111))
					AND (@FechaHasta_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) <= convert(char(10), @FechaHasta_FILTER, 111))
					AND (@HoraDesde_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) >= convert(char(8), @HoraDesde_FILTER, 108))
					AND (@HoraHasta_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) <= convert(char(8), @HoraHasta_FILTER, 108))
					AND (Viaje.IDRuta = @IDRuta1_FILTER OR Viaje.IDRuta = @IDRuta2_FILTER OR Viaje.IDRuta = @IDRuta3_FILTER OR Viaje.IDRuta = @IDRuta4_FILTER OR (@IDRuta1_FILTER IS NULL AND @IDRuta2_FILTER IS NULL AND @IDRuta3_FILTER IS NULL AND @IDRuta4_FILTER IS NULL))
					AND (@IDVehiculo_FILTER IS NULL OR Viaje.IDVehiculo = @IDVehiculo_FILTER)
					AND (@IDConductor_FILTER IS NULL OR Viaje.IDConductor = @IDConductor_FILTER)
					AND (@Personal = 0 OR Viaje.Personal = 0)
				GROUP BY DatePart(weekday, Viaje.FechaHora)) AS Viajes
			LEFT JOIN
			(SELECT DatePart(weekday, Viaje.FechaHora) AS DiaSemana, Count(ViajeDetalle.Indice) AS Ocupantes, Sum(ViajeDetalle.Importe) AS ImporteTotal
				FROM ((ViajeDetalle INNER JOIN Viaje ON ViajeDetalle.FechaHora = Viaje.FechaHora AND ViajeDetalle.IDRuta = Viaje.IDRuta)
					LEFT JOIN RutaDetalle AS RutaDetalleOrigen ON ViajeDetalle.IDRuta = RutaDetalleOrigen.IDRuta AND ViajeDetalle.IDOrigen = RutaDetalleOrigen.IDLugar)
					LEFT JOIN RutaDetalle AS RutaDetalleDestino ON ViajeDetalle.IDRuta = RutaDetalleDestino.IDRuta AND ViajeDetalle.IDDestino = RutaDetalleDestino.IDLugar
				WHERE (@OcupanteTipo_FILTER IS NULL OR ViajeDetalle.OcupanteTipo = @OcupanteTipo_FILTER)
					AND ViajeDetalle.Estado = '1CO'
					AND (ViajeDetalle.OcupanteTipo = 'CO' OR ViajeDetalle.Realizado IS NULL OR ViajeDetalle.Realizado = 1 OR ViajeDetalle.ForzarDebito = 1)
					AND Viaje.Estado <> 'CA'
					AND (@DiaSemana_FILTER IS NULL OR datepart(weekday, Viaje.FechaHora) = @DiaSemana_FILTER)
					AND (@FechaHoraDesde_FILTER IS NULL OR Viaje.FechaHora >= @FechaHoraDesde_FILTER)
					AND (@FechaHoraHasta_FILTER IS NULL OR Viaje.FechaHora <= @FechaHoraHasta_FILTER)
					AND (@FechaDesde_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) >= convert(char(10), @FechaDesde_FILTER, 111))
					AND (@FechaHasta_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) <= convert(char(10), @FechaHasta_FILTER, 111))
					AND (@HoraDesde_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) >= convert(char(8), @HoraDesde_FILTER, 108))
					AND (@HoraHasta_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) <= convert(char(8), @HoraHasta_FILTER, 108))
					AND (Viaje.IDRuta = @IDRuta1_FILTER OR Viaje.IDRuta = @IDRuta2_FILTER OR Viaje.IDRuta = @IDRuta3_FILTER OR Viaje.IDRuta = @IDRuta4_FILTER OR (@IDRuta1_FILTER IS NULL AND @IDRuta2_FILTER IS NULL AND @IDRuta3_FILTER IS NULL AND @IDRuta4_FILTER IS NULL))
					AND (@IDVehiculo_FILTER IS NULL OR Viaje.IDVehiculo = @IDVehiculo_FILTER)
					AND (@IDConductor_FILTER IS NULL OR Viaje.IDConductor = @IDConductor_FILTER)
					AND (@Realizado_FILTER IS NULL OR ViajeDetalle.Realizado = @Realizado_FILTER OR ViajeDetalle.OcupanteTipo = 'CO')
					AND (@IDOrigen_FILTER IS NULL OR ViajeDetalle.IDOrigen = @IDOrigen_FILTER)
					AND (@IDDestino_FILTER IS NULL OR ViajeDetalle.IDDestino = @IDDestino_FILTER)
					AND (@IDGrupoOrigen_FILTER IS NULL OR RutaDetalleOrigen.IDLugarGrupo = @IDGrupoOrigen_FILTER)
					AND (@IDGrupoDestino_FILTER IS NULL OR RutaDetalleDestino.IDLugarGrupo = @IDGrupoDestino_FILTER)
					AND (@Personal = 0 OR Viaje.Personal = 0)
				GROUP BY DatePart(weekday, Viaje.FechaHora)) AS Ocupantes
			ON Viajes.DiaSemana = Ocupantes.DiaSemana
		ORDER BY Viajes.DiaSemana

GO



------------------------------------------------------------------------------------------
-- REPORT_VIAJE_PASAJEROSVIAJESIMPORTE_LUGARGRUPO_DIASEMANA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name
	   FROM   sysobjects
	   WHERE  name = N'sp_Report_Viaje_PasajerosViajesImporte_LugarGrupo_DiaSemana'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_Viaje_PasajerosViajesImporte_LugarGrupo_DiaSemana
GO

CREATE PROCEDURE dbo.sp_Report_Viaje_PasajerosViajesImporte_LugarGrupo_DiaSemana
	@DiaSemana_FILTER int,
	@FechaHoraDesde_FILTER smalldatetime,
	@FechaHoraHasta_FILTER smalldatetime,
	@FechaDesde_FILTER smalldatetime,
	@FechaHasta_FILTER smalldatetime,
	@HoraDesde_FILTER smalldatetime,
	@HoraHasta_FILTER smalldatetime,
	@IDRuta_FILTER char(20),
	@IDVehiculo_FILTER int,
	@IDConductor_FILTER int,
	@OcupanteTipo_FILTER char(2),
	@Realizado_FILTER bit,
	@IDOrigen_FILTER int,
	@IDDestino_FILTER int,
	@IDGrupoOrigen_FILTER int,
	@IDGrupoDestino_FILTER int,
	@DividirViajes bit,
	@Personal bit AS

	DECLARE @LugarGrupoOrdenada table(IDLugarGrupo int NOT NULL PRIMARY KEY, Nombre varchar(100) NOT NULL, Orden int NULL)
	DECLARE @IDLugarGrupo int
	DECLARE @Nombre varchar(100)
	DECLARE @Orden int

	-- INSERTO LOS GRUPOS DE LUGARES EN LAS TABLAS TEMPORARIAS, EN EL ORDEN QUE CORRESPONDA
	DECLARE LugarGrupoCursor
		CURSOR LOCAL FORWARD_ONLY KEYSET
		FOR SELECT LugarGrupo.IDLugarGrupo, LugarGrupo.Nombre
			FROM RutaDetalle INNER JOIN LugarGrupo ON RutaDetalle.IDLugarGrupo = LugarGrupo.IDLugarGrupo
			WHERE RutaDetalle.IDRuta = @IDRuta_FILTER
			GROUP BY LugarGrupo.Nombre, LugarGrupo.IDLugarGrupo
			ORDER BY MAX(RutaDetalle.Indice)
	OPEN LugarGrupoCursor
	FETCH NEXT FROM LugarGrupoCursor INTO @IDLugarGrupo, @Nombre
	SET @Orden = 1
	WHILE @@FETCH_STATUS = 0
	BEGIN
		INSERT INTO @LugarGrupoOrdenada
			(IDLugarGrupo, Nombre, Orden)
			VALUES (@IDLugarGrupo, @Nombre, @Orden)
		SET @Orden = @Orden + 1
		FETCH NEXT FROM LugarGrupoCursor INTO @IDLugarGrupo, @Nombre
	END
	CLOSE LugarGrupoCursor
	DEALLOCATE LugarGrupoCursor
	
	-- SELECCION DE DATOS PARA EL REPORTE
	SELECT Ocupantes.LugarGrupoOrigenNombre, Ocupantes.LugarGrupoDestinoNombre, Viajes.DiaSemana, IsNull(Ocupantes.Ocupantes, 0) AS Ocupantes, IsNull(Ocupantes.ImporteTotal, 0) + IsNull(Viajes.Importe, 0) AS ImporteTotal, (CASE ISNULL(@DividirViajes, 0) WHEN 0 THEN Viajes.Viajes WHEN 1 THEN Viajes.Viajes / 2 END) AS Viajes
		FROM
			(SELECT DatePart(weekday, Viaje.FechaHora) AS DiaSemana, CONVERT(FLOAT, COUNT(CONVERT(CHAR(19), Viaje.FechaHora, 120) + Viaje.IDRuta)) AS Viajes, SUM(CASE ISNULL(@OcupanteTipo_FILTER, 'NU') WHEN 'CO' THEN 0 ELSE Viaje.Importe END) AS Importe
				FROM Viaje
				WHERE Viaje.Estado <> 'CA'
					AND (@DiaSemana_FILTER IS NULL OR datepart(weekday, Viaje.FechaHora) = @DiaSemana_FILTER)
					AND (@FechaHoraDesde_FILTER IS NULL OR Viaje.FechaHora >= @FechaHoraDesde_FILTER)
					AND (@FechaHoraHasta_FILTER IS NULL OR Viaje.FechaHora <= @FechaHoraHasta_FILTER)
					AND (@FechaDesde_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) >= convert(char(10), @FechaDesde_FILTER, 111))
					AND (@FechaHasta_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) <= convert(char(10), @FechaHasta_FILTER, 111))
					AND (@HoraDesde_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) >= convert(char(8), @HoraDesde_FILTER, 108))
					AND (@HoraHasta_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) <= convert(char(8), @HoraHasta_FILTER, 108))
					AND (@IDRuta_FILTER IS NULL OR Viaje.IDRuta = @IDRuta_FILTER)
					AND (@IDVehiculo_FILTER IS NULL OR Viaje.IDVehiculo = @IDVehiculo_FILTER)
					AND (@IDConductor_FILTER IS NULL OR Viaje.IDConductor = @IDConductor_FILTER)
					AND (@Personal = 0 OR Viaje.Personal = 0)
				GROUP BY DatePart(weekday, Viaje.FechaHora)) AS Viajes
			LEFT JOIN
			(SELECT LugarGrupoOrigen.Orden AS LugarGrupoOrigenOrden, LugarGrupoOrigen.Nombre AS LugarGrupoOrigenNombre, LugarGrupoDestino.Orden AS LugarGrupoDestinoOrden, LugarGrupoDestino.Nombre AS LugarGrupoDestinoNombre, DatePart(weekday, Viaje.FechaHora) AS DiaSemana, Count(ViajeDetalle.Indice) AS Ocupantes, Sum(ViajeDetalle.Importe) AS ImporteTotal
				FROM ((((ViajeDetalle INNER JOIN Viaje ON ViajeDetalle.FechaHora = Viaje.FechaHora AND ViajeDetalle.IDRuta = Viaje.IDRuta) INNER JOIN RutaDetalle AS RutaDetalleOrigen ON ViajeDetalle.IDRuta = RutaDetalleOrigen.IDRuta AND ViajeDetalle.IDOrigen = RutaDetalleOrigen.IDLugar) INNER JOIN RutaDetalle AS RutaDetalleDestino ON ViajeDetalle.IDRuta = RutaDetalleDestino.IDRuta AND ViajeDetalle.IDDestino = RutaDetalleDestino.IDLugar) INNER JOIN @LugarGrupoOrdenada AS LugarGrupoOrigen ON RutaDetalleOrigen.IDLugarGrupo = LugarGrupoOrigen.IDLugarGrupo) INNER JOIN @LugarGrupoOrdenada AS LugarGrupoDestino ON RutaDetalleDestino.IDLugarGrupo = LugarGrupoDestino.IDLugarGrupo
				WHERE (@OcupanteTipo_FILTER IS NULL OR ViajeDetalle.OcupanteTipo = @OcupanteTipo_FILTER)
					AND ViajeDetalle.Estado = '1CO'
					AND (ViajeDetalle.OcupanteTipo = 'CO' OR ViajeDetalle.Realizado IS NULL OR ViajeDetalle.Realizado = 1 OR ViajeDetalle.ForzarDebito = 1)
					AND Viaje.Estado <> 'CA'
					AND (@DiaSemana_FILTER IS NULL OR datepart(weekday, Viaje.FechaHora) = @DiaSemana_FILTER)
					AND (@FechaHoraDesde_FILTER IS NULL OR Viaje.FechaHora >= @FechaHoraDesde_FILTER)
					AND (@FechaHoraHasta_FILTER IS NULL OR Viaje.FechaHora <= @FechaHoraHasta_FILTER)
					AND (@FechaDesde_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) >= convert(char(10), @FechaDesde_FILTER, 111))
					AND (@FechaHasta_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) <= convert(char(10), @FechaHasta_FILTER, 111))
					AND (@HoraDesde_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) >= convert(char(8), @HoraDesde_FILTER, 108))
					AND (@HoraHasta_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) <= convert(char(8), @HoraHasta_FILTER, 108))
					AND (@IDRuta_FILTER IS NULL OR Viaje.IDRuta = @IDRuta_FILTER)
					AND (@IDVehiculo_FILTER IS NULL OR Viaje.IDVehiculo = @IDVehiculo_FILTER)
					AND (@IDConductor_FILTER IS NULL OR Viaje.IDConductor = @IDConductor_FILTER)
					AND (@Realizado_FILTER IS NULL OR ViajeDetalle.Realizado = @Realizado_FILTER OR ViajeDetalle.OcupanteTipo = 'CO')
					AND (@IDOrigen_FILTER IS NULL OR ViajeDetalle.IDOrigen = @IDOrigen_FILTER)
					AND (@IDDestino_FILTER IS NULL OR ViajeDetalle.IDDestino = @IDDestino_FILTER)
					AND (@IDGrupoOrigen_FILTER IS NULL OR RutaDetalleOrigen.IDLugarGrupo = @IDGrupoOrigen_FILTER)
					AND (@IDGrupoDestino_FILTER IS NULL OR RutaDetalleDestino.IDLugarGrupo = @IDGrupoDestino_FILTER)
					AND (@Personal = 0 OR Viaje.Personal = 0)
				GROUP BY LugarGrupoOrigen.Orden, LugarGrupoOrigen.Nombre, LugarGrupoDestino.Orden, LugarGrupoDestino.Nombre, DatePart(weekday, Viaje.FechaHora)) AS Ocupantes
			ON Viajes.DiaSemana = Ocupantes.DiaSemana
		ORDER BY Ocupantes.LugarGrupoOrigenOrden, Ocupantes.LugarGrupoDestinoOrden, Viajes.DiaSemana

GO



------------------------------------------------------------------------------------------
-- REPORT_VIAJE_PASAJEROSVIAJESIMPORTE_LUGARGRUPO_HORARIO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name
	   FROM   sysobjects
	   WHERE  name = N'sp_Report_Viaje_PasajerosViajesImporte_LugarGrupo_Horario'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_Viaje_PasajerosViajesImporte_LugarGrupo_Horario
GO

CREATE PROCEDURE dbo.sp_Report_Viaje_PasajerosViajesImporte_LugarGrupo_Horario
	@DiaSemana_FILTER int,
	@FechaHoraDesde_FILTER smalldatetime,
	@FechaHoraHasta_FILTER smalldatetime,
	@FechaDesde_FILTER smalldatetime,
	@FechaHasta_FILTER smalldatetime,
	@HoraDesde_FILTER smalldatetime,
	@HoraHasta_FILTER smalldatetime,
	@IDRuta_FILTER char(20),
	@IDVehiculo_FILTER int,
	@IDConductor_FILTER int,
	@OcupanteTipo_FILTER char(2),
	@Realizado_FILTER bit,
	@IDOrigen_FILTER int,
	@IDDestino_FILTER int,
	@IDGrupoOrigen_FILTER int,
	@IDGrupoDestino_FILTER int,
	@DividirViajes bit,
	@Personal bit AS

	DECLARE @LugarGrupoOrdenada table(IDLugarGrupo int NOT NULL PRIMARY KEY, Nombre varchar(100) NOT NULL, Orden int NULL)
	DECLARE @IDLugarGrupo int
	DECLARE @Nombre varchar(100)
	DECLARE @Orden int

	-- INSERTO LOS GRUPOS DE LUGARES EN LAS TABLAS TEMPORARIAS, EN EL ORDEN QUE CORRESPONDA
	DECLARE LugarGrupoCursor
		CURSOR LOCAL FORWARD_ONLY KEYSET
		FOR SELECT LugarGrupo.IDLugarGrupo, LugarGrupo.Nombre
			FROM RutaDetalle INNER JOIN LugarGrupo ON RutaDetalle.IDLugarGrupo = LugarGrupo.IDLugarGrupo
			WHERE RutaDetalle.IDRuta = @IDRuta_FILTER
			GROUP BY LugarGrupo.Nombre, LugarGrupo.IDLugarGrupo
			ORDER BY MAX(RutaDetalle.Indice)
	OPEN LugarGrupoCursor
	FETCH NEXT FROM LugarGrupoCursor INTO @IDLugarGrupo, @Nombre
	SET @Orden = 1
	WHILE @@FETCH_STATUS = 0
	BEGIN
		INSERT INTO @LugarGrupoOrdenada
			(IDLugarGrupo, Nombre, Orden)
			VALUES (@IDLugarGrupo, @Nombre, @Orden)
		SET @Orden = @Orden + 1
		FETCH NEXT FROM LugarGrupoCursor INTO @IDLugarGrupo, @Nombre
	END
	CLOSE LugarGrupoCursor
	DEALLOCATE LugarGrupoCursor
	
	-- SELECCION DE DATOS PARA EL REPORTE
	SELECT Ocupantes.LugarGrupoOrigenNombre, Ocupantes.LugarGrupoDestinoNombre, Viajes.Hora, IsNull(Ocupantes.Ocupantes, 0) AS Ocupantes, IsNull(Ocupantes.ImporteTotal, 0) + IsNull(Viajes.Importe, 0) AS ImporteTotal, (CASE ISNULL(@DividirViajes, 0) WHEN 0 THEN Viajes.Viajes WHEN 1 THEN Viajes.Viajes / 2 END) AS Viajes
		FROM
			(SELECT convert(char(8), Viaje.FechaHora, 108) AS Hora, CONVERT(FLOAT, COUNT(CONVERT(CHAR(19), Viaje.FechaHora, 120) + Viaje.IDRuta)) AS Viajes, SUM(CASE ISNULL(@OcupanteTipo_FILTER, 'NU') WHEN 'CO' THEN 0 ELSE Viaje.Importe END) AS Importe
				FROM Viaje
				WHERE Viaje.Estado <> 'CA'
					AND (@DiaSemana_FILTER IS NULL OR datepart(weekday, Viaje.FechaHora) = @DiaSemana_FILTER)
					AND (@FechaHoraDesde_FILTER IS NULL OR Viaje.FechaHora >= @FechaHoraDesde_FILTER)
					AND (@FechaHoraHasta_FILTER IS NULL OR Viaje.FechaHora <= @FechaHoraHasta_FILTER)
					AND (@FechaDesde_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) >= convert(char(10), @FechaDesde_FILTER, 111))
					AND (@FechaHasta_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) <= convert(char(10), @FechaHasta_FILTER, 111))
					AND (@HoraDesde_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) >= convert(char(8), @HoraDesde_FILTER, 108))
					AND (@HoraHasta_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) <= convert(char(8), @HoraHasta_FILTER, 108))
					AND (@IDRuta_FILTER IS NULL OR Viaje.IDRuta = @IDRuta_FILTER)
					AND (@IDVehiculo_FILTER IS NULL OR Viaje.IDVehiculo = @IDVehiculo_FILTER)
					AND (@IDConductor_FILTER IS NULL OR Viaje.IDConductor = @IDConductor_FILTER)
					AND (@Personal = 0 OR Viaje.Personal = 0)
				GROUP BY convert(char(8), Viaje.FechaHora, 108)) AS Viajes
			LEFT JOIN
			(SELECT LugarGrupoOrigen.Orden AS LugarGrupoOrigenOrden, LugarGrupoOrigen.Nombre AS LugarGrupoOrigenNombre, LugarGrupoDestino.Orden AS LugarGrupoDestinoOrden, LugarGrupoDestino.Nombre AS LugarGrupoDestinoNombre, convert(char(8), Viaje.FechaHora, 108) AS Hora, Count(ViajeDetalle.Indice) AS Ocupantes, Sum(ViajeDetalle.Importe) AS ImporteTotal
				FROM ((((ViajeDetalle INNER JOIN Viaje ON ViajeDetalle.FechaHora = Viaje.FechaHora AND ViajeDetalle.IDRuta = Viaje.IDRuta)
					INNER JOIN RutaDetalle AS RutaDetalleOrigen ON ViajeDetalle.IDRuta = RutaDetalleOrigen.IDRuta AND ViajeDetalle.IDOrigen = RutaDetalleOrigen.IDLugar)
					INNER JOIN RutaDetalle AS RutaDetalleDestino ON ViajeDetalle.IDRuta = RutaDetalleDestino.IDRuta AND ViajeDetalle.IDDestino = RutaDetalleDestino.IDLugar)
					INNER JOIN @LugarGrupoOrdenada AS LugarGrupoOrigen ON RutaDetalleOrigen.IDLugarGrupo = LugarGrupoOrigen.IDLugarGrupo)
					INNER JOIN @LugarGrupoOrdenada AS LugarGrupoDestino ON RutaDetalleDestino.IDLugarGrupo = LugarGrupoDestino.IDLugarGrupo
				WHERE (@OcupanteTipo_FILTER IS NULL OR ViajeDetalle.OcupanteTipo = @OcupanteTipo_FILTER)
					AND ViajeDetalle.Estado = '1CO'
					AND (ViajeDetalle.OcupanteTipo = 'CO' OR ViajeDetalle.Realizado IS NULL OR ViajeDetalle.Realizado = 1 OR ViajeDetalle.ForzarDebito = 1)
					AND Viaje.Estado <> 'CA'
					AND (@DiaSemana_FILTER IS NULL OR datepart(weekday, Viaje.FechaHora) = @DiaSemana_FILTER)
					AND (@FechaHoraDesde_FILTER IS NULL OR Viaje.FechaHora >= @FechaHoraDesde_FILTER)
					AND (@FechaHoraHasta_FILTER IS NULL OR Viaje.FechaHora <= @FechaHoraHasta_FILTER)
					AND (@FechaDesde_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) >= convert(char(10), @FechaDesde_FILTER, 111))
					AND (@FechaHasta_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) <= convert(char(10), @FechaHasta_FILTER, 111))
					AND (@HoraDesde_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) >= convert(char(8), @HoraDesde_FILTER, 108))
					AND (@HoraHasta_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) <= convert(char(8), @HoraHasta_FILTER, 108))
					AND (@IDRuta_FILTER IS NULL OR Viaje.IDRuta = @IDRuta_FILTER)
					AND (@IDVehiculo_FILTER IS NULL OR Viaje.IDVehiculo = @IDVehiculo_FILTER)
					AND (@IDConductor_FILTER IS NULL OR Viaje.IDConductor = @IDConductor_FILTER)
					AND (@Realizado_FILTER IS NULL OR ViajeDetalle.Realizado = @Realizado_FILTER OR ViajeDetalle.OcupanteTipo = 'CO')
					AND (@IDOrigen_FILTER IS NULL OR ViajeDetalle.IDOrigen = @IDOrigen_FILTER)
					AND (@IDDestino_FILTER IS NULL OR ViajeDetalle.IDDestino = @IDDestino_FILTER)
					AND (@IDGrupoOrigen_FILTER IS NULL OR RutaDetalleOrigen.IDLugarGrupo = @IDGrupoOrigen_FILTER)
					AND (@IDGrupoDestino_FILTER IS NULL OR RutaDetalleDestino.IDLugarGrupo = @IDGrupoDestino_FILTER)
					AND (@Personal = 0 OR Viaje.Personal = 0)
				GROUP BY LugarGrupoOrigen.Orden, LugarGrupoOrigen.Nombre, LugarGrupoDestino.Orden, LugarGrupoDestino.Nombre, convert(char(8), Viaje.FechaHora, 108)) AS Ocupantes
			ON Viajes.Hora = Ocupantes.Hora
		WHERE NOT Ocupantes.LugarGrupoOrigenNombre IS NULL
		ORDER BY Ocupantes.LugarGrupoOrigenOrden, Ocupantes.LugarGrupoDestinoOrden, Viajes.Hora

GO



------------------------------------------------------------------------------------------
-- REPORT_VIAJE_PASAJEROSVIAJESIMPORTE_VEHICULO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Report_Viaje_PasajerosViajesImporte_Vehiculo' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_Viaje_PasajerosViajesImporte_Vehiculo
GO

CREATE PROCEDURE dbo.sp_Report_Viaje_PasajerosViajesImporte_Vehiculo
	@DiaSemana_FILTER int,
	@FechaHoraDesde_FILTER smalldatetime,
	@FechaHoraHasta_FILTER smalldatetime,
	@FechaDesde_FILTER smalldatetime,
	@FechaHasta_FILTER smalldatetime,
	@HoraDesde_FILTER smalldatetime,
	@HoraHasta_FILTER smalldatetime,
	@IDRuta1_FILTER char(20),
	@IDRuta2_FILTER char(20),
	@IDRuta3_FILTER char(20),
	@IDRuta4_FILTER char(20),
	@IDVehiculo_FILTER int,
	@IDConductor_FILTER int,
	@OcupanteTipo_FILTER char(2),
	@Realizado_FILTER bit,
	@IDOrigen_FILTER int,
	@IDDestino_FILTER int,
	@IDGrupoOrigen_FILTER int,
	@IDGrupoDestino_FILTER int,
	@DividirViajes bit,
	@Personal bit AS

	SELECT Viajes.Vehiculo, IsNull(Ocupantes.Ocupantes, 0) AS Ocupantes, IsNull(Ocupantes.ImporteTotal, 0) + IsNull(Viajes.Importe, 0) AS ImporteTotal, (CASE ISNULL(@DividirViajes, 0) WHEN 0 THEN Viajes.Viajes WHEN 1 THEN Viajes.Viajes / 2 END) AS Viajes
		FROM
			(SELECT Viaje.IDVehiculo, Vehiculo.Nombre AS Vehiculo, CONVERT(FLOAT, COUNT(CONVERT(CHAR(19), Viaje.FechaHora, 120) + Viaje.IDRuta)) AS Viajes, SUM(CASE ISNULL(@OcupanteTipo_FILTER, 'NU') WHEN 'CO' THEN 0 ELSE Viaje.Importe END) AS Importe
				FROM Viaje INNER JOIN Vehiculo ON Viaje.IDVehiculo = Vehiculo.IDVehiculo
				WHERE Viaje.Estado <> 'CA'
					AND (@DiaSemana_FILTER IS NULL OR datepart(weekday, Viaje.FechaHora) = @DiaSemana_FILTER)
					AND (@FechaHoraDesde_FILTER IS NULL OR Viaje.FechaHora >= @FechaHoraDesde_FILTER)
					AND (@FechaHoraHasta_FILTER IS NULL OR Viaje.FechaHora <= @FechaHoraHasta_FILTER)
					AND (@FechaDesde_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) >= convert(char(10), @FechaDesde_FILTER, 111))
					AND (@FechaHasta_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) <= convert(char(10), @FechaHasta_FILTER, 111))
					AND (@HoraDesde_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) >= convert(char(8), @HoraDesde_FILTER, 108))
					AND (@HoraHasta_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) <= convert(char(8), @HoraHasta_FILTER, 108))
					AND (Viaje.IDRuta = @IDRuta1_FILTER OR Viaje.IDRuta = @IDRuta2_FILTER OR Viaje.IDRuta = @IDRuta3_FILTER OR Viaje.IDRuta = @IDRuta4_FILTER OR (@IDRuta1_FILTER IS NULL AND @IDRuta2_FILTER IS NULL AND @IDRuta3_FILTER IS NULL AND @IDRuta4_FILTER IS NULL))
					AND (@IDVehiculo_FILTER IS NULL OR Viaje.IDVehiculo = @IDVehiculo_FILTER)
					AND (@IDConductor_FILTER IS NULL OR Viaje.IDConductor = @IDConductor_FILTER)
					AND (@Personal = 0 OR Viaje.Personal = 0)
				GROUP BY Viaje.IDVehiculo, Vehiculo.Nombre) AS Viajes
			LEFT JOIN
			(SELECT Viaje.IDVehiculo, Count(ViajeDetalle.Indice) AS Ocupantes, Sum(ViajeDetalle.Importe) AS ImporteTotal
				FROM ((ViajeDetalle INNER JOIN Viaje ON ViajeDetalle.FechaHora = Viaje.FechaHora AND ViajeDetalle.IDRuta = Viaje.IDRuta) INNER JOIN RutaDetalle AS RutaDetalleOrigen ON ViajeDetalle.IDRuta = RutaDetalleOrigen.IDRuta AND ViajeDetalle.IDOrigen = RutaDetalleOrigen.IDLugar) INNER JOIN RutaDetalle AS RutaDetalleDestino ON ViajeDetalle.IDRuta = RutaDetalleDestino.IDRuta AND ViajeDetalle.IDDestino = RutaDetalleDestino.IDLugar
				WHERE (@OcupanteTipo_FILTER IS NULL OR ViajeDetalle.OcupanteTipo = @OcupanteTipo_FILTER)
					AND ViajeDetalle.Estado = '1CO'
					AND (ViajeDetalle.OcupanteTipo = 'CO' OR ViajeDetalle.Realizado IS NULL OR ViajeDetalle.Realizado = 1 OR ViajeDetalle.ForzarDebito = 1)
					AND Viaje.Estado <> 'CA'
					AND (@DiaSemana_FILTER IS NULL OR datepart(weekday, Viaje.FechaHora) = @DiaSemana_FILTER)
					AND (@FechaHoraDesde_FILTER IS NULL OR Viaje.FechaHora >= @FechaHoraDesde_FILTER)
					AND (@FechaHoraHasta_FILTER IS NULL OR Viaje.FechaHora <= @FechaHoraHasta_FILTER)
					AND (@FechaDesde_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) >= convert(char(10), @FechaDesde_FILTER, 111))
					AND (@FechaHasta_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) <= convert(char(10), @FechaHasta_FILTER, 111))
					AND (@HoraDesde_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) >= convert(char(8), @HoraDesde_FILTER, 108))
					AND (@HoraHasta_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) <= convert(char(8), @HoraHasta_FILTER, 108))
					AND (Viaje.IDRuta = @IDRuta1_FILTER OR Viaje.IDRuta = @IDRuta2_FILTER OR Viaje.IDRuta = @IDRuta3_FILTER OR Viaje.IDRuta = @IDRuta4_FILTER OR (@IDRuta1_FILTER IS NULL AND @IDRuta2_FILTER IS NULL AND @IDRuta3_FILTER IS NULL AND @IDRuta4_FILTER IS NULL))
					AND (@IDVehiculo_FILTER IS NULL OR Viaje.IDVehiculo = @IDVehiculo_FILTER)
					AND (@IDConductor_FILTER IS NULL OR Viaje.IDConductor = @IDConductor_FILTER)
					AND (@Realizado_FILTER IS NULL OR ViajeDetalle.Realizado = @Realizado_FILTER OR ViajeDetalle.OcupanteTipo = 'CO')
					AND (@IDOrigen_FILTER IS NULL OR ViajeDetalle.IDOrigen = @IDOrigen_FILTER)
					AND (@IDDestino_FILTER IS NULL OR ViajeDetalle.IDDestino = @IDDestino_FILTER)
					AND (@IDGrupoOrigen_FILTER IS NULL OR RutaDetalleOrigen.IDLugarGrupo = @IDGrupoOrigen_FILTER)
					AND (@IDGrupoDestino_FILTER IS NULL OR RutaDetalleDestino.IDLugarGrupo = @IDGrupoDestino_FILTER)
					AND (@Personal = 0 OR Viaje.Personal = 0)
				GROUP BY Viaje.IDVehiculo) AS Ocupantes
			ON Viajes.IDVehiculo = Ocupantes.IDVehiculo
		ORDER BY Viajes.Vehiculo

GO



------------------------------------------------------------------------------------------
-- REPORT_VEHICULO_MANTENIMIENTO_AVISO_LIST
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Report_Vehiculo_Mantenimiento_Aviso_List' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_Vehiculo_Mantenimiento_Aviso_List
GO

CREATE PROCEDURE dbo.sp_Report_Vehiculo_Mantenimiento_Aviso_List
	@IDVehiculo_FILTER int,
	@IDVehiculoMantenimientoGrupo_FILTER int AS

	SELECT Vehiculo.Nombre AS Vehiculo, VehiculoMantenimientoGrupo.Nombre AS Grupo, VehiculoMantenimiento.Tipo, ISNULL(Accion.KilometrajeMaximo, 0) + VehiculoMantenimiento.KilometrajeLapso AS KilometrajeMantenimiento, Vehiculo.KilometrajeEstimado AS KilometrajeActual, ISNULL(Accion.FechaHoraMaxima, VehiculoMantenimiento.FechaHoraCreacion) + VehiculoMantenimiento.DiasLapso AS DiasFechaMantenimiento, VehiculoMantenimiento.FechaFecha AS FechaFechaMantenimiento
		FROM ((VehiculoMantenimiento INNER JOIN Vehiculo ON VehiculoMantenimiento.IDVehiculo = Vehiculo.IDVehiculo) INNER JOIN VehiculoMantenimientoGrupo ON VehiculoMantenimiento.IDVehiculoMantenimientoGrupo = VehiculoMantenimientoGrupo.IDVehiculoMantenimientoGrupo)
			LEFT JOIN 
				(SELECT IDVehiculo, IDVehiculoMantenimientoGrupo, MAX(Kilometraje) AS KilometrajeMaximo, MAX(FechaHora) AS FechaHoraMaxima
					FROM VehiculoMantenimientoAccion
					GROUP BY IDVehiculo, IDVehiculoMantenimientoGrupo) AS Accion
				ON VehiculoMantenimiento.IDVehiculo = Accion.IDVehiculo AND VehiculoMantenimiento.IDVehiculoMantenimientoGrupo = Accion.IDVehiculoMantenimientoGrupo
		WHERE VehiculoMantenimiento.Activo = 1
			AND Vehiculo.Activo = 1
			AND (@IDVehiculo_FILTER IS NULL OR VehiculoMantenimiento.IDVehiculo = @IDVehiculo_FILTER)
			AND (@IDVehiculoMantenimientoGrupo_FILTER IS NULL OR VehiculoMantenimiento.IDVehiculoMantenimientoGrupo = @IDVehiculoMantenimientoGrupo_FILTER)
			AND ((VehiculoMantenimiento.Tipo = 'KI' AND (Vehiculo.KilometrajeEstimado - ISNULL(Accion.KilometrajeMaximo, 0)) >= (VehiculoMantenimiento.KilometrajeLapso - VehiculoMantenimiento.KilometrajePreaviso))
				OR (VehiculoMantenimiento.Tipo = 'DI' AND (DATEDIFF(d, ISNULL(Accion.FechaHoraMaxima, VehiculoMantenimiento.FechaHoraCreacion), getdate())) >= (VehiculoMantenimiento.DiasLapso - VehiculoMantenimiento.DiasPreaviso)))
				OR (VehiculoMantenimiento.Tipo = 'FE' AND DATEDIFF(d, getdate(), (VehiculoMantenimiento.FechaFecha - VehiculoMantenimiento.FechaPreaviso)) <= 0)
GO



------------------------------------------------------------------------------------------
-- REPORT_PERSONA_ALARMA_AVISO_LIST
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Report_Persona_Alarma_Aviso_List' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_Persona_Alarma_Aviso_List
GO

CREATE PROCEDURE dbo.sp_Report_Persona_Alarma_Aviso_List
	@IDPersona_FILTER int,
	@IDPersonaAlarmaGrupo_FILTER int AS

	SELECT Persona.Apellido + ', ' + Persona.Nombre AS Persona, PersonaAlarmaGrupo.Nombre AS Grupo, PersonaAlarma.Fecha
		FROM (PersonaAlarma INNER JOIN Persona ON PersonaAlarma.IDPersona = Persona.IDPersona) INNER JOIN PersonaAlarmaGrupo ON PersonaAlarma.IDPersonaAlarmaGrupo = PersonaAlarmaGrupo.IDPersonaAlarmaGrupo
		WHERE PersonaAlarma.Activo = 1
			AND Persona.Activo = 1
			AND (@IDPersona_FILTER IS NULL OR PersonaAlarma.IDPersona = @IDPersona_FILTER)
			AND (@IDPersonaAlarmaGrupo_FILTER IS NULL OR PersonaAlarma.IDPersonaAlarmaGrupo = @IDPersonaAlarmaGrupo_FILTER)
			AND DATEDIFF(d, getdate(), (PersonaAlarma.Fecha - PersonaAlarma.Preaviso)) <= 0
GO



------------------------------------------------------------------------------------------
-- REPORT_ALARMA_AVISO_LIST
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Report_Alarma_Aviso_List' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_Alarma_Aviso_List
GO

CREATE PROCEDURE dbo.sp_Report_Alarma_Aviso_List
	@IDAlarma_FILTER int AS

	SELECT Alarma.Nombre, Alarma.Fecha
		FROM Alarma
		WHERE Alarma.Activo = 1
			AND (@IDAlarma_FILTER IS NULL OR Alarma.IDAlarma = @IDAlarma_FILTER)
			AND DATEDIFF(d, getdate(), (Alarma.Fecha - Alarma.Preaviso)) <= 0
GO



------------------------------------------------------------------------------------------
-- REPORT_PERSONAHORARIO_LIST
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Report_PersonaHorario_List' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_PersonaHorario_List
GO

CREATE PROCEDURE dbo.sp_Report_PersonaHorario_List
	@IDPersona_FILTER int,
	@DiaSemana_FILTER int,
	@HoraDesde_FILTER smalldatetime,
	@HoraHasta_FILTER smalldatetime,
	@IDRuta_FILTER char(20),
	@IDLugarGrupoDesde_FILTER int,
	@IDLugarGrupoHasta_FILTER int,
	@Personal bit AS

	CREATE TABLE #RutaLugar
		(IDRuta char(20) COLLATE Modern_Spanish_CI_AS NOT NULL,
		IDLugar int NOT NULL)

	ALTER TABLE #RutaLugar ADD 
		CONSTRAINT PK__RutaLugar PRIMARY KEY NONCLUSTERED 
			(IDRuta, IDLugar) WITH  FILLFACTOR = 10

	DECLARE @IDRuta char(20)
	DECLARE @IndiceDesde int
	DECLARE @IndiceHasta int

	-- INSERTO LAS RUTAS CON LOS INDICES CORRESPONDIENTES
	DECLARE Rutas_Cursor
		CURSOR LOCAL FORWARD_ONLY KEYSET
		FOR SELECT IDRuta
			FROM Ruta
			WHERE (@IDRuta_FILTER IS NULL OR IDRuta = @IDRuta_FILTER)
	OPEN Rutas_Cursor
	FETCH NEXT FROM Rutas_Cursor INTO @IDRuta
	WHILE @@FETCH_STATUS = 0
	BEGIN
		
		SET @IndiceDesde = (SELECT MIN(Indice)
								FROM RutaDetalle
								WHERE IDRuta = @IDRuta
									AND (@IDLugarGrupoDesde_FILTER IS NULL OR IDLugarGrupo = @IDLugarGrupoDesde_FILTER))
		SET @IndiceHasta = (SELECT MAX(Indice)
								FROM RutaDetalle
								WHERE IDRuta = @IDRuta
									AND (@IDLugarGrupoHasta_FILTER IS NULL OR IDLugarGrupo = @IDLugarGrupoHasta_FILTER))

		IF @IndiceDesde > @IndiceHasta
			--LA DIRECCION DE LA RUTA ES AL REVES AL ORDEN DE LOS PARAMETROS
			BEGIN
			SET @IndiceDesde = (SELECT MIN(Indice)
									FROM RutaDetalle
									WHERE IDRuta = @IDRuta
										AND ((@IDLugarGrupoDesde_FILTER IS NULL AND @IDLugarGrupoHasta_FILTER IS NULL) OR IDLugarGrupo = @IDLugarGrupoHasta_FILTER))
			SET @IndiceHasta = (SELECT MAX(Indice)
									FROM RutaDetalle
									WHERE IDRuta = @IDRuta
										AND ((@IDLugarGrupoDesde_FILTER IS NULL AND @IDLugarGrupoHasta_FILTER IS NULL) OR IDLugarGrupo = @IDLugarGrupoDesde_FILTER))
			END

		INSERT INTO #RutaLugar
			(IDRuta, IDLugar)
			SELECT @IDRuta, IDLugar
				FROM RutaDetalle
				WHERE IDRuta = @IDRuta
					AND ((@IDLugarGrupoDesde_FILTER IS NULL AND @IDLugarGrupoHasta_FILTER IS NULL) OR (Indice >= @IndiceDesde AND Indice <= @IndiceHasta))
		FETCH NEXT FROM Rutas_Cursor INTO @IDRuta
	END
	CLOSE Rutas_Cursor
	DEALLOCATE Rutas_Cursor

	SELECT Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Pasajero, PersonaHorario.DiaSemana, (CASE PersonaHorario.DiaSemana WHEN 1 THEN 'Domingo' WHEN 2 THEN 'Lunes' WHEN 3 THEN 'Martes' WHEN 4 THEN 'Miércoles' WHEN 5 THEN 'Jueves' WHEN 6 THEN 'Viernes' WHEN 7 THEN 'Sábado' END) AS DiaSemanaNombre, PersonaHorario.Hora, PersonaHorario.IDRuta, PersonaHorario.FechaDesde, PersonaHorario.FechaHasta, (CASE IsNull(PersonaHorario.Sube, '') WHEN '' THEN LugarOrigen.Nombre ELSE PersonaHorario.Sube END) AS Origen, (CASE IsNull(PersonaHorario.Baja, '') WHEN '' THEN LugarDestino.Nombre ELSE PersonaHorario.Baja END) AS Destino
		FROM (((((Horario INNER JOIN PersonaHorario ON Horario.DiaSemana = PersonaHorario.DiaSemana AND Horario.Hora = PersonaHorario.Hora AND Horario.IDRuta = PersonaHorario.IDRuta) INNER JOIN Persona ON PersonaHorario.IDPersona = Persona.IDPersona) LEFT JOIN #RutaLugar AS RutaLugarOrigen ON PersonaHorario.IDRuta = RutaLugarOrigen.IDRuta AND PersonaHorario.IDOrigen = RutaLugarOrigen.IDLugar) LEFT JOIN #RutaLugar AS RutaLugarDestino ON PersonaHorario.IDRuta = RutaLugarDestino.IDRuta AND PersonaHorario.IDDestino = RutaLugarDestino.IDLugar) LEFT JOIN Lugar AS LugarOrigen ON PersonaHorario.IDOrigen = LugarOrigen.IDLugar) LEFT JOIN Lugar AS LugarDestino ON PersonaHorario.IDDestino = LugarDestino.IDLugar
		WHERE (@IDPersona_FILTER IS NULL OR Persona.IDPersona = @IDPersona_FILTER)
			AND (@DiaSemana_FILTER IS NULL OR PersonaHorario.DiaSemana = @DiaSemana_FILTER)
			AND (@HoraDesde_FILTER IS NULL OR convert(char(8), PersonaHorario.Hora, 108) >= convert(char(8), @HoraDesde_FILTER, 108))
			AND (@HoraHasta_FILTER IS NULL OR convert(char(8), PersonaHorario.Hora, 108) <= convert(char(8), @HoraHasta_FILTER, 108))
			AND (@IDRuta_FILTER IS NULL OR PersonaHorario.IDRuta = @IDRuta_FILTER)
			AND ((@IDLugarGrupoDesde_FILTER IS NULL AND @IDLugarGrupoHasta_FILTER IS NULL) OR (RutaLugarOrigen.IDLugar IS NOT NULL AND RutaLugarDestino.IDLugar IS NOT NULL))
			AND (@Personal = 0 OR Horario.Personal = @Personal)
		ORDER BY Persona.Apellido, Persona.Nombre, PersonaHorario.DiaSemana, PersonaHorario.Hora, PersonaHorario.IDRuta

GO



------------------------------------------------------------------------------------------
-- REPORT_PERSONARUTA_LIST
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Report_PersonaRuta_List' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_PersonaRuta_List
GO

CREATE PROCEDURE dbo.sp_Report_PersonaRuta_List
	@IDPersona_FILTER int,
	@IDRuta_FILTER char(20),
	@IDListaPrecio_FILTER int AS

	SELECT Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Pasajero, PersonaRuta.IDRuta, (CASE IsNull(PersonaRuta.Sube, '') WHEN '' THEN LugarOrigen.Nombre ELSE PersonaRuta.Sube END) AS Origen, (CASE IsNull(PersonaRuta.Baja, '') WHEN '' THEN LugarDestino.Nombre ELSE PersonaRuta.Baja END) AS Destino, ListaPrecio.Nombre AS ListaPrecio
		FROM (((Persona INNER JOIN PersonaRuta ON Persona.IDPersona = PersonaRuta.IDPersona) INNER JOIN Lugar AS LugarOrigen ON PersonaRuta.IDOrigen = LugarOrigen.IDLugar) INNER JOIN Lugar AS LugarDestino ON PersonaRuta.IDDestino = LugarDestino.IDLugar) INNER JOIN ListaPrecio ON PersonaRuta.IDListaPrecio = ListaPrecio.IDListaPrecio
		WHERE (@IDPersona_FILTER IS NULL OR Persona.IDPersona = @IDPersona_FILTER)
			AND (@IDRuta_FILTER IS NULL OR PersonaRuta.IDRuta = @IDRuta_FILTER)
			AND (@IDListaPrecio_FILTER IS NULL OR PersonaRuta.IDListaPrecio = @IDListaPrecio_FILTER)
		ORDER BY Persona.Apellido, Persona.Nombre, PersonaRuta.IDRuta

GO



------------------------------------------------------------------------------------------
-- REPORT_VIAJEDETALLE_LISTADO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Report_ViajeDetalle_Listado'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_ViajeDetalle_Listado
GO

CREATE PROCEDURE dbo.sp_Report_ViajeDetalle_Listado
	@FechaHoraDesde smalldatetime,
	@FechaHoraHasta smalldatetime,
	@IDRuta1 char(20),
	@IDRuta2 char(20),
	@IDRuta3 char(20),
	@IDRuta4 char(20),
	@ViajeEstado char(2),
	@OcupanteTipo char(2),
	@IDPersona int,
	@IDLugarGrupoOrigen int,
	@IDLugarGrupoDestino int,
	@IDLugarOrigen int,
	@IDLugarDestino int,
	@IDListaPrecio int,
	@Entregada bit,
	@Pagada bit,
	@Personal bit AS

	SELECT ViajeDetalle.FechaHora, ViajeDetalle.IDRuta, ViajeDetalle.Orden, ViajeDetalle.OcupanteTipo, Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Persona, ISNULL(PersonaRecibe.Apellido + (CASE ISNULL(PersonaRecibe.Nombre, '') WHEN '' THEN '' ELSE ', ' + PersonaRecibe.Nombre END), '') + ISNULL(ViajeDetalle.Recibe, '') AS Recibe, ViajeDetalle.Descripcion, ViajeDetalle.Domicilio, ViajeDetalle.Horario, ViajeDetalle.Telefono, (CASE isnull(ViajeDetalle.Sube, '') WHEN '' THEN LugarOrigen.Nombre ELSE ViajeDetalle.Sube END) AS Origen, (CASE isnull(ViajeDetalle.Baja, '') WHEN '' THEN LugarDestino.Nombre ELSE ViajeDetalle.Baja END) AS Destino, ViajeDetalle.Importe, ViajeDetalle.ImporteContado + ViajeDetalle.ImporteCuentaCorriente AS ImportePagado, ViajeDetalle.Entregada
		FROM ((((((ViajeDetalle INNER JOIN Viaje ON ViajeDetalle.FechaHora = Viaje.FechaHora AND ViajeDetalle.IDRuta = Viaje.IDRuta) INNER JOIN Persona ON ViajeDetalle.IDPersona = Persona.IDPersona) INNER JOIN Lugar AS LugarOrigen ON ViajeDetalle.IDOrigen = LugarOrigen.IDLugar) INNER JOIN Lugar AS LugarDestino ON ViajeDetalle.IDDestino = LugarDestino.IDLugar) INNER JOIN RutaDetalle AS RutaDetalleOrigen ON RutaDetalleOrigen.IDRuta = ViajeDetalle.IDRuta AND RutaDetalleOrigen.IDLugar = ViajeDetalle.IDOrigen) INNER JOIN RutaDetalle AS RutaDetalleDestino ON RutaDetalleDestino.IDRuta = ViajeDetalle.IDRuta AND RutaDetalleDestino.IDLugar = ViajeDetalle.IDDestino) LEFT JOIN Persona AS PersonaRecibe ON ViajeDetalle.IDPersonaRecibe = PersonaRecibe.IDPersona
		WHERE ViajeDetalle.Estado = '1CO'
				AND (@FechaHoraDesde IS NULL OR ViajeDetalle.FechaHora >= @FechaHoraDesde)
				AND (@FechaHoraHasta IS NULL OR ViajeDetalle.FechaHora <= @FechaHoraHasta)
				AND (Viaje.IDRuta = @IDRuta1 OR Viaje.IDRuta = @IDRuta2 OR Viaje.IDRuta = @IDRuta3 OR Viaje.IDRuta = @IDRuta4 OR (@IDRuta1 IS NULL AND @IDRuta2 IS NULL AND @IDRuta3 IS NULL AND @IDRuta4 IS NULL))
				AND (@ViajeEstado IS NULL OR ViajeDetalle.Estado = @ViajeEstado)
				AND (@OcupanteTipo IS NULL OR ViajeDetalle.OcupanteTipo = @OcupanteTipo)
				AND (@IDPersona IS NULL OR ViajeDetalle.IDPersona = @IDPersona)
				AND (@IDLugarGrupoOrigen IS NULL OR RutaDetalleOrigen.IDLugarGrupo = @IDLugarGrupoOrigen)
				AND (@IDLugarGrupoDestino IS NULL OR RutaDetalleDestino.IDLugarGrupo = @IDLugarGrupoDestino)
				AND (@IDLugarOrigen IS NULL OR ViajeDetalle.IDOrigen = @IDLugarOrigen)
				AND (@IDLugarDestino IS NULL OR ViajeDetalle.IDDestino = @IDLugarDestino)
				AND (@IDListaPrecio IS NULL OR ViajeDetalle.IDListaPrecio = @IDListaPrecio)
				AND (@Entregada IS NULL OR ViajeDetalle.Entregada = @Entregada)
				AND (@Pagada IS NULL OR (@Pagada = 1 AND ViajeDetalle.ImporteContado + ViajeDetalle.ImporteCuentaCorriente > 0) OR (@Pagada = 0 AND ViajeDetalle.ImporteContado + ViajeDetalle.ImporteCuentaCorriente = 0))
				AND (@Personal = 0 OR Viaje.Personal IS NULL OR Viaje.Personal = 0)
		ORDER BY ViajeDetalle.FechaHora, ViajeDetalle.IDRuta
GO



------------------------------------------------------------------------------------------
-- REPORT_COMISION_LISTADO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Report_Comision_Listado'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_Comision_Listado
GO

CREATE PROCEDURE dbo.sp_Report_Comision_Listado
	@FechaHoraDesde smalldatetime,
	@FechaHoraHasta smalldatetime,
	@IDRuta1 char(20),
	@IDRuta2 char(20),
	@IDRuta3 char(20),
	@IDRuta4 char(20),
	@IDPersona int,
	@IDOrigen int,
	@IDDestino int,
	@IDListaPrecio int,
	@Entregada bit,
	@Pagada bit,
	@RendicionVacia bit,
	@RendicionFechaHoraDesde smalldatetime,
	@RendicionFechaHoraHasta smalldatetime,
	@MostrarTodas bit,
	@Personal bit AS

	SELECT ViajeDetalle.FechaHora, ViajeDetalle.IDRuta, ViajeDetalle.Orden, Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Persona, ISNULL(PersonaRecibe.Apellido + (CASE ISNULL(PersonaRecibe.Nombre, '') WHEN '' THEN '' ELSE ', ' + PersonaRecibe.Nombre END), '') + ISNULL(ViajeDetalle.Recibe, '') AS Recibe, ViajeDetalle.Descripcion, ViajeDetalle.Domicilio, ViajeDetalle.Horario, ViajeDetalle.Telefono, (CASE isnull(ViajeDetalle.Sube, '') WHEN '' THEN LugarOrigen.Nombre ELSE ViajeDetalle.Sube END) AS Origen, (CASE isnull(ViajeDetalle.Baja, '') WHEN '' THEN LugarDestino.Nombre ELSE ViajeDetalle.Baja END) AS Destino, ViajeDetalle.Importe, ViajeDetalle.ImporteContado + ViajeDetalle.ImporteCuentaCorriente AS ImportePagado, ViajeDetalle.Entregada
		FROM (((((ViajeDetalle INNER JOIN Viaje ON ViajeDetalle.FechaHora = Viaje.FechaHora AND ViajeDetalle.IDRuta = Viaje.IDRuta) INNER JOIN Persona ON ViajeDetalle.IDPersona = Persona.IDPersona) INNER JOIN Lugar AS LugarOrigen ON ViajeDetalle.IDOrigen = LugarOrigen.IDLugar) INNER JOIN Lugar AS LugarDestino ON ViajeDetalle.IDDestino = LugarDestino.IDLugar) LEFT JOIN Persona AS PersonaRecibe ON ViajeDetalle.IDPersonaRecibe = PersonaRecibe.IDPersona) LEFT JOIN ViajeDetalle_Comision ON ViajeDetalle.FechaHora = ViajeDetalle_Comision.FechaHora AND ViajeDetalle.IDRuta = ViajeDetalle_Comision.IDRuta AND ViajeDetalle.Indice = ViajeDetalle_Comision.Indice
		WHERE ViajeDetalle.OcupanteTipo = 'CO' AND ViajeDetalle.Estado = '1CO'
				AND (@FechaHoraDesde IS NULL OR ViajeDetalle.FechaHora >= @FechaHoraDesde)
				AND (@FechaHoraHasta IS NULL OR ViajeDetalle.FechaHora <= @FechaHoraHasta)
				AND (Viaje.IDRuta = @IDRuta1 OR Viaje.IDRuta = @IDRuta2 OR Viaje.IDRuta = @IDRuta3 OR Viaje.IDRuta = @IDRuta4 OR (@IDRuta1 IS NULL AND @IDRuta2 IS NULL AND @IDRuta3 IS NULL AND @IDRuta4 IS NULL))
				AND (@IDPersona IS NULL OR ViajeDetalle.IDPersona = @IDPersona)
				AND (@IDOrigen IS NULL OR ViajeDetalle.IDOrigen = @IDOrigen)
				AND (@IDDestino IS NULL OR ViajeDetalle.IDDestino = @IDDestino)
				AND (@IDListaPrecio IS NULL OR ViajeDetalle.IDListaPrecio = @IDListaPrecio)
				AND (@Entregada IS NULL OR ViajeDetalle.Entregada = @Entregada)
				AND (@Pagada IS NULL OR (@Pagada = 1 AND ViajeDetalle.ImporteContado + ViajeDetalle.ImporteCuentaCorriente > 0) OR (@Pagada = 0 AND ViajeDetalle.ImporteContado + ViajeDetalle.ImporteCuentaCorriente = 0))
				AND (@RendicionVacia IS NULL OR (@RendicionVacia = 0 AND ViajeDetalle_Comision.RendicionFechaHora IS NOT NULL) OR (@RendicionVacia = 1 AND ViajeDetalle_Comision.RendicionFechaHora IS NULL))
				AND (@RendicionFechaHoraDesde IS NULL OR ViajeDetalle_Comision.RendicionFechaHora >= @RendicionFechaHoraDesde)
				AND (@RendicionFechaHoraHasta IS NULL OR ViajeDetalle_Comision.RendicionFechaHora <= @RendicionFechaHoraHasta)
				AND (@MostrarTodas = 1 OR Viaje.FechaHora BETWEEN getdate() - 30 AND getdate() + 7)
				AND (@Personal = 0 OR Viaje.Personal IS NULL OR Viaje.Personal = 0)
		ORDER BY ViajeDetalle.FechaHora, ViajeDetalle.IDRuta
GO



------------------------------------------------------------------------------------------
-- REPORT_COMISION_REMITO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Report_Comision_Remito'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_Comision_Remito
GO

CREATE PROCEDURE dbo.sp_Report_Comision_Remito
	@FechaHora_FILTER smalldatetime,
	@IDRuta_FILTER char(20),
	@Indice_FILTER int AS

	(SELECT 1 AS IDUnique, ViajeDetalle.FechaHoraCreacion AS FechaRecibida, ViajeDetalle.EntregadaFechaHora AS FechaEntregada, Conductor.Apellido + (CASE ISNULL(Conductor.Nombre, '') WHEN '' THEN '' ELSE ', ' + Conductor.Nombre END) AS Conductor, ViajeDetalle.FechaHora, Remitente.Apellido + (CASE ISNULL(Remitente.Nombre, '') WHEN '' THEN '' ELSE ', ' + Remitente.Nombre END) AS Remitente, ISNULL(PersonaRecibe.Apellido + (CASE ISNULL(PersonaRecibe.Nombre, '') WHEN '' THEN '' ELSE ', ' + PersonaRecibe.Nombre END), '') + ISNULL(ViajeDetalle.Recibe, '') AS Recibe, ViajeDetalle.PagaQuienRecibe, (CASE ISNULL(ListaPrecio_Ruta.RutaNombre, '') WHEN '' THEN Ruta.Nombre ELSE ListaPrecio_Ruta.RutaNombre END) AS Ruta, (CASE ISNULL(ListaPrecio_Ruta.OrigenNombre, '') WHEN '' THEN (CASE ISNULL(ViajeDetalle.Sube, '') WHEN '' THEN LugarOrigen.Nombre ELSE ViajeDetalle.Sube END) ELSE ListaPrecio_Ruta.OrigenNombre END) AS Origen, (CASE ISNULL(ListaPrecio_Ruta.DestinoNombre, '') WHEN '' THEN (CASE ISNULL(ViajeDetalle.Baja, '') WHEN '' THEN LugarDestino.Nombre ELSE ViajeDetalle.Baja END) ELSE ListaPrecio_Ruta.DestinoNombre END) AS Destino, ViajeDetalle.Descripcion, ViajeDetalle.Domicilio, ViajeDetalle.Horario, ViajeDetalle.Telefono, (CASE ISNULL(ViajeDetalle.DejarTraer, '') WHEN '' THEN '' WHEN 'D' THEN 'DEJAR' WHEN 'T' THEN 'TRAER' END) AS DejarTraer, ViajeDetalle.ValorDeclarado, ViajeDetalle.Importe, ViajeDetalle.Notas
		FROM (((((((ViajeDetalle INNER JOIN Ruta ON ViajeDetalle.IDRuta = Ruta.IDRuta) INNER JOIN Lugar AS LugarOrigen ON ViajeDetalle.IDOrigen = LugarOrigen.IDLugar) INNER JOIN Lugar AS LugarDestino ON ViajeDetalle.IDDestino = LugarDestino.IDLugar) INNER JOIN Persona AS Remitente ON ViajeDetalle.IDPersona = Remitente.IDPersona) INNER JOIN Viaje ON ViajeDetalle.FechaHora = Viaje.FechaHora AND ViajeDetalle.IDRuta = Viaje.IDRuta) LEFT JOIN Persona AS Conductor ON Viaje.IDConductor = Conductor.IDPersona) LEFT JOIN Persona AS PersonaRecibe ON ViajeDetalle.IDPersonaRecibe = PersonaRecibe.IDPersona) LEFT JOIN ListaPrecio_Ruta ON ViajeDetalle.IDListaPrecio = ListaPrecio_Ruta.IDListaPrecio AND ViajeDetalle.IDRuta = ListaPrecio_Ruta.IDRuta
		WHERE ViajeDetalle.FechaHora = @FechaHora_FILTER AND ViajeDetalle.IDRuta = @IDRuta_FILTER AND ViajeDetalle.Indice = @Indice_FILTER)
	UNION
	(SELECT 3 AS IDUnique, ViajeDetalle.FechaHoraCreacion AS FechaRecibida, ViajeDetalle.EntregadaFechaHora AS FechaEntregada, Conductor.Apellido + (CASE ISNULL(Conductor.Nombre, '') WHEN '' THEN '' ELSE ', ' + Conductor.Nombre END) AS Conductor, ViajeDetalle.FechaHora, Remitente.Apellido + (CASE ISNULL(Remitente.Nombre, '') WHEN '' THEN '' ELSE ', ' + Remitente.Nombre END) AS Remitente, ISNULL(PersonaRecibe.Apellido + (CASE ISNULL(PersonaRecibe.Nombre, '') WHEN '' THEN '' ELSE ', ' + PersonaRecibe.Nombre END), '') + ISNULL(ViajeDetalle.Recibe, '') AS Recibe, ViajeDetalle.PagaQuienRecibe, (CASE ISNULL(ListaPrecio_Ruta.RutaNombre, '') WHEN '' THEN Ruta.Nombre ELSE ListaPrecio_Ruta.RutaNombre END) AS Ruta, (CASE ISNULL(ListaPrecio_Ruta.OrigenNombre, '') WHEN '' THEN (CASE ISNULL(ViajeDetalle.Sube, '') WHEN '' THEN LugarOrigen.Nombre ELSE ViajeDetalle.Sube END) ELSE ListaPrecio_Ruta.OrigenNombre END) AS Origen, (CASE ISNULL(ListaPrecio_Ruta.DestinoNombre, '') WHEN '' THEN (CASE ISNULL(ViajeDetalle.Baja, '') WHEN '' THEN LugarDestino.Nombre ELSE ViajeDetalle.Baja END) ELSE ListaPrecio_Ruta.DestinoNombre END) AS Destino, ViajeDetalle.Descripcion, ViajeDetalle.Domicilio, ViajeDetalle.Horario, ViajeDetalle.Telefono, (CASE ISNULL(ViajeDetalle.DejarTraer, '') WHEN '' THEN '' WHEN 'D' THEN 'DEJAR' WHEN 'T' THEN 'TRAER' END) AS DejarTraer, ViajeDetalle.ValorDeclarado, ViajeDetalle.Importe, ViajeDetalle.Notas
		FROM (((((((ViajeDetalle INNER JOIN Ruta ON ViajeDetalle.IDRuta = Ruta.IDRuta) INNER JOIN Lugar AS LugarOrigen ON ViajeDetalle.IDOrigen = LugarOrigen.IDLugar) INNER JOIN Lugar AS LugarDestino ON ViajeDetalle.IDDestino = LugarDestino.IDLugar) INNER JOIN Persona AS Remitente ON ViajeDetalle.IDPersona = Remitente.IDPersona) INNER JOIN Viaje ON ViajeDetalle.FechaHora = Viaje.FechaHora AND ViajeDetalle.IDRuta = Viaje.IDRuta) LEFT JOIN Persona AS Conductor ON Viaje.IDConductor = Conductor.IDPersona) LEFT JOIN Persona AS PersonaRecibe ON ViajeDetalle.IDPersonaRecibe = PersonaRecibe.IDPersona) LEFT JOIN ListaPrecio_Ruta ON ViajeDetalle.IDListaPrecio = ListaPrecio_Ruta.IDListaPrecio AND ViajeDetalle.IDRuta = ListaPrecio_Ruta.IDRuta
		WHERE ViajeDetalle.FechaHora = @FechaHora_FILTER AND ViajeDetalle.IDRuta = @IDRuta_FILTER AND ViajeDetalle.Indice = @Indice_FILTER)

GO



------------------------------------------------------------------------------------------
-- REPORT_COMISION_REMITO_BLANCO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Report_Comision_Remito_Blanco'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_Comision_Remito_Blanco
GO

CREATE PROCEDURE dbo.sp_Report_Comision_Remito_Blanco AS

	(SELECT 1 AS IDUnique)
	UNION
	(SELECT 3 AS IDUnique)

GO



------------------------------------------------------------------------------------------
-- REPORT_VIAJE_ESTADOVENCIDO_LIST
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Report_Viaje_EstadoVencido_List' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_Viaje_EstadoVencido_List
GO

CREATE PROCEDURE dbo.sp_Report_Viaje_EstadoVencido_List
	@Personal bit AS

	DECLARE @IDRutaEspecial char(20)
	DECLARE @Horas_Normal int
	DECLARE @Horas_Especial int

	SET @IDRutaEspecial = (SELECT Texto FROM Parametro WHERE IDParametro = 'Ruta_ID_Otra')
	SET @Horas_Normal = (SELECT NumeroEntero FROM Parametro WHERE IDParametro = 'Viaje_Normal_Estado_Vencido_Horas')
	SET @Horas_Especial = (SELECT NumeroEntero FROM Parametro WHERE IDParametro = 'Viaje_Especial_Estado_Vencido_Horas')

	SELECT (CASE datepart(Weekday, Viaje.FechaHora) WHEN 1 THEN 'Domingo' WHEN 2 THEN 'Lunes' WHEN 3 THEN 'Martes' WHEN 4 THEN 'Miércoles' WHEN 5 THEN 'Jueves' WHEN 6 THEN 'Viernes' WHEN 7 THEN 'Sábado' END) AS DiaSemana, Viaje.FechaHora, RTRIM(Viaje.IDRuta) + (CASE Viaje.IDRuta WHEN @IDRutaEspecial THEN ': ' + Viaje.RutaOtra ELSE '' END) AS Ruta, Vehiculo.Nombre AS Vehiculo, (CASE ISNULL(Viaje.IDConductor, 1) WHEN 1 THEN '' ELSE Persona.Apellido + ', ' + Persona.Nombre END) AS Conductor, (CASE Viaje.Estado WHEN 'AC' THEN 'Activo' WHEN 'EP' Then 'En Progreso' WHEN 'FI' THEN 'Finalizado' WHEN 'CA' THEN 'Cancelado' END) AS Estado, (CASE Viaje.Charter WHEN 0 THEN '' WHEN 1 THEN 'Sí' END) AS CharterDisplay
		FROM (Viaje LEFT JOIN Persona ON Viaje.IDConductor = Persona.IDPersona) LEFT JOIN Vehiculo ON Viaje.IDVehiculo = Vehiculo.IDVehiculo
		WHERE (Viaje.Estado = 'AC' OR Viaje.Estado = 'EP')
			AND ((Viaje.FechaHora <= DATEADD(Hour, -@Horas_Normal, getdate()) AND Viaje.IDRuta <> @IDRutaEspecial) OR (Viaje.FechaHora <= DATEADD(Hour, -@Horas_Especial, getdate()) AND Viaje.IDRuta = @IDRutaEspecial))
			AND (@Personal = 0 OR Viaje.Personal = 0)
		ORDER BY Viaje.FechaHora, Viaje.IDRuta, Viaje.RutaOtra

GO



------------------------------------------------------------------------------------------
-- REPORT_PERSONA_CANTIDADVIAJES
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Report_Persona_CantidadViajes'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_Persona_CantidadViajes
GO

CREATE PROCEDURE dbo.sp_Report_Persona_CantidadViajes
	@DiaSemana_FILTER int,
	@FechaHoraDesde_FILTER smalldatetime,
	@FechaHoraHasta_FILTER smalldatetime,
	@FechaDesde_FILTER smalldatetime,
	@FechaHasta_FILTER smalldatetime,
	@HoraDesde_FILTER smalldatetime,
	@HoraHasta_FILTER smalldatetime,
	@IDRuta_FILTER char(20),
	@IDPersona_FILTER int,
	@Personal bit AS

	SELECT TOP 90 Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Persona, COUNT(ViajeDetalle.FechaHora) AS Viajes
		FROM (Persona INNER JOIN ViajeDetalle ON Persona.IDPersona = ViajeDetalle.IDPersona) INNER JOIN Viaje ON ViajeDetalle.FechaHora = Viaje.FechaHora AND ViajeDetalle.IDRuta = Viaje.IDRuta
		WHERE Viaje.Estado <> 'CA'
			AND ViajeDetalle.Estado = '1CO'
			AND ViajeDetalle.OcupanteTipo = 'PA'
			AND ViajeDetalle.Realizado = 1
			AND (@DiaSemana_FILTER IS NULL OR datepart(weekday, Viaje.FechaHora) = @DiaSemana_FILTER)
			AND (@FechaHoraDesde_FILTER IS NULL OR Viaje.FechaHora >= @FechaHoraDesde_FILTER)
			AND (@FechaHoraHasta_FILTER IS NULL OR Viaje.FechaHora <= @FechaHoraHasta_FILTER)
			AND (@FechaDesde_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) >= convert(char(10), @FechaDesde_FILTER, 111))
			AND (@FechaHasta_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) <= convert(char(10), @FechaHasta_FILTER, 111))
			AND (@HoraDesde_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) >= convert(char(8), @HoraDesde_FILTER, 108))
			AND (@HoraHasta_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) <= convert(char(8), @HoraHasta_FILTER, 108))
			AND (@IDRuta_FILTER IS NULL OR Viaje.IDRuta = @IDRuta_FILTER)
			AND (@IDPersona_FILTER IS NULL OR ViajeDetalle.IDPersona = @IDPersona_FILTER)
			AND (@Personal = 0 OR Viaje.Personal = 0)
		GROUP BY Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END)
		ORDER BY Viajes DESC, Persona

GO



------------------------------------------------------------------------------------------
-- REPORT_PERSONA_CANTIDADVIAJESCANCELADOS
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Report_Persona_CantidadViajesCancelados'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_Persona_CantidadViajesCancelados
GO

CREATE PROCEDURE dbo.sp_Report_Persona_CantidadViajesCancelados
	@DiaSemana_FILTER int,
	@FechaHoraDesde_FILTER smalldatetime,
	@FechaHoraHasta_FILTER smalldatetime,
	@FechaDesde_FILTER smalldatetime,
	@FechaHasta_FILTER smalldatetime,
	@HoraDesde_FILTER smalldatetime,
	@HoraHasta_FILTER smalldatetime,
	@IDRuta_FILTER char(20),
	@IDPersona_FILTER int,
	@Personal bit AS

	SELECT TOP 50 Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Persona, COUNT(ViajeDetalle.FechaHora) AS Viajes
		FROM (Persona INNER JOIN ViajeDetalle ON Persona.IDPersona = ViajeDetalle.IDPersona) INNER JOIN Viaje ON ViajeDetalle.FechaHora = Viaje.FechaHora AND ViajeDetalle.IDRuta = Viaje.IDRuta
		WHERE Viaje.Estado <> 'CA'
			AND ViajeDetalle.Estado = '3CA'
			AND ViajeDetalle.OcupanteTipo = 'PA'
			AND (@DiaSemana_FILTER IS NULL OR datepart(weekday, Viaje.FechaHora) = @DiaSemana_FILTER)
			AND (@FechaHoraDesde_FILTER IS NULL OR Viaje.FechaHora >= @FechaHoraDesde_FILTER)
			AND (@FechaHoraHasta_FILTER IS NULL OR Viaje.FechaHora <= @FechaHoraHasta_FILTER)
			AND (@FechaDesde_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) >= convert(char(10), @FechaDesde_FILTER, 111))
			AND (@FechaHasta_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) <= convert(char(10), @FechaHasta_FILTER, 111))
			AND (@HoraDesde_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) >= convert(char(8), @HoraDesde_FILTER, 108))
			AND (@HoraHasta_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) <= convert(char(8), @HoraHasta_FILTER, 108))
			AND (@IDRuta_FILTER IS NULL OR Viaje.IDRuta = @IDRuta_FILTER)
			AND (@IDPersona_FILTER IS NULL OR ViajeDetalle.IDPersona = @IDPersona_FILTER)
			AND (@Personal = 0 OR Viaje.Personal = 0)
		GROUP BY Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END)
		ORDER BY Viajes DESC, Persona

GO



------------------------------------------------------------------------------------------
-- REPORT_PERSONA_CANTIDADVIAJESNOREALIZADOS
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Report_Persona_CantidadViajesNoRealizados'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_Persona_CantidadViajesNoRealizados
GO

CREATE PROCEDURE dbo.sp_Report_Persona_CantidadViajesNoRealizados
	@DiaSemana_FILTER int,
	@FechaHoraDesde_FILTER smalldatetime,
	@FechaHoraHasta_FILTER smalldatetime,
	@FechaDesde_FILTER smalldatetime,
	@FechaHasta_FILTER smalldatetime,
	@HoraDesde_FILTER smalldatetime,
	@HoraHasta_FILTER smalldatetime,
	@IDRuta_FILTER char(20),
	@IDPersona_FILTER int,
	@Personal bit AS

	SELECT TOP 90 Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Persona, COUNT(ViajeDetalle.FechaHora) AS Viajes
		FROM (Persona INNER JOIN ViajeDetalle ON Persona.IDPersona = ViajeDetalle.IDPersona) INNER JOIN Viaje ON ViajeDetalle.FechaHora = Viaje.FechaHora AND ViajeDetalle.IDRuta = Viaje.IDRuta
		WHERE Viaje.Estado <> 'CA'
			AND ViajeDetalle.Estado = '1CO'
			AND ViajeDetalle.OcupanteTipo = 'PA'
			AND ViajeDetalle.Realizado = 0
			AND (@DiaSemana_FILTER IS NULL OR datepart(weekday, Viaje.FechaHora) = @DiaSemana_FILTER)
			AND (@FechaHoraDesde_FILTER IS NULL OR Viaje.FechaHora >= @FechaHoraDesde_FILTER)
			AND (@FechaHoraHasta_FILTER IS NULL OR Viaje.FechaHora <= @FechaHoraHasta_FILTER)
			AND (@FechaDesde_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) >= convert(char(10), @FechaDesde_FILTER, 111))
			AND (@FechaHasta_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) <= convert(char(10), @FechaHasta_FILTER, 111))
			AND (@HoraDesde_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) >= convert(char(8), @HoraDesde_FILTER, 108))
			AND (@HoraHasta_FILTER IS NULL OR convert(char(8), Viaje.FechaHora, 108) <= convert(char(8), @HoraHasta_FILTER, 108))
			AND (@IDRuta_FILTER IS NULL OR Viaje.IDRuta = @IDRuta_FILTER)
			AND (@IDPersona_FILTER IS NULL OR ViajeDetalle.IDPersona = @IDPersona_FILTER)
			AND (@Personal = 0 OR Viaje.Personal = 0)
		GROUP BY Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END)
		ORDER BY Viajes DESC, Persona

GO





------------------------------------------------------------------------------------------
-- REPORT_TABLEROCONTROL
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects
	   WHERE  name = N'sp_Report_TableroControl'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_TableroControl
GO

CREATE PROCEDURE dbo.sp_Report_TableroControl
	@FechaDesde smalldatetime,
	@FechaHasta smalldatetime,
	@OcupanteTipo char(2),
	@Realizado bit,
	@DividirViajes bit,
	@Personal bit AS

	SELECT Viajes.Anio, Viajes.Mes, dbo.udf_Date_GetLastDayOfMonth(Viajes.Mes, Viajes.Anio) AS DiasDelMes, IsNull(Ocupantes.Ocupantes, 0) AS Ocupantes, IsNull(Ocupantes.ImporteTotal, 0) + IsNull(Viajes.Importe, 0) AS ImporteTotal, (CASE ISNULL(@DividirViajes, 0) WHEN 0 THEN Viajes.Viajes WHEN 1 THEN Viajes.Viajes / 2 END) AS Viajes, IsNull(Combustible.LitrosCombustible, 0) AS LitrosCombustible
		FROM
			(SELECT DatePart(year, Viaje.FechaHora) AS Anio, DatePart(month, Viaje.FechaHora) AS Mes, CONVERT(FLOAT, COUNT(CONVERT(CHAR(19), Viaje.FechaHora, 120) + Viaje.IDRuta)) AS Viajes, SUM(CASE ISNULL(@OcupanteTipo, 'NU') WHEN 'CO' THEN 0 ELSE Viaje.Importe END) AS Importe
				FROM Viaje
				WHERE Viaje.Estado <> 'CA'
					AND (Viaje.FechaHora >= @FechaDesde AND Viaje.FechaHora <= @FechaHasta)
					AND (@Personal = 0 OR Viaje.Personal = 0)
				GROUP BY DatePart(year, Viaje.FechaHora), DatePart(month, Viaje.FechaHora)) AS Viajes
			LEFT JOIN
			(SELECT DatePart(year, Viaje.FechaHora) AS Anio, DatePart(month, Viaje.FechaHora) AS Mes, Count(ViajeDetalle.Indice) AS Ocupantes, Sum(ViajeDetalle.Importe) AS ImporteTotal
				FROM ((((ViajeDetalle INNER JOIN Viaje ON ViajeDetalle.FechaHora = Viaje.FechaHora AND ViajeDetalle.IDRuta = Viaje.IDRuta) INNER JOIN Lugar AS LugarOrigen ON ViajeDetalle.IDOrigen = LugarOrigen.IDLugar) INNER JOIN Lugar AS LugarDestino ON ViajeDetalle.IDDestino = LugarDestino.IDLugar) INNER JOIN RutaDetalle AS RutaDetalleOrigen ON ViajeDetalle.IDRuta = RutaDetalleOrigen.IDRuta AND ViajeDetalle.IDOrigen = RutaDetalleOrigen.IDLugar) INNER JOIN RutaDetalle AS RutaDetalleDestino ON ViajeDetalle.IDRuta = RutaDetalleDestino.IDRuta AND ViajeDetalle.IDDestino = RutaDetalleDestino.IDLugar
				WHERE (@OcupanteTipo IS NULL OR ViajeDetalle.OcupanteTipo = @OcupanteTipo)
					AND ViajeDetalle.Estado = '1CO'
					AND (ViajeDetalle.OcupanteTipo = 'CO' OR ViajeDetalle.Realizado IS NULL OR ViajeDetalle.Realizado = 1 OR ViajeDetalle.ForzarDebito = 1)
					AND Viaje.Estado <> 'CA'
					AND (Viaje.FechaHora >= @FechaDesde AND Viaje.FechaHora <= @FechaHasta)
					AND (@Realizado IS NULL OR ViajeDetalle.Realizado = @Realizado OR ViajeDetalle.OcupanteTipo = 'CO')
					AND (@Personal = 0 OR Viaje.Personal = 0)
				GROUP BY DatePart(year, Viaje.FechaHora), DatePart(month, Viaje.FechaHora)) AS Ocupantes
			ON Viajes.Anio = Ocupantes.Anio AND Viajes.Mes = Ocupantes.Mes
			LEFT JOIN
			(SELECT DatePart(year, VehiculoMantenimientoAccion.FechaHora) AS Anio, DatePart(month, VehiculoMantenimientoAccion.FechaHora) AS Mes, Sum(VehiculoMantenimientoAccion.Litros) AS LitrosCombustible
				FROM VehiculoMantenimientoAccion
				WHERE (VehiculoMantenimientoAccion.IDVehiculoMantenimientoGrupo = 2 OR VehiculoMantenimientoAccion.IDVehiculoMantenimientoGrupo = 63)
					AND (VehiculoMantenimientoAccion.FechaHora >= @FechaDesde AND VehiculoMantenimientoAccion.FechaHora <= @FechaHasta)
				GROUP BY DatePart(year, VehiculoMantenimientoAccion.FechaHora), DatePart(month, VehiculoMantenimientoAccion.FechaHora)) AS Combustible
			ON Viajes.Anio = Combustible.Anio AND Viajes.Mes = Combustible.Mes
		ORDER BY Viajes.Anio, Viajes.Mes

GO



------------------------------------------------------------------------------------------
-- REPORT_CUENTACORRIENTE_SALDO_PORSALDO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Report_CuentaCorriente_Saldo_PorSaldo'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_CuentaCorriente_Saldo_PorSaldo
GO

CREATE PROCEDURE dbo.sp_Report_CuentaCorriente_Saldo_PorSaldo
	@ImporteDesde smallmoney,
	@ImporteHasta smallmoney,
	@FechaDesde smalldatetime,
	@FechaHasta smalldatetime,
	@FechaUltimoMovimientoDesde smalldatetime,
	@FechaUltimoMovimientoHasta smalldatetime,
	@PersonaTipo char(2),
	@Personal bit AS

	SET CONCAT_NULL_YIELDS_NULL ON

	SELECT Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Persona, Localidad.Nombre AS Localidad, (CASE ISNULL(Persona.Telefono1Area, '02227') WHEN '02227' THEN '' ELSE '(' + Persona.Telefono1Area + ') ' END) + ISNULL(TelefonoTipo.DiscadoPrefijo + '-', '') + Persona.Telefono1Numero + ISNULL('-' + TelefonoTipo.DiscadoSufijo, '') AS Telefono, CuentaCorriente.Saldo, CuentaCorriente.FechaUltimoMovimiento
		FROM (Persona INNER JOIN 
			(SELECT CuentaCorriente.IDPersona, SUM(CuentaCorriente.Importe) AS Saldo, MAX(CuentaCorriente.FechaHora) AS FechaUltimoMovimiento
				FROM (((CuentaCorriente LEFT JOIN ViajeDetalle AS ViajeDetalleDebito ON CuentaCorriente.Viaje_FechaHora = ViajeDetalleDebito.FechaHora AND CuentaCorriente.Viaje_IDRuta = ViajeDetalleDebito.IDRuta AND CuentaCorriente.Viaje_Indice = ViajeDetalleDebito.Indice) LEFT JOIN Viaje AS ViajeDebito ON ViajeDetalleDebito.FechaHora = ViajeDebito.FechaHora AND ViajeDetalleDebito.IDRuta = ViajeDebito.IDRuta) LEFT JOIN ViajeDetalle AS ViajeDetalleCredito ON CuentaCorriente.Viaje_FechaHora = ViajeDetalleCredito.FechaHora AND CuentaCorriente.Viaje_IDRuta = ViajeDetalleCredito.IDRuta AND CuentaCorriente.Viaje_Indice = ViajeDetalleCredito.Indice) LEFT JOIN Viaje AS ViajeCredito ON ViajeDetalleCredito.FechaHora = ViajeCredito.FechaHora AND ViajeDetalleCredito.IDRuta = ViajeCredito.IDRuta
				WHERE (@FechaDesde IS NULL OR convert(char(10), CuentaCorriente.FechaHora, 111) >= convert(char(10), @FechaDesde, 111))
					AND (@FechaHasta IS NULL OR convert(char(10), CuentaCorriente.FechaHora, 111) <= convert(char(10), @FechaHasta, 111))
					AND (@Personal = 0 OR ViajeDebito.Personal IS NULL OR ViajeDebito.Personal = 0)
					AND (@Personal = 0 OR ViajeCredito.Personal IS NULL OR ViajeCredito.Personal = 0)
				GROUP BY CuentaCorriente.IDPersona
				HAVING SUM(CuentaCorriente.Importe) <> 0
					AND (@ImporteDesde IS NULL OR SUM(CuentaCorriente.Importe) >= @ImporteDesde)
					AND (@ImporteHasta IS NULL OR SUM(CuentaCorriente.Importe) <= @ImporteHasta)
					AND (@FechaUltimoMovimientoDesde IS NULL OR convert(char(10), MAX(CuentaCorriente.FechaHora), 111) >= convert(char(10), @FechaUltimoMovimientoDesde, 111))
					AND (@FechaUltimoMovimientoHasta IS NULL OR convert(char(10), MAX(CuentaCorriente.FechaHora), 111) <= convert(char(10), @FechaUltimoMovimientoHasta, 111))
			) AS CuentaCorriente ON Persona.IDPersona = CuentaCorriente.IDPersona)
			LEFT JOIN Localidad ON Persona.IDProvincia = Localidad.IDProvincia AND Persona.IDLocalidad = Localidad.IDLocalidad
			LEFT JOIN TelefonoTipo ON Persona.IDTelefono1Tipo = TelefonoTipo.IDTelefonoTipo
		WHERE (@PersonaTipo IS NULL OR Persona.EntidadTipo = @PersonaTipo)
		ORDER BY Saldo, Persona

	SET CONCAT_NULL_YIELDS_NULL OFF

GO



------------------------------------------------------------------------------------------
-- REPORT_CUENTACORRIENTE_SALDO_PORLOCALIDAD
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Report_CuentaCorriente_Saldo_PorLocalidad'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_CuentaCorriente_Saldo_PorLocalidad
GO

CREATE PROCEDURE dbo.sp_Report_CuentaCorriente_Saldo_PorLocalidad
	@ImporteDesde smallmoney,
	@ImporteHasta smallmoney,
	@FechaDesde smalldatetime,
	@FechaHasta smalldatetime,
	@FechaUltimoMovimientoDesde smalldatetime,
	@FechaUltimoMovimientoHasta smalldatetime,
	@PersonaTipo char(2),
	@Personal bit AS

	SET CONCAT_NULL_YIELDS_NULL ON

	SELECT Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Persona, Localidad.Nombre AS Localidad, (CASE ISNULL(Persona.Telefono1Area, '02227') WHEN '02227' THEN '' ELSE '(' + Persona.Telefono1Area + ') ' END) + ISNULL(TelefonoTipo.DiscadoPrefijo + '-', '') + Persona.Telefono1Numero + ISNULL('-' + TelefonoTipo.DiscadoSufijo, '') AS Telefono, CuentaCorriente.Saldo, CuentaCorriente.FechaUltimoMovimiento
		FROM (Persona INNER JOIN 
			(SELECT CuentaCorriente.IDPersona, SUM(CuentaCorriente.Importe) AS Saldo, MAX(CuentaCorriente.FechaHora) AS FechaUltimoMovimiento
				FROM (((CuentaCorriente LEFT JOIN ViajeDetalle AS ViajeDetalleDebito ON CuentaCorriente.Viaje_FechaHora = ViajeDetalleDebito.FechaHora AND CuentaCorriente.Viaje_IDRuta = ViajeDetalleDebito.IDRuta AND CuentaCorriente.Viaje_Indice = ViajeDetalleDebito.Indice) LEFT JOIN Viaje AS ViajeDebito ON ViajeDetalleDebito.FechaHora = ViajeDebito.FechaHora AND ViajeDetalleDebito.IDRuta = ViajeDebito.IDRuta) LEFT JOIN ViajeDetalle AS ViajeDetalleCredito ON CuentaCorriente.Viaje_FechaHora = ViajeDetalleCredito.FechaHora AND CuentaCorriente.Viaje_IDRuta = ViajeDetalleCredito.IDRuta AND CuentaCorriente.Viaje_Indice = ViajeDetalleCredito.Indice) LEFT JOIN Viaje AS ViajeCredito ON ViajeDetalleCredito.FechaHora = ViajeCredito.FechaHora AND ViajeDetalleCredito.IDRuta = ViajeCredito.IDRuta
				WHERE (@FechaDesde IS NULL OR convert(char(10), CuentaCorriente.FechaHora, 111) >= convert(char(10), @FechaDesde, 111))
					AND (@FechaHasta IS NULL OR convert(char(10), CuentaCorriente.FechaHora, 111) <= convert(char(10), @FechaHasta, 111))
					AND (@Personal = 0 OR ViajeDebito.Personal IS NULL OR ViajeDebito.Personal = 0)
					AND (@Personal = 0 OR ViajeCredito.Personal IS NULL OR ViajeCredito.Personal = 0)
				GROUP BY CuentaCorriente.IDPersona
				HAVING SUM(CuentaCorriente.Importe) <> 0
					AND (@ImporteDesde IS NULL OR SUM(CuentaCorriente.Importe) >= @ImporteDesde)
					AND (@ImporteHasta IS NULL OR SUM(CuentaCorriente.Importe) <= @ImporteHasta)
					AND (@FechaUltimoMovimientoDesde IS NULL OR convert(char(10), MAX(CuentaCorriente.FechaHora), 111) >= convert(char(10), @FechaUltimoMovimientoDesde, 111))
					AND (@FechaUltimoMovimientoHasta IS NULL OR convert(char(10), MAX(CuentaCorriente.FechaHora), 111) <= convert(char(10), @FechaUltimoMovimientoHasta, 111))
			) AS CuentaCorriente ON Persona.IDPersona = CuentaCorriente.IDPersona)
			LEFT JOIN Localidad ON Persona.IDProvincia = Localidad.IDProvincia AND Persona.IDLocalidad = Localidad.IDLocalidad
			LEFT JOIN TelefonoTipo ON Persona.IDTelefono1Tipo = TelefonoTipo.IDTelefonoTipo
		WHERE (@PersonaTipo IS NULL OR Persona.EntidadTipo = @PersonaTipo)
		ORDER BY Localidad, Saldo, Persona

	SET CONCAT_NULL_YIELDS_NULL OFF

GO



------------------------------------------------------------------------------------------
-- REPORT_CUENTACORRIENTE_SALDO_PORFECHA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Report_CuentaCorriente_Saldo_PorFecha'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_CuentaCorriente_Saldo_PorFecha
GO

CREATE PROCEDURE dbo.sp_Report_CuentaCorriente_Saldo_PorFecha
	@ImporteDesde smallmoney,
	@ImporteHasta smallmoney,
	@FechaDesde smalldatetime,
	@FechaHasta smalldatetime,
	@FechaUltimoMovimientoDesde smalldatetime,
	@FechaUltimoMovimientoHasta smalldatetime,
	@PersonaTipo char(2),
	@Personal bit AS

	SET CONCAT_NULL_YIELDS_NULL ON

	SELECT Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Persona, Localidad.Nombre AS Localidad, (CASE ISNULL(Persona.Telefono1Area, '02227') WHEN '02227' THEN '' ELSE '(' + Persona.Telefono1Area + ') ' END) + ISNULL(TelefonoTipo.DiscadoPrefijo + '-', '') + Persona.Telefono1Numero + ISNULL('-' + TelefonoTipo.DiscadoSufijo, '') AS Telefono, CuentaCorriente.Saldo, CuentaCorriente.FechaUltimoMovimiento
		FROM (Persona INNER JOIN 
			(SELECT CuentaCorriente.IDPersona, SUM(CuentaCorriente.Importe) AS Saldo, MAX(CuentaCorriente.FechaHora) AS FechaUltimoMovimiento
				FROM (((CuentaCorriente LEFT JOIN ViajeDetalle AS ViajeDetalleDebito ON CuentaCorriente.Viaje_FechaHora = ViajeDetalleDebito.FechaHora AND CuentaCorriente.Viaje_IDRuta = ViajeDetalleDebito.IDRuta AND CuentaCorriente.Viaje_Indice = ViajeDetalleDebito.Indice) LEFT JOIN Viaje AS ViajeDebito ON ViajeDetalleDebito.FechaHora = ViajeDebito.FechaHora AND ViajeDetalleDebito.IDRuta = ViajeDebito.IDRuta) LEFT JOIN ViajeDetalle AS ViajeDetalleCredito ON CuentaCorriente.Viaje_FechaHora = ViajeDetalleCredito.FechaHora AND CuentaCorriente.Viaje_IDRuta = ViajeDetalleCredito.IDRuta AND CuentaCorriente.Viaje_Indice = ViajeDetalleCredito.Indice) LEFT JOIN Viaje AS ViajeCredito ON ViajeDetalleCredito.FechaHora = ViajeCredito.FechaHora AND ViajeDetalleCredito.IDRuta = ViajeCredito.IDRuta
				WHERE (@FechaDesde IS NULL OR convert(char(10), CuentaCorriente.FechaHora, 111) >= convert(char(10), @FechaDesde, 111))
					AND (@FechaHasta IS NULL OR convert(char(10), CuentaCorriente.FechaHora, 111) <= convert(char(10), @FechaHasta, 111))
					AND (@Personal = 0 OR ViajeDebito.Personal IS NULL OR ViajeDebito.Personal = 0)
					AND (@Personal = 0 OR ViajeCredito.Personal IS NULL OR ViajeCredito.Personal = 0)
				GROUP BY CuentaCorriente.IDPersona
				HAVING SUM(CuentaCorriente.Importe) <> 0
					AND (@ImporteDesde IS NULL OR SUM(CuentaCorriente.Importe) >= @ImporteDesde)
					AND (@ImporteHasta IS NULL OR SUM(CuentaCorriente.Importe) <= @ImporteHasta)
					AND (@FechaUltimoMovimientoDesde IS NULL OR convert(char(10), MAX(CuentaCorriente.FechaHora), 111) >= convert(char(10), @FechaUltimoMovimientoDesde, 111))
					AND (@FechaUltimoMovimientoHasta IS NULL OR convert(char(10), MAX(CuentaCorriente.FechaHora), 111) <= convert(char(10), @FechaUltimoMovimientoHasta, 111))
			) AS CuentaCorriente ON Persona.IDPersona = CuentaCorriente.IDPersona)
			LEFT JOIN Localidad ON Persona.IDProvincia = Localidad.IDProvincia AND Persona.IDLocalidad = Localidad.IDLocalidad
			LEFT JOIN TelefonoTipo ON Persona.IDTelefono1Tipo = TelefonoTipo.IDTelefonoTipo
		WHERE (@PersonaTipo IS NULL OR Persona.EntidadTipo = @PersonaTipo)
		ORDER BY FechaUltimoMovimiento, Saldo, Persona

	SET CONCAT_NULL_YIELDS_NULL OFF

GO



------------------------------------------------------------------------------------------
-- REPORT_CUENTACORRIENTE_SALDOPORGRUPO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects
	   WHERE  name = N'sp_Report_CuentaCorriente_SaldoPorGrupo'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_CuentaCorriente_SaldoPorGrupo
GO

CREATE PROCEDURE dbo.sp_Report_CuentaCorriente_SaldoPorGrupo
	@FechaDesde smalldatetime,
	@FechaHasta smalldatetime,
	@Personal bit AS

	SELECT CuentaCorrienteGrupo.Nombre AS Grupo, ISNULL(Ingresos.Importe, 0) AS Ingresos, ABS(ISNULL(Egresos.Importe, 0)) AS Egresos, ISNULL(Ingresos.Importe, 0) - ABS(ISNULL(Egresos.Importe, 0)) AS Resultado
		FROM CuentaCorrienteGrupo
			LEFT JOIN (SELECT CuentaCorriente.IDCuentaCorrienteGrupo, Sum(CuentaCorriente.Importe) AS Importe
				FROM (((CuentaCorriente LEFT JOIN ViajeDetalle AS ViajeDetalleDebito ON CuentaCorriente.Viaje_FechaHora = ViajeDetalleDebito.FechaHora AND CuentaCorriente.Viaje_IDRuta = ViajeDetalleDebito.IDRuta AND CuentaCorriente.Viaje_Indice = ViajeDetalleDebito.Indice) LEFT JOIN Viaje AS ViajeDebito ON ViajeDetalleDebito.FechaHora = ViajeDebito.FechaHora AND ViajeDetalleDebito.IDRuta = ViajeDebito.IDRuta) LEFT JOIN ViajeDetalle AS ViajeDetalleCredito ON CuentaCorriente.Viaje_FechaHora = ViajeDetalleCredito.FechaHora AND CuentaCorriente.Viaje_IDRuta = ViajeDetalleCredito.IDRuta AND CuentaCorriente.Viaje_Indice = ViajeDetalleCredito.Indice) LEFT JOIN Viaje AS ViajeCredito ON ViajeDetalleCredito.FechaHora = ViajeCredito.FechaHora AND ViajeDetalleCredito.IDRuta = ViajeCredito.IDRuta
				WHERE CuentaCorriente.Importe > 0
					AND (@FechaDesde IS NULL OR convert(char(10), CuentaCorriente.FechaHora, 111) >= convert(char(10), @FechaDesde, 111))
					AND (@FechaHasta IS NULL OR convert(char(10), CuentaCorriente.FechaHora, 111) <= convert(char(10), @FechaHasta, 111))
					AND (@Personal = 0 OR ViajeDebito.Personal IS NULL OR ViajeDebito.Personal = 0)
					AND (@Personal = 0 OR ViajeCredito.Personal IS NULL OR ViajeCredito.Personal = 0)
				GROUP BY CuentaCorriente.IDCuentaCorrienteGrupo) AS Ingresos ON CuentaCorrienteGrupo.IDCuentaCorrienteGrupo = Ingresos.IDCuentaCorrienteGrupo
			LEFT JOIN (SELECT CuentaCorriente.IDCuentaCorrienteGrupo, Sum(CuentaCorriente.Importe) AS Importe
				FROM (((CuentaCorriente LEFT JOIN ViajeDetalle AS ViajeDetalleDebito ON CuentaCorriente.Viaje_FechaHora = ViajeDetalleDebito.FechaHora AND CuentaCorriente.Viaje_IDRuta = ViajeDetalleDebito.IDRuta AND CuentaCorriente.Viaje_Indice = ViajeDetalleDebito.Indice) LEFT JOIN Viaje AS ViajeDebito ON ViajeDetalleDebito.FechaHora = ViajeDebito.FechaHora AND ViajeDetalleDebito.IDRuta = ViajeDebito.IDRuta) LEFT JOIN ViajeDetalle AS ViajeDetalleCredito ON CuentaCorriente.Viaje_FechaHora = ViajeDetalleCredito.FechaHora AND CuentaCorriente.Viaje_IDRuta = ViajeDetalleCredito.IDRuta AND CuentaCorriente.Viaje_Indice = ViajeDetalleCredito.Indice) LEFT JOIN Viaje AS ViajeCredito ON ViajeDetalleCredito.FechaHora = ViajeCredito.FechaHora AND ViajeDetalleCredito.IDRuta = ViajeCredito.IDRuta
				WHERE CuentaCorriente.Importe < 0
					AND (@FechaDesde IS NULL OR convert(char(10), CuentaCorriente.FechaHora, 111) >= convert(char(10), @FechaDesde, 111))
					AND (@FechaHasta IS NULL OR convert(char(10), CuentaCorriente.FechaHora, 111) <= convert(char(10), @FechaHasta, 111))
					AND (@Personal = 0 OR ViajeDebito.Personal IS NULL OR ViajeDebito.Personal = 0)
					AND (@Personal = 0 OR ViajeCredito.Personal IS NULL OR ViajeCredito.Personal = 0)
				GROUP BY CuentaCorriente.IDCuentaCorrienteGrupo) AS Egresos ON CuentaCorrienteGrupo.IDCuentaCorrienteGrupo = Egresos.IDCuentaCorrienteGrupo
		ORDER BY CuentaCorrienteGrupo.Nombre

GO




------------------------------------------------------------------------------------------
-- REPORT_VIAJE_PASAJERO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name
	   FROM   sysobjects 
	   WHERE  name = N'sp_Report_Viaje_Pasajero'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_Viaje_Pasajero
GO

CREATE PROCEDURE dbo.sp_Report_Viaje_Pasajero
	@DiaSemana int,
	@FechaHoraDesde smalldatetime,
	@FechaHoraHasta smalldatetime,
	@FechaDesde smalldatetime,
	@FechaHasta smalldatetime,
	@HoraDesde smalldatetime,
	@HoraHasta smalldatetime,
	@IDRuta char(20),
	@IDListaPrecio int,
	@PersonaTipo char(2),
	@Personal bit AS

	SELECT (CASE datepart(weekday, Viaje.FechaHora) WHEN 1 THEN 'Domingo' WHEN 2 THEN 'Lunes' WHEN 3 THEN 'Martes' WHEN 4 THEN 'Miércoles' WHEN 5 THEN 'Jueves' WHEN 6 THEN 'Viernes' WHEN 7 THEN 'Sábado' END) AS DiaSemana, Viaje.FechaHora, Viaje.IDRuta, Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Pasajero, LugarOrigen.Nombre AS Sube, LugarDestino.Nombre AS Baja
		FROM (((Viaje INNER JOIN ViajeDetalle ON Viaje.FechaHora = ViajeDetalle.FechaHora AND Viaje.IDRuta = ViajeDetalle.IDRuta) INNER JOIN Persona ON ViajeDetalle.IDPersona = Persona.IDPersona) INNER JOIN Lugar AS LugarOrigen ON ViajeDetalle.IDOrigen = LugarOrigen.IDLugar) INNER JOIN Lugar AS LugarDestino ON ViajeDetalle.IDDestino = LugarDestino.IDLugar
		WHERE
			ViajeDetalle.Estado = '1CO' AND ViajeDetalle.OcupanteTipo = 'PA'
			AND (@DiaSemana IS NULL OR datepart(weekday, Viaje.FechaHora) = @DiaSemana)
			AND (@FechaHoraDesde IS NULL OR Viaje.FechaHora >= @FechaHoraDesde)
			AND (@FechaHoraHasta IS NULL OR Viaje.FechaHora <= @FechaHoraHasta)
			AND (@FechaDesde IS NULL OR convert(char(10), Viaje.FechaHora, 111) >= convert(char(10), @FechaDesde, 111))
			AND (@FechaHasta IS NULL OR convert(char(10), Viaje.FechaHora, 111) <= convert(char(10), @FechaHasta, 111))
			AND (@HoraDesde IS NULL OR convert(char(8), Viaje.FechaHora, 108) >= convert(char(8), @HoraDesde, 108))
			AND (@HoraHasta IS NULL OR convert(char(8), Viaje.FechaHora, 108) <= convert(char(8), @HoraHasta, 108))
			AND (@IDRuta IS NULL OR Viaje.IDRuta = @IDRuta)
			AND (@IDListaPrecio IS NULL OR ViajeDetalle.IDListaPrecio = @IDListaPrecio)
			AND (@PersonaTipo IS NULL OR Persona.EntidadTipo = @PersonaTipo)
			AND (@Personal = 0 OR Viaje.Personal = 0)
		ORDER BY Viaje.FechaHora, Viaje.IDRuta, Persona.Apellido, Persona.Nombre

GO



------------------------------------------------------------------------------------------
-- REPORT_PERSONA_VIAJE_LIST
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Report_Persona_Viaje_Listado' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_Persona_Viaje_Listado
GO

CREATE PROCEDURE dbo.sp_Report_Persona_Viaje_Listado
	@DiaSemana int,
	@FechaHoraDesde smalldatetime,
	@FechaHoraHasta smalldatetime,
	@FechaDesde smalldatetime,
	@FechaHasta smalldatetime,
	@HoraDesde smalldatetime,
	@HoraHasta smalldatetime,
	@IDRuta char(20),
	@Estado char(2),
	@IDPersona int,
	@IDListaPrecio int,
	@Charter bit,
	@Personal bit AS
	
	DECLARE @Ruta_ID_Otra char(20)
	
	SET @Ruta_ID_Otra = (SELECT Texto FROM Parametro WHERE IDParametro = 'Ruta_ID_Otra')

	SELECT Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Persona, (CASE datepart(weekday, Viaje.FechaHora) WHEN 1 THEN 'Domingo' WHEN 2 THEN 'Lunes' WHEN 3 THEN 'Martes' WHEN 4 THEN 'Miércoles' WHEN 5 THEN 'Jueves' WHEN 6 THEN 'Viernes' WHEN 7 THEN 'Sábado' END) AS DiaSemana, Viaje.FechaHora, RTRIM(Viaje.IDRuta) + (CASE RTRIM(Viaje.IDRuta) WHEN @Ruta_ID_Otra THEN ': ' + Viaje.RutaOtra ELSE '' END) AS Ruta, (CASE Viaje.Estado WHEN 'AC' THEN 'Activo' WHEN 'EP' Then 'En Progreso' WHEN 'FI' THEN 'Finalizado' WHEN 'CA' THEN 'Cancelado' END) AS Estado, ListaPrecio.Nombre AS ListaPrecio
		FROM ((Viaje INNER JOIN ViajeDetalle ON Viaje.FechaHora = ViajeDetalle.FechaHora AND Viaje.IDRuta = ViajeDetalle.IDRuta) INNER JOIN Persona ON ViajeDetalle.IDPersona = Persona.IDPersona) INNER JOIN ListaPrecio ON ViajeDetalle.IDListaPrecio = ListaPrecio.IDListaPrecio
		WHERE
			(@DiaSemana IS NULL OR datepart(weekday, Viaje.FechaHora) = @DiaSemana)
			AND (@FechaHoraDesde IS NULL OR Viaje.FechaHora >= @FechaHoraDesde)
			AND (@FechaHoraHasta IS NULL OR Viaje.FechaHora <= @FechaHoraHasta)
			AND (@FechaDesde IS NULL OR convert(char(10), Viaje.FechaHora, 111) >= convert(char(10), @FechaDesde, 111))
			AND (@FechaHasta IS NULL OR convert(char(10), Viaje.FechaHora, 111) <= convert(char(10), @FechaHasta, 111))
			AND (@HoraDesde IS NULL OR convert(char(8), Viaje.FechaHora, 108) >= convert(char(8), @HoraDesde, 108))
			AND (@HoraHasta IS NULL OR convert(char(8), Viaje.FechaHora, 108) <= convert(char(8), @HoraHasta, 108))
			AND (@IDRuta IS NULL OR Viaje.IDRuta = @IDRuta)
			AND (@Estado IS NULL OR Viaje.Estado = @Estado)
			AND (@IDPersona IS NULL OR ViajeDetalle.IDPersona = @IDPersona)
			AND (@IDListaPrecio IS NULL OR ViajeDetalle.IDListaPrecio = @IDListaPrecio)
			AND (@Charter IS NULL OR Viaje.Charter = @Charter)
			AND ViajeDetalle.OcupanteTipo = 'PA'
			AND ViajeDetalle.Estado = '1CO'
			AND (@Personal = 0 OR Viaje.Personal = 0)
		ORDER BY Persona, Viaje.FechaHora, Viaje.IDRuta, Viaje.RutaOtra

GO



------------------------------------------------------------------------------------------
-- REPORT_VIAJEDETALLE_PRECIOACTUALIZAR
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Report_ViajeDetalle_PrecioActualizar'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_ViajeDetalle_PrecioActualizar
GO

CREATE PROCEDURE dbo.sp_Report_ViajeDetalle_PrecioActualizar
	@FechaHoraDesde smalldatetime AS

	SELECT ViajeDetalle.FechaHora, ViajeDetalle.IDRuta, ViajeDetalle.Orden, Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Pasajero, ViajeDetalle.Importe AS ImporteViaje, ListaPrecioDetalle.Importe AS ImporteListaPrecio
		FROM (((ListaPrecioDetalle INNER JOIN RutaDetalle AS RutaDetalleOrigen ON ListaPrecioDetalle.IDRuta = RutaDetalleOrigen.IDRuta AND ListaPrecioDetalle.IDLugarGrupoOrigen = RutaDetalleOrigen.IDLugarGrupo)
			INNER JOIN RutaDetalle AS RutaDetalleDestino ON ListaPrecioDetalle.IDRuta = RutaDetalleDestino.IDRuta AND ListaPrecioDetalle.IDLugarGrupoDestino = RutaDetalleDestino.IDLugarGrupo)
			INNER JOIN ViajeDetalle ON ListaPrecioDetalle.IDRuta = ViajeDetalle.IDRuta AND ListaPrecioDetalle.OcupanteTipo = ViajeDetalle.OcupanteTipo AND ListaPrecioDetalle.IDListaPrecio = ViajeDetalle.IDListaPrecio AND RutaDetalleOrigen.IDLugar = ViajeDetalle.IDOrigen AND RutaDetalleDestino.IDLugar = ViajeDetalle.IDDestino)
			INNER JOIN Persona ON ViajeDetalle.IDPersona = Persona.IDPersona
		WHERE ViajeDetalle.FechaHora >= @FechaHoraDesde AND ViajeDetalle.OcupanteTipo = 'PA' AND ViajeDetalle.Importe <> ListaPrecioDetalle.Importe  AND ViajeDetalle.Estado = '1CO'
		ORDER BY ViajeDetalle.FechaHora, ViajeDetalle.IDRuta

GO



------------------------------------------------------------------------------------------
-- REPORT_COMISION_SEGURO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Report_Comision_Seguro_Totales'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_Comision_Seguro_Totales
GO

CREATE PROCEDURE dbo.sp_Report_Comision_Seguro_Totales
	@FechaHoraDesde smalldatetime,
	@FechaHoraHasta smalldatetime,
	@IDRuta char(20) AS
	
	SELECT COUNT(*) AS CantidadComisiones, SUM(ViajeDetalle.ValorDeclarado) AS ValorDeclarado, SUM(ViajeDetalle.ImporteSeguro) AS ImporteSeguro, SUM(ViajeDetalle.Importe) AS Importe
		FROM ViajeDetalle
		WHERE ViajeDetalle.OcupanteTipo = 'CO'
			AND (@FechaHoraDesde IS NULL OR ViajeDetalle.FechaHora >= @FechaHoraDesde)
			AND (@FechaHoraHasta IS NULL OR ViajeDetalle.FechaHora <= @FechaHoraHasta)
			AND (@IDRuta IS NULL OR ViajeDetalle.IDRuta = @IDRuta)

GO



------------------------------------------------------------------------------------------
-- REPORT_VIAJE_CONDUCTOR
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Report_Viaje_Conductor'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_Viaje_Conductor
GO

CREATE PROCEDURE dbo.sp_Report_Viaje_Conductor
	@Personal bit,
	@FechaDesde smalldatetime,
	@FechaHasta smalldatetime,
	@IDConductor int,
	@IDRuta char(20),
	@EstadoActivo bit,
	@EstadoEnProgreso bit,
	@EstadoFinalizado bit,
	@EstadoCancelado bit,
	@TramoCompleto bit,
	@Tramo1 bit,
	@Tramo2 bit AS

	DECLARE @Viaje_Permite_2_Conductores bit
	DECLARE @ConductorNombre varchar(150)

	SET @Viaje_Permite_2_Conductores = (SELECT SiNo FROM Parametro WHERE IDParametro = 'Viaje_Permite_2_Conductores')
	SET @ConductorNombre = (SELECT Apellido + ', ' + Nombre FROM Persona WHERE IDPersona = @IDConductor)
	
	IF @Viaje_Permite_2_Conductores = 1
		SELECT @ConductorNombre AS ConductorNombre, Viaje.FechaHora, Viaje.IDRuta, Viaje.RutaOtra, Viaje.Estado, Vehiculo.Nombre AS Vehiculo, dbo.udf_GetViajeTramoNombre(@IDConductor, Viaje.IDConductor, Viaje.IDConductor2) AS Tramo, (CASE dbo.udf_GetViajeTramoNumero(@IDConductor, Viaje.IDConductor, Viaje.IDConductor2) WHEN 0 THEN dbo.udf_GetViajeTramoImporte(ConductorRuta1.ConductorImporteTramoCompleto, Horario.ConductorImporteTramoCompleto, Ruta.ConductorImporteTramoCompleto) WHEN 1 THEN dbo.udf_GetViajeTramoImporte(ConductorRuta1.ConductorImporteTramo1, Horario.ConductorImporteTramo1, Ruta.ConductorImporteTramo1) WHEN 2 THEN dbo.udf_GetViajeTramoImporte(ConductorRuta2.ConductorImporteTramo2, Horario.ConductorImporteTramo2, Ruta.ConductorImporteTramo2) END) AS Importe
			FROM ((((Viaje LEFT JOIN Vehiculo ON Viaje.IDVehiculo = Vehiculo.IDVehiculo) INNER JOIN Ruta ON Viaje.IDRuta = Ruta.IDRuta) LEFT JOIN Horario ON Viaje.DiaSemanaBase = Horario.DiaSemana AND CONVERT(CHAR(8), Viaje.FechaHora, 108) = CONVERT(CHAR(8), Horario.Hora, 108) AND Viaje.IDRuta = Horario.IDRuta) LEFT JOIN ConductorRuta AS ConductorRuta1 ON Viaje.IDConductor = ConductorRuta1.IDPersona AND Viaje.IDRuta = ConductorRuta1.IDRuta) LEFT JOIN ConductorRuta AS ConductorRuta2 ON Viaje.IDConductor2 = ConductorRuta2.IDPersona AND Viaje.IDRuta = ConductorRuta2.IDRuta
			WHERE (@Personal = 0 OR Viaje.Personal = 0)
				AND (Viaje.FechaHora BETWEEN @FechaDesde AND @FechaHasta)
				AND (Viaje.IDConductor = @IDConductor OR Viaje.IDConductor2 = @IDConductor)
				AND (@IDRuta IS NULL OR Viaje.IDRuta = @IDRuta)
				AND (@EstadoActivo = 1 OR (@EstadoActivo = 0 AND Viaje.Estado <> 'AC'))
				AND (@EstadoEnProgreso = 1 OR (@EstadoEnProgreso = 0 AND Viaje.Estado <> 'EP'))
				AND (@EstadoFinalizado = 1 OR (@EstadoFinalizado = 0 AND Viaje.Estado <> 'FI'))
				AND (@EstadoCancelado = 1 OR (@EstadoCancelado = 0 AND Viaje.Estado <> 'CA'))
				AND (@TramoCompleto = 1 OR (@TramoCompleto = 0 AND NOT (ISNULL(Viaje.IDConductor2, 0) = 0)))
				AND (@Tramo1 = 1 OR (@Tramo1 = 0 AND NOT (ISNULL(Viaje.IDConductor2, 0) <> 0 AND Viaje.IDConductor = @IDConductor)))
				AND (@Tramo2 = 1 OR (@Tramo2 = 0 AND NOT (ISNULL(Viaje.IDConductor2, 0) = @IDConductor)))
	ELSE
		SELECT @ConductorNombre AS ConductorNombre, Viaje.FechaHora, Viaje.IDRuta, Viaje.RutaOtra, Viaje.Estado, Vehiculo.Nombre AS Vehiculo, 'Completo' AS Tramo, dbo.udf_GetViajeTramoImporte(ConductorRuta1.ConductorImporteTramoCompleto, Horario.ConductorImporteTramoCompleto, Ruta.ConductorImporteTramoCompleto) AS Importe
			FROM (((Viaje LEFT JOIN Vehiculo ON Viaje.IDVehiculo = Vehiculo.IDVehiculo) INNER JOIN Ruta ON Viaje.IDRuta = Ruta.IDRuta) LEFT JOIN Horario ON Viaje.DiaSemanaBase = Horario.DiaSemana AND CONVERT(CHAR(8), Viaje.FechaHora, 108) = CONVERT(CHAR(8), Horario.Hora, 108) AND Viaje.IDRuta = Horario.IDRuta) LEFT JOIN ConductorRuta AS ConductorRuta1 ON Viaje.IDConductor = ConductorRuta1.IDPersona AND Viaje.IDRuta = ConductorRuta1.IDRuta
			WHERE (@Personal = 0 OR Viaje.Personal = 0)
				AND (Viaje.FechaHora BETWEEN @FechaDesde AND @FechaHasta)
				AND (Viaje.IDConductor = @IDConductor OR Viaje.IDConductor2 = @IDConductor)
				AND (@IDRuta IS NULL OR Viaje.IDRuta = @IDRuta)
				AND (@EstadoActivo = 1 OR (@EstadoActivo = 0 AND Viaje.Estado <> 'AC'))
				AND (@EstadoEnProgreso = 1 OR (@EstadoEnProgreso = 0 AND Viaje.Estado <> 'EP'))
				AND (@EstadoFinalizado = 1 OR (@EstadoFinalizado = 0 AND Viaje.Estado <> 'FI'))
				AND (@EstadoCancelado = 1 OR (@EstadoCancelado = 0 AND Viaje.Estado <> 'CA'))
GO