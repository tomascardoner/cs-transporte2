------------------------------------------------------------------------------------------
-- REPORT_VIAJE_PASAJEROSVIAJESIMPORTE_MES
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name
	   FROM   sysobjects
	   WHERE  name = N'sp_Report_Viaje_PasajerosViajesImporte_Mes'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Report_Viaje_PasajerosViajesImporte_Mes
GO

CREATE PROCEDURE dbo.sp_Report_Viaje_PasajerosViajesImporte_Mes
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

	SELECT Viajes.Anio, Viajes.Mes, IsNull(Ocupantes.Ocupantes, 0) AS Ocupantes, IsNull(Ocupantes.ImporteTotal, 0) + IsNull(Viajes.Importe, 0) AS ImporteTotal, (CASE ISNULL(@DividirViajes, 0) WHEN 0 THEN Viajes.Viajes WHEN 1 THEN Viajes.Viajes / 2 END) AS Viajes
		FROM
			(SELECT DatePart(year, Viaje.FechaHora) AS Anio, DatePart(month, Viaje.FechaHora) AS Mes, CONVERT(FLOAT, COUNT(CONVERT(CHAR(19), Viaje.FechaHora, 120) + Viaje.IDRuta)) AS Viajes, SUM(CASE ISNULL(@OcupanteTipo_FILTER, 'NU') WHEN 'CO' THEN 0 ELSE Viaje.Importe END) AS Importe
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
				GROUP BY DatePart(year, Viaje.FechaHora), DatePart(month, Viaje.FechaHora)) AS Viajes
			LEFT JOIN
			(SELECT DatePart(year, Viaje.FechaHora) AS Anio, DatePart(month, Viaje.FechaHora) AS Mes, Count(ViajeDetalle.Indice) AS Ocupantes, Sum(ViajeDetalle.Importe) AS ImporteTotal
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
				GROUP BY DatePart(year, Viaje.FechaHora), DatePart(month, Viaje.FechaHora)) AS Ocupantes
			ON Viajes.Anio = Ocupantes.Anio AND Viajes.Mes = Ocupantes.Mes
		ORDER BY Viajes.Anio, Viajes.Mes

GO