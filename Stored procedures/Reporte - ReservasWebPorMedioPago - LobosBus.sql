USE CSTransporte_LobosBus
GO

-- =============================================
-- Author:		Tomás A. Cardoner
-- Create date: 2025-08-19 17:02
-- Description:	Obtiene estadísticas de las reservas web
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'ReservasWebPorMedioPago') AND type in (N'P', N'PC'))
	 DROP PROCEDURE ReservasWebPorMedioPago
GO

CREATE PROCEDURE ReservasWebPorMedioPago
	@FechaHoraDesde smalldatetime,
	@FechaHoraHasta smalldatetime
	AS

	BEGIN
		-- SET NOCOUNT ON added to prevent extra result sets from interfering with SELECT statements.
		SET NOCOUNT ON;

		SELECT CASE WHEN pr.idPago IS NULL THEN 'Oficina' ELSE 'Online' END AS LugarPago, mp.Nombre AS MedioPago, COUNT(vd.IDViajeDetalle) AS CantidadReservas, SUM(vd.Importe) AS Importe
			FROM ViajeDetalle AS vd
				INNER JOIN MedioPago AS mp ON vd.IDMedioPago = mp.IDMedioPago
                INNER JOIN Persona AS p ON vd.IDPersona = p.IDPersona
				INNER JOIN CSTransporte_Web_LobosBus.dbo.Reserva AS r ON vd.IDViajeDetalle = r.idViajeDetalle
				INNER JOIN CSTransporte_Web_LobosBus.dbo.PagoReserva AS pr ON r.idReserva = pr.idReserva
			WHERE vd.FechaHora BETWEEN @FechaHoraDesde AND @FechaHoraHasta
				AND vd.IDUsuarioCreacion = 151 AND vd.ImporteContado > 0 AND vd.Estado = '1CO' AND vd.OcupanteTipo = 'PA' AND vd.Realizado = 1
			GROUP BY CASE WHEN pr.idPago IS NULL THEN 'Oficina' ELSE 'Online' END, mp.Nombre
			ORDER BY CASE WHEN pr.idPago IS NULL THEN 'Oficina' ELSE 'Online' END, mp.Nombre
	END
GO

GRANT EXECUTE ON dbo.uspObtenerReservasWebPagadasConMP TO cstransporte
GO

USE CSTransporte_Web_LobosBus
GO
GRANT SELECT ON CSTransporte_Web_LobosBus.dbo.Reserva TO cstransporte
GO
GRANT SELECT ON CSTransporte_Web_LobosBus.dbo.PagoReserva TO cstransporte
GO
GRANT SELECT ON CSTransporte_Web_LobosBus.dbo.Pago TO cstransporte
GO