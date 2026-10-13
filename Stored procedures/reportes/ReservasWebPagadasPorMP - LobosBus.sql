USE CSTransporte_LobosBus
GO

-- =============================================
-- Author:		Tomás A. Cardoner
-- Create date: 2020-09-01 13:08
-- Updates: 2024-05-15 17:28 - Se agregó el campo importe
-- Description:	Obtiene las reservas web
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'uspObtenerReservasWebPagadasConMP') AND type in (N'P', N'PC'))
	 DROP PROCEDURE uspObtenerReservasWebPagadasConMP
GO

CREATE PROCEDURE uspObtenerReservasWebPagadasConMP
	@FechaHoraDesde smalldatetime,
	@FechaHoraHasta smalldatetime
	AS

	BEGIN
		-- SET NOCOUNT ON added to prevent extra result sets from interfering with SELECT statements.
		SET NOCOUNT ON;

		SELECT DISTINCT vd.IDViajeDetalle, vd.FechaHora, vd.IDRuta, vd.Orden, dbo.udf_GetEntidadApellidoYNombre(p.Apellido, p.Nombre) AS ApellidoNombre, vd.FacturaNumero, pa.codigoPago AS CodigoPago, pa.monto AS ImportePago
			FROM ViajeDetalle AS vd
                INNER JOIN Persona AS p ON vd.IDPersona = p.IDPersona
				INNER JOIN CSTransporte_Web_LobosBus.dbo.Reserva AS r ON vd.IDViajeDetalle = r.idViajeDetalle
				INNER JOIN CSTransporte_Web_LobosBus.dbo.PagoReserva AS pr ON r.idReserva = pr.idReserva
				INNER JOIN CSTransporte_Web_LobosBus.dbo.Pago AS pa ON pr.idPago = pa.idPago
			WHERE vd.FechaHora BETWEEN @FechaHoraDesde AND @FechaHoraHasta
				AND vd.IDUsuarioCreacion = 151 AND vd.ImporteContado > 0 AND vd.Estado = '1CO' AND vd.OcupanteTipo = 'PA'
				AND vd.IDMedioPago = 9
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