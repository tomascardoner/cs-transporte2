USE CSTransporte_LobosBus
GO

-- =============================================
-- Author:		Tomás A. Cardoner
-- Create date: 2020-09-01 21:58
-- Description:	Obtiene las reservas web que vencieron y no se pagaron
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'uspObtenerReservasWebConPagoVencido') AND type in (N'P', N'PC'))
	 DROP PROCEDURE uspObtenerReservasWebConPagoVencido
GO

CREATE PROCEDURE uspObtenerReservasWebConPagoVencido
	@FechaHoraDesde smalldatetime,
	@FechaHoraHasta smalldatetime
	AS

	BEGIN
		-- SET NOCOUNT ON added to prevent extra result sets from interfering with SELECT statements.
		SET NOCOUNT ON;

		SELECT vd.IDViajeDetalle, vd.FechaHora, vd.IDRuta, vd.Orden, dbo.udf_GetEntidadApellidoYNombre(p.Apellido, p.Nombre) AS ApellidoNombre, anu.Email
			FROM ViajeDetalle AS vd
				INNER JOIN Persona AS p ON vd.IDPersona = p.IDPersona
				INNER JOIN CSTransporte_Web_LobosBus.dbo.Reserva AS r ON vd.IDViajeDetalle = r.idViajeDetalle
				INNER JOIN CSTransporte_Web_LobosBus.dbo.AspNetUsers AS anu ON r.idAspNetUserCliente = anu.id
			WHERE vd.FechaHora BETWEEN @FechaHoraDesde AND @FechaHoraHasta
				AND vd.IDUsuarioCreacion = 151 AND vd.ImporteContado = 0 AND vd.Estado = '1CO' AND vd.OcupanteTipo = 'PA'
				AND vd.FechaHoraCreacion <= DATEADD(second, -600, GETDATE())
				AND anu.PuedePagarEnOficina = 0
	END
GO

GRANT EXECUTE ON dbo.uspObtenerReservasWebConPagoVencido TO cstransporte
GO

USE CSTransporte_Web_LobosBus
GO
GRANT SELECT ON CSTransporte_Web_LobosBus.dbo.AspNetUsers TO cstransporte
GO
GRANT SELECT ON CSTransporte_Web_LobosBus.dbo.Reserva TO cstransporte
GO