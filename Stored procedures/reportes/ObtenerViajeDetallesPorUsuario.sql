IF EXISTS (SELECT name FROM sys.procedures WHERE name = N'ObtenerViajeDetallesPorOperador')
    DROP PROCEDURE ObtenerViajeDetallesPorOperador
GO

CREATE PROCEDURE dbo.ObtenerViajeDetallesPorOperador
    @IdOperador int,
    @VentaFechaDesde smalldatetime,
    @VentaFechaHasta smalldatetime,
    @ReservaFechaDesde smalldatetime,
    @ReservaFechaHasta smalldatetime,
    @IdRuta char(20),
    @OcupanteTipo char(2),
    @Personal bit AS

	-- SET NOCOUNT ON added to prevent extra result sets from interfering with SELECT statements.
	SET NOCOUNT ON;

    DECLARE @IdRutaGrupo int = NULL

    SELECT dbo.udf_GetEntidadApellidoYNombre(pcr.Apellido, pcr.Nombre) AS Operador, vd.FechaHoraCreacion AS VentaFechaHora,
            rg.Nombre AS RutaGrupo, RTRIM(v.IDRuta) AS Ruta,
            vd.FechaHora AS ReservaFechaHora, vd.OcupanteTipo, dbo.udf_GetEntidadApellidoYNombre(pcl.Apellido, pcl.Nombre) AS Cliente, vd.Importe, dbo.udf_GetEntidadApellidoYNombre(pca.Apellido, pca.Nombre) AS CanceladoPor
        FROM ViajeDetalle AS vd
            INNER JOIN Usuario AS ucr ON vd.IDUsuarioCreacion = ucr.IDUsuario
            INNER JOIN Persona AS pcr ON ucr.IDPersona = pcr.IDPersona
            INNER JOIN Ruta AS r ON vd.IDRuta = r.IDRuta
            INNER JOIN RutaGrupo AS rg ON r.IDRutaGrupo = rg.IDRutaGrupo
            INNER JOIN Persona AS pcl ON vd.IDPersona = pcl.IDPersona
            INNER JOIN Viaje AS v ON vd.IDViaje = v.IDViaje
            LEFT JOIN Usuario AS uca ON vd.IDUsuarioCancelacion = uca.IDUsuario
            LEFT JOIN Persona AS pca ON uca.IDPersona = pca.IDPersona
        WHERE ucr.IDPersona = @IdOperador
            AND (@VentaFechaDesde IS NULL OR vd.FechaHoraCreacion >= @VentaFechaDesde)
            AND (@VentaFechaHasta IS NULL OR vd.FechaHoraCreacion < DATEADD(DAY, 1, @VentaFechaHasta))
            AND (@ReservaFechaDesde IS NULL OR vd.FechaHora >= @ReservaFechaDesde)
            AND (@ReservaFechaHasta IS NULL OR vd.FechaHora < DATEADD(DAY, 1, @ReservaFechaHasta))
            AND (@IdRutaGrupo IS NULL OR r.IDRutaGrupo = @IdRutaGrupo)
            AND (@IdRuta IS NULL OR vd.IDRuta = @IdRuta)
            AND (@OcupanteTipo IS NULL OR vd.OcupanteTipo = @OcupanteTipo)
            AND (@Personal = 0 OR v.Personal = 0)
        ORDER BY vd.FechaHoraCreacion, vd.FechaHora, vd.IDRuta, pcl.Apellido, pcl.Nombre
GO

GRANT EXECUTE ON dbo.ObtenerViajeDetallesPorOperador TO cstransporte