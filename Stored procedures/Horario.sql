------------------------------------------------------------------------------------------
-- HORARIO_LIST_DIASEMANA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Horario_List_DiaSemana' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Horario_List_DiaSemana
GO

CREATE PROCEDURE dbo.sp_Horario_List_DiaSemana
	@DiaSemana_FILTER tinyint AS

	SELECT h.DiaSemana, h.Hora, RTRIM(h.IDRuta) AS IDRuta, h.RutaOtra, h.IDPersona, h.Importe, h.Kilometro, h.Duracion, h.IDConductor, h.IDConductor2, h.IDVehiculo, h.ConductorImporteTramoCompleto, h.ConductorImporteTramo1, h.ConductorImporteTramo2, h.Notas, h.Personal, h.FechaHoraCreacion
		FROM Horario AS h
			INNER JOIN Ruta AS r ON h.IDRuta = r.IDRuta
		WHERE h.DiaSemana = @DiaSemana_FILTER AND h.Activo = 1

GO
