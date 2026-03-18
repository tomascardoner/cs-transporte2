------------------------------------------------------------------------------------------
-- VIAJEDETALLE_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ViajeDetalle_Data' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ViajeDetalle_Data
GO

CREATE PROCEDURE dbo.sp_ViajeDetalle_Data
	@FechaHora_FILTER smalldatetime,
	@IDRuta_FILTER char(20),
	@Indice_FILTER int AS

	SELECT IDViajeDetalle, IDViaje, FechaHora, IDRuta, Indice, GrupoNumero, ReservaCodigo, OcupanteTipo, Estado, Prioridad, Orden, Asiento, AsientoIdentificacion, Realizado, IDPersona, IDPersonaMenor, IDListaPrecio, IDOrigen, Sube, IDDestino, Baja, ValorDeclarado, ImporteSeguro, Importe, ImporteContado, IDMedioPago, Cuotas, Operacion, ImporteCuentaCorriente, ImprimirSaldo, IDCuentaCorrienteCaja, ForzarDebito, IDPersonaCuentaCorriente, FacturaNumero, Facturar, FacturarNotas, IDPersonaRecibe, PagaQuienRecibe, Recibe, Descripcion, Domicilio, Horario, Telefono, DejarTraer, Entregada, EntregadaFechaHora, Retira, ReservaTipo, CreadoEnProgreso, ModificadoEnProgreso, ReservadoPor, CanceladoPor, CanceladoFechaHora, Notas, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion, IDUsuarioCancelacion, AsistenciaFechaHora
		FROM ViajeDetalle
		WHERE FechaHora = @FechaHora_FILTER AND IDRuta = @IDRuta_FILTER AND Indice = @Indice_FILTER

GO



------------------------------------------------------------------------------------------
-- VIAJEDETALLE_LOADBYIDVIAJEDETALLE
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ViajeDetalle_LoadByIDViajeDetalle'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ViajeDetalle_LoadByIDViajeDetalle
GO

CREATE PROCEDURE dbo.sp_ViajeDetalle_LoadByIDViajeDetalle
	@IDViajeDetalle int AS

	SELECT IDViajeDetalle, IDViaje, FechaHora, IDRuta, Indice, GrupoNumero, ReservaCodigo, OcupanteTipo, Estado, Prioridad, Orden, Asiento, AsientoIdentificacion, Realizado, IDPersona, IDPersonaMenor, IDListaPrecio, IDOrigen, Sube, IDDestino, Baja, ValorDeclarado, ImporteSeguro, Importe, ImporteContado, IDMedioPago, Cuotas, Operacion, ImporteCuentaCorriente, ImprimirSaldo, IDCuentaCorrienteCaja, ForzarDebito, IDPersonaCuentaCorriente, FacturaNumero, Facturar, FacturarNotas, IDPersonaRecibe, PagaQuienRecibe, Recibe, Descripcion, Domicilio, Horario, Telefono, DejarTraer, Entregada, EntregadaFechaHora, Retira, ReservaTipo, CreadoEnProgreso, ModificadoEnProgreso, ReservadoPor, CanceladoPor, CanceladoFechaHora, Notas, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion, IDUsuarioCancelacion, AsistenciaFechaHora
		FROM ViajeDetalle
		WHERE IDViajeDetalle = @IDViajeDetalle

GO



------------------------------------------------------------------------------------------
-- VIAJEDETALLE_GETINDICENEW
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ViajeDetalle_GetIndiceNew'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ViajeDetalle_GetIndiceNew
GO

CREATE PROCEDURE dbo.sp_ViajeDetalle_GetIndiceNew
	@FechaHora smalldatetime,
	@IDRuta char(20),
	@Indice int OUTPUT AS

	SET XACT_ABORT ON

	BEGIN TRANSACTION

	SET @Indice = (SELECT Max(ViajeDetalle.Indice)
					FROM ViajeDetalle
					WHERE ViajeDetalle.FechaHora = @FechaHora AND ViajeDetalle.IDRuta = @IDRuta)

	SET @Indice = ISNULL(@Indice, 0) + 1

	COMMIT TRANSACTION

	SET XACT_ABORT OFF

GO



------------------------------------------------------------------------------------------
-- VIAJEDETALLE_GETPRIORIDADNEW
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ViajeDetalle_GetPrioridadNew' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ViajeDetalle_GetPrioridadNew
GO

CREATE PROCEDURE dbo.sp_ViajeDetalle_GetPrioridadNew
	@FechaHora smalldatetime,
	@IDRuta char(20),
	@Prioridad int OUTPUT AS

	SET XACT_ABORT ON

	BEGIN TRANSACTION

	SET @Prioridad = (SELECT Max(ViajeDetalle.Prioridad)
						FROM ViajeDetalle
						WHERE ViajeDetalle.FechaHora = @FechaHora AND ViajeDetalle.IDRuta = @IDRuta)

	SET @Prioridad = ISNULL(@Prioridad, 0) + 1

	COMMIT TRANSACTION

	SET XACT_ABORT OFF

GO



------------------------------------------------------------------------------------------
-- VIAJEDETALLE_INSERT
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ViajeDetalle_Insert'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ViajeDetalle_Insert
GO

CREATE PROCEDURE dbo.sp_ViajeDetalle_Insert
    @IDViajeDetalle int OUTPUT,
    @IDViaje int,
	@FechaHora smalldatetime,
	@IDRuta char(20),
	@Indice int OUTPUT,
	@GrupoNumero tinyint,
	@ReservaCodigo char(8),
	@OcupanteTipo char(2),
	@Estado char(3),
	@Asiento tinyint,
	@Realizado bit,
	@IDPersona int,
    @IDPersonaMenor int,
	@IDListaPrecio int,
	@IDOrigen int,
	@Sube varchar(50),
	@IDDestino int,
	@Baja varchar(50),
	@ValorDeclarado smallmoney,
	@ImporteSeguro smallmoney, 
	@Importe smallmoney,
	@ImporteContado smallmoney,
	@IDMedioPago tinyint,
	@Cuotas tinyint,
	@Operacion varchar(20),
	@ImporteCuentaCorriente smallmoney,
	@ImprimirSaldo bit,
	@IDCuentaCorrienteCaja int,
	@ForzarDebito bit,
	@IDPersonaCuentaCorriente int,
	@Facturar bit,
	@FacturarNotas varchar(50),
	@FacturaNumero varchar(20),
	@IDPersonaRecibe int,
	@PagaQuienRecibe bit,
	@Recibe varchar(50),
	@Descripcion varchar(100),
	@Domicilio varchar(100),
	@Horario varchar(50),
	@Telefono varchar(30),
	@DejarTraer char(1),
	@Entregada bit,
	@EntregadaFechaHora smalldatetime,
	@Retira varchar(50),
	@ReservaTipo char(2),
	@CreadoEnProgreso bit,
	@ModificadoEnProgreso bit,
	@ReservadoPorID int,
	@ReservadoPor varchar(100),
	@CanceladoPor varchar(100),
	@CanceladoFechaHora smalldatetime,
	@Notas varchar(8000),
	@IDUsuarioCreacion smallint AS

	DECLARE @Prioridad int
	DECLARE @AsistenciaFechaHora smalldatetime

	IF @ImporteContado > 0 AND @IDCuentaCorrienteCaja IS NULL
		BEGIN
		RAISERROR ('No se especificó la Caja.', 19, 1) WITH LOG
		RETURN
		END

	SET XACT_ABORT ON

	BEGIN TRANSACTION

	EXECUTE dbo.sp_ViajeDetalle_GetIndiceNew @FechaHora, @IDRuta, @Indice OUTPUT;

	IF @OcupanteTipo = 'PA'
		BEGIN
		EXECUTE dbo.sp_ViajeDetalle_GetPrioridadNew @FechaHora, @IDRuta, @Prioridad OUTPUT;
		END

	IF ISNULL(@Realizado, 0) = 0
		SET @AsistenciaFechaHora = NULL
	ELSE
		SET @AsistenciaFechaHora = getdate()

	INSERT INTO ViajeDetalle
		(IDViaje, FechaHora, IDRuta, Indice, GrupoNumero, ReservaCodigo, OcupanteTipo, Estado, Prioridad, Asiento, Realizado, IDPersona, IDPersonaMenor, IDListaPrecio, IDOrigen, Sube, IDDestino, Baja, ValorDeclarado, ImporteSeguro, Importe, ImporteContado, IDMedioPago, Cuotas, Operacion, ImporteCuentaCorriente, ImprimirSaldo, IDCuentaCorrienteCaja, ForzarDebito, IDPersonaCuentaCorriente, Facturar, FacturarNotas, FacturaNumero, IDPersonaRecibe, PagaQuienRecibe, Recibe, Descripcion, Domicilio, Horario, Telefono, DejarTraer, Entregada, EntregadaFechaHora, Retira, ReservaTipo, CreadoEnProgreso, ModificadoEnProgreso, ReservadoPor, CanceladoPor, CanceladoFechaHora, Notas, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion, AsistenciaFechaHora)
		VALUES (@IDViaje, @FechaHora, @IDRuta, @Indice, @GrupoNumero, @ReservaCodigo, @OcupanteTipo, @Estado, @Prioridad, @Asiento, @Realizado, @IDPersona, @IDPersonaMenor, @IDListaPrecio, @IDOrigen, @Sube, @IDDestino, @Baja, @ValorDeclarado, @ImporteSeguro, @Importe, @ImporteContado, @IDMedioPago, @Cuotas, @Operacion, @ImporteCuentaCorriente, @ImprimirSaldo, @IDCuentaCorrienteCaja, @ForzarDebito, @IDPersonaCuentaCorriente, @Facturar, @FacturarNotas, @FacturaNumero, @IDPersonaRecibe, @PagaQuienRecibe, @Recibe, @Descripcion, @Domicilio, @Horario, @Telefono, @DejarTraer, @Entregada, @EntregadaFechaHora, @Retira, @ReservaTipo, @CreadoEnProgreso, @ModificadoEnProgreso, @ReservadoPor, @CanceladoPor, @CanceladoFechaHora, @Notas, getdate(), @IDUsuarioCreacion, getdate(), @IDUsuarioCreacion, @AsistenciaFechaHora)
        
    SELECT @IDViajeDetalle = SCOPE_IDENTITY()

	COMMIT TRANSACTION

	SET XACT_ABORT OFF

GO



------------------------------------------------------------------------------------------
-- VIAJEDETALLE_UPDATE
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ViajeDetalle_Update'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ViajeDetalle_Update
GO

CREATE PROCEDURE dbo.sp_ViajeDetalle_Update
	@CambioDeViaje int,
	@FechaHora_Original smalldatetime,
	@IDRuta_Original char(20),
	@Indice_Original int,
    @IDViaje int,
	@FechaHora smalldatetime,
	@IDRuta char(20),
	@Indice int OUTPUT,
	@GrupoNumero tinyint,
	@ReservaCodigo char(8),
	@OcupanteTipo char(2),
	@Estado char(3),
	@Prioridad int,
	@Orden int,
	@Asiento tinyint,
	@Realizado bit,
	@IDPersona int,
    @IDPersonaMenor int,
	@IDListaPrecio int,
	@IDOrigen int,
	@Sube varchar(50),
	@IDDestino int,
	@Baja varchar(50),
	@ValorDeclarado smallmoney,
	@ImporteSeguro smallmoney, 
	@Importe smallmoney,
	@ImporteContado smallmoney,
	@IDMedioPago tinyint,
	@Cuotas tinyint,
	@Operacion varchar(20),
	@ImporteCuentaCorriente smallmoney,
	@ImprimirSaldo bit,
	@IDCuentaCorrienteCaja int,
	@ForzarDebito bit,
	@IDPersonaCuentaCorriente int,
	@Facturar bit,
	@FacturarNotas varchar(50),
	@FacturaNumero varchar(20),
	@IDPersonaRecibe int,
	@PagaQuienRecibe bit,
	@Recibe varchar(50),
	@Descripcion varchar(100),
	@Domicilio varchar(100),
	@Horario varchar(50),
	@Telefono varchar(30),
	@DejarTraer char(1),
	@Entregada bit,
	@EntregadaFechaHora smalldatetime,
	@Retira varchar(50),
	@ReservaTipo char(2),
	@CreadoEnProgreso bit,
	@ModificadoEnProgreso bit,
	@ReservadoPorID int,
	@ReservadoPor varchar(100),
	@CanceladoPor varchar(100),
	@CanceladoFechaHora smalldatetime,
	@Notas varchar(8000),
	@IDUsuarioModificacion smallint,
	@CuentaCorrienteGrupo_ID_ViajeDebito int,
	@CuentaCorrienteGrupo_ID_ViajeCredito int AS

	IF @ImporteContado > 0 AND @IDCuentaCorrienteCaja IS NULL
		BEGIN
		RAISERROR ('No se especificó la Caja.', 19, 1) WITH LOG
		RETURN
		END

	SET XACT_ABORT ON

	BEGIN TRANSACTION

	IF @CambioDeViaje = 1
		BEGIN

		EXECUTE dbo.sp_ViajeDetalle_GetIndiceNew @FechaHora, @IDRuta, @Indice OUTPUT;

		IF @OcupanteTipo = 'PA'
			BEGIN
			EXECUTE dbo.sp_ViajeDetalle_GetPrioridadNew @FechaHora, @IDRuta, @Prioridad OUTPUT;
			END

		-- UPDATE IDVIAJE TO REFLECT THE ID OF THE NEW VIAJE
		SET @IDViaje = (SELECT IDViaje FROM Viaje WHERE FechaHora = @FechaHora AND IDRuta = @IDRuta)

		END

	IF ISNULL(@Realizado, 0) = 0
		UPDATE ViajeDetalle
			SET IDViaje = @IDViaje, FechaHora = @FechaHora, IDRuta = @IDRuta, Indice = @Indice, GrupoNumero = @GrupoNumero, ReservaCodigo = @ReservaCodigo, OcupanteTipo = @OcupanteTipo, Estado = @Estado, Prioridad = @Prioridad, Orden = @Orden, Asiento = @Asiento, AsientoIdentificacion = NULL, Realizado = @Realizado, IDPersona = @IDPersona, IDPersonaMenor = @IDPersonaMenor, IDListaPrecio = @IDListaPrecio, IDOrigen = @IDOrigen, Sube = @Sube, IDDestino = @IDDestino, Baja = @Baja, ValorDeclarado = @ValorDeclarado, ImporteSeguro = @ImporteSeguro, Importe = @Importe, ImporteContado = @ImporteContado, IDMedioPago = @IDMedioPago, Cuotas = @Cuotas, Operacion = @Operacion, ImporteCuentaCorriente = @ImporteCuentaCorriente, ImprimirSaldo = @ImprimirSaldo, IDCuentaCorrienteCaja = @IDCuentaCorrienteCaja, ForzarDebito = @ForzarDebito, IDPersonaCuentaCorriente = @IDPersonaCuentaCorriente, Facturar = @Facturar, FacturarNotas = @FacturarNotas, FacturaNumero = @FacturaNumero, IDPersonaRecibe = @IDPersonaRecibe, PagaQuienRecibe = @PagaQuienRecibe, Recibe = @Recibe, Descripcion = @Descripcion, Domicilio = @Domicilio, Horario = @Horario, Telefono = @Telefono, DejarTraer = @DejarTraer, Entregada = @Entregada, EntregadaFechaHora = @EntregadaFechaHora, Retira = @Retira, ReservaTipo = @ReservaTipo, CreadoEnProgreso = @CreadoEnProgreso, ModificadoEnProgreso = @ModificadoEnProgreso, ReservadoPor = @ReservadoPor, CanceladoPor = @CanceladoPor, CanceladoFechaHora = @CanceladoFechaHora, Notas = @Notas, FechaHoraModificacion = getdate(), IDUsuarioModificacion = @IDUsuarioModificacion, AsistenciaFechaHora = NULL
			WHERE FechaHora = @FechaHora_Original AND IDRuta = @IDRuta_Original AND Indice = @Indice_Original
	ELSE
		UPDATE ViajeDetalle
			SET IDViaje = @IDViaje, FechaHora = @FechaHora, IDRuta = @IDRuta, Indice = @Indice, GrupoNumero = @GrupoNumero, ReservaCodigo = @ReservaCodigo, OcupanteTipo = @OcupanteTipo, Estado = @Estado, Prioridad = @Prioridad, Orden = @Orden, Asiento = @Asiento, Realizado = @Realizado, IDPersona = @IDPersona, IDPersonaMenor = @IDPersonaMenor, IDListaPrecio = @IDListaPrecio, IDOrigen = @IDOrigen, Sube = @Sube, IDDestino = @IDDestino, Baja = @Baja, ValorDeclarado = @ValorDeclarado, ImporteSeguro = @ImporteSeguro, Importe = @Importe, ImporteContado = @ImporteContado, IDMedioPago = @IDMedioPago, Cuotas = @Cuotas, Operacion = @Operacion, ImporteCuentaCorriente = @ImporteCuentaCorriente, ImprimirSaldo = @ImprimirSaldo, IDCuentaCorrienteCaja = @IDCuentaCorrienteCaja, ForzarDebito = @ForzarDebito, IDPersonaCuentaCorriente = @IDPersonaCuentaCorriente, Facturar = @Facturar, FacturarNotas = @FacturarNotas, FacturaNumero = @FacturaNumero, IDPersonaRecibe = @IDPersonaRecibe, PagaQuienRecibe = @PagaQuienRecibe, Recibe = @Recibe, Descripcion = @Descripcion, Domicilio = @Domicilio, Horario = @Horario, Telefono = @Telefono, DejarTraer = @DejarTraer, Entregada = @Entregada, EntregadaFechaHora = @EntregadaFechaHora, Retira = @Retira, ReservaTipo = @ReservaTipo, CreadoEnProgreso = @CreadoEnProgreso, ModificadoEnProgreso = @ModificadoEnProgreso, ReservadoPor = @ReservadoPor, CanceladoPor = @CanceladoPor, CanceladoFechaHora = @CanceladoFechaHora, Notas = @Notas, FechaHoraModificacion = getdate(), IDUsuarioModificacion = @IDUsuarioModificacion, AsistenciaFechaHora = getdate()
			WHERE FechaHora = @FechaHora_Original AND IDRuta = @IDRuta_Original AND Indice = @Indice_Original

	IF @CambioDeViaje = 1
		BEGIN
        
		UPDATE CuentaCorriente
			SET Viaje_FechaHora = @FechaHora, Viaje_IDRuta = @IDRuta, Viaje_Indice = @Indice
			WHERE (IDCuentaCorrienteGrupo = @CuentaCorrienteGrupo_ID_ViajeDebito OR IDCuentaCorrienteGrupo = @CuentaCorrienteGrupo_ID_ViajeCredito) AND Viaje_FechaHora = @FechaHora_Original AND Viaje_IDRuta = @IDRuta_Original AND Viaje_Indice = @Indice_Original

        END

	COMMIT TRANSACTION

	SET XACT_ABORT OFF

GO



------------------------------------------------------------------------------------------
-- VIAJEDETALLE_REALIZAR
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ViajeDetalle_Realizar' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ViajeDetalle_Realizar
GO

CREATE PROCEDURE dbo.sp_ViajeDetalle_Realizar
	@FechaHora_FILTER smalldatetime,
	@IDRuta_FILTER char(20),
	@Indice_FILTER int,
	@Realizado bit,
	@ForzarDebito bit,
	@Entregada bit,
	@EntregadaFechaHora smalldatetime,
	@Retira varchar(50),
	@FacturaNumero varchar(20),
	@ImporteContado smallmoney,
	@IDMedioPago tinyint,
	@Cuotas tinyint,
	@Operacion varchar(20), 
	@ImporteCuentaCorriente smallmoney,
	@IDCuentaCorrienteCaja int,
	@FechaHoraSistema smalldatetime,
	@IDUsuarioModificacion smallint AS

	IF @ImporteContado > 0 AND @IDCuentaCorrienteCaja IS NULL
		BEGIN
		RAISERROR ('No se especificó la Caja.', 19, 1) WITH LOG
		RETURN
		END

	IF ISNULL(@Realizado, 0) = 0
		UPDATE ViajeDetalle
			SET Realizado = @Realizado, AsientoIdentificacion = NULL, ForzarDebito = @ForzarDebito, Entregada = @Entregada, EntregadaFechaHora = @EntregadaFechaHora, Retira = @Retira, FacturaNumero = @FacturaNumero, ImporteContado = @ImporteContado, IDMedioPago = @IDMedioPago, Cuotas = @Cuotas, Operacion = @Operacion, ImporteCuentaCorriente = @ImporteCuentaCorriente, IDCuentaCorrienteCaja = @IDCuentaCorrienteCaja, FechaHoraModificacion = @FechaHoraSistema, IDUsuarioModificacion = @IDUsuarioModificacion, AsistenciaFechaHora = NULL
			WHERE FechaHora = @FechaHora_FILTER AND IDRuta = @IDRuta_FILTER AND Indice = @Indice_FILTER
	ELSE
		UPDATE ViajeDetalle
			SET Realizado = @Realizado, ForzarDebito = @ForzarDebito, Entregada = @Entregada, EntregadaFechaHora = @EntregadaFechaHora, Retira = @Retira, FacturaNumero = @FacturaNumero, ImporteContado = @ImporteContado, IDMedioPago = @IDMedioPago, Cuotas = @Cuotas, Operacion = @Operacion, ImporteCuentaCorriente = @ImporteCuentaCorriente, IDCuentaCorrienteCaja = @IDCuentaCorrienteCaja, FechaHoraModificacion = @FechaHoraSistema, IDUsuarioModificacion = @IDUsuarioModificacion, AsistenciaFechaHora = @FechaHoraSistema
			WHERE FechaHora = @FechaHora_FILTER AND IDRuta = @IDRuta_FILTER AND Indice = @Indice_FILTER

GO



------------------------------------------------------------------------------------------
-- VIAJEDETALLE_DELETE
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ViajeDetalle_Delete'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ViajeDetalle_Delete
GO

CREATE PROCEDURE dbo.sp_ViajeDetalle_Delete
	@FechaHora smalldatetime,
	@IDRuta char(20),
	@Indice int AS

	SET XACT_ABORT ON

	BEGIN TRANSACTION
    
    DECLARE @IDViajeDetalle int
    
    SET @IDViajeDetalle = (SELECT IDViajeDetalle FROM ViajeDetalle WHERE FechaHora = @FechaHora AND IDRuta = @IDRuta AND Indice = @Indice)

	--ELIMINO EL DETALLE DEL VIAJE
	DELETE FROM ViajeDetalle
		WHERE FechaHora = @FechaHora AND IDRuta = @IDRuta AND Indice = @Indice

	--ELIMINO LOS MOVIMIENTOS DE DE LA CUENTA CORRIENTE
	DELETE FROM CuentaCorriente
		WHERE Viaje_FechaHora = @FechaHora AND Viaje_IDRuta = @IDRuta AND Viaje_Indice = @Indice

	--ELIMINO EL DETALLE DEL VIAJE DE LA CONEXION
	DELETE FROM ViajeDetalle_Conexion
		WHERE IDViajeDetalle = @IDViajeDetalle OR Conexion_IDViajeDetalle = @IDViajeDetalle

	COMMIT TRANSACTION

	SET XACT_ABORT OFF

GO



------------------------------------------------------------------------------------------
-- VIAJEDETALLE_ASIENTOCOMPARTIDO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ViajeDetalle_AsientoCompartido' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ViajeDetalle_AsientoCompartido
GO

CREATE PROCEDURE dbo.sp_ViajeDetalle_AsientoCompartido
	@FechaHora_FILTER smalldatetime,
	@IDRuta_FILTER char(20),
	@Asiento_FILTER tinyint,
	@AsientoCompartidoCount tinyint OUTPUT AS

	SET @AsientoCompartidoCount = (SELECT COUNT(ViajeDetalle.Asiento) FROM ViajeDetalle WHERE ViajeDetalle.FechaHora = @FechaHora_FILTER AND ViajeDetalle.IDRuta = @IDRuta_FILTER AND ViajeDetalle.Asiento = @Asiento_FILTER)

GO



------------------------------------------------------------------------------------------
-- VIAJEDETALLE_LISTGRID
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ViajeDetalle_ListGrid'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ViajeDetalle_ListGrid
GO

CREATE PROCEDURE dbo.sp_ViajeDetalle_ListGrid
	@FechaHora smalldatetime,
	@IDRuta char(20),
	@FilterEstado tinyint,
	@FilterRealizado tinyint AS

	SELECT ViajeDetalle.IDViajeDetalle, ViajeDetalle.Indice, ViajeDetalle.OcupanteTipo, ViajeDetalle.Estado, ViajeDetalle.AsientoIdentificacion, ViajeDetalle.Realizado, ViajeDetalle.Orden, Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Persona, DocumentoTipo.Nombre AS DocumentoTipoNombre, Persona.DocumentoNumero, ISNULL(ViajeDetalle.ImporteSeguro, 0) + ViajeDetalle.Importe AS Importe, ViajeDetalle.ImporteContado + ViajeDetalle.ImporteCuentaCorriente AS ImportePagado, ISNULL(ViajeDetalle.ImporteSeguro, 0) + ViajeDetalle.Importe - ViajeDetalle.ImporteContado - ViajeDetalle.ImporteCuentaCorriente AS Debe, (CASE ISNULL(ViajeDetalle.Sube, '') WHEN '' THEN Lugar_Origen.Nombre ELSE ViajeDetalle.Sube END) AS Origen, (CASE ISNULL(ViajeDetalle.Baja, '') WHEN '' THEN Lugar_Destino.Nombre ELSE ViajeDetalle.Baja END) AS Destino, ViajeDetalle.ReservaTipo, ViajeDetalle.Facturar, ViajeDetalle.Notas, Persona.ListaPasajero, ViajeDetalle.CreadoEnProgreso, ViajeDetalle.ModificadoEnProgreso
		FROM (((ViajeDetalle INNER JOIN Persona ON ViajeDetalle.IDPersona = Persona.IDPersona) LEFT JOIN DocumentoTipo ON Persona.IDDocumentoTipo = DocumentoTipo.IDDocumentoTipo) INNER JOIN Lugar AS Lugar_Origen ON ViajeDetalle.IDOrigen = Lugar_Origen.IDLugar) INNER JOIN Lugar AS Lugar_Destino ON ViajeDetalle.IDDestino = Lugar_Destino.IDLugar
		WHERE ViajeDetalle.FechaHora = @FechaHora AND ViajeDetalle.IDRuta = @IDRuta
			AND ((@FilterEstado = 0) OR (@FilterEstado = 1 AND ViajeDetalle.Estado = '1CO') OR (@FilterEstado = 2 AND ViajeDetalle.Estado = '2CD') OR (@FilterEstado = 3 AND ViajeDetalle.Estado = '3CA'))
			AND ((@FilterRealizado = 0) OR (@FilterRealizado = 1 AND ViajeDetalle.Realizado IS NULL) OR (@FilterRealizado = 2 AND ViajeDetalle.Realizado = 1) OR (@FilterRealizado = 3 AND ViajeDetalle.Realizado = 0))

GO



------------------------------------------------------------------------------------------
-- VIAJEDETALLE_LISTGRID_WITHSALDO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ViajeDetalle_ListGrid_WithSaldo'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ViajeDetalle_ListGrid_WithSaldo
GO

CREATE PROCEDURE dbo.sp_ViajeDetalle_ListGrid_WithSaldo
	@FechaHora smalldatetime,
	@IDRuta char(20),
	@FilterEstado tinyint,
	@FilterRealizado tinyint AS

	SELECT ViajeDetalle.IDViajeDetalle, ViajeDetalle.Indice, ViajeDetalle.OcupanteTipo, ViajeDetalle.Estado, ViajeDetalle.AsientoIdentificacion
			, ViajeDetalle.Realizado, ViajeDetalle.Orden, Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Persona
			, DocumentoTipo.Nombre AS DocumentoTipoNombre, Persona.DocumentoNumero, ViajeDetalle.Importe, ViajeDetalle.ImporteContado + ViajeDetalle.ImporteCuentaCorriente AS ImportePagado
			, ViajeDetalle.Importe - ViajeDetalle.ImporteContado - ViajeDetalle.ImporteCuentaCorriente AS Debe, (CASE ISNULL(ViajeDetalle.Sube, '') WHEN '' THEN Lugar_Origen.Nombre ELSE ViajeDetalle.Sube END) AS Origen
			, (CASE ISNULL(ViajeDetalle.Baja, '') WHEN '' THEN Lugar_Destino.Nombre ELSE ViajeDetalle.Baja END) AS Destino, ViajeDetalle.ReservaTipo, ViajeDetalle.Facturar, ViajeDetalle.Notas
			, Persona.ListaPasajero, ViajeDetalle.CreadoEnProgreso, ViajeDetalle.ModificadoEnProgreso
			, (SELECT SUM(Importe)
					FROM CuentaCorriente
					WHERE CuentaCorriente.IDPersona = (CASE ISNULL(ViajeDetalle.IDPersonaCuentaCorriente, 0) WHEN 0 THEN ViajeDetalle.IDPersona ELSE ViajeDetalle.IDPersonaCuentaCorriente END)
					) AS SaldoActual
		FROM (((ViajeDetalle INNER JOIN Persona ON ViajeDetalle.IDPersona = Persona.IDPersona)
			LEFT JOIN DocumentoTipo ON Persona.IDDocumentoTipo = DocumentoTipo.IDDocumentoTipo)
			INNER JOIN Lugar AS Lugar_Origen ON ViajeDetalle.IDOrigen = Lugar_Origen.IDLugar)
			INNER JOIN Lugar AS Lugar_Destino ON ViajeDetalle.IDDestino = Lugar_Destino.IDLugar
		WHERE ViajeDetalle.FechaHora = @FechaHora AND ViajeDetalle.IDRuta = @IDRuta
			AND ((@FilterEstado = 0) OR (@FilterEstado = 1 AND ViajeDetalle.Estado = '1CO') OR (@FilterEstado = 2 AND ViajeDetalle.Estado = '2CD') OR (@FilterEstado = 3 AND ViajeDetalle.Estado = '3CA'))
			AND ((@FilterRealizado = 0) OR (@FilterRealizado = 1 AND ViajeDetalle.Realizado IS NULL) OR (@FilterRealizado = 2 AND ViajeDetalle.Realizado = 1) OR (@FilterRealizado = 3 AND ViajeDetalle.Realizado = 0))

GO



------------------------------------------------------------------------------------------
-- VIAJEDETALLE_LISTGRID_PAQUETE_MULTIPLESPAGOS
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ViajeDetalle_ListGrid_Paquete_MultiplesPagos'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ViajeDetalle_ListGrid_Paquete_MultiplesPagos
GO

CREATE PROCEDURE dbo.sp_ViajeDetalle_ListGrid_Paquete_MultiplesPagos
	@FechaHora smalldatetime,
	@IDRuta char(20),
	@FilterEstado tinyint,
	@FilterRealizado tinyint AS
	
	DECLARE @IDCuentaCorrienteGrupo_Credito int
	
	SET @IDCuentaCorrienteGrupo_Credito = (SELECT NumeroEntero FROM Parametro WHERE IDParametro = 'CuentaCorrienteGrupo_ID_ViajeCredito')

	SELECT ViajeDetalle.IDViajeDetalle, ViajeDetalle.Indice, ViajeDetalle.OcupanteTipo, ViajeDetalle.Estado, ViajeDetalle.AsientoIdentificacion
			, ViajeDetalle.Realizado, ViajeDetalle.Orden, Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Persona
			, DocumentoTipo.Nombre AS DocumentoTipoNombre, Persona.DocumentoNumero, ISNULL(ViajeDetalle.ImporteSeguro, 0) + ViajeDetalle.Importe AS Importe
			, (SELECT ISNULL(SUM(Importe), 0)
					FROM CuentaCorriente
					WHERE CuentaCorriente.IDPersona = (CASE ISNULL(ViajeDetalle.IDPersonaCuentaCorriente, 0) WHEN 0 THEN ViajeDetalle.IDPersona ELSE ViajeDetalle.IDPersonaCuentaCorriente END)
						AND CuentaCorriente.Viaje_FechaHora = ViajeDetalle.FechaHora
						AND CuentaCorriente.Viaje_IDRuta = ViajeDetalle.IDRuta
						AND CuentaCorriente.Viaje_Indice = ViajeDetalle.Indice
						AND CuentaCorriente.IDCuentaCorrienteGrupo = @IDCuentaCorrienteGrupo_Credito
					) AS ImportePagado
			, (CASE ISNULL(ViajeDetalle.Sube, '') WHEN '' THEN Lugar_Origen.Nombre ELSE ViajeDetalle.Sube END) AS Origen
			, (CASE ISNULL(ViajeDetalle.Baja, '') WHEN '' THEN Lugar_Destino.Nombre ELSE ViajeDetalle.Baja END) AS Destino
			, ViajeDetalle.ReservaTipo, ViajeDetalle.Facturar, ViajeDetalle.Notas, Persona.ListaPasajero, ViajeDetalle.CreadoEnProgreso, ViajeDetalle.ModificadoEnProgreso
		FROM (((ViajeDetalle INNER JOIN Persona ON ViajeDetalle.IDPersona = Persona.IDPersona)
			LEFT JOIN DocumentoTipo ON Persona.IDDocumentoTipo = DocumentoTipo.IDDocumentoTipo)
			INNER JOIN Lugar AS Lugar_Origen ON ViajeDetalle.IDOrigen = Lugar_Origen.IDLugar)
			INNER JOIN Lugar AS Lugar_Destino ON ViajeDetalle.IDDestino = Lugar_Destino.IDLugar
		WHERE ViajeDetalle.FechaHora = @FechaHora AND ViajeDetalle.IDRuta = @IDRuta
			AND ((@FilterEstado = 0) OR (@FilterEstado = 1 AND ViajeDetalle.Estado = '1CO') OR (@FilterEstado = 2 AND ViajeDetalle.Estado = '2CD') OR (@FilterEstado = 3 AND ViajeDetalle.Estado = '3CA'))
			AND ((@FilterRealizado = 0) OR (@FilterRealizado = 1 AND ViajeDetalle.Realizado IS NULL) OR (@FilterRealizado = 2 AND ViajeDetalle.Realizado = 1) OR (@FilterRealizado = 3 AND ViajeDetalle.Realizado = 0))

GO



------------------------------------------------------------------------------------------
-- VIAJEDETALLE_LISTGRID_PAQUETE_MULTIPLESPAGOS_WITHSALDO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ViajeDetalle_ListGrid_Paquete_MultiplesPagos_WithSaldo'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ViajeDetalle_ListGrid_Paquete_MultiplesPagos_WithSaldo
GO

CREATE PROCEDURE dbo.sp_ViajeDetalle_ListGrid_Paquete_MultiplesPagos_WithSaldo
	@FechaHora smalldatetime,
	@IDRuta char(20),
	@FilterEstado tinyint,
	@FilterRealizado tinyint AS

	DECLARE @IDCuentaCorrienteGrupo_Credito int
	
	SET @IDCuentaCorrienteGrupo_Credito = (SELECT NumeroEntero FROM Parametro WHERE IDParametro = 'CuentaCorrienteGrupo_ID_ViajeCredito')

	SELECT ViajeDetalle.IDViajeDetalle, ViajeDetalle.Indice, ViajeDetalle.OcupanteTipo, ViajeDetalle.Estado, ViajeDetalle.AsientoIdentificacion
		, ViajeDetalle.Realizado, ViajeDetalle.Orden
		, Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Persona, DocumentoTipo.Nombre AS DocumentoTipoNombre
		, Persona.DocumentoNumero, ISNULL(ViajeDetalle.ImporteSeguro, 0) + ViajeDetalle.Importe AS Importe
		, (SELECT ISNULL(SUM(Importe), 0)
				FROM CuentaCorriente
				WHERE CuentaCorriente.IDPersona = (CASE ISNULL(ViajeDetalle.IDPersonaCuentaCorriente, 0) WHEN 0 THEN ViajeDetalle.IDPersona ELSE ViajeDetalle.IDPersonaCuentaCorriente END)
					AND CuentaCorriente.Viaje_FechaHora = ViajeDetalle.FechaHora
					AND CuentaCorriente.Viaje_IDRuta = ViajeDetalle.IDRuta
					AND CuentaCorriente.Viaje_Indice = ViajeDetalle.Indice
					AND CuentaCorriente.IDCuentaCorrienteGrupo = @IDCuentaCorrienteGrupo_Credito
			) AS ImportePagado
		, (CASE ISNULL(ViajeDetalle.Sube, '') WHEN '' THEN Lugar_Origen.Nombre ELSE ViajeDetalle.Sube END) AS Origen
		, (CASE ISNULL(ViajeDetalle.Baja, '') WHEN '' THEN Lugar_Destino.Nombre ELSE ViajeDetalle.Baja END) AS Destino
		, ViajeDetalle.ReservaTipo, ViajeDetalle.Facturar, ViajeDetalle.Notas, Persona.ListaPasajero, ViajeDetalle.CreadoEnProgreso, ViajeDetalle.ModificadoEnProgreso
		, (SELECT SUM(Importe)
				FROM CuentaCorriente
				WHERE CuentaCorriente.IDPersona = (CASE ISNULL(ViajeDetalle.IDPersonaCuentaCorriente, 0) WHEN 0 THEN ViajeDetalle.IDPersona ELSE ViajeDetalle.IDPersonaCuentaCorriente END)
			) AS SaldoActual
		FROM (((ViajeDetalle INNER JOIN Persona ON ViajeDetalle.IDPersona = Persona.IDPersona)
			LEFT JOIN DocumentoTipo ON Persona.IDDocumentoTipo = DocumentoTipo.IDDocumentoTipo)
			INNER JOIN Lugar AS Lugar_Origen ON ViajeDetalle.IDOrigen = Lugar_Origen.IDLugar)
			INNER JOIN Lugar AS Lugar_Destino ON ViajeDetalle.IDDestino = Lugar_Destino.IDLugar
		WHERE ViajeDetalle.FechaHora = @FechaHora AND ViajeDetalle.IDRuta = @IDRuta
			AND ((@FilterEstado = 0) OR (@FilterEstado = 1 AND ViajeDetalle.Estado = '1CO') OR (@FilterEstado = 2 AND ViajeDetalle.Estado = '2CD') OR (@FilterEstado = 3 AND ViajeDetalle.Estado = '3CA'))
			AND ((@FilterRealizado = 0) OR (@FilterRealizado = 1 AND ViajeDetalle.Realizado IS NULL) OR (@FilterRealizado = 2 AND ViajeDetalle.Realizado = 1) OR (@FilterRealizado = 3 AND ViajeDetalle.Realizado = 0))

GO



------------------------------------------------------------------------------------------
-- VIAJEDETALLE_LISTPASAJERO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ViajeDetalle_ListPasajero'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ViajeDetalle_ListPasajero
GO

CREATE PROCEDURE dbo.sp_ViajeDetalle_ListPasajero
	@FechaHora smalldatetime,
	@IDRuta char(20) AS

	SELECT ViajeDetalle.IDViajeDetalle, ViajeDetalle.Indice, ViajeDetalle.Estado, ViajeDetalle.Asiento, ViajeDetalle.Orden, Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Persona, Lugar_Origen.Nombre AS Origen, ViajeDetalle.Sube, Lugar_Destino.Nombre AS Destino, ViajeDetalle.Baja, ViajeDetalle.ReservaTipo, Persona.ListaPasajero
		FROM ((ViajeDetalle INNER JOIN Persona ON ViajeDetalle.IDPersona = Persona.IDPersona) INNER JOIN Lugar AS Lugar_Origen ON ViajeDetalle.IDOrigen = Lugar_Origen.IDLugar) INNER JOIN Lugar AS Lugar_Destino ON ViajeDetalle.IDDestino = Lugar_Destino.IDLugar
		WHERE ViajeDetalle.FechaHora = @FechaHora AND ViajeDetalle.IDRuta = @IDRuta AND ViajeDetalle.OcupanteTipo = 'PA'
		ORDER BY ViajeDetalle.OcupanteTipo DESC, ViajeDetalle.Estado, ViajeDetalle.Orden

GO



------------------------------------------------------------------------------------------
-- VIAJEDETALLE_LISTBYTIPOESTADO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ViajeDetalle_ListByTipoEstado' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ViajeDetalle_ListByTipoEstado
GO

CREATE PROCEDURE dbo.sp_ViajeDetalle_ListByTipoEstado
	@FechaHora_FILTER smalldatetime,
	@IDRuta_FILTER char(20),
	@OcupanteTipo_FILTER char(2),
	@Estado_FILTER char(3) AS

	SELECT ViajeDetalle.IDViajeDetalle, ViajeDetalle.FechaHora, ViajeDetalle.IDRuta, ViajeDetalle.Indice, ViajeDetalle.OcupanteTipo, ViajeDetalle.Estado, ViajeDetalle.Prioridad, ViajeDetalle.Orden, ViajeDetalle.Asiento, ViajeDetalle.Realizado, ViajeDetalle.IDPersona, ViajeDetalle.IDListaPrecio, ViajeDetalle.IDOrigen, ViajeDetalle.Sube, ViajeDetalle.IDDestino, ViajeDetalle.Baja, ViajeDetalle.ValorDeclarado, ViajeDetalle.ImporteSeguro, ViajeDetalle.Importe, ViajeDetalle.ImporteContado, ViajeDetalle.ImporteCuentaCorriente, ViajeDetalle.ImprimirSaldo, ViajeDetalle.IDCuentaCorrienteCaja, ViajeDetalle.ForzarDebito, ViajeDetalle.IDPersonaCuentaCorriente, ViajeDetalle.FacturaNumero, ViajeDetalle.Facturar, ViajeDetalle.FacturarNotas, ViajeDetalle.IDPersonaRecibe, ViajeDetalle.PagaQuienRecibe, ViajeDetalle.Recibe, ViajeDetalle.Descripcion, ViajeDetalle.Telefono, ViajeDetalle.DejarTraer, ViajeDetalle.Entregada, ViajeDetalle.EntregadaFechaHora, ViajeDetalle.ReservaTipo, ViajeDetalle.CreadoEnProgreso, ViajeDetalle.ModificadoEnProgreso, ViajeDetalle.ReservadoPor, ViajeDetalle.CanceladoPor, ViajeDetalle.CanceladoFechaHora, ViajeDetalle.Notas, ViajeDetalle.FechaHoraCreacion, ViajeDetalle.IDUsuarioCreacion, ViajeDetalle.FechaHoraModificacion, ViajeDetalle.IDUsuarioModificacion, ViajeDetalle.IDUsuarioCancelacion
		FROM ViajeDetalle
		WHERE ViajeDetalle.FechaHora = @FechaHora_FILTER AND ViajeDetalle.IDRuta = @IDRuta_FILTER AND ViajeDetalle.OcupanteTipo = @OcupanteTipo_FILTER AND ViajeDetalle.Estado = @Estado_FILTER

GO



------------------------------------------------------------------------------------------
-- VIAJEDETALLE_FINDDUPLICATE
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ViajeDetalle_FindDuplicate' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ViajeDetalle_FindDuplicate
GO

CREATE PROCEDURE dbo.sp_ViajeDetalle_FindDuplicate
	@FechaHora_FILTER smalldatetime,
	@IDRuta_FILTER char(20),
	@IDPersona_FILTER int,
	@FechaHoraOriginal_FILTER smalldatetime,
	@IDRutaOriginal_FILTER char(20),
	@IndiceOriginal_FILTER int AS

	SELECT ViajeDetalle.FechaHora
		FROM ViajeDetalle
		WHERE convert(char(10), ViajeDetalle.FechaHora, 111) = convert(char(10), @FechaHora_FILTER, 111) 
			AND ViajeDetalle.IDRuta = @IDRuta_FILTER 
			AND ViajeDetalle.IDPersona = @IDPersona_FILTER 
			AND ViajeDetalle.Estado = '1CO'
			AND ViajeDetalle.OcupanteTipo = 'PA'
			AND NOT (ViajeDetalle.FechaHora = @FechaHoraOriginal_FILTER AND ViajeDetalle.IDRuta = @IDRutaOriginal_FILTER AND ViajeDetalle.Indice = @IndiceOriginal_FILTER)
		ORDER BY ViajeDetalle.FechaHora

GO



------------------------------------------------------------------------------------------
-- VIAJEDETALLE_BUSCAVUELTA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ViajeDetalle_BuscaVuelta'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ViajeDetalle_BuscaVuelta
GO

CREATE PROCEDURE dbo.sp_ViajeDetalle_BuscaVuelta
	@FechaHora_FILTER smalldatetime,
	@IDRuta_FILTER char(20),
	@IDPersona_FILTER int,
	@IDRutaEspecial_FILTER char(20) AS

	SELECT ViajeDetalle.FechaHora, ViajeDetalle.IDRuta
		FROM ViajeDetalle
		WHERE ViajeDetalle.IDPersona = @IDPersona_FILTER
			AND ViajeDetalle.Estado = '1CO'
			AND ViajeDetalle.OcupanteTipo = 'PA'
			AND ViajeDetalle.Realizado IS NULL
			AND ViajeDetalle.IDRuta <> @IDRutaEspecial_FILTER
			AND convert(char(10), ViajeDetalle.FechaHora, 111) = convert(char(10), @FechaHora_FILTER, 111)
			AND convert(char(8), ViajeDetalle.FechaHora, 108) > convert(char(8), @FechaHora_FILTER, 108)
		ORDER BY ViajeDetalle.FechaHora

GO



------------------------------------------------------------------------------------------
-- VIAJEDETALLE_LIST
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ViajeDetalle_List' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ViajeDetalle_List
GO

CREATE PROCEDURE dbo.sp_ViajeDetalle_List
	@FechaHora_FILTER smalldatetime,
	@IDRuta_FILTER char(20) AS

	SELECT ViajeDetalle.IDViajeDetalle, ViajeDetalle.IDViaje, ViajeDetalle.FechaHora, ViajeDetalle.IDRuta, ViajeDetalle.Indice, ViajeDetalle.OcupanteTipo, ViajeDetalle.Estado, ViajeDetalle.Prioridad, ViajeDetalle.Orden, ViajeDetalle.Asiento, ViajeDetalle.Realizado, ViajeDetalle.IDPersona, ViajeDetalle.IDListaPrecio, ViajeDetalle.IDOrigen, ViajeDetalle.Sube, ViajeDetalle.IDDestino, ViajeDetalle.Baja, ViajeDetalle.ValorDeclarado, ViajeDetalle.ImporteSeguro, ViajeDetalle.Importe, ViajeDetalle.ImporteContado, ViajeDetalle.IDMedioPago, ViajeDetalle.Cuotas, ViajeDetalle.Operacion, ViajeDetalle.ImporteCuentaCorriente, ViajeDetalle.ImprimirSaldo, ViajeDetalle.IDCuentaCorrienteCaja, ViajeDetalle.ForzarDebito, ViajeDetalle.IDPersonaCuentaCorriente, ViajeDetalle.FacturaNumero, ViajeDetalle.Facturar, ViajeDetalle.FacturarNotas, ViajeDetalle.IDPersonaRecibe, ViajeDetalle.PagaQuienRecibe, ViajeDetalle.Recibe, ViajeDetalle.Descripcion, ViajeDetalle.Domicilio, ViajeDetalle.Horario, ViajeDetalle.Telefono, ViajeDetalle.DejarTraer, ViajeDetalle.Entregada, ViajeDetalle.EntregadaFechaHora, ViajeDetalle.ReservaTipo, ViajeDetalle.CreadoEnProgreso, ViajeDetalle.ModificadoEnProgreso, ViajeDetalle.ReservadoPor, ViajeDetalle.CanceladoPor, ViajeDetalle.CanceladoFechaHora, ViajeDetalle.Notas, ViajeDetalle.FechaHoraCreacion, ViajeDetalle.IDUsuarioCreacion, ViajeDetalle.FechaHoraModificacion, ViajeDetalle.IDUsuarioModificacion, ViajeDetalle.IDUsuarioCancelacion
		FROM ViajeDetalle
		WHERE ViajeDetalle.FechaHora = @FechaHora_FILTER AND ViajeDetalle.IDRuta = @IDRuta_FILTER
		ORDER BY ViajeDetalle.Indice

GO



------------------------------------------------------------------------------------------
-- VIAJEDETALLE_ASISTENCIA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ViajeDetalle_Asistencia' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ViajeDetalle_Asistencia
GO

CREATE PROCEDURE dbo.sp_ViajeDetalle_Asistencia
	@FechaHora_FILTER smalldatetime,
	@IDRuta_FILTER char(20) AS

	SELECT ViajeDetalle.Indice, ViajeDetalle.Orden, Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Pasajero, ViajeDetalle.Importe, ViajeDetalle.ImporteContado, ViajeDetalle.Realizado
		FROM ViajeDetalle INNER JOIN Persona ON ViajeDetalle.IDPersona = Persona.IDPersona
		WHERE ViajeDetalle.FechaHora = @FechaHora_FILTER AND ViajeDetalle.IDRuta = @IDRuta_FILTER AND ViajeDetalle.Realizado IS NULL AND ViajeDetalle.OcupanteTipo = 'PA' AND ViajeDetalle.Estado = '1CO'

GO



------------------------------------------------------------------------------------------
-- VIAJEDETALLE_RESERVATIPODATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ViajeDetalle_ReservaTipoData' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ViajeDetalle_ReservaTipoData
GO

CREATE PROCEDURE dbo.sp_ViajeDetalle_ReservaTipoData
	@FechaHora_FILTER smalldatetime,
	@IDRuta_FILTER char(20),
	@OcupanteTipo_FILTER char(2),
	@IDPersona_FILTER int,
	@ReservaTipo_FILTER char(2) AS

	SELECT ViajeDetalle.IDViajeDetalle, ViajeDetalle.FechaHora, ViajeDetalle.IDRuta, ViajeDetalle.Indice, ViajeDetalle.OcupanteTipo, ViajeDetalle.Estado, ViajeDetalle.Prioridad, ViajeDetalle.Orden, ViajeDetalle.Asiento, ViajeDetalle.Realizado, ViajeDetalle.IDPersona, ViajeDetalle.IDListaPrecio, ViajeDetalle.IDOrigen, ViajeDetalle.Sube, ViajeDetalle.IDDestino, ViajeDetalle.Baja, ViajeDetalle.Importe, ViajeDetalle.ImporteContado, ViajeDetalle.ImporteCuentaCorriente, ViajeDetalle.ImprimirSaldo, ViajeDetalle.IDCuentaCorrienteCaja, ViajeDetalle.ForzarDebito, ViajeDetalle.IDPersonaCuentaCorriente, ViajeDetalle.FacturaNumero, ViajeDetalle.Facturar, ViajeDetalle.FacturarNotas, ViajeDetalle.Recibe, ViajeDetalle.Descripcion, ViajeDetalle.Telefono, ViajeDetalle.DejarTraer, ViajeDetalle.Entregada, ViajeDetalle.EntregadaFechaHora, ViajeDetalle.ReservaTipo, ViajeDetalle.CreadoEnProgreso, ViajeDetalle.ModificadoEnProgreso, ViajeDetalle.ReservadoPor, ViajeDetalle.CanceladoPor, ViajeDetalle.CanceladoFechaHora, ViajeDetalle.Notas, ViajeDetalle.FechaHoraCreacion, ViajeDetalle.IDUsuarioCreacion, ViajeDetalle.FechaHoraModificacion, ViajeDetalle.IDUsuarioModificacion, ViajeDetalle.IDUsuarioCancelacion
		FROM ViajeDetalle
		WHERE ViajeDetalle.FechaHora = @FechaHora_FILTER AND ViajeDetalle.IDRuta = @IDRuta_FILTER AND ViajeDetalle.OcupanteTipo = @OcupanteTipo_FILTER AND ViajeDetalle.IDPersona = @IDPersona_FILTER AND ViajeDetalle.ReservaTipo = @ReservaTipo_FILTER

GO



------------------------------------------------------------------------------------------
-- VIAJEDETALLE_ASIENTO_LIST
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ViajeDetalle_Asiento_List' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ViajeDetalle_Asiento_List
GO

CREATE PROCEDURE dbo.sp_ViajeDetalle_Asiento_List
	@FechaHora_FILTER smalldatetime,
	@IDRuta_FILTER char(20),
	@OcupanteTipo_FILTER char(2),
	@EstadoExclude_FILTER char(3) AS

	SELECT ViajeDetalle.IDViajeDetalle, ViajeDetalle.FechaHora, ViajeDetalle.IDRuta, ViajeDetalle.Indice, ViajeDetalle.Asiento, RutaDetalleOrigen.Indice AS IndiceOrigen, RutaDetalleDestino.Indice AS IndiceDestino, ViajeDetalle.Estado, ViajeDetalle.Realizado
		FROM RutaDetalle AS RutaDetalleOrigen INNER JOIN (RutaDetalle AS RutaDetalleDestino INNER JOIN ViajeDetalle ON RutaDetalleDestino.IDLugar = ViajeDetalle.IDDestino AND RutaDetalleDestino.IDRuta = ViajeDetalle.IDRuta) ON RutaDetalleOrigen.IDRuta = ViajeDetalle.IDRuta AND RutaDetalleOrigen.IDLugar = ViajeDetalle.IDOrigen
		WHERE ViajeDetalle.FechaHora = @FechaHora_FILTER AND ViajeDetalle.IDRuta = @IDRuta_FILTER AND (ViajeDetalle.Estado IS NULL OR ViajeDetalle.Estado <> @EstadoExclude_FILTER) AND ViajeDetalle.OcupanteTipo = @OcupanteTipo_FILTER
		ORDER BY ViajeDetalle.Prioridad, ViajeDetalle.Indice

GO



------------------------------------------------------------------------------------------
-- VIAJEDETALLE_BYPERSONAHORARIO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ViajeDetalle_ByPersonaHorario' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ViajeDetalle_ByPersonaHorario
GO

CREATE PROCEDURE dbo.sp_ViajeDetalle_ByPersonaHorario
	@IDPersona_FILTER int,
	@DiaSemana_FILTER tinyint,
	@Hora_FILTER char(8),
	@IDRuta_FILTER char(20),
	@FechaDesde_FILTER char(10),
	@FechaHasta_FILTER char(10) AS

	SELECT ViajeLista.IDViaje, ViajeLista.FechaHora, ViajeDetalleLista.IDViajeDetalle, ViajeDetalleLista.Indice
		FROM
			(SELECT Viaje.IDViaje, Viaje.FechaHora, Viaje.IDRuta
				FROM Viaje
				WHERE Viaje.DiaSemanaBase = @DiaSemana_FILTER
					AND convert(char(8), Viaje.FechaHora, 108) = @Hora_FILTER
					AND Viaje.IDRuta = @IDRuta_FILTER
					AND datediff(minute, getdate(), Viaje.FechaHora) >= 0
					AND (@FechaDesde_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) >= @FechaDesde_FILTER)
					AND (@FechaHasta_FILTER IS NULL OR convert(char(10), Viaje.FechaHora, 111) <= @FechaHasta_FILTER)
			) AS ViajeLista
			LEFT JOIN
			(SELECT ViajeDetalle.IDViajeDetalle, ViajeDetalle.FechaHora, ViajeDetalle.IDRuta, ViajeDetalle.Indice
				FROM ViajeDetalle
				WHERE ViajeDetalle.IDPersona = @IDPersona_FILTER
					AND convert(char(8), ViajeDetalle.FechaHora, 108) = @Hora_FILTER
					AND ViajeDetalle.IDRuta = @IDRuta_FILTER
					AND datediff(minute, getdate(), ViajeDetalle.FechaHora) >= 0
					AND (@FechaDesde_FILTER IS NULL OR convert(char(10), ViajeDetalle.FechaHora, 111) >= @FechaDesde_FILTER)
					AND (@FechaHasta_FILTER IS NULL OR convert(char(10), ViajeDetalle.FechaHora, 111) <= @FechaHasta_FILTER)
					AND (ViajeDetalle.ReservaTipo IS NULL OR ViajeDetalle.ReservaTipo = 'FI')
			) AS ViajeDetalleLista
			ON ViajeLista.FechaHora = ViajeDetalleLista.FechaHora AND ViajeLista.IDRuta = ViajeDetalleLista.IDRuta
		ORDER BY ViajeLista.FechaHora, ViajeLista.IDRuta

GO



------------------------------------------------------------------------------------------
-- VIAJEDETALLE_BYPERSONARUTA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ViajeDetalle_ByPersonaRuta' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ViajeDetalle_ByPersonaRuta
GO

CREATE PROCEDURE dbo.sp_ViajeDetalle_ByPersonaRuta
	@IDPersona_FILTER int,
	@IDRuta_FILTER char(20) AS

	SELECT ViajeDetalle.IDViajeDetalle, Viaje.FechaHora, ViajeDetalle.Indice, ViajeDetalle.IDOrigen, ViajeDetalle.Sube, ViajeDetalle.IDDestino, ViajeDetalle.Baja
		FROM Viaje INNER JOIN ViajeDetalle ON Viaje.FechaHora = ViajeDetalle.FechaHora AND Viaje.IDRuta = ViajeDetalle.IDRuta
		WHERE ViajeDetalle.IDPersona = @IDPersona_FILTER AND Viaje.IDRuta = @IDRuta_FILTER AND datediff(minute, getdate(), Viaje.FechaHora) >= 0
		ORDER BY Viaje.FechaHora

GO



------------------------------------------------------------------------------------------
-- VIAJEDETALLE_PAGOS
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ViajeDetalle_Pagos'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ViajeDetalle_Pagos
GO

CREATE PROCEDURE dbo.sp_ViajeDetalle_Pagos
	@FechaHora smalldatetime,
	@IDRuta char(20),
	@Indice int,
	@IDCuentaCorrienteGrupo_ViajeCredito int AS

	SELECT IDMovimiento
		FROM CuentaCorriente
		WHERE IDCuentaCorrienteGrupo = @IDCuentaCorrienteGrupo_ViajeCredito AND Viaje_FechaHora = @FechaHora AND Viaje_IDRuta = @IDRuta AND Viaje_Indice = @Indice

GO



------------------------------------------------------------------------------------------
-- VIAJEDETALLE_COMISION
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ViajeDetalle_Comision'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ViajeDetalle_Comision
GO

CREATE PROCEDURE dbo.sp_ViajeDetalle_Comision
	@FechaHora smalldatetime,
	@IDRuta char(20),
	@Indice int AS

	SELECT FechaHora, IDRuta, Indice, RendicionFechaHora, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM ViajeDetalle_Comision
		WHERE FechaHora = @FechaHora AND IDRuta = @IDRuta AND Indice = @Indice

GO