@ECHO OFF

CLS

SET ISQL_EXE="sqlcmd"
SET SERVER_NAME=localhost
SET USERID=sa
SET PASSWORD=
SET DATABASE=CSTransporte_LobosBus_Sample

ECHO ------------------------------------------
ECHO - Generando Stored Procedures de Alarmas -
ECHO ------------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "Alarma.sql"
ECHO.
ECHO.
ECHO -------------------------------------------------
ECHO - Generando Stored Procedures de Condicion IVA  -
ECHO -------------------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "CondicionIVA.sql"
ECHO.
ECHO.
ECHO -------------------------------------------------
ECHO - Generando Stored Procedures de Conductor Ruta -
ECHO -------------------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "ConductorRuta.sql"
ECHO.
ECHO.
ECHO --------------------------------------------
ECHO - Generando Stored Procedures de Contactos -
ECHO --------------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "Contacto.sql"
ECHO.
ECHO.
ECHO -----------------------------------------------------
ECHO - Generando Stored Procedures de Cuentas Corrientes -
ECHO -----------------------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "CuentaCorriente.sql"
ECHO.
ECHO.
REM ECHO ---------------------------------------------------------------
REM ECHO - Generando Stored Procedures de Cuentas Corrientes (History) -
REM ECHO ---------------------------------------------------------------
REM CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "CuentaCorriente_History.sql"
REM ECHO.
REM ECHO.
ECHO -------------------------------------------
ECHO - Generando Stored Procedures de Feriados -
ECHO -------------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "Feriado.sql"
ECHO.
ECHO.
ECHO -------------------------------------------
ECHO - Generando Stored Procedures de Francos  -
ECHO -------------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "Franco.sql"
ECHO.
ECHO.
REM ECHO ------------------------------------------
REM ECHO - Generando Stored Procedures de History -
REM ECHO ------------------------------------------
REM CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "History.sql"
REM ECHO.
REM ECHO.
ECHO -------------------------------------------
ECHO - Generando Stored Procedures de Horarios -
ECHO -------------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "Horario.sql"
ECHO.
ECHO.
ECHO ----------------------------------------------------
ECHO - Generando Stored Procedures de Listas de Precios -
ECHO ----------------------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "ListaPrecio.sql"
ECHO.
ECHO.
ECHO ------------------------------------------
ECHO - Generando Stored Procedures de Lugares -
ECHO ------------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "Lugar.sql"
ECHO.
ECHO.
ECHO -------------------------------------------------
ECHO - Generando Stored Procedures de Medios de Pago -
ECHO -------------------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "MedioPago.sql"
ECHO.
ECHO.
ECHO -------------------------------------------
ECHO - Generando Stored Procedures de Mensajes -
ECHO -------------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "Mensaje.sql"
ECHO.
ECHO.
ECHO ---------------------------------------------
ECHO - Generando Stored Procedures de Parametros -
ECHO ---------------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "Parametro.sql"
ECHO.
ECHO.
ECHO -------------------------------------------
ECHO - Generando Stored Procedures de Personas -
ECHO -------------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "Persona.sql"
ECHO.
ECHO.
ECHO ------------------------------------------------------
ECHO - Generando Stored Procedures de Alarmas de Personas -
ECHO ------------------------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "PersonaAlarma.sql"
ECHO.
ECHO.
ECHO -------------------------------------------------------
ECHO - Generando Stored Procedures de Prepagos de Personas -
ECHO -------------------------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "PersonaPrepago.sql"
ECHO.
ECHO.
ECHO -------------------------------------------
ECHO - Generando Stored Procedures de Reportes -
ECHO -------------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "Report.sql"
ECHO.
ECHO.
ECHO -----------------------------------------------------------------
ECHO - Generando Stored Procedures de Reportes - Planillas de Viajes -
ECHO -----------------------------------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "Report - ViajePlanilla.sql"
ECHO.
ECHO.
ECHO ----------------------------------------
ECHO - Generando Stored Procedures de Rutas -
ECHO ----------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "Ruta.sql"
ECHO.
ECHO.
ECHO ----------------------------------------------------
ECHO - Generando Stored Procedures de Detalles de Rutas -
ECHO ----------------------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "RutaDetalle.sql"
ECHO.
ECHO.
ECHO ----------------------------------------------------------------
ECHO - Generando Stored Procedures de Horarios de Detalles de Rutas -
ECHO ----------------------------------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "RutaDetalleHorario.sql"
ECHO.
ECHO.
ECHO ---------------------------------------------------
ECHO - Generando Stored Procedures de Semáforo General -
ECHO ---------------------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "SemaforoGeneral.sql"
ECHO.
ECHO.
ECHO ---------------------------------------------
ECHO - Generando Stored Procedures de Sucursales -
ECHO ---------------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "Sucursal.sql"
ECHO.
ECHO.
ECHO -------------------------------------------
ECHO - Generando Stored Procedures de Usuarios -
ECHO -------------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "Usuario.sql"
ECHO.
ECHO.
ECHO --------------------------------------------
ECHO - Generando Stored Procedures de Vehiculos -
ECHO --------------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "Vehiculo.sql"
ECHO.
ECHO.
ECHO -------------------------------------------------------------
ECHO - Generando Stored Procedures de Mantenimiento de Vehiculos -
ECHO -------------------------------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "VehiculoMantenimiento.sql"
ECHO.
ECHO.
ECHO -----------------------------------------
ECHO - Generando Stored Procedures de Viajes -
ECHO -----------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "Viaje.sql"
ECHO.
ECHO.
ECHO -----------------------------------------------------
ECHO - Generando Stored Procedures de Detalles de Viajes -
ECHO -----------------------------------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "ViajeDetalle.sql"
ECHO.
ECHO.
ECHO -----------------------
ECHO - Generando Functions -
ECHO -----------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "Functions.sql"
ECHO.
ECHO.
ECHO -------------------------
ECHO - Generando Permissions -
ECHO -------------------------
CALL %ISQL_EXE% -S %SERVER_NAME% -U %USERID% -P %PASSWORD% -d %DATABASE% -i "Permissions.sql"
ECHO.
ECHO.
ECHO.
ECHO /////////////////////////////////////////////////
ECHO //                                             //
ECHO // Generacion de Stored Procedures Finalizada. //
ECHO //                                             //
ECHO /////////////////////////////////////////////////
ECHO.
ECHO.
ECHO ON