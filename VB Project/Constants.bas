Attribute VB_Name = "Constants"
Option Explicit

Public Const PASSWORD_ENCRYPTION_KEY = "9A6B60C91AEAB43C682666E582B49EAB9A5F8F203C1161BD"

Public Const REPORT_DATABASE_DLL_NAME = "crdb_ado.dll"

Public Const DATABASE_IDENTIFIER = ""

'///////////////////////////////////////////////////////////////////////////
'ENTIDAD TIPO
Public Const ENTIDAD_TIPO_PERSONA_CLIENTE = "PC"
Public Const ENTIDAD_TIPO_PERSONA_CONDUCTOR = "PO"
Public Const ENTIDAD_TIPO_PERSONA_ADMINISTRATIVO = "PA"
Public Const ENTIDAD_TIPO_PERSONA_PREPAGO = "PP"
Public Const ENTIDAD_TIPO_VEHICULO = "VE"
Public Const ENTIDAD_TIPO_VIAJE = "VI"
Public Const ENTIDAD_TIPO_VIAJE_DETALLE = "VD"
Public Const ENTIDAD_TIPO_CONTACTO = "CO"

Public Const ENTIDAD_TIPO_PERSONA_CLIENTE_NOMBRE = "Cliente"
Public Const ENTIDAD_TIPO_PERSONA_CONDUCTOR_NOMBRE = "Conductor"
Public Const ENTIDAD_TIPO_PERSONA_ADMINISTRATIVO_NOMBRE = "Administrativo"
Public Const ENTIDAD_TIPO_VEHICULO_NOMBRE = "Veh�culo"
Public Const ENTIDAD_TIPO_VIAJE_NOMBRE = "Viaje"
Public Const ENTIDAD_TIPO_VIAJE_DETALLE_NOMBRE = "Detalle de Viaje"
Public Const ENTIDAD_TIPO_CONTACTO_NOMBRE = "Contacto"

'///////////////////////////////////////////////////////////////////////////
'VIAJE
Public Const VIAJE_ESTADO_ACTIVO = "AC"
Public Const VIAJE_ESTADO_EN_PROGRESO = "EP"
Public Const VIAJE_ESTADO_FINALIZADO = "FI"
Public Const VIAJE_ESTADO_CANCELADO = "CA"

Public Const VIAJE_ESTADO_ACTIVO_NOMBRE = "Activo"
Public Const VIAJE_ESTADO_EN_PROGRESO_NOMBRE = "En Progreso"
Public Const VIAJE_ESTADO_FINALIZADO_NOMBRE = "Finalizado"
Public Const VIAJE_ESTADO_CANCELADO_NOMBRE = "Cancelado"

'///////////////////////////////////////////////////////////////////////////
'VIAJE DETALLE
Public Const VIAJE_DETALLE_ESTADO_CONFIRMADO = "1CO"
Public Const VIAJE_DETALLE_ESTADO_CONDICIONAL = "2CD"
Public Const VIAJE_DETALLE_ESTADO_CANCELADO = "3CA"

Public Const VIAJE_DETALLE_ESTADO_CONFIRMADO_NOMBRE = "Confirmado"
Public Const VIAJE_DETALLE_ESTADO_CONDICIONAL_NOMBRE = "Condicional"
Public Const VIAJE_DETALLE_ESTADO_CANCELADO_NOMBRE = "Cancelado"

Public Const VIAJE_DETALLE_RESERVA_TIPO_STANDARD = "ST"
Public Const VIAJE_DETALLE_RESERVA_TIPO_FIJA = "FI"
Public Const VIAJE_DETALLE_RESERVA_TIPO_INTERNET = "IN"

Public Const VIAJE_DETALLE_RESERVA_TIPO_STANDARD_NOMBRE = "Est�ndard"
Public Const VIAJE_DETALLE_RESERVA_TIPO_FIJA_NOMBRE = "Fija"
Public Const VIAJE_DETALLE_RESERVA_TIPO_INTERNET_NOMBRE = "Internet"

Public Const VIAJE_DETALLE_REALIZADO_UNKNOWN As Long = 0
Public Const VIAJE_DETALLE_REALIZADO_SI As Long = 1
Public Const VIAJE_DETALLE_REALIZADO_NO As Long = 2

'///////////////////////////////////////////////////////////////////////////
'VEHICULO MANTENIMIENTO TIPO
Public Const VEHICULO_MATENIMIENTO_TIPO_KILOMETRAJE = "KI"
Public Const VEHICULO_MATENIMIENTO_TIPO_DIAS = "DI"
Public Const VEHICULO_MATENIMIENTO_TIPO_FECHA = "FE"
Public Const VEHICULO_MATENIMIENTO_TIPO_NINGUNO = "NI"

Public Const VEHICULO_MATENIMIENTO_TIPO_KILOMETRAJE_NOMBRE = "Kilometros"
Public Const VEHICULO_MATENIMIENTO_TIPO_DIAS_NOMBRE = "D�as"
Public Const VEHICULO_MATENIMIENTO_TIPO_FECHA_NOMBRE = "Fecha"
Public Const VEHICULO_MATENIMIENTO_TIPO_NINGUNO_NOMBRE = "Ninguno"

'///////////////////////////////////////////////////////////////////////////
'CALLER ID STRINGS
Public Const CALLERID_NOTAVAILABLE_CODE As String = ""
Public Const CALLERID_NOTAVAILABLE As String = "NO DISPONIBLE"
Public Const CALLERID_PRIVATE_CODE As String = "P"
Public Const CALLERID_PRIVATE As String = "PRIVADO"

'///////////////////////////////////////////////////////////////////////////
'USUARIOS
Public Const USUARIO_GRUPO_ID_ADMINISTRADORES = 1
Public Const USUARIO_ID_ADMINISTRATOR  As Integer = 1
Public Const USUARIO_LOGINNAME_ADMINISTRATOR As String = "administrator"
Public Const USUARIO_NOMBRE_ADMINISTRATOR As String = "Administrador del Sistema"
Public Const USUARIO_PASSWORD_ADMINISTRATOR As String = "C80976C5E33E353E6848ADDCCE00C9F3"
Public Const USUARIO_LOGINNAME_INTERNET As String = "internet"

'///////////////////////////////////////////////////////////////////////////
'TIPO DE OCUPANTE
Public Const OCUPANTE_TIPO_COMISION = "CO"
Public Const OCUPANTE_TIPO_COMISION_NOMBRE = "Comisi�n"
Public Const OCUPANTE_TIPO_PASAJERO = "PA"
Public Const OCUPANTE_TIPO_PASAJERO_NOMBRE = "Pasajero"

'///////////////////////////////////////////////////////////////////////////
'EVENTO TIPO
Public Const EVENTO_TIPO_AGREGAR = "AD"
Public Const EVENTO_TIPO_MODIFICAR = "UP"
Public Const EVENTO_TIPO_ELIMINAR = "DE"

'///////////////////////////////////////////////////////////////////////////
'DOCUMENTO FISCAL ESTADO
Public Const DOCUMENTO_FISCAL_ESTADO_WORKING = "W"      'WORKING ON
Public Const DOCUMENTO_FISCAL_ESTADO_DRAFT = "D"        'SAVED AS DRAFT
Public Const DOCUMENTO_FISCAL_ESTADO_PRINT = "P"        'WAITING FOR PRINT
Public Const DOCUMENTO_FISCAL_ESTADO_NULLED = "A"       'NULLED
Public Const DOCUMENTO_FISCAL_ESTADO_COMPLETED = "C"    'COMPLETED

'///////////////////////////////////////////////////////////////////////////
'REPORTE PARAMETRO TIPO
Public Const REPORTE_PARAMETRO_TIPO_COMPANY = "CPNY"
Public Const REPORTE_PARAMETRO_TIPO_TITLE = "TITL"
Public Const REPORTE_PARAMETRO_TIPO_PERSONAL = "PRNL"
Public Const REPORTE_PARAMETRO_TIPO_CONDITION_TEXT = "COTX"
Public Const REPORTE_PARAMETRO_TIPO_DAY_OF_WEEK = "DOW"
Public Const REPORTE_PARAMETRO_TIPO_DATE_TIME = "DATI"
Public Const REPORTE_PARAMETRO_TIPO_DATE = "DATE"
Public Const REPORTE_PARAMETRO_TIPO_TIME = "TIME"
Public Const REPORTE_PARAMETRO_TIPO_YEAR_MONTH_FROM = "YMFR"
Public Const REPORTE_PARAMETRO_TIPO_YEAR_MONTH_TO = "YMTO"
Public Const REPORTE_PARAMETRO_TIPO_PERSONA = "PERS"
Public Const REPORTE_PARAMETRO_TIPO_PERSONA_CLIENTE = "PECL"
Public Const REPORTE_PARAMETRO_TIPO_PERSONA_CONDUCTOR = "PECO"
Public Const REPORTE_PARAMETRO_TIPO_PERSONA_ADMINISTRATIVO = "PEAD"
Public Const REPORTE_PARAMETRO_TIPO_PERSONA_ALARMA_GRUPO = "PAGR"
Public Const REPORTE_PARAMETRO_TIPO_PERSONA_TIPO = "PETI"
Public Const REPORTE_PARAMETRO_TIPO_RUTA = "RUTA"
Public Const REPORTE_PARAMETRO_TIPO_VEHICULO = "VEHI"
Public Const REPORTE_PARAMETRO_TIPO_VEHICULO_MANTENIMIENTO_GRUPO = "VMGR"
Public Const REPORTE_PARAMETRO_TIPO_VIAJE_ESTADO = "VIES"
Public Const REPORTE_PARAMETRO_TIPO_CUENTACORRIENTE_GRUPO = "CCGR"
Public Const REPORTE_PARAMETRO_TIPO_CUENTACORRIENTE_CAJA = "CCCJ"
Public Const REPORTE_PARAMETRO_TIPO_ALARMA = "ALAR"
Public Const REPORTE_PARAMETRO_TIPO_BOOLEAN = "BOOL"
Public Const REPORTE_PARAMETRO_TIPO_NUMBER_LONG = "LONG"
Public Const REPORTE_PARAMETRO_TIPO_CURRENCY = "CURR"
Public Const REPORTE_PARAMETRO_TIPO_VIAJE_DETALLE_OCUPANTE_TIPO = "VDOT"
Public Const REPORTE_PARAMETRO_TIPO_LISTA_PRECIO = "LIPR"
Public Const REPORTE_PARAMETRO_TIPO_LUGAR = "LUGA"
Public Const REPORTE_PARAMETRO_TIPO_LUGAR_GRUPO = "LUGR"


'///////////////////////////////////////////////////////////////////////////
'LISTA DE PRECIOS - PREPAGO - VENCIMIENTO
Public Const LISTAPRECIO_PREPAGO_VENCIMIENTO_SEMANA1 As String = "S1"
Public Const LISTAPRECIO_PREPAGO_VENCIMIENTO_SEMANA1_NOMBRE As String = "1 Semana"
Public Const LISTAPRECIO_PREPAGO_VENCIMIENTO_SEMANA2 As String = "S2"
Public Const LISTAPRECIO_PREPAGO_VENCIMIENTO_SEMANA2_NOMBRE As String = "2 Semanas"
Public Const LISTAPRECIO_PREPAGO_VENCIMIENTO_DIA15 As String = "D15"
Public Const LISTAPRECIO_PREPAGO_VENCIMIENTO_DIA15_NOMBRE As String = "15 d�as"
Public Const LISTAPRECIO_PREPAGO_VENCIMIENTO_DIA30 As String = "D30"
Public Const LISTAPRECIO_PREPAGO_VENCIMIENTO_DIA30_NOMBRE As String = "30 d�as"
Public Const LISTAPRECIO_PREPAGO_VENCIMIENTO_DIA45 As String = "D45"
Public Const LISTAPRECIO_PREPAGO_VENCIMIENTO_DIA45_NOMBRE As String = "45 d�as"
Public Const LISTAPRECIO_PREPAGO_VENCIMIENTO_MES1 As String = "M1"
Public Const LISTAPRECIO_PREPAGO_VENCIMIENTO_MES1_NOMBRE As String = "1 Mes"
Public Const LISTAPRECIO_PREPAGO_VENCIMIENTO_MES2 As String = "M2"
Public Const LISTAPRECIO_PREPAGO_VENCIMIENTO_MES2_NOMBRE As String = "2 Meses"
Public Const LISTAPRECIO_PREPAGO_VENCIMIENTO_MES3 As String = "M3"
Public Const LISTAPRECIO_PREPAGO_VENCIMIENTO_MES3_NOMBRE As String = "3 Meses"
Public Const LISTAPRECIO_PREPAGO_VENCIMIENTO_MES6 As String = "M6"
Public Const LISTAPRECIO_PREPAGO_VENCIMIENTO_MES6_NOMBRE As String = "6 Meses"
Public Const LISTAPRECIO_PREPAGO_VENCIMIENTO_ANIO1 As String = "A1"
Public Const LISTAPRECIO_PREPAGO_VENCIMIENTO_ANIO1_NOMBRE As String = "1 A�o"