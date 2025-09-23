VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPermiso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCPermiso As Collection
Private mrecUsuarioGrupoPermiso As ADODB.Recordset

Public IDUsuarioGrupo As Long

Public CuentaCorrienteCajaWhere As String
Public CuentaCorrienteGrupoWhere As String
Public RutaWhere As String
Public ListaPrecioWhere As String

Private Sub BuildCuentaCorrienteCajaWhere()
    Dim Permiso As Permiso
    Dim PermisoStringLenght As Long
    
    Dim IncludeCount As Long
    Dim IncludeString As String
    Dim ExcludeCount As Long
    Dim ExcludeString As String
    
    CuentaCorrienteCajaWhere = ""
    
    If IDUsuarioGrupo <> USUARIO_GRUPO_ID_ADMINISTRADORES Then
        PermisoStringLenght = Len(PERMISO_CUENTA_CORRIENTE_CAJA_CAJA)
        For Each Permiso In Me
            If Left(Permiso.IDPermiso, PermisoStringLenght) = PERMISO_CUENTA_CORRIENTE_CAJA_CAJA Then
                If Permiso.Permitido Then
                    IncludeCount = IncludeCount + 1
                    IncludeString = IncludeString & IIf(IncludeString = "", "", " OR ") & "CuentaCorrienteCaja.IDCuentaCorrienteCaja = " & Val(Mid(Permiso.IDPermiso, PermisoStringLenght + 1))
                Else
                    ExcludeCount = ExcludeCount + 1
                    ExcludeString = ExcludeString & IIf(ExcludeString = "", "", " AND ") & "CuentaCorrienteCaja.IDCuentaCorrienteCaja <> " & Val(Mid(Permiso.IDPermiso, PermisoStringLenght + 1))
                End If
            End If
        Next Permiso
        If IncludeCount = 0 Then
            CuentaCorrienteCajaWhere = "0 = 1"
        ElseIf ExcludeCount = 0 Then
            CuentaCorrienteCajaWhere = ""
        ElseIf IncludeCount >= ExcludeCount Then
            CuentaCorrienteCajaWhere = ExcludeString
        Else
            CuentaCorrienteCajaWhere = IncludeString
        End If
        If CuentaCorrienteCajaWhere <> "" Then
            CuentaCorrienteCajaWhere = "(" & CuentaCorrienteCajaWhere & ")"
        End If
    End If
End Sub

Private Sub BuildCuentaCorrienteGrupoWhere()
    Dim Permiso As Permiso
    Dim PermisoStringLenght As Long
    
    Dim IncludeCount As Long
    Dim IncludeString As String
    Dim ExcludeCount As Long
    Dim ExcludeString As String
    
    CuentaCorrienteGrupoWhere = ""
    If IDUsuarioGrupo <> USUARIO_GRUPO_ID_ADMINISTRADORES Then
        PermisoStringLenght = Len(PERMISO_CUENTA_CORRIENTE_GRUPO_GRUPO)
        For Each Permiso In Me
            If Left(Permiso.IDPermiso, PermisoStringLenght) = PERMISO_CUENTA_CORRIENTE_GRUPO_GRUPO Then
                If Permiso.Permitido Then
                    IncludeCount = IncludeCount + 1
                    IncludeString = IncludeString & IIf(IncludeString = "", "", " OR ") & "CuentaCorrienteGrupo.IDCuentaCorrienteGrupo = " & Val(Mid(Permiso.IDPermiso, PermisoStringLenght + 1))
                Else
                    ExcludeCount = ExcludeCount + 1
                    ExcludeString = ExcludeString & IIf(ExcludeString = "", "", " AND ") & "CuentaCorrienteGrupo.IDCuentaCorrienteGrupo <> " & Val(Mid(Permiso.IDPermiso, PermisoStringLenght + 1))
                End If
            End If
        Next Permiso
        If IncludeCount = 0 Then
            CuentaCorrienteGrupoWhere = "0 = 1"
        ElseIf ExcludeCount = 0 Then
            CuentaCorrienteGrupoWhere = ""
        ElseIf IncludeCount >= ExcludeCount Then
            CuentaCorrienteGrupoWhere = ExcludeString
        Else
            CuentaCorrienteGrupoWhere = IncludeString
        End If
        If CuentaCorrienteGrupoWhere <> "" Then
            CuentaCorrienteGrupoWhere = "(" & CuentaCorrienteGrupoWhere & ")"
        End If
    End If
End Sub

Private Sub BuildRutaWhere()
    Dim Permiso As Permiso
    Dim PermisoStringLenght As Long
    
    Dim IncludeCount As Long
    Dim IncludeString As String
    Dim ExcludeCount As Long
    Dim ExcludeString As String
    
    RutaWhere = ""
    If IDUsuarioGrupo <> USUARIO_GRUPO_ID_ADMINISTRADORES Then
        PermisoStringLenght = Len(PERMISO_RUTA_RUTA)
        For Each Permiso In Me
            If Left(Permiso.IDPermiso, PermisoStringLenght) = PERMISO_RUTA_RUTA Then
                If Permiso.Permitido Then
                    IncludeCount = IncludeCount + 1
                    IncludeString = IncludeString & IIf(IncludeString = "", "", " OR ") & "%TABLENAME%.IDRuta = '" & Mid(Permiso.IDPermiso, PermisoStringLenght + 1) & "'"
                Else
                    ExcludeCount = ExcludeCount + 1
                    ExcludeString = ExcludeString & IIf(ExcludeString = "", "", " AND ") & "%TABLENAME%.IDRuta <> '" & Mid(Permiso.IDPermiso, PermisoStringLenght + 1) & "'"
                End If
            End If
        Next Permiso
        If IncludeCount = 0 Then
            RutaWhere = "0 = 1"
        ElseIf ExcludeCount = 0 Then
            RutaWhere = ""
        ElseIf IncludeCount >= ExcludeCount Then
            RutaWhere = ExcludeString
        Else
            RutaWhere = IncludeString
        End If
        If RutaWhere <> "" Then
            RutaWhere = "(" & RutaWhere & ")"
        End If
    End If
End Sub

Private Sub BuildListaPrecioWhere()
    Dim Permiso As Permiso
    Dim PermisoStringLenght As Long
    
    Dim IncludeCount As Long
    Dim IncludeString As String
    Dim ExcludeCount As Long
    Dim ExcludeString As String
    
    ListaPrecioWhere = ""
    If IDUsuarioGrupo <> USUARIO_GRUPO_ID_ADMINISTRADORES Then
        PermisoStringLenght = Len(PERMISO_LISTA_PRECIO_LISTA_PRECIO)
        For Each Permiso In Me
            If Left(Permiso.IDPermiso, PermisoStringLenght) = PERMISO_LISTA_PRECIO_LISTA_PRECIO Then
                If Permiso.Permitido Then
                    IncludeCount = IncludeCount + 1
                    IncludeString = IncludeString & IIf(IncludeString = "", "", " OR ") & "%TABLENAME%.IDListaPrecio = '" & Mid(Permiso.IDPermiso, PermisoStringLenght + 1) & "'"
                Else
                    ExcludeCount = ExcludeCount + 1
                    ExcludeString = ExcludeString & IIf(ExcludeString = "", "", " AND ") & "%TABLENAME%.IDListaPrecio <> '" & Mid(Permiso.IDPermiso, PermisoStringLenght + 1) & "'"
                End If
            End If
        Next Permiso
        If IncludeCount = 0 Then
            ListaPrecioWhere = "0 = 1"
        ElseIf ExcludeCount = 0 Then
            ListaPrecioWhere = ""
        ElseIf IncludeCount >= ExcludeCount Then
            ListaPrecioWhere = ExcludeString
        Else
            ListaPrecioWhere = IncludeString
        End If
        If ListaPrecioWhere <> "" Then
            ListaPrecioWhere = "(" & ListaPrecioWhere & ")"
        End If
    End If
End Sub

Public Property Get Item(IndexKey As Variant) As Permiso
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "200"
    Set Item = mCPermiso(IndexKey)
End Property

Public Property Get Count() As Long
    Count = mCPermiso.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCPermiso.[_NewEnum]
End Property

Public Sub Load()
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    
    Screen.MousePointer = vbHourglass
    
    Add PERMISO_OPCIONES_SYSTEM, "Opciones del Sistema", "Ver y Modificar"
    Add PERMISO_OPCIONES_WORKSTATION, "Opciones de la Estación de Trabajo", "Ver y Modificar"
    
    Add PERMISO_USUARIO, "Usuarios", "Ver"
    Add PERMISO_USUARIO_ADD, "Usuarios", "Agregar"
    Add PERMISO_USUARIO_MODIFY, "Usuarios", "Modificar"
    Add PERMISO_USUARIO_DELETE, "Usuarios", "Eliminar"
    
    Add PERMISO_USUARIO_GRUPO, "Grupo de Usuarios", "Ver"
    Add PERMISO_USUARIO_GRUPO_ADD, "Grupo de Usuarios", "Agregar"
    Add PERMISO_USUARIO_GRUPO_MODIFY, "Grupo de Usuarios", "Modificar"
    Add PERMISO_USUARIO_GRUPO_DELETE, "Grupo de Usuarios", "Eliminar"
    Add PERMISO_USUARIO_GRUPO_PERMISSION, "Grupo de Usuarios", "Ver Permisos"
    Add PERMISO_USUARIO_GRUPO_PERMISSION_MODIFY, "Grupo de Usuarios", "Modificar Permisos"
    
    Add PERMISO_LUGAR, "Lugares", "Ver"
    Add PERMISO_LUGAR_ADD, "Lugares", "Agregar"
    Add PERMISO_LUGAR_MODIFY, "Lugares", "Modificar"
    Add PERMISO_LUGAR_DELETE, "Lugares", "Eliminar"
    
    Add PERMISO_LUGAR_GRUPO, "Grupos de Lugares", "Ver"
    Add PERMISO_LUGAR_GRUPO_ADD, "Grupos de Lugares", "Agregar"
    Add PERMISO_LUGAR_GRUPO_MODIFY, "Grupos de Lugares", "Modificar"
    Add PERMISO_LUGAR_GRUPO_DELETE, "Grupos de Lugares", "Eliminar"
    
    Add PERMISO_RUTA, "Rutas", "Ver"
    Add PERMISO_RUTA_ADD, "Rutas", "Agregar"
    Add PERMISO_RUTA_MODIFY, "Rutas", "Modificar"
    Add PERMISO_RUTA_DELETE, "Rutas", "Eliminar"
    
    'RUTAS
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "usp_Ruta_AllData"
    cmdData.CommandType = adCmdStoredProc
    Set recData = New ADODB.Recordset
    recData.Open cmdData, , adOpenForwardOnly, adLockReadOnly
    Set cmdData = Nothing
    Do While Not recData.EOF
        Add PERMISO_RUTA_RUTA & RTrim(recData("IDRuta").value), "Rutas", "Ver: " & recData("IDRuta").value
        recData.MoveNext
    Loop
    recData.Close
    Set recData = Nothing
    
    BuildRutaWhere
    
    Add PERMISO_RUTA_DETALLE, "Detalle de Rutas", "Ver"
    Add PERMISO_RUTA_DETALLE_ADD, "Detalle de Rutas", "Agregar"
    Add PERMISO_RUTA_DETALLE_MODIFY, "Detalle de Rutas", "Modificar"
    Add PERMISO_RUTA_DETALLE_DELETE, "Detalle de Rutas", "Eliminar"
    
    Add PERMISO_RUTA_DETALLE_HORARIO, "Horarios de Detalle de Ruta", "Ver"
    Add PERMISO_RUTA_DETALLE_HORARIO_ADD, "Horarios de Detalle de Ruta", "Agregar"
    Add PERMISO_RUTA_DETALLE_HORARIO_MODIFY, "Horarios de Detalle de Ruta", "Modificar"
    Add PERMISO_RUTA_DETALLE_HORARIO_DELETE, "Horarios de Detalle de Ruta", "Eliminar"
    
    Add PERMISO_RUTALUGARGRUPO, "Rutas-Grupos de Lugares", "Ver"
    Add PERMISO_RUTALUGARGRUPO_ADD, "Rutas-Grupos de Lugares", "Agregar"
    Add PERMISO_RUTALUGARGRUPO_MODIFY, "Rutas-Grupos de Lugares", "Modificar"
    Add PERMISO_RUTALUGARGRUPO_DELETE, "Rutas-Grupos de Lugares", "Eliminar"
    
    Add PERMISO_LISTA_PRECIO, "Listas de Precios", "Ver"
    Add PERMISO_LISTA_PRECIO_ADD, "Listas de Precios", "Agregar"
    Add PERMISO_LISTA_PRECIO_MODIFY, "Listas de Precios", "Modificar"
    Add PERMISO_LISTA_PRECIO_DELETE, "Listas de Precios", "Eliminar"
    
    'LISTA DE PRECIO
    Set recData = New ADODB.Recordset
    recData.Source = "SELECT IDListaPrecio, Nombre FROM ListaPrecio ORDER BY IDListaPrecio"
    recData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not recData.EOF
        Add PERMISO_LISTA_PRECIO_LISTA_PRECIO & recData("IDListaPrecio").value, "Listas de Precios", "Ver: " & recData("Nombre").value
        recData.MoveNext
    Loop
    recData.Close
    Set recData = Nothing
    
    BuildListaPrecioWhere
    
    Add PERMISO_LISTA_PRECIO_DETALLE, "Detalle de Lista de Precios", "Ver"
    Add PERMISO_LISTA_PRECIO_DETALLE_ADD, "Detalle de Lista de Precios", "Agregar"
    Add PERMISO_LISTA_PRECIO_DETALLE_MODIFY, "Detalle de Lista de Precios", "Modificar"
    Add PERMISO_LISTA_PRECIO_DETALLE_DELETE, "Detalle de Lista de Precios", "Eliminar"
    
    Add PERMISO_VEHICULO, "Vehículos", "Ver"
    Add PERMISO_VEHICULO_ADD, "Vehículos", "Agregar"
    Add PERMISO_VEHICULO_MODIFY, "Vehículos", "Modificar"
    Add PERMISO_VEHICULO_DELETE, "Vehículos", "Eliminar"
    Add PERMISO_VEHICULO_UTILIZACION, "Vehículos", "Utilización"
    
    Add PERMISO_HORARIO, "Horarios", "Ver"
    Add PERMISO_HORARIO_ADD, "Horarios", "Agregar"
    Add PERMISO_HORARIO_MODIFY, "Horarios", "Modificar"
    Add PERMISO_HORARIO_DELETE, "Horarios", "Eliminar"
    
    Add PERMISO_PERSONA, "Personas", "Ver"
    Add PERMISO_PERSONA_ADD, "Personas", "Agregar"
    Add PERMISO_PERSONA_ADD_ALLTYPE, "Personas", "Agregar cualquier tipo"
    Add PERMISO_PERSONA_MODIFY, "Personas", "Modificar"
    Add PERMISO_PERSONA_DELETE, "Personas", "Eliminar"
    Add PERMISO_PERSONA_HABILITACION_VIAJAR_ESTABLECER, "Personas", "Habilitación a Viajar: Establecer"
    Add PERMISO_PERSONA_HABILITACION_VIAJAR_IGNORAR, "Personas", "Habilitación a Viajar: Ignorar"
    
    Add PERMISO_PERSONA_HORARIO, "Personas", "Ver Horarios"
    Add PERMISO_PERSONA_HORARIO_ADD, "Personas", "Agregar Horario"
    Add PERMISO_PERSONA_HORARIO_MODIFY, "Personas", "Modificar Horario"
    Add PERMISO_PERSONA_HORARIO_DELETE, "Personas", "Eliminar Horario"
    
    Add PERMISO_PERSONA_RUTA, "Personas", "Ver Rutas"
    Add PERMISO_PERSONA_RUTA_ADD, "Personas", "Agregar Ruta"
    Add PERMISO_PERSONA_RUTA_MODIFY, "Personas", "Modificar Ruta"
    Add PERMISO_PERSONA_RUTA_DELETE, "Personas", "Eliminar Ruta"
    
    Add PERMISO_PERSONA_PREPAGO, "Personas", "Ver Prepagos"
    Add PERMISO_PERSONA_PREPAGO_ADD, "Personas", "Agregar Prepago"
    Add PERMISO_PERSONA_PREPAGO_MODIFY, "Personas", "Modificar Prepago"
    Add PERMISO_PERSONA_PREPAGO_DELETE, "Personas", "Eliminar Prepago"
    
    Add PERMISO_PERSONA_INFO, "Personas", "Ver Información"
    
    Add PERMISO_PERSONA_RESPUESTA, "Personas", "Ver Respuestas"
    Add PERMISO_PERSONA_RESPUESTA_ADD, "Personas", "Agregar Respuesta"
    Add PERMISO_PERSONA_RESPUESTA_MODIFY, "Personas", "Modificar Respuesta"
    Add PERMISO_PERSONA_RESPUESTA_DELETE, "Personas", "Eliminar Respuesta"
    Add PERMISO_PERSONA_RESPUESTA_ACTIVATE_ALL, "Personas", "Activar/Desactivar todas las Respuestas"
    
    Add PERMISO_CONDUCTOR_RUTA, "Conductores", "Ver Rutas"
    Add PERMISO_CONDUCTOR_RUTA_ADD, "Conductores", "Agregar Ruta"
    Add PERMISO_CONDUCTOR_RUTA_MODIFY, "Conductores", "Modificar Ruta"
    Add PERMISO_CONDUCTOR_RUTA_DELETE, "Conductores", "Eliminar Ruta"
    
    Add PERMISO_VIAJE, "Viajes", "Ver"
    Add PERMISO_VIAJE_HISTORICO, "Viajes", "Ver Históricos"
    Add PERMISO_VIAJE_GENERATE, "Viajes", "Generar"
    Add PERMISO_VIAJE_ADD, "Viajes", "Agregar"
    Add PERMISO_VIAJE_MODIFY, "Viajes", "Modificar"
    Add PERMISO_VIAJE_MODIFY_HORA, "Viajes", "Modificar Hora"
    Add PERMISO_VIAJE_MODIFY_FECHA, "Viajes", "Modificar Fecha"
    Add PERMISO_VIAJE_MODIFY_RUTA, "Viajes", "Modificar Ruta"
    Add PERMISO_VIAJE_DELETE, "Viajes", "Eliminar"
    Add PERMISO_VIAJE_CHANGE_STATUS, "Viajes", "Cambiar Estado"
    Add PERMISO_VIAJE_CHANGE_STATUS_CANCEL, "Viajes", "Cambiar Estado: Cancelar"
    Add PERMISO_VIAJE_CHANGE_STATUS_SPECIAL, "Viajes", "Cambiar Estado: Cualquier Estado"
    Add PERMISO_VIAJE_SEND_EMAIL, "Viajes", "Enviar Planilla por E-mail"
    Add PERMISO_VIAJE_CONDUCTOR, "Viajes", "Por Conductor"
    Add PERMISO_VIAJE_ASISTENCIA, "Viajes", "Dar Asistencia a varios Viajes"
    
    Add PERMISO_VIAJE_DETALLE, "Detalle de Viajes", "Ver"
    Add PERMISO_VIAJE_DETALLE_ADD, "Detalle de Viajes", "Agregar"
    Add PERMISO_VIAJE_DETALLE_ADD_FECHAHORARUTA, "Detalle de Viajes", "Cambiar Fecha/Hora y Ruta al Agregar"
    Add PERMISO_VIAJE_DETALLE_MODIFY, "Detalle de Viajes", "Modificar"
    Add PERMISO_VIAJE_DETALLE_MODIFY_FECHAHORARUTA, "Detalle de Viajes", "Cambiar Fecha/Hora y Ruta al Modificar"
    Add PERMISO_VIAJE_DETALLE_MODIFY_IMPORTE_OTRO_USUARIO, "Detalle de Viajes", "Modificar Importe desde otro Usuario"
    Add PERMISO_VIAJE_DETALLE_COMISION_MODIFY_IMPORTE_PAGOCONTADO_FINALIZADO, "Detalle de Viajes", "Modificar Importe y/o Pago Contado de Comisión en Viaje finalizado."
    Add PERMISO_VIAJE_DETALLE_COMISION_RENDIR, "Detalle de Viajes", "Rendir Comisión"
    Add PERMISO_VIAJE_DETALLE_PASAJERO_IMPORTE_ALLOWMODIFY, "Detalle de Viajes", "Permite modificar el Importe de la Reserva del Pasajero."
    Add PERMISO_VIAJE_DETALLE_COMISION_IMPORTE_ALLOWMODIFY, "Detalle de Viajes", "Permite modificar el Importe de la Comisión."
    Add PERMISO_VIAJE_DETALLE_PAQUETE_PASAJERO_IMPORTE_ALLOWMODIFY, "Detalle de Viajes", "Permite modificar el Importe de la Reserva del Pasajero en un Paquete."
    Add PERMISO_VIAJE_DETALLE_MODIFY_CAJA_OTRO_USUARIO, "Detalle de Viajes", "Modificar Caja desde otro Usuario"
    Add PERMISO_VIAJE_DETALLE_INASISTENCIA_NODEBITAR, "Detalle de Viajes", "No Debitar Viaje No Realizado"
    Add PERMISO_VIAJE_DETALLE_DELETE, "Detalle de Viajes", "Eliminar"
    Add PERMISO_VIAJE_DETALLE_PRINT, "Detalle de Viajes", "Imprimir"
    Add PERMISO_VIAJE_DETALLE_CHANGE_STATUS, "Detalle de Viajes", "Cambiar Estado (Reserva estándar)"
    Add PERMISO_VIAJE_DETALLE_WEB_CHANGE_STATUS, "Detalle de Viajes", "Cambiar Estado (Reserva Web)"
    Add PERMISO_VIAJE_DETALLE_CHANGE_STATUS_AFTER_LIMIT, "Detalle de Viajes", "Cambiar Estado Después del Límite"
    Add PERMISO_VIAJE_DETALLE_TRANSFER, "Detalle de Viajes", "Transferencia de Pasajeros"
    Add PERMISO_VIAJE_DETALLE_TRANSFER_ALLOWOUTOFRANGE, "Detalle de Viajes", "Transferencia de Pasajeros entre Viajes fuera del rango horario establecido."
    Add PERMISO_VIAJE_DETALLE_TRANSFER_ALLOWDIFFERENTROUTE, "Detalle de Viajes", "Transferencia de Pasajeros entre Viajes de distinta Ruta."
    Add PERMISO_VIAJE_DETALLE_FINALIZADO_MODIFY, "Detalle de Viajes", "Modificar con Viaje Finalizado"
    Add PERMISO_VIAJE_DETALLE_FINALIZADO_MODIFY_MEDIOPAGO_DEBITO, "Detalle de Viajes", "Modificar medio de pago y debitar a"
    Add PERMISO_VIAJE_DETALLE_SUBEBAJA_MOSTRAR, "Detalle de Viajes", "Mostrar campos sube/baja"
    
    Add PERMISO_DOCUMENTOFISCAL, "Documentos Fiscales", "Ver"
    Add PERMISO_DOCUMENTOFISCAL_ADD, "Documentos Fiscales", "Agregar"
    Add PERMISO_DOCUMENTOFISCAL_MODIFY, "Documentos Fiscales", "Modificar"
    Add PERMISO_DOCUMENTOFISCAL_DELETE, "Documentos Fiscales", "Eliminar"
    Add PERMISO_DOCUMENTOFISCAL_PRINT, "Documentos Fiscales", "Imprimir"
    
    Add PERMISO_CUENTA_CORRIENTE, "Cuenta Corriente", "Ver"
    Add PERMISO_CUENTA_CORRIENTE_HISTORICO, "Cuenta Corriente", "Ver Históricos"
    Add PERMISO_CUENTA_CORRIENTE_CONDUCTOR_SELECT, "Cuenta Corriente", "Seleccionar Personas de tipo Conductores"
    Add PERMISO_CUENTA_CORRIENTE_ADMINISTRATIVO_SELECT, "Cuenta Corriente", "Seleccionar Personas de tipo Administrativos"
    Add PERMISO_CUENTA_CORRIENTE_ADD_ANTERIOR, "Cuenta Corriente", "Agregar Movimientos con Fecha anterior"
    Add PERMISO_CUENTA_CORRIENTE_ADD_ACTUAL, "Cuenta Corriente", "Agregar Movimientos con Fecha actual"
    Add PERMISO_CUENTA_CORRIENTE_MODIFY_ANTERIOR, "Cuenta Corriente", "Modificar Movimientos Anteriores"
    Add PERMISO_CUENTA_CORRIENTE_MODIFY_ACTUAL, "Cuenta Corriente", "Modificar Movimientos Actuales"
    Add PERMISO_CUENTA_CORRIENTE_MODIFY_TRANSFER_ANTERIOR, "Cuenta Corriente", "Modificar Transferencias Anteriores"
    Add PERMISO_CUENTA_CORRIENTE_MODIFY_TRANSFER_ACTUAL, "Cuenta Corriente", "Modificar Transferencias Actuales"
    
    Dim CuentaCorrienteCaja As CuentaCorrienteCaja
    
    Set CuentaCorrienteCaja = New CuentaCorrienteCaja
    With CuentaCorrienteCaja
        .IDCuentaCorrienteCaja = pParametro.CuentaCorrienteCaja_ID_ViajeDebito
        .NoMatchRaiseError = False
        .Load
        If .NoMatch Then
            Add PERMISO_CUENTA_CORRIENTE_MODIFY_CAJA_OFICINA, "Cuenta Corriente", "Modificar en la Caja de Débitos de Viaje"
        Else
            Add PERMISO_CUENTA_CORRIENTE_MODIFY_CAJA_OFICINA, "Cuenta Corriente", "Modificar en la Caja '" & .Nombre & "'."
        End If
    End With
    Set CuentaCorrienteCaja = Nothing
    
    Add PERMISO_CUENTA_CORRIENTE_MODIFY_NOTES_ANTERIOR, "Cuenta Corriente", "Modificar Notas de Movimientos Anteriores"
    Add PERMISO_CUENTA_CORRIENTE_MODIFY_NOTES_ACTUAL, "Cuenta Corriente", "Modificar Notas de Movimientos Actuales"
    Add PERMISO_CUENTA_CORRIENTE_DELETE, "Cuenta Corriente", "Eliminar"
    Add PERMISO_CUENTA_CORRIENTE_DELETE_DEBITO_CREDITO, "Cuenta Corriente", "Eliminar Débitos o Créditos de Viajes"
    
    Add PERMISO_CUENTA_CORRIENTE_CAJA, "Cajas de Cuenta Corriente", "Ver"
    Add PERMISO_CUENTA_CORRIENTE_CAJA_ADD, "Cajas de Cuenta Corriente", "Agregar"
    Add PERMISO_CUENTA_CORRIENTE_CAJA_MODIFY, "Cajas de Cuenta Corriente", "Modificar"
    Add PERMISO_CUENTA_CORRIENTE_CAJA_DELETE, "Cajas de Cuenta Corriente", "Eliminar"
    Add PERMISO_CUENTA_CORRIENTE_CAJA_TRANSFER, "Cajas de Cuenta Corriente", "Transferencias"
    Add PERMISO_CUENTA_CORRIENTE_CAJA_TRANSFER_GRUPO_CAMBIAR, "Cajas de Cuenta Corriente", "Transferencias: Modificar Grupo"
    Add PERMISO_CUENTA_CORRIENTE_CAJA_SALDOOCULTO_VIEW, "Cajas de Cuenta Corriente", "Ver Saldo Oculto"
    Add PERMISO_CUENTA_CORRIENTE_CAJA_VIEW_ALL, "Cajas de Cuenta Corriente", "Ver Todas"
    
    'CAJAS DE CUENTA CORRIENTE
    Set recData = New ADODB.Recordset
    recData.Source = "SELECT IDCuentaCorrienteCaja, Nombre FROM CuentaCorrienteCaja ORDER BY Nombre"
    recData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not recData.EOF
        Add PERMISO_CUENTA_CORRIENTE_CAJA_CAJA & recData("IDCuentaCorrienteCaja").value, "Cajas de Cuenta Corriente", "Ver Movimientos de: " & recData("Nombre").value
        recData.MoveNext
    Loop
    recData.Close
    Set recData = Nothing
        
    Add PERMISO_CUENTA_CORRIENTE_GRUPO, "Grupos de Cuenta Corriente", "Ver"
    Add PERMISO_CUENTA_CORRIENTE_GRUPO_ADD, "Grupos de Cuenta Corriente", "Agregar"
    Add PERMISO_CUENTA_CORRIENTE_GRUPO_MODIFY, "Grupos de Cuenta Corriente", "Modificar"
    Add PERMISO_CUENTA_CORRIENTE_GRUPO_DELETE, "Grupos de Cuenta Corriente", "Eliminar"
    Add PERMISO_CUENTA_CORRIENTE_GRUPO_HIDDEN_SHOW, "Grupos de Cuenta Corriente", "Ver Ocultos"
    
    'GRUPOS DE CUENTA CORRIENTE
    Set recData = New ADODB.Recordset
    recData.Source = "SELECT IDCuentaCorrienteGrupo, Nombre FROM CuentaCorrienteGrupo ORDER BY Nombre"
    recData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not recData.EOF
        Add PERMISO_CUENTA_CORRIENTE_GRUPO_GRUPO & recData("IDCuentaCorrienteGrupo").value, "Grupos de Cuenta Corriente", "Mostrar el Grupo: " & recData("Nombre").value
        recData.MoveNext
    Loop
    recData.Close
    Set recData = Nothing
    
    Add PERMISO_MEDIOPAGO, "Medios de Pago", "Ver"
    Add PERMISO_MEDIOPAGO_ADD, "Medios de Pago", "Agregar"
    Add PERMISO_MEDIOPAGO_MODIFY, "Medios de Pago", "Modificar"
    Add PERMISO_MEDIOPAGO_DELETE, "Medios de Pago", "Eliminar"
    
    Add PERMISO_FERIADO, "Feriados", "Ver"
    Add PERMISO_FERIADO_ADD, "Feriados", "Agregar"
    Add PERMISO_FERIADO_MODIFY, "Feriados", "Modificar"
    Add PERMISO_FERIADO_DELETE, "Feriados", "Eliminar"
    
    Add PERMISO_VEHICULO_MANTENIMIENTO_GRUPO, "Grupos de Mantenimiento de Vehículos", "Ver"
    Add PERMISO_VEHICULO_MANTENIMIENTO_GRUPO_ADD, "Grupos de Mantenimiento de Vehículos", "Agregar"
    Add PERMISO_VEHICULO_MANTENIMIENTO_GRUPO_MODIFY, "Grupos de Mantenimiento de Vehículos", "Modificar"
    Add PERMISO_VEHICULO_MANTENIMIENTO_GRUPO_DELETE, "Grupos de Mantenimiento de Vehículos", "Eliminar"
    
    Add PERMISO_VEHICULO_MANTENIMIENTO, "Mantenimiento de Vehículos", "Ver"
    Add PERMISO_VEHICULO_MANTENIMIENTO_ADD, "Mantenimiento de Vehículos", "Agregar"
    Add PERMISO_VEHICULO_MANTENIMIENTO_MODIFY, "Mantenimiento de Vehículos", "Modificar"
    Add PERMISO_VEHICULO_MANTENIMIENTO_DELETE, "Mantenimiento de Vehículos", "Eliminar"
    
    Add PERMISO_VEHICULO_MANTENIMIENTO_ACCION, "Acciones de Mantenimiento de Vehículos", "Ver"
    Add PERMISO_VEHICULO_MANTENIMIENTO_ACCION_HISTORICO, "Acciones de Mantenimiento de Vehículos", "Ver Históricos"
    Add PERMISO_VEHICULO_MANTENIMIENTO_ACCION_ADD, "Acciones de Mantenimiento de Vehículos", "Agregar"
    Add PERMISO_VEHICULO_MANTENIMIENTO_ACCION_MODIFY, "Acciones de Mantenimiento de Vehículos", "Modificar"
    Add PERMISO_VEHICULO_MANTENIMIENTO_ACCION_DELETE, "Acciones de Mantenimiento de Vehículos", "Eliminar"
    
    Add PERMISO_DOCUMENTO_TIPO, "Tipos de Documento", "Ver"
    Add PERMISO_DOCUMENTO_TIPO_ADD, "Tipos de Documento", "Agregar"
    Add PERMISO_DOCUMENTO_TIPO_MODIFY, "Tipos de Documento", "Modificar"
    Add PERMISO_DOCUMENTO_TIPO_DELETE, "Tipos de Documento", "Eliminar"
    
    Add PERMISO_TELEFONO_TIPO, "Tipos de Teléfono", "Ver"
    Add PERMISO_TELEFONO_TIPO_ADD, "Tipos de Teléfono", "Agregar"
    Add PERMISO_TELEFONO_TIPO_MODIFY, "Tipos de Teléfono", "Modificar"
    Add PERMISO_TELEFONO_TIPO_DELETE, "Tipos de Teléfono", "Eliminar"
    
    Add PERMISO_PERSONA_ALARMA_GRUPO, "Grupos de Alarmas de Personas", "Ver"
    Add PERMISO_PERSONA_ALARMA_GRUPO_ADD, "Grupos de Alarmas de Personas", "Agregar"
    Add PERMISO_PERSONA_ALARMA_GRUPO_MODIFY, "Grupos de Alarmas de Personas", "Modificar"
    Add PERMISO_PERSONA_ALARMA_GRUPO_DELETE, "Grupos de Alarmas de Personas", "Eliminar"
    
    Add PERMISO_PERSONA_ALARMA, "Alarmas de Personas", "Ver"
    Add PERMISO_PERSONA_ALARMA_ADD, "Alarmas de Personas", "Agregar"
    Add PERMISO_PERSONA_ALARMA_MODIFY, "Alarmas de Personas", "Modificar"
    Add PERMISO_PERSONA_ALARMA_DELETE, "Alarmas de Personas", "Eliminar"
    
    Add PERMISO_ALARMA, "Alarmas", "Ver"
    Add PERMISO_ALARMA_ADD, "Alarmas", "Agregar"
    Add PERMISO_ALARMA_MODIFY, "Alarmas", "Modificar"
    Add PERMISO_ALARMA_DELETE, "Alarmas", "Eliminar"
    
    Add PERMISO_CONTACTO_GRUPO, "Grupos de Contactos", "Ver"
    Add PERMISO_CONTACTO_GRUPO_ADD, "Grupos de Contactos", "Agregar"
    Add PERMISO_CONTACTO_GRUPO_MODIFY, "Grupos de Contactos", "Modificar"
    Add PERMISO_CONTACTO_GRUPO_DELETE, "Grupos de Contactos", "Eliminar"
    
    Add PERMISO_CONTACTO, "Contactos", "Ver"
    Add PERMISO_CONTACTO_ADD, "Contactos", "Agregar"
    Add PERMISO_CONTACTO_MODIFY, "Contactos", "Modificar"
    Add PERMISO_CONTACTO_DELETE, "Contactos", "Eliminar"
    
    Add PERMISO_LISTAPASAJERO, "Lista de Pasajeros", "Generar"
    
    Add PERMISO_REGISTROLLAMADA, "Registro de Llamadas", "Ver"
    
    Add PERMISO_FRANCO, "Francos", "Ver"
    Add PERMISO_FRANCO_ADD, "Francos", "Agregar"
    Add PERMISO_FRANCO_MODIFY, "Francos", "Modificar"
    Add PERMISO_FRANCO_DELETE, "Francos", "Eliminar"
    
    Add PERMISO_APPEXT_MSWORD, "Aplicaciones", "Microsoft Word"
    Add PERMISO_APPEXT_MSEXCEL, "Aplicaciones", "Microsoft Excel"
    
    Add PERMISO_UTILIDAD, "Utilidades", "Mostrar"
    Add PERMISO_UTILIDAD_EDITOR_DIRECTO_PARAMETRO, "Utilidades", "Editor Directo de Parámetros"
    Add PERMISO_UTILIDAD_ACTUALIZAR_PRECIO_RESERVA, "Utilidades", "Actualizar Precio de Reservas"
    Add PERMISO_UTILIDAD_ACTUALIZAR_SUELDO_VIAJE, "Utilidades", "Actualizar Sueldo de Viajes"
    
    Add PERMISO_MENSAJE, "Mensajes", "Ver"
    Add PERMISO_MENSAJE_ADD, "Mensajes", "Agregar"
    Add PERMISO_MENSAJE_MODIFY, "Mensajes", "Modificar"
    Add PERMISO_MENSAJE_DELETE, "Mensajes", "Eliminar"
    
    Add PERMISO_MESSENGER, "Messenger", "Habilitar"
    
    Add PERMISO_REPORTE, "Reportes", "Ver"
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    'REPORTES
    Set recData = New ADODB.Recordset
    recData.Source = "SELECT IDReporte, Tipo, Nombre FROM Reporte ORDER BY Tipo, Nombre"
    recData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not recData.EOF
        Add PERMISO_REPORTE_REPORTE & RTrim(recData("IDReporte").value), "Reportes", recData("Tipo").value & ": " & recData("Nombre").value
        recData.MoveNext
    Loop
    recData.Close
    Set recData = Nothing
    
    BuildCuentaCorrienteCajaWhere
    BuildCuentaCorrienteGrupoWhere
    BuildListaPrecioWhere
    BuildRutaWhere
    
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorHandler:
    ShowErrorMessage "Classes.CPermiso.Load", "Error al obtener la Lista de Reportes."
End Sub

Public Sub Add(ByVal IDPermiso As String, ByVal Categoria As String, ByVal Descripcion As String)
    Dim Permiso As Permiso
    
    Set Permiso = New Permiso
    With Permiso
        .IDPermiso = IDPermiso
        .Categoria = Categoria
        .Descripcion = Descripcion
        .Permitido = ReadPermission(IDPermiso)
    End With
    mCPermiso.Add Permiso, Permiso.IDPermiso
    
    Set Permiso = Nothing
End Sub

Private Function ReadPermission(ByVal IDPermiso As String) As Boolean
    If IDUsuarioGrupo = USUARIO_GRUPO_ID_ADMINISTRADORES Then
        ReadPermission = True
        Exit Function
    End If
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    mrecUsuarioGrupoPermiso.Filter = "IDPermiso = '" & IDPermiso & "'"
    If mrecUsuarioGrupoPermiso.EOF Then
        ReadPermission = False
    Else
        ReadPermission = True
    End If
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.CPermiso.ReadPermission", "Error al obtener el Permiso del Grupo de Usuarios." & vbCr & vbCr & "IDUsuarioGrupo: " & IDUsuarioGrupo & vbCr & "IDPermiso: " & IDPermiso
End Function

Private Function Exist(ByVal IDPermiso As String) As Boolean
    Dim DummyPermiso As Permiso
    
    On Error GoTo ErrorHandler
    
    Set DummyPermiso = mCPermiso(IDPermiso)
    Set DummyPermiso = Nothing
    Exist = True
    Exit Function

ErrorHandler:
    If Err.Number = 5 Then
        MsgBox "El Reporte que intenta ver es un Reporte Nuevo." & vbCr & "Para poder verlo, deberá salir del sistema y volver a ingresar.", vbExclamation, App.Title
        Exist = False
    End If
End Function

Public Function GotPermission(ByVal IDPermiso As String, Optional ByVal ShowMessage As Boolean = True) As Boolean
    If Left(IDPermiso, Len(PERMISO_REPORTE_REPORTE)) = PERMISO_REPORTE_REPORTE Then
        If Not Exist(IDPermiso) Then
            GotPermission = False
            Exit Function
        End If
    End If
    If Me(IDPermiso).Permitido Then
        GotPermission = True
    Else
        GotPermission = False
        If ShowMessage Then
            MsgBox "No está autorizado a realizar esta acción.", vbExclamation, App.Title
        End If
    End If
End Function

Public Sub OpenRecordset()
    Dim cmdData As ADODB.command

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
        
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "usp_UsuarioGrupoPermiso_AllData"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDUsuarioGrupo", adInteger, adParamInput, , IDUsuarioGrupo)
    Set mrecUsuarioGrupoPermiso = New ADODB.Recordset
    mrecUsuarioGrupoPermiso.Open cmdData, , adOpenForwardOnly, adLockReadOnly
    Set cmdData = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Classes.CPermiso.Initialize", "Error al Abrir la Tabla de Permisos."
End Sub

Private Sub Class_Initialize()
    Set mCPermiso = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCPermiso = Nothing
    
    If Not mrecUsuarioGrupoPermiso Is Nothing Then
        If mrecUsuarioGrupoPermiso.State = adStateOpen Then
            mrecUsuarioGrupoPermiso.Close
        End If
        Set mrecUsuarioGrupoPermiso = Nothing
    End If
End Sub

