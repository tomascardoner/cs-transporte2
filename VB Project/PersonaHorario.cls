VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PersonaHorario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public IDPersona As Long
Public DiaSemana As Byte
Public Hora As Date
Public IDRuta As String
Public FechaDesde As Date
Public FechaHasta As Date
Public IDOrigen As Long
Public Sube As String
Public IDDestino As Long
Public Baja As String
Public FechaHoraCreacion As Date
Public IDUsuarioCreacion As Integer
Public FechaHoraModificacion As Date
Public IDUsuarioModificacion As Integer

Public NoMatchRaiseError As Boolean
Public NoMatch As Boolean

Public RefreshList As Boolean

Private mrecData As ADODB.Recordset

Public Property Get Hora_Formatted() As String
    Hora_Formatted = Format(Hora, "Short Time")
End Property

Public Property Get Persona() As Persona
    Dim PersonaObject As Persona
    
    Set PersonaObject = New Persona
    PersonaObject.IDPersona = IDPersona
    Call PersonaObject.Load
    Set Persona = PersonaObject
End Property

Private Function OpenRecordset() As Boolean
    Dim cmdData As ADODB.command

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
        
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_PersonaHorario_Data"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDPersona_FILTER", adInteger, adParamInput, , IDPersona)
    cmdData.Parameters.Append cmdData.CreateParameter("DiaSemana_FILTER", adTinyInt, adParamInput, , DiaSemana)
    cmdData.Parameters.Append cmdData.CreateParameter("Hora_FILTER", adChar, adParamInput, 8, Format(Hora, "hh:nn:ss"))
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta_FILTER", adChar, adParamInput, 20, IDRuta)
    Set mrecData = New ADODB.Recordset
    mrecData.Open cmdData, , adOpenKeyset, adLockOptimistic
    mrecData.Properties("Update Criteria").value = adCriteriaKey
    Set cmdData = Nothing
    
    Screen.MousePointer = vbDefault
    OpenRecordset = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.PersonaHorario.OpenRecordset", "Error al abrir la tabla de Horarios de las Personas."
End Function

Public Function Load() As Boolean
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    If Not OpenRecordset() Then
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    If (Not NoMatchRaiseError) And mrecData.EOF Then
        NoMatch = True
        
        InitializeValues
    Else
        NoMatch = False
        
        FechaDesde = IIf(IsNull(mrecData("FechaDesde").value), DATE_TIME_FIELD_NULL_VALUE, mrecData("FechaDesde").value)
        FechaHasta = IIf(IsNull(mrecData("FechaHasta").value), DATE_TIME_FIELD_NULL_VALUE, mrecData("FechaHasta").value)
        IDOrigen = Val(mrecData("IDOrigen").value & "")
        Sube = mrecData("Sube").value & ""
        IDDestino = Val(mrecData("IDDestino").value & "")
        Baja = mrecData("Baja").value & ""
        FechaHoraCreacion = mrecData("FechaHoraCreacion").value
        IDUsuarioCreacion = mrecData("IDUsuarioCreacion").value
        FechaHoraModificacion = mrecData("FechaHoraModificacion").value
        IDUsuarioModificacion = mrecData("IDUsuarioModificacion").value
    End If
    
    Screen.MousePointer = vbDefault
    Load = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.PersonaHorario.Load", "Error al obtener los datos del Horario de la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona & vbCr & "DiaSemana: " & DiaSemana & vbCr & "Hora: " & Hora_Formatted & vbCr & "IDRuta: " & IDRuta
End Function

Public Function AddNew() As Boolean
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    If mrecData Is Nothing Then
        If Not OpenRecordset() Then
            Exit Function
        End If
    End If
    
    Screen.MousePointer = vbHourglass
        
    mrecData.AddNew
    mrecData("IDPersona").value = IDPersona
    mrecData("DiaSemana").value = DiaSemana
    mrecData("Hora").value = DATE_TIME_FIELD_NULL_VALUE & " " & Hora
    mrecData("IDRuta").value = IDRuta
    mrecData("FechaDesde").value = IIf(FechaDesde = DATE_TIME_FIELD_NULL_VALUE, Null, FechaDesde)
    mrecData("FechaHasta").value = IIf(FechaHasta = DATE_TIME_FIELD_NULL_VALUE, Null, FechaHasta)
    mrecData("IDOrigen").value = IIf(IDOrigen = 0, Null, IDOrigen)
    mrecData("Sube").value = IIf(Trim(Sube) = "", Null, Sube)
    mrecData("IDDestino").value = IIf(IDDestino = 0, Null, IDDestino)
    mrecData("Baja").value = IIf(Trim(Baja) = "", Null, Baja)
    mrecData("FechaHoraCreacion").value = Now
    mrecData("IDUsuarioCreacion").value = pUsuario.IDUsuario
    mrecData("FechaHoraModificacion").value = Now
    mrecData("IDUsuarioModificacion").value = pUsuario.IDUsuario
    mrecData.Update
    
    Call GenerarReservasFijas
    
    If RefreshList Then
        RefreshList_RefreshPersonaHorario IDPersona, DiaSemana, Hora, IDRuta
        RefreshList_RefreshViajeDetalle Date, "", 0, True
        RefreshList_RefreshViaje Date, ""
        RefreshList_RefreshCuentaCorriente 0
    End If
    
    Screen.MousePointer = vbDefault
    AddNew = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe el Horario para el Pasajero.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.PersonaHorario.AddNew", "Error al Agregar el Horario de la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona & vbCr & "DiaSemana: " & DiaSemana & vbCr & "Hora: " & Hora_Formatted & vbCr & "IDRuta: " & IDRuta
        End If
    Else
        ShowErrorMessage "Classes.PersonaHorario.AddNew", "Error al Agregar el Horario de la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona & vbCr & "DiaSemana: " & DiaSemana & vbCr & "Hora: " & Hora_Formatted & vbCr & "IDRuta: " & IDRuta
    End If
    If mrecData.EditMode = adEditInProgress Then
        mrecData.CancelUpdate
    End If
End Function

Public Function Delete() As Boolean
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    mrecData.Delete
    
    Call EliminarReservasFijas
    
    If RefreshList Then
        RefreshList_RefreshPersonaHorario IDPersona, 0, Date, ""
        RefreshList_RefreshViajeDetalle Date, "", 0, True
        RefreshList_RefreshViaje Date, ""
        RefreshList_RefreshCuentaCorriente 0
    End If
    
    Screen.MousePointer = vbDefault
    Delete = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.PersonaHorario.Delete", "Error al Eliminar el Horario de la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona & vbCr & "DiaSemana: " & DiaSemana & vbCr & "Hora: " & Hora_Formatted & vbCr & "IDRuta: " & IDRuta
    mrecData.CancelUpdate
End Function

Private Sub InitializeValues()
    FechaDesde = DATE_TIME_FIELD_NULL_VALUE
    FechaHasta = DATE_TIME_FIELD_NULL_VALUE
    IDOrigen = 0
    Sube = ""
    IDDestino = 0
    Baja = ""
    FechaHoraCreacion = DATE_TIME_FIELD_NULL_VALUE
    IDUsuarioCreacion = 0
    FechaHoraModificacion = DATE_TIME_FIELD_NULL_VALUE
    IDUsuarioModificacion = 0
End Sub

Private Sub Class_Initialize()
    InitializeValues
    NoMatchRaiseError = True
    RefreshList = True
End Sub

Private Sub Class_Terminate()
    If Not mrecData Is Nothing Then
        If mrecData.State = adStateOpen Then
            If Not (mrecData.BOF Or mrecData.EOF) Then
                If mrecData.EditMode <> adEditNone Then
                    mrecData.CancelUpdate
                End If
            End If
            mrecData.Close
        End If
        Set mrecData = Nothing
    End If
End Sub

'============================================================================
' Genera las Reservas a Futuro de un Pasajero, dado un Horario
'============================================================================
Private Function GenerarReservasFijas() As Boolean
    Dim ViajeDetalle As ViajeDetalle
    Dim Viaje As Viaje
    
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_ViajeDetalle_ByPersonaHorario"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDPersona_FILTER", adInteger, adParamInput, , IDPersona)
    cmdData.Parameters.Append cmdData.CreateParameter("DiaSemanaBase_FILTER", adTinyInt, adParamInput, , DiaSemana)
    cmdData.Parameters.Append cmdData.CreateParameter("Hora_FILTER", adChar, adParamInput, 8, Format(Hora, "hh:nn:ss"))
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta_FILTER", adChar, adParamInput, 20, IDRuta)
    cmdData.Parameters.Append cmdData.CreateParameter("FechaDesde_FILTER", adChar, adParamInput, 10, IIf(FechaDesde = DATE_TIME_FIELD_NULL_VALUE, Null, Format(FechaDesde, "yyyy/mm/dd")))
    cmdData.Parameters.Append cmdData.CreateParameter("FechaHasta_FILTER", adChar, adParamInput, 10, IIf(FechaHasta = DATE_TIME_FIELD_NULL_VALUE, Null, Format(FechaHasta, "yyyy/mm/dd")))
    Set recData = New ADODB.Recordset
    recData.Open cmdData, , adOpenForwardOnly, adLockReadOnly
    Set cmdData = Nothing
    
    Do While Not recData.EOF
        'Si no existe la Reserva, la genero
        If IsNull(recData("Indice").value) Then
            Set ViajeDetalle = New ViajeDetalle
            With ViajeDetalle
                .RefreshListSkip = True
                .VerifyEstadoCondicionalOnAddNew = True
            
                .IDViaje = recData("IDViaje").value
                .FechaHora = recData("FechaHora").value
                .IDRuta = IDRuta
                .OcupanteTipo = OCUPANTE_TIPO_PASAJERO
                .IDPersona = IDPersona
                .ReservaTipo = VIAJE_DETALLE_RESERVA_TIPO_FIJA
                Set Viaje = New Viaje
                Viaje.FechaHora = recData("FechaHora").value
                Viaje.IDRuta = IDRuta
                If Viaje.Load() Then
                    If .GetNewValues(Viaje.DiaSemanaBase) Then
                        Viaje.IDOrigen = .IDOrigen
                        Viaje.IDDestino = .IDDestino
                        .Asiento = Viaje.Asiento_Asignar_GetAsiento(0)
                        
                        ViajeDetalle.ReservadoPor = "Reserva Fija Automática"
                        ViajeDetalle.Update
                    End If
                End If
                Set Viaje = Nothing
            End With
            
            Set ViajeDetalle = Nothing
        End If
        
        recData.MoveNext
    Loop
    
    recData.Close
    Set recData = Nothing
    
    Set ViajeDetalle = Nothing
    
    Screen.MousePointer = vbDefault
    GenerarReservasFijas = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.PersonaHorario.GenerarReservasFijas", "Error al Generar las Reservas Fijas del Pasajero en el Horario." & vbCr & vbCr & "IDPersona: " & IDPersona & vbCr & "Hora: " & Format(Hora, "Short Time") & vbCr & "IDRuta: " & IDRuta
End Function

'============================================================================
' Elimina las Reservas a Futuro de un Pasajero, dado un Horario
'============================================================================
Public Function EliminarReservasFijas() As Boolean
    Dim ViajeDetalle As ViajeDetalle
    
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_ViajeDetalle_ByPersonaHorario"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDPersona_FILTER", adInteger, adParamInput, , IDPersona)
    cmdData.Parameters.Append cmdData.CreateParameter("DiaSemanaBase_FILTER", adTinyInt, adParamInput, , DiaSemana)
    cmdData.Parameters.Append cmdData.CreateParameter("Hora_FILTER", adChar, adParamInput, 8, Format(Hora, "hh:nn:ss"))
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta_FILTER", adChar, adParamInput, 20, IDRuta)
    cmdData.Parameters.Append cmdData.CreateParameter("FechaDesde_FILTER", adChar, adParamInput, 10, IIf(FechaDesde = DATE_TIME_FIELD_NULL_VALUE, Null, Format(FechaDesde, "yyyy/mm/dd")))
    cmdData.Parameters.Append cmdData.CreateParameter("FechaHasta_FILTER", adChar, adParamInput, 10, IIf(FechaHasta = DATE_TIME_FIELD_NULL_VALUE, Null, Format(FechaHasta, "yyyy/mm/dd")))
    Set recData = New ADODB.Recordset
    recData.Open cmdData, , adOpenForwardOnly, adLockReadOnly
    Set cmdData = Nothing
    
    Do While Not recData.EOF
        'Si existe la Reserva, la elimino
        If Not IsNull(recData("Indice").value) Then
            Set ViajeDetalle = New ViajeDetalle
            ViajeDetalle.RefreshListSkip = True
            
            ViajeDetalle.FechaHora = recData("FechaHora").value
            ViajeDetalle.IDRuta = IDRuta
            ViajeDetalle.Indice = recData("Indice").value
            If ViajeDetalle.Load Then
                Call ViajeDetalle.Delete
            End If
            
            Set ViajeDetalle = Nothing
        End If
        
        recData.MoveNext
    Loop
    
    recData.Close
    Set recData = Nothing
    
    Set ViajeDetalle = Nothing
    
    Screen.MousePointer = vbDefault
    EliminarReservasFijas = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.PersonaHorario.EliminarReservasFijas", "Error al Eliminar las Reservas Fijas del Pasajero en el Horario." & vbCr & vbCr & "IDPersona: " & IDPersona & vbCr & "Hora: " & Format(Hora, "Short Time") & vbCr & "IDRuta: " & IDRuta
End Function
