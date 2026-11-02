VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Horario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'PROPERTIES VARIABLES
Private mDiaSemana As Integer
Private mHora As Date
Private mIDRuta As String
Private mRutaOtra As String
Private mIDPersona As Long
Private mKilometro As Integer
Private mDuracion As Integer
Private mImporte As Currency
Private mIDConductor As Long
Private mIDConductor2 As Long
Private mIDVehiculo As Long
Private mConductorImporteTramoCompleto As Currency
Private mConductorImporteTramo1 As Currency
Private mConductorImporteTramo2 As Currency
Private mNotas As String
Private mActivo As Boolean
Private mPersonal As Boolean
Private mFechaHoraCreacion As Date
Private mIDUsuarioCreacion As Integer
Private mFechaHoraModificacion As Date
Private mIDUsuarioModificacion As Integer

'STATE VARIABLES
Private mIsNew As Boolean
Private mIsCopy As Boolean
Private mIsDirty As Boolean
Private mNoMatch As Boolean

'BEHAVIOR VARIABLES
Private mNoMatchRaiseError As Boolean
Private mRefreshListSkip As Boolean

'DATABASE VARIABLES
Private mDatabase As CSC_Database_ADO_SQL

'INTERNAL VARIABLES
Private mrecData As ADODB.Recordset

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get DiaSemana() As Integer
    DiaSemana = mDiaSemana
End Property

Public Property Let DiaSemana(ByVal value As Integer)
    If value <> mDiaSemana Then
        mIsDirty = True
    End If
    mDiaSemana = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Hora() As Date
    Hora = mHora
End Property

Public Property Get Hora_Formatted() As String
    Hora_Formatted = Format(mHora, "Short Time")
End Property

Public Property Let Hora(ByVal value As Date)
    If value <> mHora Then
        mIsDirty = True
    End If
    mHora = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDRuta() As String
    IDRuta = mIDRuta
End Property

Public Property Let IDRuta(ByVal value As String)
    If value <> mIDRuta Then
        mIsDirty = True
    End If
    mIDRuta = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get RutaOtra() As String
    RutaOtra = mRutaOtra
End Property

Public Property Let RutaOtra(ByVal value As String)
    If value <> mRutaOtra Then
        mIsDirty = True
    End If
    mRutaOtra = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Ruta_DisplayName() As String
    Ruta_DisplayName = mIDRuta & IIf(mIDRuta = pParametro.Ruta_ID_Otra Or mIDRuta = pParametro.Ruta_Paquete_ID, ": " & mRutaOtra, "")
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDPersona() As Long
    IDPersona = mIDPersona
End Property

Public Property Let IDPersona(ByVal value As Long)
    If value <> mIDPersona Then
        mIsDirty = True
    End If
    mIDPersona = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Kilometro() As Integer
    Kilometro = mKilometro
End Property

Public Property Let Kilometro(ByVal value As Integer)
    If value <> mKilometro Then
        mIsDirty = True
    End If
    mKilometro = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Duracion() As Integer
    Duracion = mDuracion
End Property

Public Property Let Duracion(ByVal value As Integer)
    If value <> mDuracion Then
        mIsDirty = True
    End If
    mDuracion = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Importe() As Currency
    Importe = mImporte
End Property

Public Property Get Importe_Formatted() As String
    Importe_Formatted = Format(mImporte, "Currency")
End Property

Public Property Let Importe(ByVal value As Currency)
    If value <> mImporte Then
        mIsDirty = True
    End If
    mImporte = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDConductor() As Long
    IDConductor = mIDConductor
End Property

Public Property Let IDConductor(ByVal value As Long)
    If value <> mIDConductor Then
        mIsDirty = True
    End If
    mIDConductor = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDConductor2() As Long
    IDConductor2 = mIDConductor2
End Property

Public Property Let IDConductor2(ByVal value As Long)
    If value <> mIDConductor2 Then
        mIsDirty = True
    End If
    mIDConductor2 = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDVehiculo() As Long
    IDVehiculo = mIDVehiculo
End Property

Public Property Let IDVehiculo(ByVal value As Long)
    If value <> mIDVehiculo Then
        mIsDirty = True
    End If
    mIDVehiculo = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ConductorImporteTramoCompleto() As Currency
    ConductorImporteTramoCompleto = mConductorImporteTramoCompleto
End Property

Public Property Get ConductorImporteTramoCompleto_FormattedAsString() As String
    ConductorImporteTramoCompleto_FormattedAsString = IIf(mConductorImporteTramoCompleto = -1, "", Format(mConductorImporteTramoCompleto, "Currency"))
End Property

Public Property Let ConductorImporteTramoCompleto(ByVal value As Currency)
    If value <> mConductorImporteTramoCompleto Then
        mIsDirty = True
    End If
    mConductorImporteTramoCompleto = value
End Property

Public Property Let ConductorImporteTramoCompleto_FormattedAsString(ByVal value As String)
    If IsNumeric(value) Then
        ConductorImporteTramoCompleto = CCur(value)
    Else
        ConductorImporteTramoCompleto = -1
    End If
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ConductorImporteTramo1() As Currency
    ConductorImporteTramo1 = mConductorImporteTramo1
End Property

Public Property Get ConductorImporteTramo1_FormattedAsString() As String
    ConductorImporteTramo1_FormattedAsString = IIf(mConductorImporteTramo1 = -1, "", Format(mConductorImporteTramo1, "Currency"))
End Property

Public Property Let ConductorImporteTramo1(ByVal value As Currency)
    If value <> mConductorImporteTramo1 Then
        mIsDirty = True
    End If
    mConductorImporteTramo1 = value
End Property

Public Property Let ConductorImporteTramo1_FormattedAsString(ByVal value As String)
    If IsNumeric(value) Then
        ConductorImporteTramo1 = CCur(value)
    Else
        ConductorImporteTramo1 = -1
    End If
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ConductorImporteTramo2() As Currency
    ConductorImporteTramo2 = mConductorImporteTramo2
End Property

Public Property Get ConductorImporteTramo2_FormattedAsString() As String
    ConductorImporteTramo2_FormattedAsString = IIf(mConductorImporteTramo2 = -1, "", Format(mConductorImporteTramo2, "Currency"))
End Property

Public Property Let ConductorImporteTramo2(ByVal value As Currency)
    If value <> mConductorImporteTramo2 Then
        mIsDirty = True
    End If
    mConductorImporteTramo2 = value
End Property

Public Property Let ConductorImporteTramo2_FormattedAsString(ByVal value As String)
    If IsNumeric(value) Then
        ConductorImporteTramo2 = CCur(value)
    Else
        ConductorImporteTramo2 = -1
    End If
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Notas() As String
    Notas = mNotas
End Property

Public Property Let Notas(ByVal value As String)
    If value <> mNotas Then
        mIsDirty = True
    End If
    mNotas = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Activo() As Boolean
    Activo = mActivo
End Property

Public Property Let Activo(ByVal value As Boolean)
    If value <> mActivo Then
        mIsDirty = True
    End If
    mActivo = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Personal() As Boolean
    Personal = mPersonal
End Property

Public Property Let Personal(ByVal value As Boolean)
    If value <> mPersonal Then
        mIsDirty = True
    End If
    mPersonal = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get FechaHoraCreacion() As Date
    FechaHoraCreacion = mFechaHoraCreacion
End Property

Public Property Let FechaHoraCreacion(ByVal value As Date)
    If value <> mFechaHoraCreacion Then
        mIsDirty = True
    End If
    mFechaHoraCreacion = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDUsuarioCreacion() As String
    IDUsuarioCreacion = mIDUsuarioCreacion
End Property

Public Property Let IDUsuarioCreacion(ByVal value As String)
    If value <> mIDUsuarioCreacion Then
        mIsDirty = True
    End If
    mIDUsuarioCreacion = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get FechaHoraModificacion() As Date
    FechaHoraModificacion = mFechaHoraModificacion
End Property

Public Property Let FechaHoraModificacion(ByVal value As Date)
    If value <> mFechaHoraModificacion Then
        mIsDirty = True
    End If
    mFechaHoraModificacion = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDUsuarioModificacion() As String
    IDUsuarioModificacion = mIDUsuarioModificacion
End Property

Public Property Let IDUsuarioModificacion(ByVal value As String)
    If value <> mIDUsuarioModificacion Then
        mIsDirty = True
    End If
    mIDUsuarioModificacion = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IsNew() As Boolean
    IsNew = mIsNew
End Property

Public Property Get IsCopy() As Boolean
    IsCopy = mIsCopy
End Property

Public Property Get IsDirty() As Boolean
    IsDirty = mIsDirty
End Property

Public Property Get NoMatch() As Boolean
    NoMatch = mNoMatch
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get NoMatchRaiseError() As Boolean
    NoMatchRaiseError = mNoMatchRaiseError
End Property

Public Property Let NoMatchRaiseError(ByVal value As Boolean)
    mNoMatchRaiseError = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get RefreshListSkip() As Boolean
    RefreshListSkip = mRefreshListSkip
End Property

Public Property Let RefreshListSkip(ByVal value As Boolean)
    mRefreshListSkip = value
End Property

Public Sub RefreshList()
    RefreshList_Module.RefreshList_RefreshHorario DiaSemana, Hora, IDRuta
End Sub

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Database() As CSC_Database_ADO_SQL
    Set Database = mDatabase
End Property

Public Property Set Database(ByRef value As CSC_Database_ADO_SQL)
    Set mDatabase = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function OpenRecordset() As Boolean
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
        
    Screen.MousePointer = vbHourglass
      
    Set mrecData = New ADODB.Recordset
    Set mrecData.ActiveConnection = pDatabase.Connection
    mrecData.CursorType = adOpenKeyset
    mrecData.LockType = adLockOptimistic
    mrecData.Source = "SELECT * FROM Horario"
    mrecData.Open , , , , adCmdText
    mrecData.Properties("Update Criteria").value = adCriteriaKey

    Screen.MousePointer = vbDefault
    OpenRecordset = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.Horario.OpenRecordset", "Error al abrir la tabla de Horario."
    If Not mrecData Is Nothing Then
        Set mrecData = Nothing
    End If
End Function

Public Function Load() As Boolean
    If mrecData Is Nothing Then
        If Not OpenRecordset() Then
            Exit Function
        End If
    Else
        If mrecData.State <> adStateOpen Then
            If Not OpenRecordset() Then
                Exit Function
            End If
        End If
    End If

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass

    Call InitializeValues

    mrecData.Filter = "DiaSemana = " & mDiaSemana & " AND Hora = #" & DATE_TIME_FIELD_NULL_VALUE & " " & mHora & "# AND IDRuta = '" & mIDRuta & "'"

    If (Not mNoMatchRaiseError) And mrecData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False
    
        mDiaSemana = mrecData("DiaSemana").value
        mHora = Format(mrecData("Hora").value, "Long Time")
        
        mIDRuta = RTrim(mrecData("IDRuta").value)
        Select Case mIDRuta
            Case pParametro.Ruta_ID_Otra
                mRutaOtra = mrecData("RutaOtra").value & ""
                mIDPersona = Val(mrecData("IDPersona").value & "")
                mImporte = IIf(IsNull(mrecData("Importe").value), 0, mrecData("Importe").value)
            Case pParametro.Ruta_Paquete_ID
                mRutaOtra = mrecData("RutaOtra").value & ""
                mIDPersona = 0
                mImporte = IIf(IsNull(mrecData("Importe").value), 0, mrecData("Importe").value)
            Case Else
                mRutaOtra = ""
                mIDPersona = 0
                mImporte = 0
        End Select
        mKilometro = Val(mrecData("Kilometro").value & "")
        mDuracion = Val(mrecData("Duracion").value & "")
            
        mIDConductor = Val(mrecData("IDConductor").value & "")
        mIDConductor2 = Val(mrecData("IDConductor2").value & "")
        mIDVehiculo = Val(mrecData("IDVehiculo").value & "")
        mConductorImporteTramoCompleto = IIf(IsNull(mrecData("ConductorImporteTramoCompleto").value), -1, mrecData("ConductorImporteTramoCompleto").value)
        mConductorImporteTramo1 = IIf(IsNull(mrecData("ConductorImporteTramo1").value), -1, mrecData("ConductorImporteTramo1").value)
        mConductorImporteTramo2 = IIf(IsNull(mrecData("ConductorImporteTramo2").value), -1, mrecData("ConductorImporteTramo2").value)
        mNotas = mrecData("Notas").value & ""
        mActivo = mrecData("Activo").value
        mPersonal = mrecData("Personal").value
        mFechaHoraCreacion = mrecData("FechaHoraCreacion").value
        mIDUsuarioCreacion = mrecData("IDUsuarioCreacion").value
        mFechaHoraModificacion = mrecData("FechaHoraModificacion").value
        mIDUsuarioModificacion = mrecData("IDUsuarioModificacion").value
    End If
    
    Screen.MousePointer = vbDefault
    Load = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Horario.Load", "Error al obtener los datos del Horario." & vbCr & vbCr & "DiaSemana: " & DiaSemana & vbCr & "Hora: " & Hora & vbCr & "IDRuta: " & IDRuta
End Function

Public Function Copy() As Boolean
    If Load() Then
        Copy = True

        mDiaSemana = 0
        mHora = DATE_TIME_FIELD_NULL_VALUE
        mIDRuta = ""

        mIsNew = True
        mIsCopy = True
        mIsDirty = True
    End If
End Function

Public Function Update() As Boolean
    If Not mIsDirty Then
        Update = True
        Exit Function
    End If

    If mrecData Is Nothing Then
        If Not OpenRecordset() Then
            Exit Function
        End If
    Else
        If mrecData.State <> adStateOpen Then
            If Not OpenRecordset() Then
                Exit Function
            End If
        End If
    End If
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    If mIsNew Then
        mrecData.AddNew
    End If
    mrecData("DiaSemana").value = mDiaSemana
    mrecData("Hora").value = DATE_TIME_FIELD_NULL_VALUE & " " & mHora
    
    mrecData("IDRuta").value = mIDRuta
    Select Case mIDRuta
        Case pParametro.Ruta_ID_Otra
            mrecData("RutaOtra").value = IIf(mRutaOtra = "", Null, mRutaOtra)
            mrecData("IDPersona").value = IIf(mIDPersona = 0, Null, mIDPersona)
            mrecData("Importe").value = IIf(mImporte = 0, Null, mImporte)
        Case pParametro.Ruta_Paquete_ID
            mrecData("RutaOtra").value = IIf(mRutaOtra = "", Null, mRutaOtra)
            mrecData("IDPersona").value = Null
            mrecData("Importe").value = IIf(mImporte = 0, Null, mImporte)
        Case Else
            mrecData("RutaOtra").value = Null
            mrecData("IDPersona").value = Null
            mrecData("Importe").value = Null
    End Select
    mrecData("Kilometro").value = IIf(mKilometro = 0, Null, mKilometro)
    mrecData("Duracion").value = IIf(mDuracion = 0, Null, mDuracion)
    
    mrecData("IDConductor").value = IIf(mIDConductor = 0, Null, mIDConductor)
    mrecData("IDConductor2").value = IIf(mIDConductor2 = 0, Null, mIDConductor2)
    mrecData("IDVehiculo").value = IIf(mIDVehiculo = 0, Null, mIDVehiculo)
    mrecData("ConductorImporteTramoCompleto").value = IIf(mConductorImporteTramoCompleto = -1, Null, mConductorImporteTramoCompleto)
    mrecData("ConductorImporteTramo1").value = IIf(mConductorImporteTramo1 = -1, Null, mConductorImporteTramo1)
    mrecData("ConductorImporteTramo2").value = IIf(mConductorImporteTramo2 = -1, Null, mConductorImporteTramo2)
    mrecData("Notas").value = IIf(Trim(mNotas) = "", Null, mNotas)
    mrecData("Activo").value = mActivo
    mrecData("Personal").value = mPersonal
    If mIsNew Then
        mrecData("FechaHoraCreacion").value = mFechaHoraCreacion
        mrecData("IDUsuarioCreacion").value = pUsuario.IDUsuario
    End If
    mrecData("FechaHoraModificacion").value = Now
    mrecData("IDUsuarioModificacion").value = pUsuario.IDUsuario
    mrecData.Update

    mIsNew = False
    mIsCopy = False
    mIsDirty = False
    
    If Not mRefreshListSkip Then
        Call RefreshList
    End If
    
    Screen.MousePointer = vbDefault
    Update = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe un Horario con el mismo nombre.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.Horario.Update", "Error al actualizar el Horario." & vbCr & vbCr & "DiaSemana: " & DiaSemana & vbCr & "Hora: " & Hora & vbCr & "IDRuta: " & IDRuta
        End If
        If mrecData.EditMode = adEditAdd Or mrecData.EditMode = adEditInProgress Then
            mrecData.CancelUpdate
        End If
    Else
        ShowErrorMessage "Classes.Horario.Update", "Error al actualizar el Horario." & vbCr & vbCr & "DiaSemana: " & DiaSemana & vbCr & "Hora: " & Hora & vbCr & "IDRuta: " & IDRuta
    End If
End Function

Public Function Delete() As Boolean
    If mrecData Is Nothing Then
        If Not OpenRecordset() Then
            Exit Function
        End If
    Else
        If mrecData.State <> adStateOpen Then
            If Not OpenRecordset() Then
                Exit Function
            End If
        End If
    End If
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass
    
    mrecData.Delete
    
    If Not mRefreshListSkip Then
        Call RefreshList
    End If
    
    Screen.MousePointer = vbDefault
    Delete = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_RELATED_RECORDS Then
            Screen.MousePointer = vbDefault
            MsgBox "No se puede eliminar el Horario debido a que tiene datos relacionados.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.Horario.Delete", "Error al eliminar el Horario." & vbCr & vbCr & "DiaSemana: " & DiaSemana & vbCr & "Hora: " & Hora & vbCr & "IDRuta: " & IDRuta
        End If
        If mrecData.EditMode = adEditDelete Then
            mrecData.CancelUpdate
        End If
    Else
        ShowErrorMessage "Classes.Horario.Delete", "Error al eliminar el Horario." & vbCr & vbCr & "DiaSemana: " & DiaSemana & vbCr & "Hora: " & Hora & vbCr & "IDRuta: " & IDRuta
    End If
End Function

Private Sub InitializeValues()
    mRutaOtra = ""
    mIDPersona = 0
    mKilometro = 0
    mDuracion = 0
    mImporte = 0
    mIDConductor = 0
    mIDConductor2 = 0
    mIDVehiculo = 0
    mConductorImporteTramoCompleto = -1
    mConductorImporteTramo1 = -1
    mConductorImporteTramo2 = -1
    mNotas = ""
    mActivo = True
    mPersonal = False
    mFechaHoraCreacion = DATE_TIME_FIELD_NULL_VALUE
    mIDUsuarioCreacion = 0
    mFechaHoraModificacion = DATE_TIME_FIELD_NULL_VALUE
    mIDUsuarioModificacion = 0

    mIsNew = True
    mIsCopy = False
    mIsDirty = False
End Sub

Private Sub Class_Initialize()
    InitializeValues
    mNoMatchRaiseError = True
    mRefreshListSkip = False
End Sub

Private Sub Class_Terminate()
    If Not mrecData Is Nothing Then
        If mrecData.State = adStateOpen Then
            mrecData.Close
        End If
        Set mrecData = Nothing
    End If
End Sub
