VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmViajePropiedadMultiple 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades de Múltiples Viajes"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ViajePropiedadMultiple.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3225
   ScaleWidth      =   6090
   Begin VB.CheckBox chkConductor2Modificar 
      Caption         =   "Modificar"
      Height          =   210
      Left            =   1200
      TabIndex        =   11
      Top             =   2040
      Width           =   1035
   End
   Begin VB.TextBox txtConductor2 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   2340
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   " Selección Múltiple"
      Top             =   1980
      Width           =   3630
   End
   Begin VB.TextBox txtConductor 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   2340
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   " Selección Múltiple"
      Top             =   1500
      Width           =   3630
   End
   Begin VB.CheckBox chkConductorModificar 
      Caption         =   "Modificar"
      Height          =   210
      Left            =   1200
      TabIndex        =   6
      Top             =   1560
      Width           =   1035
   End
   Begin VB.CheckBox chkVehiculoModificar 
      Caption         =   "Modificar"
      Height          =   210
      Left            =   1200
      TabIndex        =   1
      Top             =   1080
      Width           =   1035
   End
   Begin VB.TextBox txtVehiculo 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   2340
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   " Selección Múltiple"
      Top             =   1020
      Width           =   3630
   End
   Begin VB.CommandButton cmdConductor 
      Caption         =   "..."
      Height          =   315
      Left            =   5700
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Personas"
      Top             =   1500
      Width           =   255
   End
   Begin VB.CommandButton cmdVehiculo 
      Caption         =   "..."
      Height          =   315
      Left            =   5700
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Vehículos"
      Top             =   1020
      Width           =   255
   End
   Begin VB.Frame fraLine 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   75
      Left            =   120
      TabIndex        =   18
      Top             =   780
      Width           =   5835
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3420
      TabIndex        =   16
      Top             =   2700
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2100
      TabIndex        =   15
      Top             =   2700
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo datcboVehiculo 
      Height          =   330
      Left            =   2340
      TabIndex        =   3
      Top             =   1020
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   582
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo datcboConductor 
      Height          =   330
      Left            =   2340
      TabIndex        =   8
      Top             =   1500
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   582
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo datcboConductor2 
      Height          =   330
      Left            =   2340
      TabIndex        =   13
      Top             =   1980
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   582
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdConductor2 
      Caption         =   "..."
      Height          =   315
      Left            =   5700
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Personas"
      Top             =   1980
      Width           =   255
   End
   Begin VB.Label lblConductor2 
      AutoSize        =   -1  'True
      Caption         =   "&Conductor 2:"
      Height          =   210
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   930
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   2
      Left            =   600
      Picture         =   "ViajePropiedadMultiple.frx":054A
      Top             =   120
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   1
      Left            =   360
      Picture         =   "ViajePropiedadMultiple.frx":0E2C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblConductor 
      AutoSize        =   -1  'True
      Caption         =   "&Conductor:"
      Height          =   210
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   795
   End
   Begin VB.Label lblVehiculo 
      AutoSize        =   -1  'True
      Caption         =   "&Vehículo:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   675
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Modifique aquí los Datos de los Viajes"
      Height          =   210
      Left            =   1740
      TabIndex        =   17
      Top             =   300
      Width           =   2730
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "ViajePropiedadMultiple.frx":170E
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmViajePropiedadMultiple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCViaje As Collection

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByVal CFechaHora As Collection, ByVal CIDRuta As Collection)
    Dim Index As Long
    Dim FechaHora As Variant
    Dim IDRuta As Variant
    Dim Viaje As Viaje
    
    Dim IDVehiculo As Long
    Dim IDConductor As Long
    Dim IDConductor2 As Long
    
    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    Set mCViaje = New Collection
    
    For Index = 1 To CFechaHora.Count
        FechaHora = CFechaHora(Index)
        IDRuta = CIDRuta(Index)
        
        Set Viaje = New Viaje
        With Viaje
            .FechaHora = FechaHora
            .IDRuta = IDRuta
            If Not .Load() Then
                Unload Me
                Exit Sub
            End If
            If Index = 1 Then
                IDVehiculo = .IDVehiculo
                IDConductor = .IDConductor
                IDConductor2 = .IDConductor2
            Else
                If IDVehiculo <> .IDVehiculo Then
                    IDVehiculo = -1
                End If
                If IDConductor <> .IDConductor Then
                    IDConductor = -1
                End If
                If IDConductor2 <> .IDConductor2 Then
                    IDConductor2 = -1
                End If
            End If
            
            mCViaje.Add Viaje
        End With
    Next Index
    
    If Not CSM_Control_DataCombo.FillFromSQL(datcboVehiculo, "(SELECT -1 AS IDVehiculo, '------------------' AS Nombre, 1 AS Orden FROM Vehiculo) UNION (SELECT IDVehiculo, Nombre, 2 AS Orden FROM Vehiculo WHERE Activo = 1) ORDER BY Orden, Nombre", "IDVehiculo", "Nombre", "Vehículos", cscpItemOrFirst, IDVehiculo) Then
        Unload Me
        Exit Sub
    End If
    
    If Not CSM_Control_DataCombo.FillFromSQL(datcboConductor, "(SELECT -1 AS IDPersona, '------------------' AS ApellidoNombre, 1 AS Orden FROM Persona) UNION (SELECT IDPersona, dbo.udf_GetEntidadApellidoYNombre(Apellido, Nombre) AS ApellidoNombre, 2 AS Orden FROM Persona WHERE Activo = 1 AND EntidadTipo = '" & ENTIDAD_TIPO_PERSONA_CONDUCTOR & "') ORDER BY Orden, ApellidoNombre", "IDPersona", "ApellidoNombre", "Conductores", cscpItemOrFirst, IDConductor) Then
        Unload Me
        Exit Sub
    End If
    
    If Not CSM_Control_DataCombo.FillFromSQL(datcboConductor2, "(SELECT -1 AS IDPersona, '------------------' AS ApellidoNombre, 1 AS Orden FROM Persona) UNION (SELECT IDPersona, dbo.udf_GetEntidadApellidoYNombre(Apellido, Nombre) AS ApellidoNombre, 2 AS Orden FROM Persona WHERE Activo = 1 AND EntidadTipo = '" & ENTIDAD_TIPO_PERSONA_CONDUCTOR & "') ORDER BY Orden, ApellidoNombre", "IDPersona", "ApellidoNombre", "Conductores", cscpItemOrFirst, IDConductor2) Then
        Unload Me
        Exit Sub
    End If
    
    chkVehiculoModificar.value = IIf(IDVehiculo = -1, vbUnchecked, vbChecked)
    chkVehiculoModificar.Visible = (IDVehiculo = -1)
    txtVehiculo.Visible = (IDVehiculo = -1)
    datcboVehiculo.Visible = (IDVehiculo <> -1)
    cmdVehiculo.Visible = (IDVehiculo <> -1)
    
    chkConductorModificar.value = IIf(IDConductor = -1, vbUnchecked, vbChecked)
    chkConductorModificar.Visible = (IDConductor = -1)
    txtConductor.Visible = (IDConductor = -1)
    datcboConductor.Visible = (IDConductor <> -1)
    cmdConductor.Visible = (IDConductor <> -1)
    
    chkConductor2Modificar.value = IIf(IDConductor2 = -1, vbUnchecked, vbChecked)
    chkConductor2Modificar.Visible = (pParametro.Viaje_Permite_2_Conductores And IDConductor2 = -1)
    txtConductor2.Visible = (pParametro.Viaje_Permite_2_Conductores And IDConductor = -1)
    datcboConductor2.Visible = (pParametro.Viaje_Permite_2_Conductores And IDConductor <> -1)
    cmdConductor2.Visible = (pParametro.Viaje_Permite_2_Conductores And IDConductor <> -1)
    
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Public Sub FillComboBoxVehiculo()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboVehiculo.BoundText)
    Set recData = datcboVehiculo.RowSource
    recData.Requery
    Set recData = Nothing
    datcboVehiculo.BoundText = KeySave
End Sub

Public Sub FillComboBoxConductor()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboConductor.BoundText)
    Set recData = datcboConductor.RowSource
    recData.Requery
    Set recData = Nothing
    datcboConductor.BoundText = KeySave

    If pParametro.Viaje_Permite_2_Conductores Then
        KeySave = Val(datcboConductor.BoundText)
        Set recData = datcboConductor.RowSource
        recData.Requery
        Set recData = Nothing
        datcboConductor.BoundText = KeySave
    End If
End Sub

Private Sub chkConductorModificar_Click()
    txtConductor.Visible = (chkConductorModificar.value = vbUnchecked)
    datcboConductor.Visible = (chkConductorModificar.value = vbChecked)
    cmdConductor.Visible = (chkConductorModificar.value = vbChecked)
End Sub

Private Sub chkConductor2Modificar_Click()
    txtConductor2.Visible = (chkConductor2Modificar.value = vbUnchecked)
    datcboConductor2.Visible = (chkConductor2Modificar.value = vbChecked)
    cmdConductor2.Visible = (chkConductor2Modificar.value = vbChecked)
End Sub

Private Sub chkVehiculoModificar_Click()
    txtVehiculo.Visible = (chkVehiculoModificar.value = vbUnchecked)
    datcboVehiculo.Visible = (chkVehiculoModificar.value = vbChecked)
    cmdVehiculo.Visible = (chkVehiculoModificar.value = vbChecked)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim Viaje As Viaje
    
    If chkVehiculoModificar.value = vbChecked Or chkConductorModificar.value = vbChecked Then
        If pParametro.Viaje_Permite_2_Conductores Then
            If Val(datcboConductor.BoundText) = -1 And Val(datcboConductor2.BoundText) > -1 Then
                MsgBox "Si selecciona el Conductor N° 2, debe seleccionar el Conductor N° 1.", vbInformation, App.Title
                datcboConductor.SetFocus
                Exit Sub
            End If
        End If
    
        For Each Viaje In mCViaje
            With Viaje
                .RefreshListSkip = True
                If chkVehiculoModificar.value = vbChecked Then
                    .IDVehiculo = IIf(Val(datcboVehiculo.BoundText) = -1, 0, Val(datcboVehiculo.BoundText))
                End If
                If chkConductorModificar.value = vbChecked Then
                    .IDConductor = IIf(Val(datcboConductor.BoundText) = -1, 0, Val(datcboConductor.BoundText))
                End If
                If chkConductor2Modificar.value = vbChecked Then
                    .IDConductor2 = IIf(Val(datcboConductor2.BoundText) = -1, 0, Val(datcboConductor2.BoundText))
                End If
                
                If Not .Update Then
                    RefreshList_RefreshViaje Now, ""
                    Exit Sub
                End If
            End With
        Next Viaje
        
        RefreshList_RefreshViaje Now, ""
    End If
    
    Unload Me
End Sub

Private Sub cmdConductor_Click()
    If pCPermiso.GotPermission(PERMISO_PERSONA) Then
        Call frmPersona.FindAndShowItem(Val(datcboConductor.BoundText), UCase(Left(datcboConductor.Text, 1)), Me.Name, "", "")
    End If
End Sub

Private Sub cmdConductor2_Click()
    If pCPermiso.GotPermission(PERMISO_PERSONA) Then
        Call frmPersona.FindAndShowItem(Val(datcboConductor2.BoundText), UCase(Left(datcboConductor2.Text, 1)), Me.Name, "", "")
    End If
End Sub

Private Sub cmdVehiculo_Click()
    If pCPermiso.GotPermission(PERMISO_VEHICULO) Then
        Screen.MousePointer = vbHourglass
        frmVehiculo.Show
        On Error Resume Next
        Set frmVehiculo.lvwData.SelectedItem = frmVehiculo.lvwData.ListItems(KEY_STRINGER & datcboVehiculo.BoundText)
        frmVehiculo.lvwData.SelectedItem.EnsureVisible
        If frmVehiculo.WindowState = vbMinimized Then
            frmVehiculo.WindowState = vbNormal
        End If
        frmVehiculo.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mCViaje = Nothing
    Set frmViajePropiedadMultiple = Nothing
End Sub

