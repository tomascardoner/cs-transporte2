VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmHorarioPropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9540
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "HorarioPropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   9540
   Begin VB.TextBox txtDuracion 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3180
      MaxLength       =   5
      TabIndex        =   17
      Top             =   3720
      Width           =   795
   End
   Begin VB.TextBox txtKilometro 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1020
      MaxLength       =   4
      TabIndex        =   15
      Top             =   3720
      Width           =   795
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1020
      MaxLength       =   20
      TabIndex        =   13
      Top             =   3180
      Width           =   1455
   End
   Begin VB.TextBox txtPersona 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2820
      Width           =   2955
   End
   Begin VB.CommandButton cmdPersona 
      Height          =   315
      Left            =   3960
      Picture         =   "HorarioPropiedad.frx":054A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Buscar..."
      Top             =   2820
      Width           =   315
   End
   Begin VB.CommandButton cmdPersonaClear 
      Height          =   315
      Left            =   4320
      Picture         =   "HorarioPropiedad.frx":0AD4
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Borrar"
      Top             =   2820
      Width           =   315
   End
   Begin VB.TextBox txtRutaOtra 
      Height          =   315
      Left            =   1020
      MaxLength       =   50
      TabIndex        =   7
      Top             =   2280
      Width           =   3615
   End
   Begin VB.CommandButton cmdAuditoria 
      Caption         =   "Auditoría"
      Height          =   615
      Left            =   8400
      Picture         =   "HorarioPropiedad.frx":105E
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   60
      Width           =   975
   End
   Begin VB.CheckBox chkActivo 
      Alignment       =   1  'Right Justify
      Caption         =   "&Activo"
      Height          =   210
      Left            =   4860
      TabIndex        =   36
      Top             =   5520
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.Frame fraConductorImporte 
      Caption         =   "Importe a acreditar al conductor:"
      Height          =   1695
      Left            =   4860
      TabIndex        =   26
      Top             =   2580
      Width           =   4515
      Begin VB.TextBox txtConductorImporteTramoCompleto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1620
         MaxLength       =   20
         TabIndex        =   28
         Tag             =   "CURRENCY|EMPTY|ZERO|POSITIVE"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtConductorImporteTramo1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1620
         MaxLength       =   20
         TabIndex        =   30
         Tag             =   "CURRENCY|EMPTY|ZERO|POSITIVE"
         Top             =   780
         Width           =   1455
      End
      Begin VB.TextBox txtConductorImporteTramo2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1620
         MaxLength       =   20
         TabIndex        =   32
         Tag             =   "CURRENCY|EMPTY|ZERO|POSITIVE"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblConductorImporteTramoCompleto 
         AutoSize        =   -1  'True
         Caption         =   "Tramo completo:"
         Height          =   210
         Left            =   180
         TabIndex        =   27
         Top             =   420
         Width           =   1185
      End
      Begin VB.Label lblConductorImporteTramo1 
         AutoSize        =   -1  'True
         Caption         =   "Tramo 1:"
         Height          =   210
         Left            =   180
         TabIndex        =   29
         Top             =   840
         Width           =   630
      End
      Begin VB.Label lblConductorImporteTramo2 
         AutoSize        =   -1  'True
         Caption         =   "Tramo 2:"
         Height          =   210
         Left            =   180
         TabIndex        =   31
         Top             =   1260
         Width           =   630
      End
   End
   Begin VB.CheckBox chkPersonal 
      Height          =   195
      Left            =   5400
      TabIndex        =   34
      Top             =   4500
      Width           =   195
   End
   Begin VB.ComboBox cboDiaSemana 
      Height          =   330
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1020
      Width           =   3510
   End
   Begin VB.CommandButton cmdVehiculo 
      Caption         =   "..."
      Height          =   315
      Left            =   9120
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Vehículos"
      Top             =   1020
      Width           =   255
   End
   Begin VB.CommandButton cmdRuta 
      Caption         =   "..."
      Height          =   315
      Left            =   4380
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Rutas"
      Top             =   1860
      Width           =   255
   End
   Begin MSComCtl2.DTPicker dtpHora 
      Height          =   315
      Left            =   1020
      TabIndex        =   3
      Top             =   1440
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "HH:mm"
      Format          =   119668739
      CurrentDate     =   36494
   End
   Begin VB.TextBox txtNotas 
      Height          =   945
      Left            =   5760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   35
      Top             =   4440
      Width           =   3615
   End
   Begin VB.Frame fraLine 
      Height          =   75
      Left            =   120
      TabIndex        =   40
      Top             =   780
      Width           =   9255
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   8160
      TabIndex        =   38
      Top             =   5820
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   6840
      TabIndex        =   37
      Top             =   5820
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo datcboRuta 
      Height          =   330
      Left            =   1020
      TabIndex        =   5
      Top             =   1860
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
   Begin MSDataListLib.DataCombo datcboVehiculo 
      Height          =   330
      Left            =   5880
      TabIndex        =   20
      Top             =   1020
      Width           =   3195
      _ExtentX        =   5636
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
      Left            =   5880
      TabIndex        =   23
      Top             =   1620
      Width           =   3495
      _ExtentX        =   6165
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
      Left            =   5880
      TabIndex        =   25
      Top             =   2040
      Width           =   3495
      _ExtentX        =   6165
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
   Begin VB.Label lblDuracionMinutos 
      AutoSize        =   -1  'True
      Caption         =   "minutos"
      Height          =   210
      Left            =   4080
      TabIndex        =   42
      Top             =   3780
      Width           =   555
   End
   Begin VB.Label lblDuracion 
      AutoSize        =   -1  'True
      Caption         =   "Duración:"
      Height          =   210
      Left            =   2280
      TabIndex        =   16
      Top             =   3780
      Width           =   690
   End
   Begin VB.Label lblKilometro 
      AutoSize        =   -1  'True
      Caption         =   "&Kms.:"
      Height          =   210
      Left            =   120
      TabIndex        =   14
      Top             =   3780
      Width           =   405
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   4740
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label lblImporte 
      AutoSize        =   -1  'True
      Caption         =   "&Importe:"
      Height          =   210
      Left            =   120
      TabIndex        =   12
      Top             =   3240
      Width           =   570
   End
   Begin VB.Label lblPersona 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
      Height          =   210
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   525
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4740
      Y1              =   2700
      Y2              =   2700
   End
   Begin VB.Label lblRutaOtra 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
      Height          =   210
      Left            =   120
      TabIndex        =   6
      Top             =   2340
      Width           =   600
   End
   Begin VB.Line Line1 
      X1              =   4740
      X2              =   4740
      Y1              =   1020
      Y2              =   5760
   End
   Begin VB.Label lblConductor2 
      AutoSize        =   -1  'True
      Caption         =   "Conductor 2:"
      Height          =   210
      Left            =   4860
      TabIndex        =   24
      Top             =   2100
      Width           =   930
   End
   Begin VB.Label lblConductor 
      AutoSize        =   -1  'True
      Caption         =   "&Conductor:"
      Height          =   210
      Left            =   4860
      TabIndex        =   22
      Top             =   1680
      Width           =   795
   End
   Begin VB.Label lblVehiculo 
      AutoSize        =   -1  'True
      Caption         =   "&Vehículo:"
      Height          =   210
      Left            =   4860
      TabIndex        =   19
      Top             =   1080
      Width           =   675
   End
   Begin VB.Label lblRuta 
      AutoSize        =   -1  'True
      Caption         =   "&Ruta:"
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblNotas 
      AutoSize        =   -1  'True
      Caption         =   "Notas:"
      Height          =   210
      Left            =   4860
      TabIndex        =   33
      Top             =   4500
      Width           =   465
   End
   Begin VB.Label lblHora 
      AutoSize        =   -1  'True
      Caption         =   "&Hora:"
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   1500
      Width           =   390
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Ingrese aquí los Datos del Horario"
      Height          =   210
      Left            =   780
      TabIndex        =   39
      Top             =   300
      Width           =   2430
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "HorarioPropiedad.frx":1688
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblDiaSemana 
      AutoSize        =   -1  'True
      Caption         =   "&Día:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   270
   End
End
Attribute VB_Name = "frmHorarioPropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mHorario As Horario
Private mNew As Boolean
Private mKeyDecimal As Boolean
Private mPermite2Conductores As Boolean

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef Horario As Horario)
    Dim Persona As Persona
    Set mHorario = Horario
    mNew = (mHorario.DiaSemana = 0)
    
    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    With mHorario
        If Not CSM_Control_DataCombo.FillFromSQL(datcboRuta, "SELECT RTRIM(IDRuta) AS IDRuta FROM Ruta WHERE Activo = 1 OR IDRuta = '" & ReplaceQuote(.IDRuta) & "'" & IIf(pCPermiso.RutaWhere <> "", " AND " & Replace(pCPermiso.RutaWhere, "%TABLENAME%", "Ruta"), "") & " ORDER BY IDRuta", "IDRuta", "IDRuta", "Rutas") Then
            Unload Me
            Exit Sub
        End If
        datcboRuta.BoundText = .IDRuta
        Select Case datcboRuta.BoundText
            Case pParametro.Ruta_ID_Otra
                txtRutaOtra.Text = .RutaOtra
                
                txtPersona.Tag = .IDPersona
                If .IDPersona > 0 Then
                    Set Persona = New Persona
                    Persona.IDPersona = .IDPersona
                    If Persona.Load() Then
                        txtPersona.Text = Persona.ApellidoNombre
                    End If
                    Set Persona = Nothing
                End If
    
                txtImporte.Text = .Importe_Formatted
                        
            Case pParametro.Ruta_Paquete_ID
                txtRutaOtra.Text = .RutaOtra
                
                txtPersona.Tag = 0
                txtPersona.Text = ""
                
                txtImporte.Text = .Importe_Formatted
            
            Case Else
                txtRutaOtra.Text = ""
                txtPersona.Tag = 0
                txtPersona.Text = ""
                
                txtImporte.Text = Format(0, "Currency")
        End Select
        datcboRuta_Change
        
        txtKilometro.Text = IIf(.Kilometro = 0, "", .Kilometro)
        txtDuracion.Text = IIf(.Duracion = 0, "", .Duracion)
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboVehiculo, "(SELECT -1 AS IDVehiculo, '------------------' AS Nombre, 1 AS Orden FROM Vehiculo) UNION (SELECT IDVehiculo, Nombre, 2 AS Orden FROM Vehiculo WHERE Activo = 1 OR IDVehiculo = " & .IDVehiculo & ") ORDER BY Orden, Nombre", "IDVehiculo", "Nombre", "Vehículos") Then
            Unload Me
            Exit Sub
        End If
        datcboVehiculo.BoundText = .IDVehiculo
        If Val(datcboVehiculo.BoundText) = 0 Then
            datcboVehiculo.BoundText = -1
        End If
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboConductor, "(SELECT -1 AS IDPersona, '------------------' AS ApellidoNombre, 1 AS Orden FROM Persona) UNION (SELECT IDPersona, dbo.udf_GetEntidadApellidoYNombre(Apellido, Nombre) AS ApellidoNombre, 2 AS Orden FROM Persona WHERE (Activo = 1 OR IDPersona = " & .IDConductor & ") AND EntidadTipo = '" & ENTIDAD_TIPO_PERSONA_CONDUCTOR & "') ORDER BY Orden, ApellidoNombre", "IDPersona", "ApellidoNombre", "Conductores") Then
            Unload Me
            Exit Sub
        End If
        datcboConductor.BoundText = .IDConductor
        If Val(datcboConductor.BoundText) = 0 Then
            datcboConductor.BoundText = -1
        End If
        
        If pParametro.Viaje_Permite_2_Conductores Then
            If Not CSM_Control_DataCombo.FillFromSQL(datcboConductor2, "(SELECT -1 AS IDPersona, '------------------' AS ApellidoNombre, 1 AS Orden FROM Persona) UNION (SELECT IDPersona, dbo.udf_GetEntidadApellidoYNombre(Apellido, Nombre) AS ApellidoNombre, 2 AS Orden FROM Persona WHERE (Activo = 1 OR IDPersona = " & .IDConductor2 & ") AND EntidadTipo = '" & ENTIDAD_TIPO_PERSONA_CONDUCTOR & "') ORDER BY Orden, ApellidoNombre", "IDPersona", "ApellidoNombre", "Conductores") Then
                Unload Me
                Exit Sub
            End If
            datcboConductor2.BoundText = .IDConductor2
            If Val(datcboConductor2.BoundText) = 0 Then
                datcboConductor2.BoundText = -1
            End If
        End If
    
        cboDiaSemana.ListIndex = .DiaSemana - 1
        dtpHora.value = .Hora_Formatted
        SetCaption
        
        txtConductorImporteTramoCompleto.Text = .ConductorImporteTramoCompleto_FormattedAsString
        If mPermite2Conductores Then
            txtConductorImporteTramo1.Text = .ConductorImporteTramo1_FormattedAsString
            txtConductorImporteTramo2.Text = .ConductorImporteTramo2_FormattedAsString
        End If
        
        txtNotas.Text = .Notas
        chkActivo.value = IIf(.Activo, vbChecked, vbUnchecked)
        chkPersonal.value = IIf(.Personal, vbChecked, vbUnchecked)
    End With
    
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Private Sub cmdPersona_Click()
    Screen.MousePointer = vbHourglass
    If CSM_Forms.IsLoaded("frmPersona") Then
        frmPersona.FormKeepOpenOnSelect = True
    End If
    Load frmPersona
    'frmPersona.cboFilterTipo.ListIndex = 1
    frmPersona.Show
    If frmPersona.WindowState = vbMinimized Then
        frmPersona.WindowState = vbNormal
    End If
    frmPersona.FormWaitingForSelect = Me.Name
    'frmPersona.SelectTypeFilter = ENTIDAD_TIPO_PERSONA_CLIENTE
    frmPersona.SetFocus
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPersonaClear_Click()
    txtPersona.Tag = 0
    txtPersona.Text = ""
End Sub

Private Sub txtImporte_GotFocus()
    CSM_Control_TextBox.SelAllText txtImporte
End Sub

Private Sub txtImporte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDecimal Then
        mKeyDecimal = True
    End If
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    If ((KeyAscii > 32 And KeyAscii < 48) Or KeyAscii > 57) And KeyAscii <> Asc(pRegionalSettings.CurrencyDecimalSymbol) And KeyAscii <> Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
        KeyAscii = 0
    End If
    If mKeyDecimal Then
        KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol)
        mKeyDecimal = False
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol) Then
        If InStr(1, txtImporte.Text, pRegionalSettings.CurrencyDecimalSymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
        If InStr(1, txtImporte.Text, pRegionalSettings.CurrencyCurrencySymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtImporte_LostFocus()
    If Not IsNumeric(txtImporte.Text) Then
        txtImporte.Text = Val(txtImporte.Text)
    End If
    txtImporte.Text = Format(CCur(txtImporte.Text), "Currency")
End Sub

Private Sub Form_Load()
    Dim Weekday As Byte
    
    For Weekday = 1 To 7
        cboDiaSemana.AddItem WeekdayName(Weekday)
    Next Weekday
    
    lblConductor2.Visible = pParametro.Viaje_Permite_2_Conductores
    datcboConductor2.Visible = pParametro.Viaje_Permite_2_Conductores
    
    lblConductorImporteTramo1.Visible = pParametro.Viaje_Permite_2_Conductores
    txtConductorImporteTramo1.Visible = pParametro.Viaje_Permite_2_Conductores
    lblConductorImporteTramo2.Visible = pParametro.Viaje_Permite_2_Conductores
    txtConductorImporteTramo2.Visible = pParametro.Viaje_Permite_2_Conductores
End Sub

Private Sub cboDiaSemana_Change()
    SetCaption
End Sub

Private Sub dtpHora_Change()
    SetCaption
End Sub

Private Sub datcboRuta_Change()
    SetCaption
    ShowControls
End Sub

Private Sub cmdRuta_Click()
    If pCPermiso.GotPermission(PERMISO_RUTA) Then
        Screen.MousePointer = vbHourglass
        frmRuta.Show
        On Error Resume Next
        Set frmRuta.lvwData.SelectedItem = frmRuta.lvwData.ListItems(KEY_STRINGER & datcboRuta.BoundText)
        frmRuta.lvwData.SelectedItem.EnsureVisible
        If frmRuta.WindowState = vbMinimized Then
            frmRuta.WindowState = vbNormal
        End If
        frmRuta.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub txtKilometro_GotFocus()
    CSM_Control_TextBox.SelAllText txtKilometro
End Sub

Private Sub txtKilometro_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtKilometro_LostFocus()
    txtKilometro.Text = Val(txtKilometro.Text)
    If txtKilometro.Text = 0 Then
        txtKilometro.Text = ""
    End If
End Sub

Private Sub txtDuracion_GotFocus()
    CSM_Control_TextBox.SelAllText txtDuracion
End Sub

Private Sub txtDuracion_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtDuracion_LostFocus()
    txtDuracion.Text = Val(txtDuracion.Text)
    If txtDuracion.Text = 0 Then
        txtDuracion.Text = ""
    End If
End Sub

Private Sub cmdVehiculo_Click()
    If pCPermiso.GotPermission(PERMISO_VEHICULO) Then
        Screen.MousePointer = vbHourglass
        frmVehiculo.Show
        On Error Resume Next
        Set frmVehiculo.lvwData.SelectedItem = frmVehiculo.lvwData.ListItems(KEY_STRINGER & Val(datcboVehiculo.BoundText))
        frmVehiculo.lvwData.SelectedItem.EnsureVisible
        If frmPersona.WindowState = vbMinimized Then
            frmPersona.WindowState = vbNormal
        End If
        frmVehiculo.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub txtConductorImporteTramoCompleto_GotFocus()
    CSM_Control_TextBox.SelAllText txtConductorImporteTramoCompleto
End Sub

Private Sub txtConductorImporteTramoCompleto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDecimal Then
        mKeyDecimal = True
    End If
End Sub

Private Sub txtConductorImporteTramoCompleto_KeyPress(KeyAscii As Integer)
    Call CSM_Control_TextBox.CheckKeyPress(txtConductorImporteTramoCompleto, KeyAscii, mKeyDecimal)
End Sub

Private Sub txtConductorImporteTramoCompleto_LostFocus()
    Call CSM_Control_TextBox.FormatValue_ByTag(txtConductorImporteTramoCompleto)
End Sub

Private Sub txtConductorImporteTramo1_GotFocus()
    CSM_Control_TextBox.SelAllText txtConductorImporteTramo1
End Sub

Private Sub txtConductorImporteTramo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDecimal Then
        mKeyDecimal = True
    End If
End Sub

Private Sub txtConductorImporteTramo1_KeyPress(KeyAscii As Integer)
    Call CSM_Control_TextBox.CheckKeyPress(txtConductorImporteTramo1, KeyAscii, mKeyDecimal)
End Sub

Private Sub txtConductorImporteTramo1_LostFocus()
    Call CSM_Control_TextBox.FormatValue_ByTag(txtConductorImporteTramo1)
End Sub

Private Sub txtConductorImporteTramo2_GotFocus()
    CSM_Control_TextBox.SelAllText txtConductorImporteTramo2
End Sub

Private Sub txtConductorImporteTramo2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDecimal Then
        mKeyDecimal = True
    End If
End Sub

Private Sub txtConductorImporteTramo2_KeyPress(KeyAscii As Integer)
    Call CSM_Control_TextBox.CheckKeyPress(txtConductorImporteTramo2, KeyAscii, mKeyDecimal)
End Sub

Private Sub txtConductorImporteTramo2_LostFocus()
    Call CSM_Control_TextBox.FormatValue_ByTag(txtConductorImporteTramo2)
End Sub

Private Sub txtNotas_GotFocus()
    CSM_Control_TextBox.SelAllText txtNotas
    cmdOK.Default = False
End Sub
    
Private Sub txtNotas_LostFocus()
    cmdOK.Default = True
End Sub

Private Sub cmdOK_Click()
    If cboDiaSemana.ListIndex = -1 Then
        MsgBox "Debe seleccionar el Día.", vbInformation, App.Title
        cboDiaSemana.SetFocus
        Exit Sub
    End If
    If datcboRuta.BoundText = "" Then
        MsgBox "Debe seleccionar la Ruta.", vbInformation, App.Title
        datcboRuta.SetFocus
        Exit Sub
    End If
    If (datcboRuta.BoundText = pParametro.Ruta_ID_Otra Or datcboRuta.BoundText = pParametro.Ruta_Paquete_ID) And Trim(txtRutaOtra.Text) = "" Then
        MsgBox "Debe ingresar el nombre de la Ruta '" & datcboRuta.Text & "'.", vbInformation, App.Title
        txtRutaOtra.SetFocus
        Exit Sub
    End If
    
        If Not IsNumeric(txtImporte.Text) Then
        MsgBox "El Importe ingresado es incorrecto.", vbInformation, App.Title
        txtImporte.SetFocus
        Exit Sub
    End If
    If CCur(txtImporte.Text) < 0 Then
        MsgBox "El Importe debe ser mayor o igual a cero.", vbInformation, App.Title
        txtImporte.SetFocus
        Exit Sub
    End If

    If mPermite2Conductores Then
        If Val(datcboConductor.BoundText) = -1 And Val(datcboConductor2.BoundText) > -1 Then
            MsgBox "Si selecciona el Conductor N° 2, debe seleccionar el Conductor N° 1.", vbInformation, App.Title
            datcboConductor.SetFocus
            Exit Sub
        End If
        If Val(datcboConductor.BoundText) > -1 And Val(datcboConductor.BoundText) = Val(datcboConductor2.BoundText) Then
            MsgBox "El Conductor N° 1 y el Conductor N° 2 no pueden ser el mismo.", vbInformation, App.Title
            datcboConductor2.SetFocus
            Exit Sub
        End If
    End If
    
    Call txtConductorImporteTramoCompleto_LostFocus
    If mPermite2Conductores Then
        Call txtConductorImporteTramo1_LostFocus
        Call txtConductorImporteTramo2_LostFocus
    End If

    If datcboRuta.BoundText = pParametro.Ruta_ID_Otra And Val(txtPersona.Tag) = 0 Then
        MsgBox "Debe seleccionar el Cliente.", vbInformation, App.Title
        txtPersona.SetFocus
        Exit Sub
    End If

    With mHorario
        .DiaSemana = cboDiaSemana.ListIndex + 1
        .Hora = dtpHora.value
        .IDRuta = datcboRuta.BoundText
        
        Select Case datcboRuta.BoundText
            Case pParametro.Ruta_ID_Otra
                .RutaOtra = txtRutaOtra.Text
                .IDPersona = Val(txtPersona.Tag)
                .Importe = CCur(txtImporte.Text)
                
            Case pParametro.Ruta_Paquete_ID
                .RutaOtra = txtRutaOtra.Text
                .IDPersona = 0
                .Importe = CCur(txtImporte.Text)
                
            Case Else
                .RutaOtra = ""
                .IDPersona = 0
                .Importe = 0
        End Select
        .Kilometro = Val(txtKilometro.Text)
        .Duracion = Val(txtDuracion.Text)
        
        .IDVehiculo = IIf(Val(datcboVehiculo.BoundText) = -1, 0, Val(datcboVehiculo.BoundText))
        .IDConductor = IIf(Val(datcboConductor.BoundText) = -1, 0, Val(datcboConductor.BoundText))
        .ConductorImporteTramoCompleto_FormattedAsString = txtConductorImporteTramoCompleto.Text
        If mPermite2Conductores Then
            .IDConductor2 = IIf(Val(datcboConductor2.BoundText) = -1, 0, Val(datcboConductor2.BoundText))
            .ConductorImporteTramo1_FormattedAsString = txtConductorImporteTramo1.Text
            .ConductorImporteTramo2_FormattedAsString = txtConductorImporteTramo2.Text
        Else
            .IDConductor2 = 0
            .ConductorImporteTramo1 = -1
            .ConductorImporteTramo2 = -1
        End If
        .Notas = txtNotas.Text
        .Activo = (chkActivo.value = vbChecked)
        .Personal = (chkPersonal.value = vbChecked)
        If Not .Update Then
            Exit Sub
        End If
    End With
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdAuditoria_Click()
    frmAuditoriaGenerico.LoadDataAndShow mHorario
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mHorario = Nothing
    Set frmHorarioPropiedad = Nothing
End Sub

Public Sub FillComboBoxRuta()
    Dim KeySave As String
    Dim recData As ADODB.Recordset
    
    KeySave = datcboRuta.BoundText
    Set recData = datcboRuta.RowSource
    recData.Requery
    Set recData = Nothing
    datcboRuta.BoundText = KeySave
End Sub

Public Sub FillComboBoxVehiculo()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboVehiculo.BoundText)
    Set recData = datcboVehiculo.RowSource
    recData.Requery
    Set recData = Nothing
    datcboVehiculo.BoundText = KeySave
    If Val(datcboVehiculo.BoundText) = 0 Then
        datcboVehiculo.BoundText = -1
    End If
End Sub

Public Sub FillComboBoxPersona()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboConductor.BoundText)
    Set recData = datcboConductor.RowSource
    recData.Requery
    Set recData = Nothing
    datcboConductor.BoundText = KeySave
    If Val(datcboConductor.BoundText) = 0 Then
        datcboConductor.BoundText = -1
    End If

    If pParametro.Viaje_Permite_2_Conductores Then
        KeySave = Val(datcboConductor2.BoundText)
        Set recData = datcboConductor2.RowSource
        recData.Requery
        Set recData = Nothing
        datcboConductor2.BoundText = KeySave
        If Val(datcboConductor2.BoundText) = 0 Then
            datcboConductor2.BoundText = -1
        End If
    End If
End Sub

Public Sub PersonaSelected(ByVal IDPersona As Long, ByVal Tag As String)
    If IDPersona = 0 Then
        Exit Sub
    End If
    
    txtPersona.Tag = IDPersona
    txtPersona.Text = frmMDI.cboPersona.Text
    
    On Error Resume Next
    txtKilometro.SetFocus
End Sub

Private Sub SetCaption()
    Dim CaptionTemp As String
    
    If cboDiaSemana.ListIndex < -1 Then
        CaptionTemp = cboDiaSemana.Text
    End If
    If dtpHora.value <> "" Then
        CaptionTemp = CaptionTemp & IIf(CaptionTemp = "", "", " - ") & dtpHora.value
    End If
    If datcboRuta.BoundText <> "" Then
        CaptionTemp = CaptionTemp & IIf(CaptionTemp = "", "", " - ") & datcboRuta.Text
    End If
    Caption = "Propiedades" & IIf(CaptionTemp = "", "", " ") & CaptionTemp
End Sub

Private Sub ShowControls()
    Dim Ruta As Ruta
    
    If datcboRuta.BoundText <> "" Then
        Set Ruta = New Ruta
        Ruta.IDRuta = datcboRuta.BoundText
        If Ruta.Load() Then
            mPermite2Conductores = (pParametro.Viaje_Permite_2_Conductores And Ruta.Permite2Conductores)
            lblConductorImporteTramo1.Visible = mPermite2Conductores
            txtConductorImporteTramo1.Visible = mPermite2Conductores
            lblConductorImporteTramo2.Visible = mPermite2Conductores
            txtConductorImporteTramo2.Visible = mPermite2Conductores
        End If
        Set Ruta = Nothing
    End If
End Sub
