VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmVehiculoMantenimientoAccionPropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   Icon            =   "VehiculoMantenimientoAccionPropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   4800
   Begin VB.CommandButton cmdAuditoria 
      Caption         =   "Auditoría"
      Height          =   615
      Left            =   3660
      Picture         =   "VehiculoMantenimientoAccionPropiedad.frx":054A
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdVehiculoKilometraje 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2100
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "Kilometraje Actual"
      Top             =   3060
      Width           =   255
   End
   Begin VB.TextBox txtLitros 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1020
      MaxLength       =   7
      TabIndex        =   13
      Top             =   3480
      Width           =   1035
   End
   Begin VB.CommandButton cmdConductor 
      Caption         =   "..."
      Height          =   315
      Left            =   4380
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "Personas"
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton cmdSiguiente 
      Height          =   315
      Left            =   4020
      Picture         =   "VehiculoMantenimientoAccionPropiedad.frx":0B74
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Siguiente"
      Top             =   2220
      Width           =   300
   End
   Begin VB.CommandButton cmdAnterior 
      Height          =   315
      Left            =   1020
      Picture         =   "VehiculoMantenimientoAccionPropiedad.frx":10FE
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Anterior"
      Top             =   2220
      Width           =   300
   End
   Begin VB.CommandButton cmdHoy 
      Height          =   315
      Left            =   4320
      Picture         =   "VehiculoMantenimientoAccionPropiedad.frx":1688
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Hoy"
      Top             =   2220
      Width           =   315
   End
   Begin VB.TextBox txtDiaSemana 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2220
      Width           =   1050
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1020
      MaxLength       =   20
      TabIndex        =   15
      Top             =   3960
      Width           =   1515
   End
   Begin VB.TextBox txtKilometraje 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1020
      MaxLength       =   7
      TabIndex        =   11
      Top             =   3060
      Width           =   1035
   End
   Begin VB.CommandButton cmdVehiculoMantenimientoGrupo 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4380
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Grupos"
      Top             =   1380
      Width           =   255
   End
   Begin VB.CommandButton cmdVehiculo 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4380
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Vehículos"
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox txtNotas 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   1020
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   4440
      Width           =   3615
   End
   Begin VB.Frame fraLine 
      Height          =   75
      Left            =   120
      TabIndex        =   28
      Top             =   780
      Width           =   4515
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3420
      TabIndex        =   19
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2100
      TabIndex        =   18
      Top             =   5640
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo datcboVehiculo 
      Height          =   330
      Left            =   1020
      TabIndex        =   1
      Top             =   960
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
   Begin MSDataListLib.DataCombo datcboVehiculoMantenimientoGrupo 
      Height          =   330
      Left            =   1020
      TabIndex        =   3
      Top             =   1380
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
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   315
      Left            =   2400
      TabIndex        =   7
      Top             =   2220
      Width           =   1635
      _ExtentX        =   2884
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
      Format          =   119668737
      CurrentDate     =   36950
   End
   Begin MSComCtl2.DTPicker dtpHora 
      Height          =   315
      Left            =   1020
      TabIndex        =   9
      Top             =   2640
      Width           =   1395
      _ExtentX        =   2461
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
      Format          =   119668738
      CurrentDate     =   36494
   End
   Begin MSDataListLib.DataCombo datcboConductor 
      Height          =   330
      Left            =   1020
      TabIndex        =   5
      Top             =   1800
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
   Begin VB.Label lblLitros 
      AutoSize        =   -1  'True
      Caption         =   "Litros:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   12
      Top             =   3540
      Width           =   450
   End
   Begin VB.Label lblConductor 
      AutoSize        =   -1  'True
      Caption         =   "&Conductor:"
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   1860
      Width           =   795
   End
   Begin VB.Label lblHora 
      AutoSize        =   -1  'True
      Caption         =   "&Hora:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   8
      Top             =   2700
      Width           =   390
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      Caption         =   "&Fecha:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lblImporte 
      AutoSize        =   -1  'True
      Caption         =   "Importe:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   14
      Top             =   4020
      Width           =   570
   End
   Begin VB.Label lblKilometraje 
      AutoSize        =   -1  'True
      Caption         =   "Kilometraje:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   825
   End
   Begin VB.Label lblVehiculoMantenimientoGrupo 
      AutoSize        =   -1  'True
      Caption         =   "Grupo:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label lblVehiculo 
      AutoSize        =   -1  'True
      Caption         =   "&Vehículo:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   1020
      Width           =   675
   End
   Begin VB.Label lblNotas 
      AutoSize        =   -1  'True
      Caption         =   "Notas:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   16
      Top             =   4500
      Width           =   465
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Datos del Mantenimiento del Vehículo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   780
      TabIndex        =   27
      Top             =   300
      Width           =   2670
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "VehiculoMantenimientoAccionPropiedad.frx":17D2
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmVehiculoMantenimientoAccionPropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mVehiculoMantenimientoAccion As VehiculoMantAccion
Private mNew As Boolean

Private mKeyDecimal As Boolean


Private Sub cmdAuditoria_Click()
    frmAuditoriaGenerico.LoadDataAndShow mVehiculoMantenimientoAccion
End Sub

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef VehiculoMantenimientoAccion As VehiculoMantAccion)
    Set mVehiculoMantenimientoAccion = VehiculoMantenimientoAccion
    mNew = (mVehiculoMantenimientoAccion.IDVehiculoMantenimientoAccion = 0)
    
    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    With mVehiculoMantenimientoAccion
        If Not CSM_Control_DataCombo.FillFromSQL(datcboVehiculo, "SELECT IDVehiculo, Nombre FROM Vehiculo WHERE Activo = 1 OR IDVehiculo = " & .IDVehiculo & " ORDER BY Nombre", "IDVehiculo", "Nombre", "Vehículos", cscpItemOrNone, .IDVehiculo) Then
            Unload Me
            Exit Sub
        End If
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboVehiculoMantenimientoGrupo, "SELECT IDVehiculoMantenimientoGrupo, Nombre FROM VehiculoMantenimientoGrupo WHERE Activo = 1 OR IDVehiculoMantenimientoGrupo = " & .IDVehiculoMantenimientoGrupo & " ORDER BY Nombre", "IDVehiculoMantenimientoGrupo", "Nombre", "Grupos de Mantenimiento de Vehículos", cscpItemOrNone, .IDVehiculoMantenimientoGrupo) Then
            Unload Me
            Exit Sub
        End If
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboConductor, "(SELECT 0 AS IDPersona, '------------------' AS ApellidoNombre, 1 AS Orden FROM Persona) UNION (SELECT IDPersona, dbo.udf_GetEntidadApellidoYNombre(Apellido, Nombre) AS ApellidoNombre, 2 AS Orden FROM Persona WHERE Activo = 1 AND EntidadTipo = '" & ENTIDAD_TIPO_PERSONA_CONDUCTOR & "') ORDER BY Orden, ApellidoNombre", "IDPersona", "ApellidoNombre", "Conductores", cscpItemOrFirst, .IDConductor) Then
            Unload Me
            Exit Sub
        End If
        
        If mNew Then
            dtpFecha.value = Date
            dtpHora.value = Time
        Else
            dtpFecha.value = Format(.FechaHora, "Short Date")
            dtpHora.value = Format(.FechaHora, "Short Time")
        End If
        dtpFecha_Change
        txtKilometraje.Text = IIf(.Kilometraje = 0, "", .Kilometraje)
        txtLitros.Text = IIf(.Litros = 0, "", .Litros)
        txtLitros_LostFocus
        txtImporte.Text = IIf(.Importe = 0, "", .Importe)
        txtImporte_LostFocus
        txtNotas.Text = .Notas
    End With
    
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

Public Sub FillComboBoxVehiculoMantenimientoGrupo()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboVehiculoMantenimientoGrupo.BoundText)
    Set recData = datcboVehiculoMantenimientoGrupo.RowSource
    recData.Requery
    Set recData = Nothing
    datcboVehiculoMantenimientoGrupo.BoundText = KeySave
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdConductor_Click()
    If pCPermiso.GotPermission(PERMISO_PERSONA) Then
        Call frmPersona.FindAndShowItem(Val(datcboConductor.BoundText), UCase(Left(datcboConductor.Text, 1)), Me.Name, "", "")
    End If
End Sub

Private Sub cmdOK_Click()
    Dim VehiculoMantenimiento As VehiculoMantenimiento
    
    If Val(datcboVehiculo.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Vehículo.", vbInformation, App.Title
        datcboVehiculo.SetFocus
        Exit Sub
    End If
    If Val(datcboVehiculoMantenimientoGrupo.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Grupo.", vbInformation, App.Title
        datcboVehiculoMantenimientoGrupo.SetFocus
        Exit Sub
    End If
    If dtpFecha.value > Date Then
        If MsgBox("Está ingresando una acción con Fecha posterior a hoy." & vbCr & vbCr & "¿Desea continuar?", vbExclamation + vbYesNo, App.Title) = vbNo Then
            dtpFecha.SetFocus
            Exit Sub
        End If
    End If
    
    Set VehiculoMantenimiento = New VehiculoMantenimiento
    VehiculoMantenimiento.IDVehiculo = Val(datcboVehiculo.BoundText)
    VehiculoMantenimiento.IDVehiculoMantenimientoGrupo = Val(datcboVehiculoMantenimientoGrupo.BoundText)
    VehiculoMantenimiento.NoMatchRaiseError = False
    If Not VehiculoMantenimiento.Load() Then
        Set VehiculoMantenimiento = Nothing
        Exit Sub
    End If
    If VehiculoMantenimiento.NoMatch Then
        MsgBox "No existe el Mantenimiento para el Vehículo.", vbInformation, App.Title
        datcboVehiculoMantenimientoGrupo.SetFocus
        Set VehiculoMantenimiento = Nothing
        Exit Sub
    End If
    If VehiculoMantenimiento.Tipo = VEHICULO_MATENIMIENTO_TIPO_KILOMETRAJE Then
        If txtKilometraje.Text = "" Then
            MsgBox "Debe ingresar el Kilometraje porque este mantenimiento es de tipo Kilómetros.", vbInformation, App.Title
            txtKilometraje.SetFocus
            Set VehiculoMantenimiento = Nothing
            Exit Sub
        End If
    End If
    Set VehiculoMantenimiento = Nothing
    
    If txtKilometraje.Text <> "" Then
        If Not IsNumeric(txtKilometraje.Text) Then
            MsgBox "El Kiometraje debe ser un valor numérico.", vbInformation, App.Title
            txtKilometraje.SetFocus
            Exit Sub
        End If
        If CLng(txtKilometraje.Text) <= 0 Then
            MsgBox "El Kilometraje debe ser mayor a cero.", vbInformation, App.Title
            txtKilometraje.SetFocus
            Exit Sub
        End If
    End If
    If txtLitros.Text <> "" Then
        If Not IsNumeric(txtLitros.Text) Then
            MsgBox "Los Litros deben ser un valor numérico.", vbInformation, App.Title
            txtLitros.SetFocus
            Exit Sub
        End If
        If CDbl(txtLitros.Text) <= 0 Then
            MsgBox "Los Litros deben ser mayores a cero.", vbInformation, App.Title
            txtLitros.SetFocus
            Exit Sub
        End If
    End If
    If txtImporte.Text <> "" Then
        If Not IsNumeric(txtImporte.Text) Then
            MsgBox "El Importe debe ser un valor numérico.", vbInformation, App.Title
            txtImporte.SetFocus
            Exit Sub
        End If
        If CCur(txtImporte.Text) < 0 Then
            MsgBox "El Importe debe ser mayor o igual a cero.", vbInformation, App.Title
            txtImporte.SetFocus
            Exit Sub
        End If
    End If
    
    With mVehiculoMantenimientoAccion
        .IDVehiculo = Val(datcboVehiculo.BoundText)
        .IDVehiculoMantenimientoGrupo = Val(datcboVehiculoMantenimientoGrupo.BoundText)
        .IDConductor = Val(datcboConductor.BoundText)
        .FechaHora = CDate(Format(dtpFecha.value, "Short Date") & " " & Format(dtpHora.value, "Short Time"))
        If txtKilometraje.Text = "" Then
            .Kilometraje = 0
        Else
            .Kilometraje = CLng(txtKilometraje.Text)
        End If
        If txtLitros.Text = "" Then
            .Litros = 0
        Else
            .Litros = CDbl(txtLitros.Text)
        End If
        If txtImporte.Text = "" Then
            .Importe = 0
        Else
            .Importe = CCur(txtImporte.Text)
        End If
        .Notas = txtNotas.Text
        If mNew Then
            If Not .AddNew() Then
                Exit Sub
            End If
        Else
            If Not .Update Then
                Exit Sub
            End If
        End If
    End With
    
    Unload Me
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

Private Sub cmdVehiculoKilometraje_Click()
    Dim Vehiculo As Vehiculo
    
    If Val(datcboVehiculo.BoundText) > 0 Then
        Set Vehiculo = New Vehiculo
        Vehiculo.IDVehiculo = Val(datcboVehiculo.BoundText)
        If Vehiculo.Load() Then
            txtKilometraje.Text = Vehiculo.KilometrajeEstimado
        End If
        Set Vehiculo = Nothing
    End If
End Sub

Private Sub cmdVehiculoMantenimientoGrupo_Click()
    If pCPermiso.GotPermission(PERMISO_VEHICULO_MANTENIMIENTO_GRUPO) Then
        Screen.MousePointer = vbHourglass
        frmVehiculoMantenimientoGrupo.Show
        On Error Resume Next
        Set frmVehiculoMantenimientoGrupo.lvwData.SelectedItem = frmVehiculoMantenimientoGrupo.lvwData.ListItems(KEY_STRINGER & datcboVehiculoMantenimientoGrupo.BoundText)
        frmVehiculoMantenimientoGrupo.lvwData.SelectedItem.EnsureVisible
        If frmVehiculoMantenimientoGrupo.WindowState = vbMinimized Then
            frmVehiculoMantenimientoGrupo.WindowState = vbNormal
        End If
        frmVehiculoMantenimientoGrupo.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub dtpFecha_Change()
    txtDiaSemana.Text = WeekdayName(Weekday(dtpFecha.value))
End Sub

Private Sub cmdAnterior_Click()
    dtpFecha.value = DateAdd("d", -1, dtpFecha.value)
    dtpFecha.SetFocus
    dtpFecha_Change
End Sub

Private Sub cmdSiguiente_Click()
    dtpFecha.value = DateAdd("d", 1, dtpFecha.value)
    dtpFecha.SetFocus
    dtpFecha_Change
End Sub

Private Sub cmdHoy_Click()
    Dim OldValue As Date
    
    OldValue = dtpFecha.value
    dtpFecha.value = Date
    dtpFecha.SetFocus
    If OldValue <> dtpFecha.value Then
        dtpFecha_Change
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mVehiculoMantenimientoAccion = Nothing
    Set frmVehiculoMantenimientoAccionPropiedad = Nothing
End Sub

Private Sub txtNotas_GotFocus()
    CSM_Control_TextBox.SelAllText txtNotas
    cmdOK.Default = False
End Sub

Private Sub txtNotas_LostFocus()
    cmdOK.Default = True
End Sub

Private Sub txtKilometraje_GotFocus()
    CSM_Control_TextBox.SelAllText txtKilometraje
End Sub

Private Sub txtKilometraje_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtKilometraje_LostFocus()
    txtKilometraje.Text = Val(txtKilometraje.Text)
    If CLng(txtKilometraje.Text) = 0 Then
        txtKilometraje.Text = ""
    End If
End Sub

Private Sub txtLitros_GotFocus()
    CSM_Control_TextBox.SelAllText txtLitros
End Sub

Private Sub txtLitros_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDecimal Then
        mKeyDecimal = True
    End If
End Sub

Private Sub txtLitros_KeyPress(KeyAscii As Integer)
    If ((KeyAscii > 32 And KeyAscii < 48) Or KeyAscii > 57) And KeyAscii <> Asc(pRegionalSettings.NumberDecimalSymbol) And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
    If mKeyDecimal Then
        KeyAscii = Asc(pRegionalSettings.NumberDecimalSymbol)
        mKeyDecimal = False
    End If
    If KeyAscii = Asc(pRegionalSettings.NumberDecimalSymbol) Then
        If InStr(1, txtLitros.Text, pRegionalSettings.NumberDecimalSymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtLitros_LostFocus()
    If Not IsNumeric(txtLitros.Text) Then
        txtLitros.Text = Val(txtLitros.Text)
    End If
    If CDbl(txtLitros.Text) = 0 Then
        txtLitros.Text = ""
    Else
        txtLitros.Text = CDbl(txtLitros.Text)
    End If
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
    If ((KeyAscii > 32 And KeyAscii < 48) Or KeyAscii > 57) And KeyAscii <> Asc(pRegionalSettings.CurrencyDecimalSymbol) And KeyAscii <> Asc(pRegionalSettings.CurrencyCurrencySymbol) And KeyAscii <> 45 Then
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
    If CCur(txtImporte.Text) = 0 Then
        txtImporte.Text = ""
    Else
        txtImporte.Text = Format(CCur(txtImporte.Text), "Currency")
    End If
End Sub

Public Sub FillComboBoxConductor()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboConductor.BoundText)
    Set recData = datcboConductor.RowSource
    recData.Requery
    Set recData = Nothing
    datcboConductor.BoundText = KeySave
End Sub
