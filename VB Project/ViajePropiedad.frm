VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmViajePropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ViajePropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5685
   ScaleWidth      =   9510
   Begin VB.ComboBox cboCuotas 
      Height          =   330
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   3900
      Width           =   630
   End
   Begin VB.TextBox txtOperacion 
      Height          =   315
      Left            =   2520
      MaxLength       =   20
      TabIndex        =   20
      Top             =   3900
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.CheckBox chkAcreditaSueldo2 
      Height          =   210
      Left            =   9000
      TabIndex        =   37
      Top             =   2160
      Width           =   195
   End
   Begin VB.CheckBox chkPersonal 
      Height          =   210
      Left            =   5400
      TabIndex        =   56
      Top             =   3810
      Width           =   195
   End
   Begin VB.CheckBox chkAcreditaSueldo 
      Height          =   210
      Left            =   9000
      TabIndex        =   34
      Top             =   1740
      Width           =   195
   End
   Begin VB.TextBox txtImporteContado 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1020
      MaxLength       =   20
      TabIndex        =   15
      Top             =   3540
      Width           =   1455
   End
   Begin VB.CommandButton cmdCuentaCorrienteCaja 
      Caption         =   "..."
      Height          =   315
      Left            =   4380
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Cajas"
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdPersonaClear 
      Height          =   315
      Left            =   4320
      Picture         =   "ViajePropiedad.frx":054A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Borrar"
      Top             =   2820
      Width           =   315
   End
   Begin VB.CommandButton cmdPersona 
      Height          =   315
      Left            =   3960
      Picture         =   "ViajePropiedad.frx":0AD4
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Buscar..."
      Top             =   2820
      Width           =   315
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
   Begin VB.CommandButton cmdAuditoria 
      Caption         =   "Auditoría"
      Height          =   615
      Left            =   8400
      Picture         =   "ViajePropiedad.frx":105E
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   60
      Width           =   975
   End
   Begin VB.CheckBox chkCharter 
      Alignment       =   1  'Right Justify
      Caption         =   "Charter:"
      Height          =   210
      Left            =   120
      TabIndex        =   28
      Top             =   5340
      Width           =   1095
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
   Begin VB.TextBox txtKilometro 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1020
      MaxLength       =   4
      TabIndex        =   25
      Top             =   4860
      Width           =   795
   End
   Begin VB.TextBox txtDuracion 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3180
      MaxLength       =   5
      TabIndex        =   27
      Top             =   4860
      Width           =   795
   End
   Begin VB.TextBox txtRutaOtra 
      Height          =   315
      Left            =   1020
      MaxLength       =   50
      TabIndex        =   7
      Top             =   2280
      Width           =   3615
   End
   Begin VB.TextBox txtEstado 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   2640
      Width           =   2715
   End
   Begin VB.ComboBox cboDiaSemana 
      Height          =   330
      Left            =   5760
      Style           =   2  'Dropdown List
      TabIndex        =   40
      Top             =   3180
      Width           =   3630
   End
   Begin VB.TextBox txtDiaSemana 
      Height          =   315
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   1020
      Width           =   1050
   End
   Begin VB.CommandButton cmdHoy 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4320
      Picture         =   "ViajePropiedad.frx":1688
      Style           =   1  'Graphical
      TabIndex        =   48
      TabStop         =   0   'False
      ToolTipText     =   "Hoy"
      Top             =   1020
      Width           =   315
   End
   Begin VB.CommandButton cmdAnterior 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1020
      Picture         =   "ViajePropiedad.frx":17D2
      Style           =   1  'Graphical
      TabIndex        =   45
      TabStop         =   0   'False
      ToolTipText     =   "Anterior"
      Top             =   1020
      Width           =   300
   End
   Begin VB.CommandButton cmdCopyFrom 
      Caption         =   "Copiar desde..."
      Height          =   375
      Left            =   3240
      TabIndex        =   50
      Top             =   240
      Width           =   1395
   End
   Begin VB.CommandButton cmdVehiculo 
      Caption         =   "..."
      Height          =   315
      Left            =   9120
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Vehículos"
      Top             =   1020
      Width           =   255
   End
   Begin VB.CommandButton cmdRuta 
      Caption         =   "..."
      Height          =   315
      Left            =   4380
      TabIndex        =   49
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
      UpDown          =   -1  'True
      CurrentDate     =   36494
   End
   Begin VB.TextBox txtNotas 
      Height          =   1245
      Left            =   5760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   42
      Top             =   3720
      Width           =   3615
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
      TabIndex        =   52
      Top             =   780
      Width           =   9255
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   8160
      TabIndex        =   44
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   6840
      TabIndex        =   43
      Top             =   5160
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Top             =   1020
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
   Begin VB.CommandButton cmdSiguiente 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4020
      Picture         =   "ViajePropiedad.frx":1D5C
      Style           =   1  'Graphical
      TabIndex        =   47
      TabStop         =   0   'False
      ToolTipText     =   "Siguiente"
      Top             =   1020
      Width           =   300
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
      Left            =   5760
      TabIndex        =   30
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
      Left            =   5760
      TabIndex        =   33
      Top             =   1680
      Width           =   3135
      _ExtentX        =   5530
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
   Begin MSDataListLib.DataCombo datcboCuentaCorrienteCaja 
      Height          =   330
      Left            =   1020
      TabIndex        =   22
      Top             =   4320
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
      Left            =   5760
      TabIndex        =   36
      Top             =   2100
      Width           =   3135
      _ExtentX        =   5530
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
   Begin MSDataListLib.DataCombo datcboMedioPago 
      Height          =   330
      Left            =   2520
      TabIndex        =   16
      Top             =   3540
      Width           =   2115
      _ExtentX        =   3731
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
   Begin VB.Label lblCuotas 
      AutoSize        =   -1  'True
      Caption         =   "Cuotas:"
      Height          =   210
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Width           =   555
   End
   Begin VB.Label lblOperacion 
      AutoSize        =   -1  'True
      Caption         =   "Operación:"
      Height          =   210
      Left            =   1680
      TabIndex        =   19
      Top             =   3960
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblConductor2 
      AutoSize        =   -1  'True
      Caption         =   "Conduct. 2:"
      Height          =   210
      Left            =   4860
      TabIndex        =   35
      Top             =   2160
      Width           =   825
   End
   Begin VB.Label lblConductorSueldo 
      AutoSize        =   -1  'True
      Caption         =   "Sueldo:"
      Height          =   210
      Left            =   8820
      TabIndex        =   57
      Top             =   1440
      Width           =   540
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   4740
      Y1              =   4740
      Y2              =   4740
   End
   Begin VB.Label lblImporteContado 
      AutoSize        =   -1  'True
      Caption         =   "P. Contado:"
      Height          =   210
      Left            =   120
      TabIndex        =   14
      Top             =   3600
      Width           =   825
   End
   Begin VB.Label lblCuentaCorrienteCaja 
      AutoSize        =   -1  'True
      Caption         =   "Caja:"
      Height          =   210
      Left            =   120
      TabIndex        =   21
      Top             =   4380
      Width           =   360
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
   Begin VB.Label lblImporte 
      AutoSize        =   -1  'True
      Caption         =   "&Importe:"
      Height          =   210
      Left            =   120
      TabIndex        =   12
      Top             =   3240
      Width           =   570
   End
   Begin VB.Line Line1 
      X1              =   4740
      X2              =   4740
      Y1              =   1020
      Y2              =   5580
   End
   Begin VB.Label lblKilometro 
      AutoSize        =   -1  'True
      Caption         =   "&Kms.:"
      Height          =   210
      Left            =   120
      TabIndex        =   24
      Top             =   4920
      Width           =   405
   End
   Begin VB.Label lblDuracion 
      AutoSize        =   -1  'True
      Caption         =   "Duración:"
      Height          =   210
      Left            =   2280
      TabIndex        =   26
      Top             =   4920
      Width           =   690
   End
   Begin VB.Label lblDuracionMinutos 
      AutoSize        =   -1  'True
      Caption         =   "minutos"
      Height          =   210
      Left            =   4080
      TabIndex        =   54
      Top             =   4920
      Width           =   555
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
   Begin VB.Image imgEstadoActivo 
      Height          =   480
      Left            =   5760
      Picture         =   "ViajePropiedad.frx":22E6
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEstadoFinalizado 
      Height          =   480
      Left            =   5760
      Picture         =   "ViajePropiedad.frx":2BC8
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEstadoCancelado 
      Height          =   480
      Left            =   5760
      Picture         =   "ViajePropiedad.frx":3492
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEstadoEnProgreso 
      Height          =   480
      Left            =   5760
      Picture         =   "ViajePropiedad.frx":3D5C
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblEstado 
      AutoSize        =   -1  'True
      Caption         =   "&Estado:"
      Height          =   210
      Left            =   4860
      TabIndex        =   38
      Top             =   2700
      Width           =   540
   End
   Begin VB.Label lblDiaSemana 
      AutoSize        =   -1  'True
      Caption         =   "Basado en:"
      Height          =   210
      Left            =   4860
      TabIndex        =   39
      Top             =   3240
      Width           =   825
   End
   Begin VB.Label lblConductor 
      AutoSize        =   -1  'True
      Caption         =   "&Conductor:"
      Height          =   210
      Left            =   4860
      TabIndex        =   32
      Top             =   1740
      Width           =   795
   End
   Begin VB.Label lblVehiculo 
      AutoSize        =   -1  'True
      Caption         =   "&Vehículo:"
      Height          =   210
      Left            =   4860
      TabIndex        =   29
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
      TabIndex        =   41
      Top             =   3780
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
      Caption         =   "Ingrese aquí los Datos del Viaje"
      Height          =   210
      Left            =   780
      TabIndex        =   51
      Top             =   300
      Width           =   2265
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "ViajePropiedad.frx":4626
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      Caption         =   "&Fecha:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   495
   End
End
Attribute VB_Name = "frmViajePropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mViaje As Viaje

Private mKeyDecimal As Boolean

Private mMedioPago As MedioPago

Private mNew As Boolean
Private mPermite2Conductores As Boolean

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef Viaje As Viaje)
    Dim Persona As Persona
    Dim PermisoModifyHora As Boolean
    Dim PermisoModifyFecha As Boolean
    
    Set mViaje = Viaje
    Set Viaje = Nothing
    mNew = (mViaje.IDRuta = "")
    
    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    PermisoModifyHora = pCPermiso.GotPermission(PERMISO_VIAJE_MODIFY_HORA, False)
    PermisoModifyFecha = pCPermiso.GotPermission(PERMISO_VIAJE_MODIFY_FECHA, False)
    
    With mViaje
        cmdCopyFrom.Visible = mNew
        
        cmdAnterior.Visible = (mNew Or PermisoModifyFecha)
        dtpFecha.Enabled = (mNew Or PermisoModifyFecha)
        cmdSiguiente.Visible = (mNew Or PermisoModifyFecha)
        cmdHoy.Visible = (mNew Or PermisoModifyFecha)
        dtpHora.Enabled = (mNew Or PermisoModifyFecha Or PermisoModifyHora)
        
        dtpFecha.value = .FechaHora_FormattedAsDate
        dtpFecha_Change
        dtpHora.value = .FechaHora_FormattedAsTime
        
        txtImporte.Text = Format(0, "Currency")
        txtImporteContado.Text = Format(0, "Currency")
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboRuta, "SELECT RTRIM(IDRuta) AS IDRuta FROM Ruta WHERE Activo = 1 OR IDRuta = '" & ReplaceQuote(.IDRuta) & "' " & IIf(pCPermiso.RutaWhere <> "", " AND " & Replace(pCPermiso.RutaWhere, "%TABLENAME%", "Ruta"), "") & "ORDER BY IDRuta", "IDRuta", "IDRuta", "Rutas", cscpItemOrNone, .IDRuta) Then
            Unload Me
            Exit Sub
        End If
        datcboRuta.Enabled = (mNew Or pCPermiso.GotPermission(PERMISO_VIAJE_MODIFY_RUTA, False))
                
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
                txtImporteContado.Text = .ImporteContado_Formatted
                
                'MEDIO DE PAGO
                If Not CSM_Control_DataCombo.FillFromSQL(datcboMedioPago, "SELECT IDMedioPago, Nombre FROM MedioPago WHERE Activo = 1 OR IDMedioPago = " & .IDMedioPago & " ORDER BY Nombre", "IDMedioPago", "Nombre", "Medios de Pago", cscpItemOrFirst, IIf(.IDMedioPago = 0, pParametro.MedioPago_Predeterminado_ID, .IDMedioPago)) Then
                    Unload Me
                    Exit Sub
                End If
                datcboMedioPago_Change
                cboCuotas.ListIndex = CSM_Control_ComboBox.GetListIndexByText(cboCuotas, .Cuotas_Formatted, cscpItemOrFirst)
                txtOperacion.Text = .Operacion
        
            Case pParametro.Ruta_Paquete_ID
                txtRutaOtra.Text = .RutaOtra
                
                txtPersona.Tag = 0
                txtPersona.Text = ""
                
                txtImporte.Text = .Importe_Formatted
                txtImporteContado.Text = Format(0, "Currency")
                
                'MEDIO DE PAGO
                If Not CSM_Control_DataCombo.FillFromSQL(datcboMedioPago, "SELECT IDMedioPago, Nombre FROM MedioPago WHERE Activo = 1 OR IDMedioPago = " & .IDMedioPago & " ORDER BY Nombre", "IDMedioPago", "Nombre", "Medios de Pago", cscpItemOrFirst, pParametro.MedioPago_Predeterminado_ID) Then
                    Unload Me
                    Exit Sub
                End If
                datcboMedioPago_Change
                cboCuotas.ListIndex = -1
                txtOperacion.Text = ""
        
            Case Else
                txtRutaOtra.Text = ""
                txtPersona.Tag = 0
                txtPersona.Text = ""
                
                txtImporte.Text = Format(0, "Currency")
                txtImporteContado.Text = Format(0, "Currency")
        
                'MEDIO DE PAGO
                If Not CSM_Control_DataCombo.FillFromSQL(datcboMedioPago, "SELECT IDMedioPago, Nombre FROM MedioPago WHERE Activo = 1 OR IDMedioPago = " & .IDMedioPago & " ORDER BY Nombre", "IDMedioPago", "Nombre", "Medios de Pago", cscpItemOrFirst, pParametro.MedioPago_Predeterminado_ID) Then
                    Unload Me
                    Exit Sub
                End If
                datcboMedioPago_Change
                cboCuotas.ListIndex = -1
                txtOperacion.Text = ""
        End Select
        datcboRuta_Change
        
        If Not mNew Then
            txtKilometro.Text = IIf(.Kilometro = 0, "", .Kilometro)
            txtDuracion.Text = IIf(.Duracion = 0, "", .Duracion)
        End If
        
        If pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_CAJA_VIEW_ALL, False) Then
            If Not CSM_Control_DataCombo.FillFromSQL(datcboCuentaCorrienteCaja, "SELECT IDCuentaCorrienteCaja, Nombre FROM CuentaCorrienteCaja WHERE Activo = 1 OR IDCuentaCorrienteCaja = " & .IDCuentaCorrienteCaja & " ORDER BY Nombre", "IDCuentaCorrienteCaja", "Nombre", "Cajas", cscpItemOrFirstIfUnique, .IDCuentaCorrienteCaja) Then
                Unload Me
                Exit Sub
            End If
        Else
            If Not CSM_Control_DataCombo.FillFromSQL(datcboCuentaCorrienteCaja, "SELECT CuentaCorrienteCaja.IDCuentaCorrienteCaja, CuentaCorrienteCaja.Nombre FROM CuentaCorrienteCaja LEFT JOIN Persona ON CuentaCorrienteCaja.IDPersona = Persona.IDPersona WHERE (CuentaCorrienteCaja.Activo = 1 OR CuentaCorrienteCaja.IDCuentaCorrienteCaja = " & .IDCuentaCorrienteCaja & ") AND (CuentaCorrienteCaja.IDCuentaCorrienteCaja = " & pUsuario.IDCuentaCorrienteCaja & " OR Persona.IDPersona = " & mViaje.IDConductor & " OR CuentaCorrienteCaja.IDCuentaCorrienteCaja = " & .IDCuentaCorrienteCaja & ") ORDER BY CuentaCorrienteCaja.Nombre", "IDCuentaCorrienteCaja", "Nombre", "Cajas", cscpItemOrFirstIfUnique, IIf(.IDCuentaCorrienteCaja = 0, pUsuario.IDCuentaCorrienteCaja, .IDCuentaCorrienteCaja)) Then
                Unload Me
                Exit Sub
            End If
        End If
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboVehiculo, "(SELECT -1 AS IDVehiculo, '------------------' AS Nombre, 1 AS Orden FROM Vehiculo) UNION (SELECT IDVehiculo, Nombre, 2 AS Orden FROM Vehiculo WHERE Activo = 1 OR IDVehiculo = " & .IDVehiculo & ") ORDER BY Orden, Nombre", "IDVehiculo", "Nombre", "Vehículos", cscpItemOrFirst, .IDVehiculo) Then
            Unload Me
            Exit Sub
        End If
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboConductor, "(SELECT -1 AS IDPersona, '------------------' AS ApellidoNombre, 1 AS Orden FROM Persona) UNION (SELECT IDPersona, dbo.udf_GetEntidadApellidoYNombre(Apellido, Nombre) AS ApellidoNombre, 2 AS Orden FROM Persona WHERE (Activo = 1 OR IDPersona = " & .IDConductor & ") AND EntidadTipo = '" & ENTIDAD_TIPO_PERSONA_CONDUCTOR & "') ORDER BY Orden, ApellidoNombre", "IDPersona", "ApellidoNombre", "Conductores", cscpItemOrFirst, .IDConductor) Then
            Unload Me
            Exit Sub
        End If
        chkAcreditaSueldo.value = IIf(.AcreditaSueldo, vbChecked, vbUnchecked)
        
        If pParametro.Viaje_Permite_2_Conductores Then
            If Not CSM_Control_DataCombo.FillFromSQL(datcboConductor2, "(SELECT -1 AS IDPersona, '------------------' AS ApellidoNombre, 1 AS Orden FROM Persona) UNION (SELECT IDPersona, dbo.udf_GetEntidadApellidoYNombre(Apellido, Nombre) AS ApellidoNombre, 2 AS Orden FROM Persona WHERE (Activo = 1 OR IDPersona = " & .IDConductor2 & ") AND EntidadTipo = '" & ENTIDAD_TIPO_PERSONA_CONDUCTOR & "') ORDER BY Orden, ApellidoNombre", "IDPersona", "ApellidoNombre", "Conductores", cscpItemOrFirst, IIf(mPermite2Conductores, .IDConductor2, -1)) Then
                Unload Me
                Exit Sub
            End If
            chkAcreditaSueldo2.value = IIf(mPermite2Conductores = False Or .AcreditaSueldo2, vbChecked, vbUnchecked)
        End If
        
        chkCharter.value = IIf(.Charter, vbChecked, vbUnchecked)
        
        cboDiaSemana.Enabled = (mNew)
        
        txtEstado.Text = .Estado_ToString
        imgEstadoActivo.Visible = (.Estado = VIAJE_ESTADO_ACTIVO)
        imgEstadoEnProgreso.Visible = (.Estado = VIAJE_ESTADO_EN_PROGRESO)
        imgEstadoFinalizado.Visible = (.Estado = VIAJE_ESTADO_FINALIZADO)
        imgEstadoCancelado.Visible = (.Estado = VIAJE_ESTADO_CANCELADO)
        
        txtNotas.Text = .Notas
        chkPersonal.value = IIf(.Personal, vbChecked, vbUnchecked)
    End With
    
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Private Sub Form_Load()
    Dim Weekday As Byte
    
    For Weekday = 1 To 7
        cboDiaSemana.AddItem WeekdayName(Weekday)
    Next Weekday
    
    lblConductor2.Visible = (pParametro.Viaje_Permite_2_Conductores)
    datcboConductor2.Visible = (pParametro.Viaje_Permite_2_Conductores)
    chkAcreditaSueldo2.Visible = (pParametro.Viaje_Permite_2_Conductores)
End Sub

Private Sub cmdCopyFrom_Click()
    Screen.MousePointer = vbHourglass
    frmHorario.Show
    If frmHorario.WindowState = vbMinimized Then
        frmHorario.WindowState = vbNormal
    End If
    frmHorario.FormWaitingForSelect = Me.Name
    frmHorario.AllowMultipleSelect = False
    frmHorario.SetFocus
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdAuditoria_Click()
    frmViajePropiedadAuditoria.LoadDataAndShow mViaje
End Sub

Private Sub cmdAnterior_Click()
    dtpFecha.value = DateAdd("d", -1, dtpFecha.value)
    dtpFecha.SetFocus
    dtpFecha_Change
End Sub

Private Sub txtDiaSemana_GotFocus()
    On Error Resume Next
    dtpFecha.SetFocus
End Sub

Private Sub dtpFecha_Change()
    SetCaption
    txtDiaSemana.Text = WeekdayName(Weekday(dtpFecha.value))
    If cboDiaSemana.Enabled Then
        cboDiaSemana.ListIndex = Weekday(dtpFecha.value) - 1
    End If
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

Private Sub dtpHora_Change()
    SetCaption
End Sub

Private Sub datcboRuta_Change()
    Dim rutaSeleccionada As Ruta
    
    SetCaption
    ShowControls
    
    If datcboRuta.BoundText <> "" Then
        Set rutaSeleccionada = New Ruta
        rutaSeleccionada.IDRuta = datcboRuta.BoundText
        If rutaSeleccionada.Load() Then
            txtKilometro.Text = CSM_Function.IfIsZeroLenghtString_Null(rutaSeleccionada.Kilometro)
            txtDuracion.Text = CSM_Function.IfIsZeroLenghtString_Null(rutaSeleccionada.Duracion)
        End If
        Set rutaSeleccionada = Nothing
    End If
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

Private Sub txtRutaOtra_GotFocus()
    CSM_Control_TextBox.SelAllText txtRutaOtra
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

Private Sub txtImporteContado_GotFocus()
    CSM_Control_TextBox.SelAllText txtImporteContado
End Sub

Private Sub txtImporteContado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDecimal Then
        mKeyDecimal = True
    End If
End Sub

Private Sub txtImporteContado_KeyPress(KeyAscii As Integer)
    If ((KeyAscii > 32 And KeyAscii < 48) Or KeyAscii > 57) And KeyAscii <> Asc(pRegionalSettings.CurrencyDecimalSymbol) And KeyAscii <> Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
        KeyAscii = 0
    End If
    If mKeyDecimal Then
        KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol)
        mKeyDecimal = False
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol) Then
        If InStr(1, txtImporteContado.Text, pRegionalSettings.CurrencyDecimalSymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
        If InStr(1, txtImporteContado.Text, pRegionalSettings.CurrencyCurrencySymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtImporteContado_LostFocus()
    If Not IsNumeric(txtImporteContado.Text) Then
        txtImporteContado.Text = Val(txtImporteContado.Text)
    End If
    txtImporteContado.Text = Format(CCur(txtImporteContado.Text), "Currency")

    lblCuentaCorrienteCaja.Visible = (CCur(txtImporteContado.Text) > 0)
    datcboCuentaCorrienteCaja.Visible = (CCur(txtImporteContado.Text) > 0)
    cmdCuentaCorrienteCaja.Visible = (CCur(txtImporteContado.Text) > 0)
End Sub

Private Sub datcboMedioPago_Change()
    If Val(datcboMedioPago.BoundText) > 0 Then
        Set mMedioPago = New MedioPago
        mMedioPago.IDMedioPago = Val(datcboMedioPago.BoundText)
        If mMedioPago.Load() Then
            lblCuotas.Visible = mMedioPago.UtilizaOperacion
            If mMedioPago.UtilizaOperacion Then
                If mMedioPago.MedioPagoPlan.LoadCuotas Then
                    CSM_Control_ComboBox.FillFromCollection cboCuotas, mMedioPago.MedioPagoPlan.CCuotas, "Cuota", "Cuota", cscpCurrentOrFirst
                End If
            End If
            cboCuotas.Visible = mMedioPago.UtilizaOperacion
            lblOperacion.Visible = mMedioPago.UtilizaOperacion
            txtOperacion.Visible = mMedioPago.UtilizaOperacion
        
            lblCuentaCorrienteCaja.Visible = (CCur(txtImporteContado.Text) > 0 And mMedioPago.IDCuentaCorrienteCaja = 0)
            datcboCuentaCorrienteCaja.Visible = (CCur(txtImporteContado.Text) > 0 And mMedioPago.IDCuentaCorrienteCaja = 0)
            cmdCuentaCorrienteCaja.Visible = (CCur(txtImporteContado.Text) > 0 And mMedioPago.IDCuentaCorrienteCaja = 0)
        End If
    Else
        lblCuotas.Visible = False
        cboCuotas.Visible = False
        lblOperacion.Visible = False
        txtOperacion.Visible = False
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
        Set frmVehiculo.lvwData.SelectedItem = frmVehiculo.lvwData.ListItems(KEY_STRINGER & datcboVehiculo.BoundText)
        frmVehiculo.lvwData.SelectedItem.EnsureVisible
        If frmVehiculo.WindowState = vbMinimized Then
            frmVehiculo.WindowState = vbNormal
        End If
        frmVehiculo.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub txtNotas_GotFocus()
    CSM_Control_TextBox.SelAllText txtNotas
    cmdOK.Default = False
End Sub

Private Sub txtNotas_LostFocus()
    cmdOK.Default = True
End Sub

Private Sub cmdOK_Click()
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
    If Not mNew And (datcboRuta.BoundText <> mViaje.IDRuta) Then
        If MsgBox("Ha modificado la Ruta del Viaje. Utilice esta opción sólo en casos realmente necesarios." & vbCr & vbCr & "¿Desea continuar?", vbExclamation + vbYesNo, App.Title) = vbNo Then
            datcboRuta.SetFocus
            Exit Sub
        End If
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
    If Not IsNumeric(txtImporteContado.Text) Then
        MsgBox "El Importe de Contado ingresado es incorrecto.", vbInformation, App.Title
        txtImporteContado.SetFocus
        Exit Sub
    End If
    If CCur(txtImporteContado.Text) < 0 Then
        MsgBox "El Importe de Contado debe ser mayor o igual a cero.", vbInformation, App.Title
        txtImporteContado.SetFocus
        Exit Sub
    End If
    If CCur(txtImporteContado.Text) > 0 And Val(datcboCuentaCorrienteCaja.BoundText) = 0 Then
        MsgBox "Debe seleccionar la Caja.", vbInformation, App.Title
        On Error Resume Next
        datcboCuentaCorrienteCaja.SetFocus
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
    
    If mNew Then
        If DateDiff("d", Date, dtpFecha.value) < 0 Then
            If MsgBox("Está por Crear un Viaje correspondientes a una Fecha anterior a Hoy." & vbCr & "Las Reservas Fijas no serán Generadas." & vbCr & vbCr & "¿Desea continuar de todos modos?", vbExclamation + vbYesNo, App.Title) = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    If datcboRuta.BoundText = pParametro.Ruta_ID_Otra And Val(txtPersona.Tag) = 0 Then
        MsgBox "Debe seleccionar el Cliente.", vbInformation, App.Title
        txtPersona.SetFocus
        Exit Sub
    End If
    
    With mViaje
        .FechaHora = dtpFecha.value & " " & dtpHora.value
        .IDRuta = datcboRuta.BoundText
        Select Case datcboRuta.BoundText
            Case pParametro.Ruta_ID_Otra
                .RutaOtra = txtRutaOtra.Text
                .IDPersona = Val(txtPersona.Tag)
                .Importe = CCur(txtImporte.Text)
                .ImporteContado = CCur(txtImporteContado.Text)
                .IDMedioPago = Val(datcboMedioPago.BoundText)
                If mMedioPago.UtilizaOperacion Then
                    .Cuotas = Val(cboCuotas.Text)
                    .Operacion = txtOperacion.Text
                Else
                    .Cuotas = 0
                    .Operacion = ""
                End If
                .IDCuentaCorrienteCaja = Val(datcboCuentaCorrienteCaja.BoundText)
                
            Case pParametro.Ruta_Paquete_ID
                .RutaOtra = txtRutaOtra.Text
                .IDPersona = 0
                .Importe = CCur(txtImporte.Text)
                .ImporteContado = 0
                .IDMedioPago = 0
                .Cuotas = 0
                .Operacion = ""
                .IDCuentaCorrienteCaja = 0
                
            Case Else
                .RutaOtra = ""
                .IDPersona = 0
                .Importe = 0
                .ImporteContado = 0
                .IDMedioPago = 0
                .Cuotas = 0
                .Operacion = ""
                .IDCuentaCorrienteCaja = 0
        End Select
        .Kilometro = Val(txtKilometro.Text)
        .Duracion = Val(txtDuracion.Text)
        .Charter = (chkCharter.value = vbChecked)
        .IDVehiculo = IIf(Val(datcboVehiculo.BoundText) = -1, 0, Val(datcboVehiculo.BoundText))
        .IDConductor = IIf(Val(datcboConductor.BoundText) = -1, 0, Val(datcboConductor.BoundText))
        .AcreditaSueldo = (chkAcreditaSueldo.value = vbChecked)
        If mPermite2Conductores Then
            .IDConductor2 = IIf(Val(datcboConductor2.BoundText) = -1, 0, Val(datcboConductor2.BoundText))
            .AcreditaSueldo2 = (chkAcreditaSueldo2.value = vbChecked)
        Else
            .IDConductor2 = 0
            .AcreditaSueldo2 = True
        End If
        .Notas = txtNotas.Text
        .Personal = (chkPersonal.value = vbChecked)
        If mNew Then
            .DiaSemanaBase = cboDiaSemana.ListIndex + 1
        End If
        If Not .Update Then
            Exit Sub
        End If
    End With
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mViaje = Nothing
    Set frmViajePropiedad = Nothing
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
        KeySave = Val(datcboConductor2.BoundText)
        Set recData = datcboConductor2.RowSource
        recData.Requery
        Set recData = Nothing
        datcboConductor2.BoundText = KeySave
    End If
End Sub

Public Sub HorarioSelected(ByVal DiaSemana As Byte, ByVal Hora As Date, ByVal IDRuta As String)
    Dim Horario As Horario
    Dim DaysDiff As Integer
    
    Set Horario = New Horario
    Horario.DiaSemana = DiaSemana
    Horario.Hora = Hora
    Horario.IDRuta = IDRuta
    If Horario.Load() Then
        If Not Horario.Activo Then
            MsgBox "No se puede seleccionar este Horario porque está inactivo.", vbExclamation, App.Title
            Set Horario = Nothing
            Exit Sub
        End If
        DaysDiff = DiaSemana - Weekday(Date)
        If DaysDiff < 0 Then
            DaysDiff = 7 + DaysDiff
        Else
            If DaysDiff = 0 And DateDiff("s", Time, Hora) < 0 Then
                DaysDiff = 7
            End If
        End If
        dtpFecha.value = DateAdd("d", DaysDiff, Date)
        dtpFecha_Change
        dtpHora.value = Format(Hora, "Short Time")
        datcboRuta.BoundText = IDRuta
        datcboVehiculo.BoundText = Horario.IDVehiculo
        If Val(datcboVehiculo.BoundText) = 0 Then
            datcboVehiculo.BoundText = -1
        End If
        
        datcboConductor.BoundText = Horario.IDConductor
        If Val(datcboConductor.BoundText) = 0 Then
            datcboConductor.BoundText = -1
        End If
        If mPermite2Conductores Then
            datcboConductor2.BoundText = Horario.IDConductor2
            If Val(datcboConductor2.BoundText) = 0 Then
                datcboConductor2.BoundText = -1
            End If
        End If
        txtNotas.Text = Horario.Notas
        chkPersonal.value = IIf(Horario.Personal, vbChecked, vbUnchecked)
    End If
    Set Horario = Nothing
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
    
    CaptionTemp = CaptionTemp & IIf(CaptionTemp = "", "", " - ") & dtpFecha.value & " " & dtpHora.value
    If datcboRuta.BoundText <> "" Then
        CaptionTemp = CaptionTemp & IIf(CaptionTemp = "", "", " - ") & datcboRuta.Text
    End If
    Caption = "Propiedades" & IIf(CaptionTemp = "", "", " ") & CaptionTemp
End Sub

Private Sub ShowControls()
    Dim ViajeEspecial As Boolean
    Dim ViajePaquete As Boolean
    Dim Ruta As Ruta
    
    ViajeEspecial = (datcboRuta.BoundText = pParametro.Ruta_ID_Otra)
    ViajePaquete = (datcboRuta.BoundText = pParametro.Ruta_Paquete_ID)
    
    If datcboRuta.BoundText <> "" Then
        Set Ruta = New Ruta
        Ruta.IDRuta = datcboRuta.BoundText
        If Ruta.Load() Then
            mPermite2Conductores = (pParametro.Viaje_Permite_2_Conductores And Ruta.Permite2Conductores)
            lblConductor2.Visible = mPermite2Conductores
            datcboConductor2.Visible = mPermite2Conductores
        End If
        Set Ruta = Nothing
    End If
    
    lblRutaOtra.Visible = (ViajeEspecial Or ViajePaquete)
    txtRutaOtra.Visible = (ViajeEspecial Or ViajePaquete)
    
    lblPersona.Visible = ViajeEspecial
    txtPersona.Visible = ViajeEspecial
    cmdPersona.Visible = ViajeEspecial
    cmdPersonaClear.Visible = ViajeEspecial
    
    lblImporte.Visible = (ViajeEspecial Or ViajePaquete)
    txtImporte.Visible = (ViajeEspecial Or ViajePaquete)
    
    lblImporteContado.Visible = ViajeEspecial
    txtImporteContado.Visible = ViajeEspecial
    
    datcboMedioPago.Visible = ViajeEspecial
    
    If Not mMedioPago Is Nothing Then
        lblCuotas.Visible = (ViajeEspecial And mMedioPago.UtilizaOperacion)
        cboCuotas.Visible = (ViajeEspecial And mMedioPago.UtilizaOperacion)
        lblOperacion.Visible = (ViajeEspecial And mMedioPago.UtilizaOperacion)
        txtOperacion.Visible = (ViajeEspecial And mMedioPago.UtilizaOperacion)
    Else
        lblCuotas.Visible = ViajeEspecial
        cboCuotas.Visible = ViajeEspecial
        lblOperacion.Visible = ViajeEspecial
        txtOperacion.Visible = ViajeEspecial
    End If
    
    lblCuentaCorrienteCaja.Visible = (ViajeEspecial And CCur(txtImporteContado.Text) > 0)
    datcboCuentaCorrienteCaja.Visible = (ViajeEspecial And CCur(txtImporteContado.Text) > 0)
    cmdCuentaCorrienteCaja.Visible = (ViajeEspecial And CCur(txtImporteContado.Text) > 0)
    
    lblConductorSueldo.Visible = Not (ViajeEspecial Or ViajePaquete)
    chkAcreditaSueldo.Visible = Not (ViajeEspecial Or ViajePaquete)
    chkAcreditaSueldo2.Visible = (pParametro.Viaje_Permite_2_Conductores And Not (ViajeEspecial Or ViajePaquete))
    
    lblDiaSemana.Visible = Not (ViajeEspecial Or ViajePaquete)
    cboDiaSemana.Visible = Not (ViajeEspecial Or ViajePaquete)
End Sub

Public Sub FillComboBoxMedioPago()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboMedioPago.BoundText)
    Set recData = datcboMedioPago.RowSource
    recData.Requery
    Set recData = Nothing
    datcboMedioPago.BoundText = KeySave
End Sub
