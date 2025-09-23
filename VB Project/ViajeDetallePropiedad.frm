VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmViajeDetallePropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   7215
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
   Icon            =   "ViajeDetallePropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   9510
   Begin VB.PictureBox picImporte 
      BorderStyle     =   0  'None
      Height          =   1755
      Left            =   4860
      ScaleHeight     =   1755
      ScaleWidth      =   4575
      TabIndex        =   108
      Top             =   1380
      Width           =   4575
      Begin VB.ComboBox cboCuotas 
         Height          =   330
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   360
         Width           =   630
      End
      Begin VB.TextBox txtImporteContado 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   900
         MaxLength       =   20
         TabIndex        =   55
         Top             =   0
         Width           =   1395
      End
      Begin VB.TextBox txtSaldoActual 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   2760
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1455
      End
      Begin VB.TextBox txtImporteCuentaCorriente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   900
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   112
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1455
      End
      Begin VB.CommandButton cmdCuentaCorrienteCaja 
         Caption         =   "..."
         Height          =   315
         Left            =   4260
         TabIndex        =   111
         TabStop         =   0   'False
         ToolTipText     =   "Cajas"
         Top             =   1440
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox chkImprimirSaldo 
         Caption         =   "Imprimir Saldo Actual"
         Height          =   210
         Left            =   2760
         TabIndex        =   61
         Top             =   780
         Width           =   1905
      End
      Begin VB.CommandButton cmdSaldoActual 
         Caption         =   "..."
         Height          =   315
         Left            =   4260
         TabIndex        =   110
         ToolTipText     =   "Ver Movimientos..."
         Top             =   1020
         Width           =   255
      End
      Begin VB.TextBox txtOperacion 
         Height          =   315
         Left            =   2400
         MaxLength       =   20
         TabIndex        =   60
         Top             =   360
         Visible         =   0   'False
         Width           =   2115
      End
      Begin MSDataListLib.DataCombo datcboCuentaCorrienteCaja 
         Height          =   330
         Left            =   900
         TabIndex        =   62
         Top             =   1440
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
      Begin MSDataListLib.DataCombo datcboMedioPago 
         Height          =   330
         Left            =   2400
         TabIndex        =   56
         Top             =   0
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
         Left            =   0
         TabIndex        =   57
         Top             =   420
         Width           =   555
      End
      Begin VB.Line linPagos 
         X1              =   0
         X2              =   4500
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label lblImporteContado 
         AutoSize        =   -1  'True
         Caption         =   "P. Actual:"
         Height          =   210
         Left            =   0
         TabIndex        =   54
         Top             =   60
         Width           =   690
      End
      Begin VB.Label lblImporteCuentaCorriente 
         AutoSize        =   -1  'True
         Caption         =   "P. Cta. Cte.:"
         Height          =   210
         Left            =   0
         TabIndex        =   115
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label lblCuentaCorrienteCaja 
         AutoSize        =   -1  'True
         Caption         =   "Caja:"
         Height          =   210
         Left            =   0
         TabIndex        =   114
         Top             =   1500
         Width           =   360
      End
      Begin VB.Label lblOperacion 
         AutoSize        =   -1  'True
         Caption         =   "Operación:"
         Height          =   210
         Left            =   1560
         TabIndex        =   59
         Top             =   420
         Visible         =   0   'False
         Width           =   795
      End
   End
   Begin VB.CheckBox chkRutaConexion 
      Caption         =   "Combinar Ruta:"
      Height          =   210
      Left            =   120
      TabIndex        =   31
      Top             =   6240
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.TextBox txtImporteTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8100
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   960
      Width           =   1275
   End
   Begin VB.TextBox txtImporteSeguro 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   8100
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   540
      Width           =   1275
   End
   Begin VB.TextBox txtValorDeclarado 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5760
      MaxLength       =   20
      TabIndex        =   39
      Top             =   540
      Width           =   1395
   End
   Begin VB.TextBox txtHorario 
      Height          =   315
      Left            =   1020
      MaxLength       =   50
      TabIndex        =   19
      Top             =   3780
      Width           =   3615
   End
   Begin VB.TextBox txtDomicilio 
      Height          =   315
      Left            =   1020
      MaxLength       =   100
      TabIndex        =   17
      Top             =   3420
      Width           =   3615
   End
   Begin VB.OptionButton optPagaRecibe 
      Height          =   210
      Left            =   780
      TabIndex        =   10
      Top             =   2760
      Width           =   195
   End
   Begin VB.OptionButton optPagaEnvia 
      Height          =   210
      Left            =   780
      TabIndex        =   5
      Top             =   2220
      Width           =   195
   End
   Begin VB.CommandButton cmdPersonaRecibe 
      Height          =   315
      Left            =   3720
      Picture         =   "ViajeDetallePropiedad.frx":054A
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Buscar..."
      Top             =   2700
      Width           =   315
   End
   Begin VB.TextBox txtPersonaRecibe 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2700
      Width           =   2715
   End
   Begin VB.CommandButton cmdPersonaRecibeUltimo 
      Caption         =   "&Ultimo"
      Height          =   315
      Left            =   4080
      TabIndex        =   13
      Top             =   2700
      Width           =   555
   End
   Begin VB.TextBox txtCanceladoPor 
      Height          =   315
      Left            =   1800
      TabIndex        =   107
      Top             =   7320
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtRetira 
      Height          =   315
      Left            =   5760
      MaxLength       =   50
      TabIndex        =   82
      Top             =   5580
      Width           =   3615
   End
   Begin VB.CheckBox chkForzarDebito 
      Height          =   210
      Left            =   7620
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   5220
      Width           =   195
   End
   Begin VB.TextBox txtFacturaNumero 
      Height          =   315
      Left            =   5760
      MaxLength       =   20
      TabIndex        =   71
      Top             =   4200
      Width           =   2115
   End
   Begin VB.CheckBox chkEntregada 
      Alignment       =   1  'Right Justify
      Caption         =   "Entregada:"
      Height          =   210
      Left            =   4830
      TabIndex        =   78
      Top             =   5220
      Width           =   1125
   End
   Begin VB.ComboBox cboDejarTraer 
      Height          =   330
      Left            =   3780
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   4140
      Width           =   855
   End
   Begin VB.TextBox txtTelefono 
      Height          =   315
      Left            =   1020
      MaxLength       =   30
      TabIndex        =   21
      Top             =   4140
      Width           =   2295
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   315
      Left            =   1020
      MaxLength       =   100
      TabIndex        =   15
      Top             =   3060
      Width           =   3615
   End
   Begin VB.ComboBox cboRealizado 
      Height          =   330
      Left            =   5760
      Style           =   2  'Dropdown List
      TabIndex        =   75
      Top             =   5160
      Width           =   615
   End
   Begin VB.TextBox txtFacturarNotas 
      Height          =   315
      Left            =   6120
      MaxLength       =   50
      TabIndex        =   69
      Top             =   3780
      Width           =   3255
   End
   Begin VB.CheckBox chkFacturar 
      Height          =   210
      Left            =   5790
      TabIndex        =   68
      Top             =   3840
      Width           =   225
   End
   Begin VB.CommandButton cmdVerificarAsiento 
      Caption         =   "Verif. Lugar"
      Height          =   615
      Left            =   4020
      TabIndex        =   35
      Top             =   6240
      Width           =   615
   End
   Begin VB.TextBox txtBaja 
      Height          =   315
      Left            =   1020
      MaxLength       =   50
      TabIndex        =   30
      Top             =   5820
      Width           =   3615
   End
   Begin VB.TextBox txtSube 
      Height          =   315
      Left            =   1020
      MaxLength       =   50
      TabIndex        =   26
      Top             =   5040
      Width           =   3615
   End
   Begin VB.CommandButton cmdAuditoria 
      Height          =   375
      Left            =   6780
      Picture         =   "ViajeDetallePropiedad.frx":0AD4
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   6720
      Width           =   435
   End
   Begin VB.PictureBox picTipo 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1020
      ScaleHeight     =   195
      ScaleWidth      =   2655
      TabIndex        =   1
      Top             =   1260
      Width           =   2655
      Begin VB.OptionButton optTipoComision 
         Caption         =   "Comisión"
         Height          =   210
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
      Begin VB.OptionButton optTipoPasajero 
         Caption         =   "Pasajero"
         Height          =   210
         Left            =   1260
         TabIndex        =   3
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.TextBox txtAsiento 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   2580
      Locked          =   -1  'True
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   1620
      Width           =   615
   End
   Begin VB.CommandButton cmdListaPrecio 
      Caption         =   "..."
      Height          =   315
      Left            =   9120
      TabIndex        =   103
      TabStop         =   0   'False
      ToolTipText     =   "Listas de Precios"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txtReservadoPor 
      Height          =   315
      Left            =   5760
      MaxLength       =   50
      TabIndex        =   73
      Top             =   4740
      Width           =   3615
   End
   Begin VB.CommandButton cmdSiguiente 
      Height          =   315
      Left            =   4020
      Picture         =   "ViajeDetallePropiedad.frx":10FE
      Style           =   1  'Graphical
      TabIndex        =   96
      TabStop         =   0   'False
      ToolTipText     =   "Siguiente"
      Top             =   120
      Width           =   300
   End
   Begin VB.CommandButton cmdRuta 
      Caption         =   "..."
      Height          =   315
      Left            =   4380
      TabIndex        =   98
      TabStop         =   0   'False
      ToolTipText     =   "Rutas"
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton cmdAnterior 
      Height          =   315
      Left            =   1020
      Picture         =   "ViajeDetallePropiedad.frx":1688
      Style           =   1  'Graphical
      TabIndex        =   95
      TabStop         =   0   'False
      ToolTipText     =   "Anterior"
      Top             =   120
      Width           =   300
   End
   Begin VB.CommandButton cmdHoy 
      Height          =   315
      Left            =   4320
      Picture         =   "ViajeDetallePropiedad.frx":1C12
      Style           =   1  'Graphical
      TabIndex        =   97
      TabStop         =   0   'False
      ToolTipText     =   "Hoy"
      Top             =   120
      Width           =   315
   End
   Begin VB.TextBox txtDiaSemana 
      Height          =   315
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   120
      Width           =   1050
   End
   Begin VB.TextBox txtPersonaCuentaCorriente 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   3360
      Width           =   2955
   End
   Begin VB.CommandButton cmdPersonaCuentaCorriente 
      Height          =   315
      Left            =   8700
      Picture         =   "ViajeDetallePropiedad.frx":1D5C
      Style           =   1  'Graphical
      TabIndex        =   65
      ToolTipText     =   "Buscar..."
      Top             =   3360
      Width           =   315
   End
   Begin VB.CommandButton cmdPersonaCuentaCorrienteClear 
      Height          =   315
      Left            =   9060
      Picture         =   "ViajeDetallePropiedad.frx":22E6
      Style           =   1  'Graphical
      TabIndex        =   66
      ToolTipText     =   "Borrar"
      Top             =   3360
      Width           =   315
   End
   Begin VB.CommandButton cmdPersonaUltimo 
      Caption         =   "&Ultimo"
      Height          =   315
      Left            =   4080
      TabIndex        =   8
      Top             =   2160
      Width           =   555
   End
   Begin VB.CommandButton cmdLugarOrigen 
      Caption         =   "..."
      Height          =   315
      Left            =   4380
      TabIndex        =   101
      TabStop         =   0   'False
      ToolTipText     =   "Lugares"
      Top             =   4680
      Width           =   255
   End
   Begin VB.CommandButton cmdLugarDestino 
      Caption         =   "..."
      Height          =   315
      Left            =   4380
      TabIndex        =   102
      TabStop         =   0   'False
      ToolTipText     =   "Lugares"
      Top             =   5460
      Width           =   255
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   960
      Width           =   1395
   End
   Begin VB.TextBox txtPersona 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2715
   End
   Begin VB.TextBox txtOrden 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   1620
      Width           =   435
   End
   Begin VB.CommandButton cmdPersona 
      Height          =   315
      Left            =   3720
      Picture         =   "ViajeDetallePropiedad.frx":2870
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Buscar..."
      Top             =   2160
      Width           =   315
   End
   Begin VB.TextBox txtNotas 
      Height          =   585
      Left            =   5760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   84
      Top             =   6000
      Width           =   3615
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   8160
      TabIndex        =   86
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4860
      TabIndex        =   85
      Top             =   6720
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   315
      Left            =   2400
      TabIndex        =   89
      Top             =   120
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
      Format          =   58720257
      CurrentDate     =   36950
   End
   Begin MSDataListLib.DataCombo datcboHora 
      Height          =   330
      Left            =   1020
      TabIndex        =   91
      Top             =   480
      Width           =   975
      _ExtentX        =   1720
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
   Begin MSDataListLib.DataCombo datcboRuta 
      Height          =   330
      Left            =   1020
      TabIndex        =   93
      Top             =   840
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
   Begin MSDataListLib.DataCombo datcboListaPrecio 
      Height          =   330
      Left            =   5760
      TabIndex        =   37
      Top             =   120
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
   Begin MSDataListLib.DataCombo datcboOrigen 
      Height          =   330
      Left            =   1020
      TabIndex        =   24
      Top             =   4680
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
   Begin MSDataListLib.DataCombo datcboDestino 
      Height          =   330
      Left            =   1020
      TabIndex        =   28
      Top             =   5460
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
   Begin MSComCtl2.DTPicker dtpEntregadaHora 
      Height          =   315
      Left            =   7980
      TabIndex        =   80
      Top             =   5160
      Visible         =   0   'False
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
      Format          =   58720259
      UpDown          =   -1  'True
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpEntregadaFecha 
      Height          =   315
      Left            =   6240
      TabIndex        =   79
      Top             =   5160
      Visible         =   0   'False
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
      Format          =   58720257
      CurrentDate     =   36950
   End
   Begin MSDataListLib.DataCombo datcboRutaConexion 
      Height          =   330
      Left            =   120
      TabIndex        =   32
      Top             =   6480
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
   Begin MSDataListLib.DataCombo datcboViajeConexion 
      Height          =   330
      Left            =   1920
      TabIndex        =   34
      Top             =   6480
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
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
   Begin VB.PictureBox picPagos 
      BorderStyle     =   0  'None
      Height          =   1755
      Left            =   4860
      ScaleHeight     =   1755
      ScaleWidth      =   4515
      TabIndex        =   109
      Top             =   1380
      Width           =   4515
      Begin VB.TextBox txtPagosSaldo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3180
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1275
      End
      Begin VB.TextBox txtPagosTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1275
      End
      Begin VB.CommandButton cmdPagoAgregar 
         Height          =   390
         Left            =   4140
         Picture         =   "ViajeDetallePropiedad.frx":2DFA
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Agregar"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdPagoEditar 
         Height          =   390
         Left            =   4140
         Picture         =   "ViajeDetallePropiedad.frx":3384
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Editar"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdPagoEliminar 
         Height          =   390
         Left            =   4140
         Picture         =   "ViajeDetallePropiedad.frx":390E
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Borrar"
         Top             =   960
         Width           =   375
      End
      Begin MSComctlLib.ListView lvwPagos 
         Height          =   1395
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   2461
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "IDMovimientoCuentaCorriente"
            Text            =   "IDMovimientoCuentaCorriente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "Fecha"
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Key             =   "Importe"
            Text            =   "Importe"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "MedioPago"
            Text            =   "Medio de Pago"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "Caja"
            Text            =   "Caja"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblPagosSaldo 
         AutoSize        =   -1  'True
         Caption         =   "Saldo:"
         Height          =   210
         Left            =   2640
         TabIndex        =   52
         Top             =   1500
         Width           =   450
      End
      Begin VB.Label lblPagosTotal 
         AutoSize        =   -1  'True
         Caption         =   "Total Pagos:"
         Height          =   210
         Left            =   60
         TabIndex        =   50
         Top             =   1500
         Width           =   885
      End
   End
   Begin VB.Image imgInformacion 
      Height          =   360
      Left            =   7320
      Picture         =   "ViajeDetallePropiedad.frx":3E98
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   360
   End
   Begin VB.Line linImporte 
      X1              =   4740
      X2              =   9360
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label lblViajeConexion 
      AutoSize        =   -1  'True
      Caption         =   "Conexión con Viaje:"
      Height          =   210
      Left            =   1920
      TabIndex        =   33
      Top             =   6240
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label lblImporteTotal 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      Height          =   210
      Left            =   7380
      TabIndex        =   44
      Top             =   1020
      Width           =   390
   End
   Begin VB.Label lblImporteSeguro 
      AutoSize        =   -1  'True
      Caption         =   "Seguro:"
      Height          =   210
      Left            =   7380
      TabIndex        =   40
      Top             =   600
      Width           =   570
   End
   Begin VB.Label lblValorDeclarado 
      AutoSize        =   -1  'True
      Caption         =   "Valor Decl.:"
      Height          =   210
      Left            =   4860
      TabIndex        =   38
      Top             =   600
      Width           =   840
   End
   Begin VB.Label lblHorario 
      AutoSize        =   -1  'True
      Caption         =   "Horario:"
      Height          =   210
      Left            =   120
      TabIndex        =   18
      Top             =   3840
      Width           =   570
   End
   Begin VB.Label lblDomicilio 
      AutoSize        =   -1  'True
      Caption         =   "Domicilio:"
      Height          =   210
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Width           =   660
   End
   Begin VB.Label lblRetira 
      AutoSize        =   -1  'True
      Caption         =   "Retira:"
      Height          =   210
      Left            =   4860
      TabIndex        =   81
      Top             =   5640
      Width           =   465
   End
   Begin VB.Label lblForzarDebito 
      AutoSize        =   -1  'True
      Caption         =   "Debitar Viaje:"
      Height          =   210
      Left            =   6540
      TabIndex        =   76
      Top             =   5220
      Width           =   960
   End
   Begin VB.Label lblFacturaNumero 
      AutoSize        =   -1  'True
      Caption         =   "Factura Nº:"
      Height          =   210
      Left            =   4860
      TabIndex        =   70
      Top             =   4260
      Width           =   825
   End
   Begin VB.Line Line6 
      X1              =   4740
      X2              =   120
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line5 
      X1              =   4740
      X2              =   120
      Y1              =   2580
      Y2              =   2580
   End
   Begin VB.Line Line4 
      X1              =   4740
      X2              =   120
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line3 
      X1              =   4740
      X2              =   9360
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line2 
      X1              =   4740
      X2              =   9360
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label lblTelefono 
      AutoSize        =   -1  'True
      Caption         =   "Teléfono:"
      Height          =   210
      Left            =   120
      TabIndex        =   20
      Top             =   4200
      Width           =   675
   End
   Begin VB.Label lblDescripcion 
      AutoSize        =   -1  'True
      Caption         =   "Descripc.:"
      Height          =   210
      Left            =   120
      TabIndex        =   14
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label lblPersonaRecibe 
      AutoSize        =   -1  'True
      Caption         =   "&Recibe:"
      Height          =   210
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   540
   End
   Begin VB.Label lblFacturar 
      AutoSize        =   -1  'True
      Caption         =   "Facturar:"
      Height          =   210
      Left            =   4860
      TabIndex        =   67
      Top             =   3840
      Width           =   660
   End
   Begin VB.Label lblRealizado 
      AutoSize        =   -1  'True
      Caption         =   "Realizado:"
      Height          =   210
      Left            =   4860
      TabIndex        =   74
      Top             =   5220
      Width           =   750
   End
   Begin VB.Label lblBaja 
      AutoSize        =   -1  'True
      Caption         =   "Baja:"
      Height          =   210
      Left            =   120
      TabIndex        =   29
      Top             =   5880
      Width           =   360
   End
   Begin VB.Label lblSube 
      AutoSize        =   -1  'True
      Caption         =   "Sube:"
      Height          =   210
      Left            =   120
      TabIndex        =   25
      Top             =   5100
      Width           =   420
   End
   Begin VB.Label lblAsiento 
      AutoSize        =   -1  'True
      Caption         =   "Asiento N°"
      Height          =   210
      Left            =   1680
      TabIndex        =   106
      Top             =   1680
      Width           =   765
   End
   Begin VB.Image imgPasajeroConfirmado 
      Height          =   480
      Left            =   4080
      Picture         =   "ViajeDetallePropiedad.frx":4119
      Top             =   1500
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblListaPrecio 
      AutoSize        =   -1  'True
      Caption         =   "Lista Precio:"
      Height          =   210
      Left            =   4860
      TabIndex        =   36
      Top             =   180
      Width           =   885
   End
   Begin VB.Label lblReservadoPor 
      AutoSize        =   -1  'True
      Caption         =   "Reservado:"
      Height          =   210
      Left            =   4860
      TabIndex        =   72
      Top             =   4800
      Width           =   840
   End
   Begin VB.Label lblEstado 
      AutoSize        =   -1  'True
      Caption         =   "Estado:"
      Height          =   210
      Left            =   3420
      TabIndex        =   105
      Top             =   1680
      Width           =   540
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   4740
      X2              =   4740
      Y1              =   120
      Y2              =   7080
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      Caption         =   "&Fecha:"
      Height          =   210
      Left            =   120
      TabIndex        =   87
      Top             =   180
      Width           =   495
   End
   Begin VB.Label lblHora 
      AutoSize        =   -1  'True
      Caption         =   "&Hora:"
      Height          =   210
      Left            =   120
      TabIndex        =   90
      Top             =   540
      Width           =   390
   End
   Begin VB.Label lblRuta 
      AutoSize        =   -1  'True
      Caption         =   "&Ruta:"
      Height          =   210
      Left            =   120
      TabIndex        =   92
      Top             =   900
      Width           =   375
   End
   Begin VB.Label lblPersonaCuentaCorriente 
      AutoSize        =   -1  'True
      Caption         =   "Debitar a:"
      Height          =   210
      Left            =   4860
      TabIndex        =   63
      Top             =   3420
      Width           =   690
   End
   Begin VB.Label lblOrigen 
      AutoSize        =   -1  'True
      Caption         =   "&Origen:"
      Height          =   210
      Left            =   120
      TabIndex        =   23
      Top             =   4740
      Width           =   525
   End
   Begin VB.Label lblDestino 
      AutoSize        =   -1  'True
      Caption         =   "&Destino:"
      Height          =   210
      Left            =   120
      TabIndex        =   27
      Top             =   5520
      Width           =   585
   End
   Begin VB.Label lblOrden 
      AutoSize        =   -1  'True
      Caption         =   "Orden:"
      Height          =   210
      Left            =   120
      TabIndex        =   104
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblTipo 
      AutoSize        =   -1  'True
      Caption         =   "&Tipo:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   1260
      Width           =   345
   End
   Begin VB.Label lblNotas 
      AutoSize        =   -1  'True
      Caption         =   "Notas:"
      Height          =   210
      Left            =   4860
      TabIndex        =   83
      Top             =   6000
      Width           =   465
   End
   Begin VB.Label lblImporte 
      AutoSize        =   -1  'True
      Caption         =   "&Importe:"
      Height          =   210
      Left            =   4860
      TabIndex        =   42
      Top             =   1020
      Width           =   570
   End
   Begin VB.Label lblPersona 
      AutoSize        =   -1  'True
      Caption         =   "&Pasajero:"
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   2220
      Width           =   675
   End
   Begin VB.Image imgPasajeroCancelado 
      Height          =   480
      Left            =   4080
      Picture         =   "ViajeDetallePropiedad.frx":49E3
      Top             =   1500
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPasajeroCondicional 
      Height          =   480
      Left            =   4080
      Picture         =   "ViajeDetallePropiedad.frx":52AD
      Top             =   1500
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgComisionConfirmado 
      Height          =   480
      Left            =   4080
      Picture         =   "ViajeDetallePropiedad.frx":5B77
      Top             =   1500
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgComisionCancelado 
      Height          =   480
      Left            =   4080
      Picture         =   "ViajeDetallePropiedad.frx":6441
      Top             =   1500
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmViajeDetallePropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mViajeDetalle As ViajeDetalle
Private mNew As Boolean

Private mKeyDecimal As Boolean
Private mIDDestino As Long
Private mRutaDetalleIndiceMaximo As Long

Private mEsRutaEspecial As Boolean
Private mEsRutaPaquete As Boolean

Private mTramo1_Tramo2_IDLugar As Long

Private mMedioPago As MedioPago

Private mCPagosToAdd As Collection
Private mCPagosToUpdate As Collection
Private mCPagosToDelete As Collection

Private mTramo2_IDRuta As String
Private mRutaDetalleOrigen As RutaDetalle
Private mRutaDetalleDestino As RutaDetalle

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef ViajeDetalle As ViajeDetalle)
    Dim Persona As Persona
    Dim Feriado As Feriado
    Dim CuentaCorrienteCaja As CuentaCorrienteCaja
    Dim ShowSaldo As Boolean
    
    Set mViajeDetalle = ViajeDetalle
    Set ViajeDetalle = Nothing
    mNew = (mViajeDetalle.Indice = 0)
    
    Set mCPagosToAdd = New Collection
    Set mCPagosToUpdate = New Collection
    Set mCPagosToDelete = New Collection
    
    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    cmdAuditoria.Visible = (Not mNew)
    
    optTipoComision.Enabled = mNew
    optTipoPasajero.Enabled = mNew
    
    EnableControls True
    
    chkForzarDebito.Enabled = pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_INASISTENCIA_NODEBITAR, False)

    With mViajeDetalle
        dtpFecha.value = .FechaHora_FormattedAsDate
        dtpFecha_Change
        datcboHora.BoundText = Format(.FechaHora, "HH:nn")

        datcboRuta.BoundText = .IDRuta
        
        Dim permiteModificar As Boolean
        If mNew Then
            permiteModificar = pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_ADD_FECHAHORARUTA, False)
        Else
            permiteModificar = pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_MODIFY_FECHAHORARUTA, False)
        End If
        cmdAnterior.Visible = permiteModificar
        dtpFecha.Enabled = permiteModificar
        cmdSiguiente.Visible = permiteModificar
        cmdHoy.Visible = permiteModificar
        datcboHora.Enabled = permiteModificar
        datcboRuta.Enabled = permiteModificar
        cmdRuta.Visible = permiteModificar
        
        If mNew Then
            optTipoPasajero.value = True
        Else
            Select Case .OcupanteTipo
                Case OCUPANTE_TIPO_COMISION
                    optTipoComision.value = True
                Case OCUPANTE_TIPO_PASAJERO
                    optTipoPasajero.value = True
            End Select
        End If
        txtOrden.Text = IIf(.Orden = 0, "", .Orden)
        txtPersona.Tag = .IDPersona
        
        If mNew Then
            txtPersona.Text = ""
        Else
            Set Persona = New Persona
            Persona.IDPersona = .IDPersona
            If Persona.Load() Then
                txtPersona.Text = Persona.ApellidoNombre
            End If
            Set Persona = Nothing
        End If
        
        If .OcupanteTipo = OCUPANTE_TIPO_COMISION Then
            txtPersonaRecibe.Tag = .IDPersonaRecibe
            If .IDPersonaRecibe = 0 Then
                txtPersonaRecibe.Text = .Recibe
            Else
                Set Persona = New Persona
                Persona.IDPersona = .IDPersonaRecibe
                If Persona.Load() Then
                    txtPersonaRecibe.Text = Persona.ApellidoNombre
                End If
            End If
            optPagaEnvia.value = Not .PagaQuienRecibe
            optPagaRecibe.value = .PagaQuienRecibe
            
            txtDescripcion.Text = .Descripcion
            txtDomicilio.Text = .Domicilio
            txtHorario.Text = .Horario
            txtTelefono.Text = .Telefono
            cboDejarTraer.ListIndex = IIf(.DejarTraer = "", 0, IIf(.DejarTraer = "D", 1, 2))
        Else
            txtPersonaRecibe.Text = ""
            optPagaEnvia.value = True
            optPagaRecibe.value = False
            txtDescripcion.Text = ""
            txtDomicilio.Text = ""
            txtHorario.Text = ""
            txtTelefono.Text = ""
            cboDejarTraer.ListIndex = 0
        End If
        
        If Not mNew Then
            SetLastPersona .IDPersona, txtPersona.Text
        End If
        
        If mNew Then
            If Not CSM_Control_DataCombo.FillFromSQL(datcboListaPrecio, "SELECT IDListaPrecio, Nombre FROM ListaPrecio WHERE Activo = 1" & IIf(pCPermiso.ListaPrecioWhere <> "", " AND " & Replace(pCPermiso.ListaPrecioWhere, "%TABLENAME%", "ListaPrecio"), "") & " ORDER BY Nombre", "IDListaPrecio", "Nombre", "Listas de Precios", cscpItemOrFirst, .IDListaPrecio) Then
                Unload Me
                Exit Sub
            End If
        Else
            If Not CSM_Control_DataCombo.FillFromSQL(datcboListaPrecio, "SELECT IDListaPrecio, Nombre FROM ListaPrecio WHERE Activo = 1 OR IDListaPrecio = " & .IDListaPrecio & " ORDER BY Nombre", "IDListaPrecio", "Nombre", "Listas de Precios", cscpItemOrFirst, .IDListaPrecio) Then
                Unload Me
                Exit Sub
            End If
        End If
                
        If mNew Then
            If mEsRutaPaquete Then
                txtImporte.Text = .Viaje.Importe_Formatted
            End If
        Else
            datcboOrigen.BoundText = .IDOrigen
            txtSube.Text = .Sube
            datcboDestino.BoundText = .IDDestino
            txtBaja.Text = .Baja
            mIDDestino = .IDDestino
            If .OcupanteTipo = OCUPANTE_TIPO_COMISION Then
                txtValorDeclarado.Text = .ValorDeclarado_Formatted
                txtImporteSeguro.Text = .ImporteSeguro_Formatted
                txtImporteTotal.Text = Format(.ImporteSeguro + .Importe, "Currency")
            Else
                txtValorDeclarado.Text = ""
                txtImporteSeguro.Text = ""
                txtImporteTotal.Text = .Importe_Formatted
            End If
            txtImporte.Text = .Importe_Formatted
        End If
        
        Set mMedioPago = New MedioPago
        txtImporteContado.Text = .ImporteContado_Formatted
        txtImporteContado_LostFocus
        
        'MEDIO DE PAGO
        If Not CSM_Control_DataCombo.FillFromSQL(datcboMedioPago, "SELECT IDMedioPago, Nombre FROM MedioPago WHERE Activo = 1 OR IDMedioPago = " & .IDMedioPago & " ORDER BY Nombre", "IDMedioPago", "Nombre", "Medios de Pago", cscpItemOrFirst, IIf(.IDMedioPago = 0, pParametro.MedioPago_Predeterminado_ID, .IDMedioPago)) Then
            Unload Me
            Exit Sub
        End If
        cboCuotas.ListIndex = CSM_Control_ComboBox.GetListIndexByText(cboCuotas, .Cuotas_Formatted, cscpItemOrFirst)
        txtOperacion.Text = .Operacion
        
        txtImporteCuentaCorriente.Text = .ImporteCuentaCorriente_Formatted
        
        If pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_CAJA_VIEW_ALL, False) Then
            If Not CSM_Control_DataCombo.FillFromSQL(datcboCuentaCorrienteCaja, "SELECT IDCuentaCorrienteCaja, Nombre FROM CuentaCorrienteCaja WHERE Activo = 1 OR IDCuentaCorrienteCaja = " & .IDCuentaCorrienteCaja & " ORDER BY Nombre", "IDCuentaCorrienteCaja", "Nombre", "Cajas", cscpItemOrFirstIfUnique, .IDCuentaCorrienteCaja) Then
                Unload Me
                Exit Sub
            End If
        Else
            If Not CSM_Control_DataCombo.FillFromSQL(datcboCuentaCorrienteCaja, "SELECT CuentaCorrienteCaja.IDCuentaCorrienteCaja, CuentaCorrienteCaja.Nombre FROM CuentaCorrienteCaja LEFT JOIN Persona ON CuentaCorrienteCaja.IDPersona = Persona.IDPersona WHERE (CuentaCorrienteCaja.Activo = 1 OR CuentaCorrienteCaja.IDCuentaCorrienteCaja = " & .IDCuentaCorrienteCaja & ") AND (CuentaCorrienteCaja.MostrarSiempre = 1 OR CuentaCorrienteCaja.IDCuentaCorrienteCaja = " & pUsuario.IDCuentaCorrienteCaja & " OR Persona.IDPersona = " & .Viaje.IDConductor & " OR Persona.IDPersona = " & .Viaje.IDConductor2 & " OR CuentaCorrienteCaja.IDCuentaCorrienteCaja = " & .IDCuentaCorrienteCaja & ") ORDER BY CuentaCorrienteCaja.Nombre", "IDCuentaCorrienteCaja", "Nombre", "Cajas", cscpItemOrFirstIfUnique, IIf(.IDCuentaCorrienteCaja = 0, pUsuario.IDCuentaCorrienteCaja, .IDCuentaCorrienteCaja)) Then
                Unload Me
                Exit Sub
            End If
        End If
        
        If Not mNew Then
            If .ImporteContado <> 0 Then
                If .Viaje.IDConductor <> 0 Then
                    Set CuentaCorrienteCaja = New CuentaCorrienteCaja
                    CuentaCorrienteCaja.IDPersona = .Viaje.IDConductor
                    Call CuentaCorrienteCaja.LoadByPersona
                    If .IDCuentaCorrienteCaja <> pUsuario.IDCuentaCorrienteCaja And .IDCuentaCorrienteCaja <> CuentaCorrienteCaja.IDCuentaCorrienteCaja Then
                        If Not pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_MODIFY_IMPORTE_OTRO_USUARIO, False) Then
                            txtImporteContado.Enabled = False
                            datcboMedioPago.Enabled = False
                            cboCuotas.Enabled = False
                            txtOperacion.Enabled = False
                        End If
                        If Not pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_MODIFY_CAJA_OTRO_USUARIO, False) Then
                            datcboCuentaCorrienteCaja.Enabled = False
                        End If
                    End If
                    Set CuentaCorrienteCaja = Nothing
                Else
                    If .IDCuentaCorrienteCaja <> pUsuario.IDCuentaCorrienteCaja Then
                        If Not pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_MODIFY_IMPORTE_OTRO_USUARIO, False) Then
                            txtImporteContado.Enabled = False
                            datcboMedioPago.Enabled = False
                            cboCuotas.Enabled = False
                            txtOperacion.Enabled = False
                        End If
                        If Not pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_MODIFY_CAJA_OTRO_USUARIO, False) Then
                            datcboCuentaCorrienteCaja.Enabled = False
                        End If
                    End If
                End If
            End If
        End If
        
        If pParametro.ViajeDetalle_Paquete_Permite_Multiples_Pagos And pParametro.Ruta_Paquete_ID <> "" And mEsRutaPaquete Then
            Call .LoadPagos
            Call FillListView_MultiplesPagos
        End If
        
        cboRealizado.ListIndex = .Realizado
        
        chkForzarDebito.value = IIf(.ForzarDebito, vbChecked, vbUnchecked)
        txtCanceladoPor.Text = .CanceladoPor
        
        chkEntregada.value = IIf(.Entregada, vbChecked, vbUnchecked)
        If .Entregada Then
            dtpEntregadaFecha.value = .EntregadaFechaHora
            dtpEntregadaHora.value = .EntregadaFechaHora
            txtRetira.Text = .Retira
        Else
            dtpEntregadaFecha.value = Date
            dtpEntregadaHora.value = Time
            txtRetira.Text = ""
        End If
        
        txtPersonaCuentaCorriente.Tag = .IDPersonaCuentaCorriente
        Set Persona = New Persona
        If .IDPersonaCuentaCorriente = 0 Then
            txtPersonaCuentaCorriente.Text = ""
            
            Persona.IDPersona = .IDPersona
            If Not mNew Then
                Call Persona.Load
            End If
        Else
            Persona.IDPersona = .IDPersonaCuentaCorriente
            If Persona.Load() Then
                txtPersonaCuentaCorriente.Text = Persona.ApellidoNombre
            End If
        End If
        Persona.ViajeActual_FechaHora = .FechaHora
        Persona.ViajeActual_IDRuta = .IDRuta
        Persona.ViajeActual_Indice = .Indice
        Persona.LoadSaldoActual
        txtSaldoActual.Tag = Persona.SaldoActual
        chkImprimirSaldo.value = IIf(.ImprimirSaldo, vbChecked, vbUnchecked)
        Call CalcularImporteCuentaCorriente
        
        'AVISA QUE DEBE
        If (Not mNew) And Persona.SaldoActual < 0 And Not Persona.PermiteViajarSinPagar Then
            Load frmPersonaSaldo
            frmPersonaSaldo.txtPersona.Tag = Persona.IDPersona
            frmPersonaSaldo.txtPersona.Text = " " & Persona.ApellidoNombre
            frmPersonaSaldo.txtSaldo.Text = Persona.SaldoActual_Formatted
            frmPersonaSaldo.FillListView
            frmPersonaSaldo.Show
            ShowSaldo = True
        End If
        
        Set Persona = Nothing
        
        '//////////////////////////////////////////////////////////
        'FACTURACION
        chkFacturar.value = IIf(.Facturar, vbChecked, vbUnchecked)
        txtFacturarNotas.Text = .FacturarNotas
        txtFacturaNumero.Text = .FacturaNumero
        
        imgPasajeroConfirmado.Visible = (.OcupanteTipo = OCUPANTE_TIPO_PASAJERO And .Estado = VIAJE_DETALLE_ESTADO_CONFIRMADO)
        imgPasajeroCondicional.Visible = (.OcupanteTipo = OCUPANTE_TIPO_PASAJERO And .Estado = VIAJE_DETALLE_ESTADO_CONDICIONAL)
        imgPasajeroCancelado.Visible = (.OcupanteTipo = OCUPANTE_TIPO_PASAJERO And .Estado = VIAJE_DETALLE_ESTADO_CANCELADO)
        
        imgComisionConfirmado.Visible = (.OcupanteTipo = OCUPANTE_TIPO_COMISION And .Estado = VIAJE_DETALLE_ESTADO_CONFIRMADO)
        imgComisionCancelado.Visible = (.OcupanteTipo = OCUPANTE_TIPO_COMISION And .Estado = VIAJE_DETALLE_ESTADO_CANCELADO)
        
        lblAsiento.Visible = (.Estado = VIAJE_DETALLE_ESTADO_CONFIRMADO)
        txtAsiento.Visible = (.Estado = VIAJE_DETALLE_ESTADO_CONFIRMADO)
        txtAsiento.Text = IIf(.Asiento = -1, "", .Asiento)
        txtNotas.Text = .Notas
        txtReservadoPor.Text = .ReservadoPor
        
        If mNew Then
            .ReservaTipo = VIAJE_DETALLE_RESERVA_TIPO_STANDARD
        End If
    
        Set Feriado = New Feriado
        Feriado.VerificarReservasDelPasajero .IDPersona
        Set Feriado = Nothing
        
        If .Viaje.Estado = VIAJE_ESTADO_FINALIZADO Then
            If Not pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_FINALIZADO_MODIFY, False) Then
                EnableControls False
                If .OcupanteTipo = OCUPANTE_TIPO_COMISION And Not .Entregada Then
                    If pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_COMISION_MODIFY_IMPORTE_PAGOCONTADO_FINALIZADO, False) Then
                        optPagaEnvia.Enabled = True
                        optPagaRecibe.Enabled = True
                        txtImporte.Enabled = True
                        txtImporteContado.Enabled = True
                        datcboMedioPago.Enabled = True
                        cboCuotas.Enabled = True
                        txtOperacion.Enabled = True
                        datcboCuentaCorrienteCaja.Enabled = True
                        cmdOK.Visible = True
                        cmdCancel.Caption = "Cancelar"
                    End If
                End If
                If .OcupanteTipo = OCUPANTE_TIPO_PASAJERO Then
                    If pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_FINALIZADO_MODIFY_MEDIOPAGO_DEBITO, False) Then
                        txtImporteContado.Enabled = True
                        datcboMedioPago.Enabled = True
                        cboCuotas.Enabled = True
                        txtOperacion.Enabled = True
                        datcboCuentaCorrienteCaja.Enabled = True
                        cmdPersonaCuentaCorriente.Visible = True
                        cmdPersonaCuentaCorrienteClear.Visible = True
                        cmdOK.Visible = True
                        cmdCancel.Caption = "Cancelar"
                    End If
                End If
            End If
        End If
    End With
    
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
    
    If ShowSaldo Then
        frmPersonaSaldo.SetFocus
    End If
End Sub

Private Sub FillListView_MultiplesPagos()
    Dim Pago As CuentaCorriente
    Dim ListItem As MSComctlLib.ListItem
    
    Dim Importe As Currency
    Dim TotalPagos As Currency
    
    lvwPagos.ListItems.Clear
    TotalPagos = 0
    
    For Each Pago In mViajeDetalle.Pagos
        Set ListItem = lvwPagos.ListItems.Add()
        ListItem.SubItems(1) = Pago.FechaHora_Formatted
        ListItem.SubItems(2) = Pago.Importe_Formatted
        ListItem.SubItems(3) = Pago.MedioPago.Nombre
        ListItem.SubItems(4) = Pago.CuentaCorrienteCaja.Nombre
        TotalPagos = TotalPagos + Pago.Importe
    Next Pago
    
    If IsNumeric(txtImporte.Text) Then
        Importe = CCur(txtImporte.Text)
    End If
    txtPagosTotal.Text = Format(TotalPagos, "Currency")
    txtPagosSaldo.Text = Format(Importe - TotalPagos, "Currency")
End Sub

Private Sub Form_Load()
    dtpFecha.CalendarTitleBackColor = vbDesktop
    dtpEntregadaFecha.CalendarTitleBackColor = vbDesktop
    
    cboDejarTraer.AddItem "--"
    cboDejarTraer.AddItem "Dejar"
    cboDejarTraer.AddItem "Traer"
    
    cboRealizado.AddItem "--"
    cboRealizado.AddItem "Sí"
    cboRealizado.AddItem "No"
    
    Dim SubeBaja As Boolean
    SubeBaja = pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_SUBEBAJA_MOSTRAR, False)
    lblSube.Visible = SubeBaja
    txtSube.Visible = SubeBaja
    lblBaja.Visible = SubeBaja
    txtBaja.Visible = SubeBaja
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CSM_Forms.IsLoaded("frmPersonaSaldo") Then
        Unload frmPersonaSaldo
    End If
    Set mRutaDetalleOrigen = Nothing
    Set mRutaDetalleDestino = Nothing
    Set mCPagosToAdd = Nothing
    Set mCPagosToUpdate = Nothing
    Set mCPagosToDelete = Nothing
    Set mViajeDetalle = Nothing
    Set mMedioPago = Nothing
End Sub

Private Sub dtpFecha_Change()
    txtDiaSemana.Text = WeekdayName(Weekday(dtpFecha.value))
    Call CSM_Control_DataCombo.FillFromSQL(datcboHora, "SELECT DISTINCT convert(char(5), FechaHora, 108) AS Hora FROM Viaje WHERE convert(char(10), FechaHora, 111) = '" & Format(dtpFecha.value, "yyyy/mm/dd") & "' ORDER BY convert(char(5), FechaHora, 108)", "Hora", "Hora", "Horas", cscpItemOrFirst, datcboHora.BoundText)
End Sub

Private Sub datcboHora_Change()
    Call CSM_Control_DataCombo.FillFromSQL(datcboRuta, "SELECT RTRIM(IDRuta) AS IDRuta, Nombre = RTRIM(IDRuta) + CASE IDRuta WHEN '" & ReplaceQuote(pParametro.Ruta_ID_Otra) & "' THEN ': ' + RutaOtra WHEN '" & ReplaceQuote(pParametro.Ruta_Paquete_ID) & "' THEN ': ' + RutaOtra ELSE '' END FROM Viaje WHERE convert(char(10), FechaHora, 111) = '" & Format(dtpFecha.value, "yyyy/mm/dd") & "' AND convert(char(5), FechaHora, 108) = '" & datcboHora.Text & "'" & IIf(pCPermiso.RutaWhere <> "", " AND " & Replace(pCPermiso.RutaWhere, "%TABLENAME%", "Viaje"), "") & " ORDER BY IDRuta", "IDRuta", "Nombre", "Rutas", cscpItemOrFirst, datcboRuta.BoundText)
End Sub

Private Sub cboRealizado_Click()
    lblForzarDebito.Visible = (optTipoPasajero.value And cboRealizado.ListIndex = VIAJE_DETALLE_REALIZADO_NO)
    chkForzarDebito.Visible = (optTipoPasajero.value And cboRealizado.ListIndex = VIAJE_DETALLE_REALIZADO_NO)
    If cboRealizado.ListIndex = VIAJE_DETALLE_REALIZADO_NO Then
        chkForzarDebito.value = vbChecked
    End If
End Sub

Private Sub chkEntregada_Click()
    dtpEntregadaFecha.Visible = (optTipoComision.value And chkEntregada.value = vbChecked)
    dtpEntregadaHora.Visible = (optTipoComision.value And chkEntregada.value = vbChecked)
    lblRetira.Visible = (optTipoComision.value And chkEntregada.value = vbChecked)
    txtRetira.Visible = (optTipoComision.value And chkEntregada.value = vbChecked)
End Sub

Private Sub cmdAnterior_Click()
    dtpFecha.value = DateAdd("d", -1, dtpFecha.value)
    dtpFecha.SetFocus
    dtpFecha_Change
End Sub

Private Sub cmdAuditoria_Click()
    mViajeDetalle.ForzarDebito = (chkForzarDebito.value = vbChecked)
    mViajeDetalle.CanceladoPor = txtCanceladoPor.Text
    
    frmViajeDetallePropiedadAuditoria.LoadDataAndShow mViajeDetalle

    chkForzarDebito.value = IIf(mViajeDetalle.ForzarDebito, vbChecked, vbUnchecked)
    txtCanceladoPor.Text = mViajeDetalle.CanceladoPor
End Sub

Private Sub imgInformacion_Click()
    With frmViajeDetallePropiedadInformacion
        .txtIDViaje.Text = mViajeDetalle.IDViaje_Formatted
        .txtIDViajeDetalle.Text = mViajeDetalle.IDViajeDetalle_Formatted
        .txtFechaHora.Text = mViajeDetalle.FechaHora_Formatted
        .txtIDRuta.Text = mViajeDetalle.IDRuta
        .txtIndice.Text = mViajeDetalle.Indice
        .txtReservaCodigo.Text = mViajeDetalle.ReservaCodigo
        .txtIDPersona.Text = mViajeDetalle.IDPersona_Formatted
        
        .Show vbModal, frmMDI
    End With
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

Private Sub CalcularImporteCuentaCorriente()
    Dim ImporteRestante As Currency
    
    On Error Resume Next
    
    If CCur(txtImporteContado.Text) >= CCur(txtImporte.Text) Then
        ImporteRestante = 0
    Else
        ImporteRestante = CCur(txtImporte.Text) - CCur(txtImporteContado.Text)
    End If
    
    If CCur(txtSaldoActual.Tag) > 0 And CCur(txtImporte.Text) > 0 Then
        If CCur(txtSaldoActual.Tag) >= ImporteRestante Then
            txtImporteCuentaCorriente.Text = Format(ImporteRestante, "Currency")
        Else
            txtImporteCuentaCorriente.Text = Format(CCur(txtSaldoActual.Tag), "Currency")
        End If
        txtSaldoActual.Text = Format(CCur(txtSaldoActual.Tag) - CCur(txtImporteCuentaCorriente.Text), "Currency")
    Else
        txtImporteCuentaCorriente.Text = Format(0, "Currency")
        txtSaldoActual.Text = Format(CCur(txtSaldoActual.Tag), "Currency")
    End If
End Sub

Private Sub cmdPagoAgregar_Click()
    Dim Pago As CuentaCorriente
    
    Set Pago = New CuentaCorriente
    frmViajeDetallePropiedadPago.LoadData mViajeDetalle, Pago
    frmViajeDetallePropiedadPago.Show vbModal, frmMDI
    If frmViajeDetallePropiedadPago.Tag = "OK" Then
        'AGREGO EL PAGO EN LA COLECCION DE PAGOS A CREAR EN LA BASE DE DATOS
        Pago_Add Pago
        'AGREGO EL PAGO A LA COLECCION DE PAGOS DEL OBJETO VIAJE DETALLE PARA QUE APAREZCA EN LA GRILLA
        mViajeDetalle.Pago_Add Pago
        FillListView_MultiplesPagos
    End If
    Set Pago = Nothing
    Unload frmViajeDetallePropiedadPago
    Set frmViajeDetallePropiedadPago = Nothing
End Sub

Private Sub cmdPagoEditar_Click()
    Dim Pago As CuentaCorriente
    
    If lvwPagos.SelectedItem Is Nothing Then
        MsgBox "No hay ningún Pago seleccionado para modificar.", vbInformation, App.Title
        lvwPagos.SetFocus
        Exit Sub
    End If
    Set Pago = mViajeDetalle.Pagos(lvwPagos.SelectedItem.Index)
    frmViajeDetallePropiedadPago.LoadData mViajeDetalle, Pago
    frmViajeDetallePropiedadPago.Show vbModal, frmMDI
    If frmViajeDetallePropiedadPago.Tag = "OK" Then
        Pago_Update Pago
        FillListView_MultiplesPagos
    End If
    Set Pago = Nothing
    Unload frmViajeDetallePropiedadPago
    Set frmViajeDetallePropiedadPago = Nothing
End Sub

Private Sub cmdPagoEliminar_Click()
    Dim Pago As CuentaCorriente
    
    If lvwPagos.SelectedItem Is Nothing Then
        MsgBox "No hay ningún Pago seleccionado para eliminar.", vbInformation, App.Title
        lvwPagos.SetFocus
        Exit Sub
    End If
    Set Pago = mViajeDetalle.Pagos(lvwPagos.SelectedItem.Index)
    If MsgBox("¿Desea eliminar el pago?" & vbCr & vbCr & "Fecha/Hora: " & Pago.FechaHora_Formatted & vbCr & "Importe: " & Pago.Importe_Formatted & vbCr & "Medio de Pago: " & Pago.MedioPago.Nombre, vbExclamation + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
        'AGREGO EL PAGO A LA COLECCIÓN DE PAGOS A ELIMINAR DE LA BASE DE DATOS
        Pago_Delete Pago
        'ELIMINO EL PAGO DE LA COLECCION DE PAGOS QUE TIENE EL OBJETO DEL DETALLE DEL VIAJE PARA QUE NO APAREZCA EN LA GRILLA
        mViajeDetalle.Pago_Delete Pago
        FillListView_MultiplesPagos
    End If
    Set Pago = Nothing
End Sub

Private Sub cmdListaPrecio_Click()
    If pCPermiso.GotPermission(PERMISO_LISTA_PRECIO) Then
        Screen.MousePointer = vbHourglass
        frmListaPrecio.Show
        On Error Resume Next
        Set frmListaPrecio.lvwData.SelectedItem = frmListaPrecio.lvwData.ListItems(KEY_STRINGER & Val(datcboListaPrecio.BoundText))
        frmListaPrecio.lvwData.SelectedItem.EnsureVisible
        If frmListaPrecio.WindowState = vbMinimized Then
            frmListaPrecio.WindowState = vbNormal
        End If
        frmListaPrecio.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdPersonaCuentaCorriente_Click()
    If pCPermiso.GotPermission(PERMISO_PERSONA) Then
        Call frmPersona.FindAndShowItem(Val(txtPersonaCuentaCorriente.Tag), UCase(Left(txtPersonaCuentaCorriente.Text, 1)), Me.Name, ENTIDAD_TIPO_PERSONA_CLIENTE, "PF")
    End If
End Sub

Private Sub cmdPersonaCuentaCorrienteClear_Click()
    Dim Persona As Persona
    
    If Val(txtPersonaCuentaCorriente.Tag) <> 0 Then
        txtPersonaCuentaCorriente.Tag = 0
        txtPersonaCuentaCorriente.Text = ""
        
        Set Persona = New Persona
        Persona.IDPersona = Val(txtPersona.Tag)
        Persona.LoadSaldoActual
        txtSaldoActual.Text = Format(Persona.SaldoActual, "Currency")
        Set Persona = Nothing
    End If
    
    On Error Resume Next
    txtPersonaRecibe.SetFocus
End Sub

Private Sub cmdSaldoActual_Click()
    Dim Persona As Persona
    
    If Val(txtPersona.Tag) > 0 Then
        If Not pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE) Then
            Exit Sub
        End If
        
        Set Persona = New Persona
        If Val(txtPersonaCuentaCorriente.Tag) = 0 Then
            Persona.IDPersona = Val(txtPersona.Tag)
        Else
            Persona.IDPersona = Val(txtPersonaCuentaCorriente.Tag)
        End If
        
        If Persona.Load() Then
            Select Case Persona.EntidadTipo
                Case ENTIDAD_TIPO_PERSONA_CLIENTE
                Case ENTIDAD_TIPO_PERSONA_CONDUCTOR
                    If Not pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_CONDUCTOR_SELECT, False) Then
                        MsgBox "No puede ver los Movimientos de Personas de tipo Conductor.", vbExclamation, App.Title
                        Set Persona = Nothing
                        Exit Sub
                    End If
                Case ENTIDAD_TIPO_PERSONA_ADMINISTRATIVO
                    If Not pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_ADMINISTRATIVO_SELECT, False) Then
                        MsgBox "No puede ver los Movimientos de Personas de tipo Administrativo.", vbExclamation, App.Title
                        Set Persona = Nothing
                        Exit Sub
                    End If
            End Select
        End If
        Set Persona = Nothing
        
        Screen.MousePointer = vbHourglass
        Load frmCuentaCorriente
        frmCuentaCorriente.DatabaseName = pParametro.Database_Database
        frmCuentaCorriente.IsHistory = False
        frmCuentaCorriente.Caption = "Cuenta Corriente Actual"
        If Val(txtPersonaCuentaCorriente.Tag) = 0 Then
            frmCuentaCorriente.txtPersona.Tag = Val(txtPersona.Tag)
            frmCuentaCorriente.txtPersona.Text = txtPersona.Text
        Else
            frmCuentaCorriente.txtPersona.Tag = Val(txtPersonaCuentaCorriente.Tag)
            frmCuentaCorriente.txtPersona.Text = txtPersonaCuentaCorriente.Text
        End If
        frmCuentaCorriente.cboFecha.ListIndex = 2
        frmCuentaCorriente.dtpFechaDesde.value = DateAdd("d", -30, Date)
        frmCuentaCorriente.LoadDataAndShow
        On Error Resume Next
        If frmCuentaCorriente.WindowState = vbMinimized Then
            frmCuentaCorriente.WindowState = vbNormal
        End If
        frmCuentaCorriente.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdSiguiente_Click()
    dtpFecha.value = DateAdd("d", 1, dtpFecha.value)
    dtpFecha.SetFocus
    dtpFecha_Change
End Sub

Private Sub cmdPersonaUltimo_Click()
    If frmMDI.cboPersona.ListIndex > -1 Then
        PersonaSelected Val(frmMDI.cboPersona.ItemData(frmMDI.cboPersona.ListIndex)), "PP"
    End If
    cmdPersona.SetFocus
    If CSM_Forms.IsLoaded("frmPersonaSaldo") Then
        frmPersonaSaldo.SetFocus
    End If
End Sub

Private Sub cmdPersonaRecibeUltimo_Click()
    If frmMDI.cboPersona.ListIndex > -1 Then
        PersonaSelected Val(frmMDI.cboPersona.ItemData(frmMDI.cboPersona.ListIndex)), "PR"
    End If
    cmdPersonaRecibe.SetFocus
    If CSM_Forms.IsLoaded("frmPersonaSaldo") Then
        frmPersonaSaldo.SetFocus
    End If
End Sub

Private Sub cmdLugarDestino_Click()
    If pCPermiso.GotPermission(PERMISO_LUGAR) Then
        Screen.MousePointer = vbHourglass
        frmLugar.Show
        On Error Resume Next
        Set frmLugar.lvwData.SelectedItem = frmLugar.lvwData.ListItems(KEY_STRINGER & datcboDestino.BoundText)
        frmLugar.lvwData.SelectedItem.EnsureVisible
        If frmLugar.WindowState = vbMinimized Then
            frmLugar.WindowState = vbNormal
        End If
        frmLugar.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdLugarOrigen_Click()
    If pCPermiso.GotPermission(PERMISO_LUGAR) Then
        Screen.MousePointer = vbHourglass
        frmLugar.Show
        On Error Resume Next
        Set frmLugar.lvwData.SelectedItem = frmLugar.lvwData.ListItems(KEY_STRINGER & datcboOrigen.BoundText)
        frmLugar.lvwData.SelectedItem.EnsureVisible
        If frmLugar.WindowState = vbMinimized Then
            frmLugar.WindowState = vbNormal
        End If
        frmLugar.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdPersona_Click()
    If pCPermiso.GotPermission(PERMISO_PERSONA) Then
        Call frmPersona.FindAndShowItem(Val(txtPersona.Tag), UCase(Left(txtPersona.Text, 1)), Me.Name, "", "PP")
    End If
End Sub

Private Sub cmdPersonaRecibe_Click()
    If pCPermiso.GotPermission(PERMISO_PERSONA) Then
        Call frmPersona.FindAndShowItem(Val(txtPersonaRecibe.Tag), UCase(Left(txtPersonaRecibe.Text, 1)), Me.Name, "", "PR")
    End If
End Sub

Private Sub cmdVerificarAsiento_Click()
    Dim Viaje As Viaje
    Dim RutaConexion As RutaConexion
    
    Dim AsientoAsignado As Long
    Dim AsientoAsignadoCombinado As Long
    Dim Tramo1_Tramo2_IDLugar As Long
    
    If datcboHora.Text = "" Then
        MsgBox "Debe seleccionar la Hora.", vbInformation, App.Title
        datcboHora.SetFocus
        Exit Sub
    End If
    If Val(datcboOrigen.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Origen.", vbInformation, App.Title
        datcboOrigen.SetFocus
        Exit Sub
    End If
    If Val(datcboDestino.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Destino.", vbInformation, App.Title
        datcboDestino.SetFocus
        Exit Sub
    End If
    
    AsientoAsignado = -3
    AsientoAsignadoCombinado = -3
    
    If chkRutaConexion.Visible = False Or chkRutaConexion.value = vbUnchecked Then
        '//////////////////////////////////
        'RUTA SIMPLE
        '//////////////////////////////////
        Set Viaje = New Viaje
        Viaje.FechaHora = CDate(dtpFecha.value & " " & datcboHora.Text)
        Viaje.IDRuta = datcboRuta.BoundText
        Viaje.IDOrigen = Val(datcboOrigen.BoundText)
        Viaje.IDDestino = Val(datcboDestino.BoundText)
        AsientoAsignado = Viaje.Asiento_Asignar_GetAsiento(mViajeDetalle.Indice)
        Set Viaje = Nothing
        
        MsgBox VerificarAsientoMensaje(AsientoAsignado, ""), vbExclamation, App.Title
    Else
        '//////////////////////////////////
        'COMBINACION DE RUTAS
        '//////////////////////////////////
        If datcboRutaConexion.BoundText = "" Then
            MsgBox "Debe seleccionar la Conexión de Rutas.", vbInformation, App.Title
            datcboRutaConexion.SetFocus
            Exit Sub
        End If
        If datcboViajeConexion.BoundText = "" Then
            MsgBox "Debe seleccionar el Viaje para la Conexión.", vbInformation, App.Title
            datcboViajeConexion.SetFocus
            Exit Sub
        End If
        
        Set RutaConexion = New RutaConexion
        RutaConexion.IDRutaConexion = Val(datcboRutaConexion.BoundText)
        If RutaConexion.Load() Then
            Tramo1_Tramo2_IDLugar = RutaConexion.Tramo1_Tramo2_IDLugar
        End If
        Set RutaConexion = Nothing
        
        'PRIMERO VERIFICO LA RUTA SIMPLE
        Set Viaje = New Viaje
        Viaje.FechaHora = CDate(dtpFecha.value & " " & datcboHora.Text)
        Viaje.IDRuta = datcboRuta.BoundText
        Viaje.IDOrigen = Val(datcboOrigen.BoundText)
        Viaje.IDDestino = Tramo1_Tramo2_IDLugar
        AsientoAsignado = Viaje.Asiento_Asignar_GetAsiento(mViajeDetalle.Indice)
        Set Viaje = Nothing
        
        'AHORA BUSCO LA RUTA COMBINADA
        Set Viaje = New Viaje
        Viaje.FechaHora = CDate(Left(datcboViajeConexion.BoundText, 16))
        Viaje.IDRuta = Mid(datcboViajeConexion.BoundText, 18)
        Viaje.IDOrigen = Tramo1_Tramo2_IDLugar
        Viaje.IDDestino = Val(datcboDestino.BoundText)
        AsientoAsignadoCombinado = Viaje.Asiento_Asignar_GetAsiento(0)
        Set Viaje = Nothing
                
        MsgBox VerificarAsientoMensaje(AsientoAsignado, Format(CDate(dtpFecha.value & " " & datcboHora.Text), "hh:nn") & " - " & RTrim(datcboRuta.Text) & " --> ") & vbCr & vbCr & VerificarAsientoMensaje(AsientoAsignado, datcboViajeConexion.Text & " --> "), vbExclamation, App.Title
    End If
End Sub

Private Sub datcboListaPrecio_Change()
    CalcularImporte
End Sub

Private Sub datcboRuta_Change()
    Dim Viaje As Viaje
    Dim ViajeDetalle As ViajeDetalle
    Dim Ruta As Ruta
    Dim recRutaConexion As ADODB.Recordset
    
    ShowControls
    
    mEsRutaEspecial = (datcboRuta.BoundText = pParametro.Ruta_ID_Otra)
    mEsRutaPaquete = (datcboRuta.BoundText = pParametro.Ruta_Paquete_ID)
    
    If datcboRuta.BoundText <> "" Then
        Set Ruta = New Ruta
        Ruta.IDRuta = datcboRuta.BoundText
        If Not Ruta.GetStatistics(0, 0, mRutaDetalleIndiceMaximo) Then
            Set Ruta = Nothing
            Exit Sub
        End If
        Set Ruta = Nothing
    End If
    
    If mEsRutaEspecial Or mEsRutaPaquete Then
        Call CSM_Control_DataCombo.FillFromSQL(datcboOrigen, "SELECT IDLugar, Nombre FROM Lugar WHERE IDLugar = " & pParametro.Lugar_ID_Otro, "IDLugar", "Nombre", "Orígenes", cscpFirst)
        Call CSM_Control_DataCombo.FillFromSQL(datcboDestino, "SELECT IDLugar, Nombre FROM Lugar WHERE IDLugar = " & pParametro.Lugar_ID_Otro, "IDLugar", "Nombre", "Destinos", cscpFirst)
        Exit Sub
    End If
        
    If pParametro.Viaje_Permite_RutaConexion And mViajeDetalle.IsNew Then
        If datcboRuta.BoundText <> "" Then
            Call CSM_Control_DataCombo.FillFromSQL(datcboRutaConexion, "SELECT DISTINCT RutaConexion.IDRutaConexion, RutaConexion.Nombre FROM RutaConexion INNER JOIN RutaConexionDetalle ON RutaConexion.IDRutaConexion = RutaConexionDetalle.IDRutaConexion WHERE RutaConexion.Activo = 1 AND RutaConexionDetalle.Tramo1_IDRuta = '" & datcboRuta.Text & "'", "IDRutaConexion", "Nombre", "Conexiones de Ruta", cscpCurrentOrFirst)
            Set recRutaConexion = datcboRutaConexion.RowSource
            If recRutaConexion.RecordCount > 0 Then
                chkRutaConexion.Visible = True
                datcboRutaConexion.Visible = (chkRutaConexion.Visible And chkRutaConexion.value = vbChecked)
                lblViajeConexion.Visible = (chkRutaConexion.Visible And chkRutaConexion.value = vbChecked And datcboRutaConexion.BoundText <> "")
                datcboViajeConexion.Visible = (chkRutaConexion.Visible And chkRutaConexion.value = vbChecked And datcboRutaConexion.BoundText <> "")
            Else
                chkRutaConexion.Visible = False
                datcboRutaConexion.Visible = False
                lblViajeConexion.Visible = False
                datcboViajeConexion.Visible = False
            End If
        Else
            Set datcboRutaConexion.RowSource = Nothing
            chkRutaConexion.Visible = False
            datcboRutaConexion.Visible = False
            lblViajeConexion.Visible = False
            datcboViajeConexion.Visible = False
        End If
    End If
    
    Call CSM_Control_DataCombo.FillFromSQL(datcboOrigen, "SELECT RutaDetalle.IDLugar, Lugar.Nombre FROM RutaDetalle INNER JOIN Lugar ON RutaDetalle.IDLugar = Lugar.IDLugar WHERE RutaDetalle.IDRuta = '" & ReplaceQuote(datcboRuta.BoundText) & "' AND RutaDetalle.Indice < " & mRutaDetalleIndiceMaximo & " AND (Lugar.Activo = 1 OR Lugar.IDLugar = " & mViajeDetalle.IDOrigen & ") ORDER BY RutaDetalle.Indice, Lugar.Nombre", "IDLugar", "Nombre", "Orígenes", cscpFirst)
    Call datcboOrigen_Change
    Call CSM_Control_DataCombo.FindItem(datcboDestino, "Destinos", cscpLast)
    
    If datcboRuta.BoundText <> "" And Val(txtPersona.Tag) > 0 Then
        Set ViajeDetalle = New ViajeDetalle
        ViajeDetalle.FechaHora = CDate(dtpFecha.value & " " & datcboHora.BoundText)
        ViajeDetalle.IDPersona = Val(txtPersona.Tag)
        ViajeDetalle.IDRuta = datcboRuta.BoundText
        ViajeDetalle.OcupanteTipo = IIf(optTipoComision.value, OCUPANTE_TIPO_COMISION, OCUPANTE_TIPO_PASAJERO)
        Set Viaje = New Viaje
        Viaje.FechaHora = CDate(dtpFecha.value & " " & datcboHora.BoundText)
        Viaje.IDRuta = datcboRuta.BoundText
        If Viaje.Load() Then
            If Not ViajeDetalle.GetNewValues(Viaje.DiaSemanaBase) Then
                Exit Sub
            End If
        End If
        Set Viaje = Nothing
        txtPersonaCuentaCorriente.Tag = ViajeDetalle.IDPersonaCuentaCorriente
        datcboListaPrecio.BoundText = ViajeDetalle.IDListaPrecio
        datcboOrigen.BoundText = ViajeDetalle.IDOrigen
        datcboDestino.BoundText = ViajeDetalle.IDDestino
        txtImporte.Text = Format(ViajeDetalle.Importe, "Currency")
        Set ViajeDetalle = Nothing
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

Private Sub chkRutaConexion_Click()
    datcboRutaConexion.Visible = (chkRutaConexion.Visible And chkRutaConexion.value = vbChecked)
    lblViajeConexion.Visible = (chkRutaConexion.Visible And chkRutaConexion.value = vbChecked And datcboRutaConexion.BoundText <> "")
    datcboViajeConexion.Visible = (chkRutaConexion.Visible And chkRutaConexion.value = vbChecked And datcboRutaConexion.BoundText <> "")
    datcboOrigen_Change
    Call FillComboBoxViajes
End Sub

Private Sub datcboRutaConexion_Change()
    datcboOrigen_Change
    
    lblViajeConexion.Visible = (datcboRutaConexion.BoundText <> "")
    datcboViajeConexion.Visible = (datcboRutaConexion.BoundText <> "")
    Call FillComboBoxViajes
End Sub

Private Sub datcboOrigen_Change()
    Dim RutaConexion As RutaConexion
    Dim RutaConexionDetalle As RutaConexionDetalle
    
    If mEsRutaEspecial Or mEsRutaPaquete Then
        Exit Sub
    End If
    
    If Val(datcboOrigen.BoundText) = 0 Then
        Exit Sub
    End If
    
    If chkRutaConexion.Visible And chkRutaConexion.value = vbChecked Then
        'Se utiliza una conexión de Ruta, por lo tanto se llenará el ComboBox de Destino a partir de la Conexión
        Set RutaConexion = New RutaConexion
        RutaConexion.IDRutaConexion = Val(datcboRutaConexion.BoundText)
        If RutaConexion.Load() Then
            mTramo1_Tramo2_IDLugar = RutaConexion.Tramo1_Tramo2_IDLugar
        End If
        Set RutaConexion = Nothing
        
        Set RutaConexionDetalle = New RutaConexionDetalle
        RutaConexionDetalle.IDRutaConexion = Val(datcboRutaConexion.BoundText)
        RutaConexionDetalle.Tramo1_IDRuta = datcboRuta.BoundText
        If RutaConexionDetalle.LoadFirstWhereTramo1() Then
            mTramo2_IDRuta = RutaConexionDetalle.Tramo2_IDRuta
        End If
        Set RutaConexionDetalle = Nothing
        
        Set mRutaDetalleOrigen = New RutaDetalle
        mRutaDetalleOrigen.IDRuta = mTramo2_IDRuta
        mRutaDetalleOrigen.IDLugar = mTramo1_Tramo2_IDLugar
        If mRutaDetalleOrigen.Load() Then
        End If
        
        Call CSM_Control_DataCombo.FillFromSQL(datcboDestino, "SELECT RutaDetalle.IDLugar, Lugar.Nombre FROM RutaDetalle INNER JOIN Lugar ON RutaDetalle.IDLugar = Lugar.IDLugar WHERE RutaDetalle.IDRuta = '" & ReplaceQuote(mTramo2_IDRuta) & "' AND RutaDetalle.Indice > " & mRutaDetalleOrigen.Indice & " AND (Lugar.Activo = 1 OR Lugar.IDLugar = " & mIDDestino & ") ORDER BY RutaDetalle.Indice, Lugar.Nombre", "IDLugar", "Nombre", "Destinos", cscpCurrentOrLast)
    Else
        mTramo1_Tramo2_IDLugar = 0
        
        'Busco el Detalle de la Ruta para filtrar el ComboBox de Destino a partir del Origen
        Set mRutaDetalleOrigen = New RutaDetalle
        mRutaDetalleOrigen.IDRuta = datcboRuta.BoundText
        mRutaDetalleOrigen.IDLugar = Val(datcboOrigen.BoundText)
        mRutaDetalleOrigen.NoMatchRaiseError = False
        If mRutaDetalleOrigen.Load() Then
        End If
        
        Call CSM_Control_DataCombo.FillFromSQL(datcboDestino, "SELECT RutaDetalle.IDLugar, Lugar.Nombre FROM RutaDetalle INNER JOIN Lugar ON RutaDetalle.IDLugar = Lugar.IDLugar WHERE RutaDetalle.IDRuta = '" & ReplaceQuote(datcboRuta.BoundText) & "' AND RutaDetalle.Indice > " & mRutaDetalleOrigen.Indice & " AND (Lugar.Activo = 1 OR Lugar.IDLugar = " & mIDDestino & ") ORDER BY RutaDetalle.Indice, Lugar.Nombre", "IDLugar", "Nombre", "Destinos", cscpCurrentOrLast)
    End If
    
    CalcularImporte
End Sub

Private Sub txtSube_GotFocus()
    CSM_Control_TextBox.SelAllText txtSube
End Sub

Private Sub datcboDestino_Change()
    If mEsRutaEspecial Or mEsRutaPaquete Then
        Exit Sub
    End If
    
    If Val(datcboDestino.BoundText) = 0 Then
        Exit Sub
    End If
    
    If chkRutaConexion.Visible And chkRutaConexion.value = vbChecked Then
        Set mRutaDetalleDestino = New RutaDetalle
        mRutaDetalleDestino.IDRuta = mTramo2_IDRuta
        mRutaDetalleDestino.IDLugar = Val(datcboDestino.BoundText)
        mRutaDetalleOrigen.NoMatchRaiseError = False
        If mRutaDetalleDestino.Load() Then
        End If
    Else
        Set mRutaDetalleDestino = New RutaDetalle
        mRutaDetalleDestino.IDRuta = datcboRuta.BoundText
        mRutaDetalleDestino.IDLugar = Val(datcboDestino.BoundText)
        mRutaDetalleDestino.NoMatchRaiseError = False
        If mRutaDetalleDestino.Load() Then
        End If
    End If
    
    CalcularImporte
End Sub

Private Sub txtBaja_GotFocus()
    CSM_Control_TextBox.SelAllText txtBaja
End Sub

Private Sub optPagaEnvia_Click()
    Call CambioPersona
End Sub

Private Sub optPagaRecibe_Click()
    Call CambioPersona
End Sub

Private Sub optTipoComision_Click()
    ShowControls
    SetCaption
    CalcularImporte
    If pParametro.Comision_Seguro_ValorDeclaradoMinimo_Preasignar Then
        If txtValorDeclarado.Text = "" Then
            txtValorDeclarado.Text = Format(pParametro.Comision_Seguro_ValorDeclaradoMinimo, "Currency")
        End If
    Else
        txtValorDeclarado.Text = Format(0, "Currency")
    End If
End Sub

Private Sub optTipoPasajero_Click()
    ShowControls
    SetCaption
    CalcularImporte
End Sub

Private Sub txtDiaSemana_GotFocus()
    On Error Resume Next
    dtpFecha.SetFocus
End Sub

Private Sub txtFacturarNotas_GotFocus()
    CSM_Control_TextBox.SelAllText txtFacturarNotas
End Sub

Private Sub txtFacturaNumero_GotFocus()
    CSM_Control_TextBox.SelAllText txtFacturaNumero
End Sub

Private Sub txtRetira_GotFocus()
    CSM_Control_TextBox.SelAllText txtRetira
End Sub

Private Sub txtRetira_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRetira_LostFocus()
    txtRetira.Text = UCase(txtRetira.Text)
    txtRetira.Text = CleanInvalidSpaces(txtRetira.Text)
End Sub

Private Sub txtDescripcion_GotFocus()
    CSM_Control_TextBox.SelAllText txtDescripcion
End Sub

Private Sub txtDescripcion_LostFocus()
    txtDescripcion.Text = CleanInvalidSpaces(txtDescripcion.Text)
End Sub

Private Sub txtDomicilio_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilio
End Sub

Private Sub txtDomicilio_LostFocus()
    txtDomicilio.Text = CleanInvalidSpaces(txtDomicilio.Text)
End Sub

Private Sub txtHorario_GotFocus()
    CSM_Control_TextBox.SelAllText txtHorario
End Sub

Private Sub txtHorario_LostFocus()
    txtHorario.Text = CleanInvalidSpaces(txtHorario.Text)
End Sub

Private Sub txtTelefono_GotFocus()
    CSM_Control_TextBox.SelAllText txtTelefono
End Sub

Private Sub txtTelefono_LostFocus()
    txtTelefono.Text = CleanInvalidSpaces(txtTelefono.Text)
End Sub

Private Sub txtValorDeclarado_GotFocus()
    CSM_Control_TextBox.SelAllText txtValorDeclarado
End Sub

Private Sub txtValorDeclarado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDecimal Then
        mKeyDecimal = True
    End If
End Sub

Private Sub txtValorDeclarado_KeyPress(KeyAscii As Integer)
    If ((KeyAscii > 32 And KeyAscii < 48) Or KeyAscii > 57) And KeyAscii <> Asc(pRegionalSettings.CurrencyDecimalSymbol) And KeyAscii <> Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
        KeyAscii = 0
    End If
    If mKeyDecimal Then
        KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol)
        mKeyDecimal = False
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol) Then
        If InStr(1, txtValorDeclarado.Text, pRegionalSettings.CurrencyDecimalSymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
        If InStr(1, txtValorDeclarado.Text, pRegionalSettings.CurrencyCurrencySymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtValorDeclarado_Change()
    Dim BaseImponible As Currency
    
    If IsNumeric(txtValorDeclarado.Text) Then
        'ES UN VALOR NUMERICO
        If pParametro.Comision_Seguro_PorcentajeAplicarSobreTotal Then
            'TOMO COMO BASE IMPONIBLE EL VALOR DECLARADO COMPLETO
            BaseImponible = CCur(txtValorDeclarado.Text)
        Else
            'TOMO COMO BASE IMPONIBLE EL VALOR DECLARADO MENOS EL MINIMO
            BaseImponible = CCur(txtValorDeclarado.Text) - pParametro.Comision_Seguro_ValorDeclaradoMinimo
        End If
        If CCur(txtValorDeclarado.Text) > pParametro.Comision_Seguro_ValorDeclaradoMinimo Then
            txtImporteSeguro.Text = Format(Round(BaseImponible * (pParametro.Comision_Seguro_Porcentaje / 100), pParametro.Comision_Seguro_RedondeoDecimales), "Currency")
        Else
            txtImporteSeguro.Text = Format(0, "Currency")
        End If
    Else
        txtImporteSeguro.Text = ""
    End If
    
    Call CalcularImporteTotal
End Sub

Private Sub txtValorDeclarado_LostFocus()
    If Not IsNumeric(txtValorDeclarado.Text) Then
        txtValorDeclarado.Text = Val(txtValorDeclarado.Text)
    End If
    txtValorDeclarado.Text = Format(CCur(txtValorDeclarado.Text), "Currency")
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

Private Sub txtImporte_Change()
    Call CalcularImporteTotal
    Call CalcularImporteCuentaCorriente
End Sub

Private Sub txtImporte_LostFocus()
    If Not txtImporte.Locked Then
        If Not IsNumeric(txtImporte.Text) Then
            txtImporte.Text = Val(txtImporte.Text)
        End If
        txtImporte.Text = Format(CCur(txtImporte.Text), "Currency")
    End If
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

Private Sub txtImporteContado_Change()
    CalcularImporteCuentaCorriente
End Sub

Private Sub txtImporteContado_LostFocus()
    If Not IsNumeric(txtImporteContado.Text) Then
        txtImporteContado.Text = Val(txtImporteContado.Text)
    End If
    txtImporteContado.Text = Format(CCur(txtImporteContado.Text), "Currency")

    lblCuentaCorrienteCaja.Visible = (CCur(txtImporteContado.Text) > 0 And mMedioPago.IDCuentaCorrienteCaja = 0)
    datcboCuentaCorrienteCaja.Visible = (CCur(txtImporteContado.Text) > 0 And mMedioPago.IDCuentaCorrienteCaja = 0)
    cmdCuentaCorrienteCaja.Visible = (CCur(txtImporteContado.Text) > 0 And mMedioPago.IDCuentaCorrienteCaja = 0)
End Sub

Private Sub datcboMedioPago_Change()
    If (Not mEsRutaEspecial) And Val(datcboMedioPago.BoundText) > 0 Then
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

Private Sub txtOperacion_GotFocus()
    CSM_Control_TextBox.SelAllText txtOperacion
End Sub

Private Sub txtOperacion_LostFocus()
    txtOperacion.Text = CleanInvalidSpaces(txtOperacion.Text)
End Sub

Private Sub txtNotas_GotFocus()
    CSM_Control_TextBox.SelAllText txtNotas
    cmdOK.Default = False
End Sub

Private Sub txtNotas_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNotas_LostFocus()
    txtNotas.Text = UCase(txtNotas.Text)
    cmdOK.Default = True
End Sub

Private Sub txtReservadoPor_GotFocus()
    CSM_Control_TextBox.SelAllText txtReservadoPor
End Sub

Private Sub cmdOK_Click()
    Dim Viaje As Viaje
    Dim ViajeDetalle As ViajeDetalle
    Dim ViajeDetalle_Conexion As ViajeDetalle_Conexion
    Dim RutaConexion As RutaConexion
    Dim Tramo1_Tramo2_IDLugar As Long
    Dim Tramo1_IDListaPrecio As Long
    Dim Tramo1_Importe As Currency
    Dim Tramo2_IDListaPrecio As Long
    Dim Tramo2_Importe As Currency
    Dim AsientoAsignado As Long
    Dim Conexion_Tramo1_IDDestino As Long
    Dim Conexion_Tramo1_Baja As String
    
    ' UPDATE 2018-04-21
    ' Cargo el Viaje actual
    Set Viaje = New Viaje
    Viaje.FechaHora = CDate(dtpFecha.value & " " & datcboHora.Text)
    Viaje.IDRuta = datcboRuta.BoundText
    If Not Viaje.Load() Then
        MsgBox "No se pudieron obtener los datos del Viaje.", vbCritical, App.Title
        Exit Sub
    End If
    
    If datcboHora.BoundText = "" Then
        MsgBox "Debe seleccionar la Hora.", vbInformation, App.Title
        datcboHora.SetFocus
        Exit Sub
    End If
    If datcboRuta.BoundText = "" Then
        MsgBox "Debe seleccionar la Ruta.", vbInformation, App.Title
        datcboRuta.SetFocus
        Exit Sub
    End If
    
    If (Not optTipoComision.value) And (Not optTipoPasajero.value) Then
        MsgBox "Debe seleccionar el Tipo.", vbInformation, App.Title
        optTipoComision.SetFocus
        Exit Sub
    End If
    If Val(txtPersona.Tag) = 0 Then
        MsgBox IIf(optTipoComision.value, "Debe indicar quién envía la Comisión.", "Debe seleccionar el Pasajero."), vbInformation, App.Title
        cmdPersona.SetFocus
        Exit Sub
    End If
    If optTipoPasajero.value And Val(datcboListaPrecio.BoundText) = 0 Then
        MsgBox "Debe seleccionar la Lista de Precios.", vbInformation, App.Title
        datcboListaPrecio.SetFocus
        Exit Sub
    End If
    
    If chkRutaConexion.Visible = True And chkRutaConexion.value = vbChecked Then
        If datcboRutaConexion.BoundText = "" Then
            MsgBox "Debe seleccionar la Conexión de Rutas.", vbInformation, App.Title
            datcboRutaConexion.SetFocus
            Exit Sub
        End If
        If datcboViajeConexion.BoundText = "" Then
            MsgBox "Debe seleccionar el Viaje para la Conexión.", vbInformation, App.Title
            datcboViajeConexion.SetFocus
            Exit Sub
        End If
    End If
    
    If Not mEsRutaEspecial Then
        If chkRutaConexion.Visible = False Or chkRutaConexion.value = vbUnchecked Then
            If txtImporte.Text = "" Then
                MsgBox "No se puede cargar una Reserva con el Importe Vacío.", vbInformation, App.Title
                datcboListaPrecio.SetFocus
                Exit Sub
            End If
            If Not IsNumeric(txtImporte.Text) Then
                MsgBox "El Importe ingresado es incorrecto.", vbInformation, App.Title
                txtImporte.SetFocus
                Exit Sub
            End If
            If CCur(txtImporte.Text) = 0 Then
                If optTipoComision.value Then
                    If mNew Then
                        If MsgBox("Esta Comisión tiene Importe Cero." & vbCr & "¿Desea cargar esta Comisión de todos modos?", vbExclamation + vbYesNo, App.Title) = vbNo Then
                            txtImporte.SetFocus
                            Exit Sub
                        End If
                    Else
                        If mViajeDetalle.Importe <> CCur(txtImporte.Text) Then
                            If MsgBox("Esta Comisión tiene Importe Cero." & vbCr & "¿Desea modificar esta Comisión de todos modos?", vbExclamation + vbYesNo, App.Title) = vbNo Then
                                txtImporte.SetFocus
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
            
            Tramo1_IDListaPrecio = Val(datcboListaPrecio.BoundText)
            Tramo1_Importe = CCur(txtImporte.Text)
        Else
            'VERIFICO LAS LISTA DE PRECIOS ESPECIFICADAS PARA LA CONEXION
            Set RutaConexion = New RutaConexion
            RutaConexion.IDRutaConexion = Val(datcboRutaConexion.BoundText)
            If RutaConexion.Load() Then
                Tramo1_Tramo2_IDLugar = RutaConexion.Tramo1_Tramo2_IDLugar
                Tramo1_IDListaPrecio = RutaConexion.Tramo1_IDListaPrecio
                Tramo2_IDListaPrecio = RutaConexion.Tramo2_IDListaPrecio
            End If
            Set RutaConexion = Nothing
        
            If Tramo1_IDListaPrecio > 0 Then
                'ESTÁ ESPECIFICADA UNA LISTA DE PRECIO DIFERENTE
                Tramo1_Importe = CalcularImporte_Function(Tramo1_IDListaPrecio, datcboRuta.Text, Val(datcboOrigen.BoundText), Tramo1_Tramo2_IDLugar)
                If Tramo1_Importe = -1 Then
                    MsgBox "No se puede cargar la Reserva porque no se ha especificado el Importe en la Lista de Precios del 1° Tramo de la Combinación de Rutas.", vbInformation, App.Title
                    Exit Sub
                End If
                If Tramo1_Importe = 0 Then
                    If optTipoComision.value Then
                        If mNew Then
                            If MsgBox("Esta Comisión tiene Importe Cero." & vbCr & "¿Desea cargar esta Comisión de todos modos?", vbExclamation + vbYesNo, App.Title) = vbNo Then
                                txtImporte.SetFocus
                                Exit Sub
                            End If
                        Else
                            If mViajeDetalle.Importe <> Tramo1_Importe Then
                                If MsgBox("Esta Comisión tiene Importe Cero." & vbCr & "¿Desea modificar esta Comisión de todos modos?", vbExclamation + vbYesNo, App.Title) = vbNo Then
                                    txtImporte.SetFocus
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                'NO HAY UNA LISTA DE PRECIOS ESPECIFICADA, USO LA ELEGIDA POR EL USUARIO
                If txtImporte.Text = "" Then
                    MsgBox "No se puede cargar una Reserva con el Importe Vacío.", vbInformation, App.Title
                    datcboListaPrecio.SetFocus
                    Exit Sub
                End If
                If Not IsNumeric(txtImporte.Text) Then
                    MsgBox "El Importe ingresado es incorrecto.", vbInformation, App.Title
                    txtImporte.SetFocus
                    Exit Sub
                End If
                If CCur(txtImporte.Text) = 0 Then
                    If optTipoComision.value Then
                        If mNew Then
                            If MsgBox("Esta Comisión tiene Importe Cero." & vbCr & "¿Desea cargar esta Comisión de todos modos?", vbExclamation + vbYesNo, App.Title) = vbNo Then
                                txtImporte.SetFocus
                                Exit Sub
                            End If
                        Else
                            If mViajeDetalle.Importe <> CCur(txtImporte.Text) Then
                                If MsgBox("Esta Comisión tiene Importe Cero." & vbCr & "¿Desea modificar esta Comisión de todos modos?", vbExclamation + vbYesNo, App.Title) = vbNo Then
                                    txtImporte.SetFocus
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
            
                Tramo1_IDListaPrecio = Val(datcboListaPrecio.BoundText)
                Tramo1_Importe = CCur(txtImporte.Text)
            End If
        
            If Tramo2_IDListaPrecio = 0 Then
                Tramo2_IDListaPrecio = Val(datcboListaPrecio.BoundText)
            End If
            If optTipoComision.value Then
                Tramo2_Importe = 0
            Else
                Tramo2_Importe = CalcularImporte_Function(Tramo2_IDListaPrecio, Mid(datcboViajeConexion.BoundText, 18), Tramo1_Tramo2_IDLugar, Val(datcboDestino.BoundText))
            End If
            If Tramo2_Importe = -1 Then
                MsgBox "No se puede cargar la Reserva porque no se ha especificado el Importe en la Lista de Precios del 2° Tramo de la Combinación de Rutas.", vbInformation, App.Title
                Exit Sub
            End If
        End If
    Else
        Tramo1_IDListaPrecio = Val(datcboListaPrecio.BoundText)
    End If
    
    If optTipoComision.value And pParametro.Comision_Seguro_Habilitar Then
        If Trim(txtValorDeclarado.Text) = "" Then
            MsgBox "Debe ingresar el Valor Declarado.", vbInformation, App.Title
            txtValorDeclarado.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txtValorDeclarado.Text) Then
            MsgBox "El Valor Declarado ingresado es incorrecto.", vbInformation, App.Title
            txtValorDeclarado.SetFocus
            Exit Sub
        End If
        If CCur(txtValorDeclarado.Text) < 0 Then
            MsgBox "El Valor Declarado debe ser mayor o igual a cero.", vbInformation, App.Title
            txtValorDeclarado.SetFocus
            Exit Sub
        End If
        If CCur(txtValorDeclarado.Text) > pParametro.Comision_Seguro_ValorDeclaradoMaximo Then
            MsgBox "El Valor Declarado no puede ser mayor a " & Format(pParametro.Comision_Seguro_ValorDeclaradoMaximo, "Currency"), vbInformation, App.Title
            txtValorDeclarado.SetFocus
            Exit Sub
        End If
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
    If CCur(txtImporteContado.Text) > 0 And mMedioPago.IDCuentaCorrienteCaja = 0 And Val(datcboCuentaCorrienteCaja.BoundText) = 0 Then
        MsgBox "Debe seleccionar la Caja.", vbInformation, App.Title
        Call txtImporteContado_LostFocus
        On Error Resume Next
        datcboCuentaCorrienteCaja.SetFocus
        Exit Sub
    End If
    
    If (Not mEsRutaEspecial) And Val(datcboMedioPago.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Medio de Pago.", vbInformation, App.Title
        datcboMedioPago.SetFocus
        Exit Sub
    End If
    If cboCuotas.Visible And cboCuotas.ListIndex = -1 Then
        MsgBox "Debe especificar las Cuotas.", vbInformation, App.Title
        cboCuotas.SetFocus
        Exit Sub
    End If
    If pParametro.ViajeDetalle_MedioPago_UtilizaOperacion_ObligaFacturaNumero Then
        If mMedioPago.UtilizaOperacion And Len(Trim(txtFacturaNumero.Text)) = 0 Then
            MsgBox "Debe especificar el Nº de Factura.", vbInformation, App.Title
            txtFacturaNumero.SetFocus
            Exit Sub
        End If
    End If
    
    ' UPDATE 2020-06-20
    ' Verifico el Origen y el Destino para ver si están disponibles en ese horario
    If Val(datcboOrigen.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Origen.", vbInformation, App.Title
        datcboOrigen.SetFocus
        Exit Sub
    End If
    If Not mEsRutaEspecial Then
        If Not VerificarLugarDisponiblePorHorario(datcboRuta.Text, datcboOrigen.BoundText, Weekday(dtpFecha.value), datcboHora.Text) Then
            MsgBox "El Origen no está disponible para este Horario.", vbExclamation, App.Title
            datcboOrigen.SetFocus
            Exit Sub
        End If
    End If
    If Val(datcboDestino.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Destino.", vbInformation, App.Title
        datcboDestino.SetFocus
        Exit Sub
    End If
    If Not mEsRutaEspecial Then
        If Not VerificarLugarDisponiblePorHorario(datcboRuta.Text, datcboDestino.BoundText, Weekday(dtpFecha.value), datcboHora.Text) Then
            MsgBox "El Destino no está disponible para este Horario.", vbExclamation, App.Title
            datcboDestino.SetFocus
            Exit Sub
        End If
    End If
        
    If Val(txtPersona.Tag) = Val(txtPersonaCuentaCorriente.Tag) Then
        MsgBox "No se puede especificar la misma Persona para Facturar.", vbExclamation, App.Title
        cmdPersonaCuentaCorriente.SetFocus
        Exit Sub
    End If
    
    If (Not pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_FINALIZADO_MODIFY, False)) And (Not pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_COMISION_MODIFY_IMPORTE_PAGOCONTADO_FINALIZADO, False)) Then
        Select Case Viaje.Estado
            Case VIAJE_ESTADO_CANCELADO
                MsgBox "No se pueden realizar cambios en este Viaje porque está Cancelado.", vbInformation, App.Title
                Set Viaje = Nothing
                Exit Sub
            Case VIAJE_ESTADO_FINALIZADO
                MsgBox "No se pueden realizar cambios en este Viaje porque está Finalizado.", vbInformation, App.Title
                Set Viaje = Nothing
                Exit Sub
        End Select
    End If
    
    'CONFIRMACION
    With frmViajeDetalleConfirmacion
        .txtTipo.Text = IIf(optTipoComision.value, OCUPANTE_TIPO_COMISION_NOMBRE, OCUPANTE_TIPO_PASAJERO_NOMBRE)
        .txtDia.Text = UCase(txtDiaSemana.Text)
        .txtFecha.Text = dtpFecha.value
        .txtHora.Text = datcboHora.Text
        .txtRuta.Text = datcboRuta.Text
        
        .lblPasajeroEnvia.Caption = IIf(optTipoComision.value, "Envía:", "Pasajero:")
        .txtPasajeroEnvia.Text = txtPersona.Text
        
        .lblRecibe.Visible = optTipoComision.value
        .txtRecibe.Visible = optTipoComision.value
        .txtRecibe.Text = txtPersonaRecibe.Text
        
        .lblCombinacion.Visible = (chkRutaConexion.Visible And chkRutaConexion.value = vbChecked)
        .txtCombinacion.Visible = (chkRutaConexion.Visible And chkRutaConexion.value = vbChecked)
        .txtCombinacion.Text = datcboViajeConexion.Text
        
        .Show vbModal, frmMDI
        If .Tag = "CANCEL" Then
            Unload frmViajeDetalleConfirmacion
            Set frmViajeDetalleConfirmacion = Nothing
            Exit Sub
        Else
            Unload frmViajeDetalleConfirmacion
            Set frmViajeDetalleConfirmacion = Nothing
        End If
    End With
    
        
'    If optTipoPasajero.Value Then
'        If MsgBox("¿Confirma que desea guardar los cambios de esta Reserva?" & vbCr & vbCr & "Tipo: Pasajero" & vbCr & "Día: " & txtDiaSemana.Text & vbCr & "Fecha/Hora: " & dtpFecha.Value & " " & Format(datcboHora.Text, "Short Time") & vbCr & "Ruta: " & datcboRuta.Text & vbCr & "Pasajero: " & txtPersona.Text & IIf(chkRutaConexion.Visible And chkRutaConexion.Value = vbChecked, vbCr & vbCr & "En Combinación con el Viaje: " & datcboViajeConexion.Text, ""), vbQuestion + vbYesNo, App.Title) = vbNo Then
'            Exit Sub
'        End If
'    Else
'        If MsgBox("¿Confirma que desea guardar los cambios de esta Reserva?" & vbCr & vbCr & "Tipo: Comisión" & vbCr & "Día: " & txtDiaSemana.Text & vbCr & "Fecha/Hora: " & dtpFecha.Value & " " & Format(datcboHora.Text, "Short Time") & vbCr & "Ruta: " & datcboRuta.Text & vbCr & "Envía: " & txtPersona.Text & vbCr & "Recibe: " & txtPersonaRecibe.Text & IIf(chkRutaConexion.Visible And chkRutaConexion.Value = vbChecked, vbCr & vbCr & "En Combinación con el Viaje: " & datcboViajeConexion.Text, ""), vbQuestion + vbYesNo, App.Title) = vbNo Then
'            Exit Sub
'        End If
'    End If
    
    If (Not mNew) And cboRealizado.ListIndex = VIAJE_DETALLE_REALIZADO_NO Then
        Call ViajeDetalle_ShowViajeVuelta(mViajeDetalle, "Hay %1 Reserva(s) sin Asistencia para este Pasajero en el mismo Día.")
    End If
    
    If chkRutaConexion.Visible = False Or chkRutaConexion.value = vbUnchecked Then
        Conexion_Tramo1_IDDestino = Val(datcboDestino.BoundText)
        Conexion_Tramo1_Baja = txtBaja.Text
    Else
        Conexion_Tramo1_IDDestino = Tramo1_Tramo2_IDLugar
        Conexion_Tramo1_Baja = ""
    End If
    
    'SI ES NUEVO O CAMBIO LA FECHA-HORA/RUTA O PASO DE NO REALIZADO A OTRO ESTADO DE REALIZADO,
    'Y NO PERMITE RESERVAS CONDICIONALES, VERIFICO QUE HAYA LUGAR
    If optTipoPasajero.value And Not pParametro.Permitir_Reservas_Condicionales And Not mEsRutaPaquete Then
        If ((mNew And cboRealizado.ListIndex <> VIAJE_DETALLE_REALIZADO_NO) Or mViajeDetalle.FechaHora <> CDate(dtpFecha.value & " " & datcboHora.Text) Or mViajeDetalle.IDRuta <> datcboRuta.BoundText Or mViajeDetalle.IDOrigen <> Val(datcboOrigen.BoundText) Or mViajeDetalle.IDDestino <> Val(datcboDestino.BoundText) Or (mViajeDetalle.Realizado = VIAJE_DETALLE_REALIZADO_NO And cboRealizado.ListIndex <> VIAJE_DETALLE_REALIZADO_NO)) Then
            Viaje.IDOrigen = Val(datcboOrigen.BoundText)
            Viaje.IDDestino = Conexion_Tramo1_IDDestino
            '#If USARDLLKUBALSOFT Then
                'AsientoAsignado = 99
            '#Else
                AsientoAsignado = Viaje.Asiento_Asignar_GetAsiento(mViajeDetalle.Indice)
            '#End If
            Select Case AsientoAsignado
                Case -1
                    MsgBox "No se puede " & IIf(mNew, "tomar", "actualizar") & " la Reserva porque no hay más lugar en el Viaje de origen.", vbExclamation, App.Title
                    Set Viaje = Nothing
                    Exit Sub
                Case -2
                Case -3
                Case Else
            End Select
        End If
        
        'SI ES CONEXION, VERIFICO EL ASIENTO DE LA CONEXION
        If mNew And datcboRutaConexion.BoundText <> "" And datcboViajeConexion.BoundText <> "" Then
            'AHORA BUSCO LA RUTA COMBINADA
            Dim ViajeConexion As New Viaje
            ViajeConexion.FechaHora = CDate(Left(datcboViajeConexion.BoundText, 16))
            ViajeConexion.IDRuta = Mid(datcboViajeConexion.BoundText, 18)
            ViajeConexion.IDOrigen = Tramo1_Tramo2_IDLugar
            ViajeConexion.IDDestino = Val(datcboDestino.BoundText)
            AsientoAsignado = ViajeConexion.Asiento_Asignar_GetAsiento(0)
            Select Case AsientoAsignado
                Case -1
                    MsgBox "No se puede tomar la Reserva porque no hay más lugar en el Viaje a combinar.", vbExclamation, App.Title
                    Set ViajeConexion = Nothing
                    Exit Sub
                Case -2
                Case -3
                Case Else
            End Select
            Set ViajeConexion = Nothing
        End If
    End If
    
    '//////////////////////////////////////////////
    'EL ORIGINAL
    With mViajeDetalle
        If Not mNew Then
            Call .Load
        End If
        
        .FechaHora = Viaje.FechaHora
        .IDRuta = Viaje.IDRuta
        .IDViaje = Viaje.IDViaje
        
        If mNew Then
            .OcupanteTipo = IIf(optTipoComision.value, OCUPANTE_TIPO_COMISION, OCUPANTE_TIPO_PASAJERO)
            If cboRealizado.ListIndex <> VIAJE_DETALLE_REALIZADO_NO Then
                .Asiento = AsientoAsignado
            Else
                .Asiento = -2
            End If
        End If
        .IDPersona = Val(txtPersona.Tag)
        
        .IDPersonaRecibe = IIf(optTipoComision.value, Val(txtPersonaRecibe.Tag), 0)
        .PagaQuienRecibe = IIf(optTipoComision.value, optPagaRecibe.value, False)
        .Descripcion = IIf(optTipoComision.value, txtDescripcion.Text, "")
        .Domicilio = IIf(optTipoComision.value, txtDomicilio.Text, "")
        .Horario = IIf(optTipoComision.value, txtHorario.Text, "")
        .Telefono = IIf(optTipoComision.value, txtTelefono.Text, "")
        .DejarTraer = IIf(optTipoComision.value, IIf(cboDejarTraer.ListIndex = 0, "", IIf(cboDejarTraer.ListIndex = 1, "D", "T")), "")
        
        .IDOrigen = Val(datcboOrigen.BoundText)
        .Sube = txtSube.Text
        .IDDestino = Conexion_Tramo1_IDDestino
        .Baja = Conexion_Tramo1_Baja
        
        .IDListaPrecio = Tramo1_IDListaPrecio
        If optTipoComision.value And pParametro.Comision_Seguro_Habilitar Then
            .ValorDeclarado = CCur(txtValorDeclarado.Text)
            .ImporteSeguro = CCur(txtImporteSeguro.Text)
        Else
            .ValorDeclarado = 0
            .ImporteSeguro = 0
        End If
        If Not mEsRutaEspecial Then
            .Importe = Tramo1_Importe
        Else
            .Importe = 0
        End If
        If picImporte.Visible Then
            .ImporteContado = CCur(txtImporteContado.Text)
            .IDMedioPago = Val(datcboMedioPago.BoundText)
            If mMedioPago.UtilizaOperacion Then
                .Cuotas = Val(cboCuotas.Text)
                .Operacion = txtOperacion.Text
            Else
                .Cuotas = 0
                .Operacion = ""
            End If
            If mEsRutaEspecial Then
                .ImporteCuentaCorriente = 0
            Else
                .ImporteCuentaCorriente = CCur(txtImporteCuentaCorriente.Text)
            End If
            .ImprimirSaldo = (chkImprimirSaldo.value = vbChecked)
            If CCur(txtImporteContado.Text) > 0 Then
                If mMedioPago.IDCuentaCorrienteCaja = 0 Then
                    .IDCuentaCorrienteCaja = Val(datcboCuentaCorrienteCaja.BoundText)
                Else
                    .IDCuentaCorrienteCaja = mMedioPago.IDCuentaCorrienteCaja
                End If
            Else
                .IDCuentaCorrienteCaja = 0
            End If
        Else
            .ImporteContado = 0
            .IDMedioPago = 0
            .Cuotas = 0
            .Operacion = ""
            .ImporteCuentaCorriente = 0
            .ImprimirSaldo = False
            .IDCuentaCorrienteCaja = 0
        End If
        .Realizado = cboRealizado.ListIndex
        .ForzarDebito = (chkForzarDebito.value = vbChecked)
        .CanceladoPor = (txtCanceladoPor.Text)
        .Entregada = (chkEntregada.value = vbChecked)
        If optTipoComision.value Then
            If .Entregada Then
                .EntregadaFechaHora = CDate(Format(dtpEntregadaFecha.value, "Short Date") & " " & Format(dtpEntregadaHora.value, "Short Time"))
                .Retira = txtRetira.Text
            Else
                .EntregadaFechaHora = DATE_TIME_FIELD_NULL_VALUE
                .Retira = ""
            End If
            .Asiento = -1
        Else
            .Retira = ""
        End If
        
        .IDPersonaCuentaCorriente = Val(txtPersonaCuentaCorriente.Tag)
        .Facturar = (chkFacturar.value = vbChecked)
        .FacturarNotas = txtFacturarNotas.Text
        .FacturaNumero = txtFacturaNumero.Text
        .Notas = txtNotas.Text
        If txtReservadoPor.Text = txtPersona.Text Then
            .ReservadoPorID = Val(txtPersona.Tag)
        End If
        .ReservadoPor = txtReservadoPor.Text
        
        .RefreshListSkip = True
        If Not .Update() Then
            Exit Sub
        End If
    End With
    
    '//////////////////////////////////////////////
    'LA CONEXION
    If chkRutaConexion.Visible And chkRutaConexion.value = vbChecked Then
        Set ViajeDetalle = New ViajeDetalle
        With ViajeDetalle
            .FechaHora = Left(datcboViajeConexion.BoundText, 16)
            .IDRuta = Mid(datcboViajeConexion.BoundText, 18)
            .OcupanteTipo = IIf(optTipoComision.value, OCUPANTE_TIPO_COMISION, OCUPANTE_TIPO_PASAJERO)
            If cboRealizado.ListIndex <> VIAJE_DETALLE_REALIZADO_NO Then
                .Asiento = AsientoAsignado
            Else
                .Asiento = -2
            End If
            .IDPersona = Val(txtPersona.Tag)
            
            .IDPersonaRecibe = IIf(optTipoComision.value, Val(txtPersonaRecibe.Tag), 0)
            .PagaQuienRecibe = IIf(optTipoComision.value, optPagaRecibe.value, False)
            .Descripcion = IIf(optTipoComision.value, txtDescripcion.Text, "")
            .Domicilio = IIf(optTipoComision.value, txtDomicilio.Text, "")
            .Horario = IIf(optTipoComision.value, txtHorario.Text, "")
            .Telefono = IIf(optTipoComision.value, txtTelefono.Text, "")
            .DejarTraer = IIf(optTipoComision.value, IIf(cboDejarTraer.ListIndex = 0, "", IIf(cboDejarTraer.ListIndex = 1, "D", "T")), "")
            
            .IDOrigen = Tramo1_Tramo2_IDLugar
            .Sube = ""
            .IDDestino = Val(datcboDestino.BoundText)
            .Baja = txtBaja.Text
            
            .IDListaPrecio = Tramo2_IDListaPrecio
            If optTipoComision.value And pParametro.Comision_Seguro_Habilitar Then
                .ValorDeclarado = CCur(txtValorDeclarado.Text)
                .ImporteSeguro = CCur(txtImporteSeguro.Text)
            Else
                .ValorDeclarado = 0
                .ImporteSeguro = 0
            End If
            If datcboRuta.BoundText <> pParametro.Ruta_ID_Otra Then
                .Importe = Tramo2_Importe
            Else
                .Importe = 0
            End If
            .ImporteContado = CCur(txtImporteContado.Text)
            .IDMedioPago = Val(datcboMedioPago.BoundText)
            .Operacion = txtOperacion.Text
            .ImporteCuentaCorriente = CCur(txtImporteCuentaCorriente.Text)
            .ImprimirSaldo = (chkImprimirSaldo.value = vbChecked)
            .IDCuentaCorrienteCaja = IIf(datcboCuentaCorrienteCaja.Visible, Val(datcboCuentaCorrienteCaja.BoundText), 0)
            .Realizado = cboRealizado.ListIndex
            .ForzarDebito = (chkForzarDebito.value = vbChecked)
            .CanceladoPor = (txtCanceladoPor.Text)
            .Entregada = (chkEntregada.value = vbChecked)
            If optTipoComision.value Then
                If .Entregada Then
                    .EntregadaFechaHora = CDate(Format(dtpEntregadaFecha.value, "Short Date") & " " & Format(dtpEntregadaHora.value, "Short Time"))
                    .Retira = txtRetira.Text
                Else
                    .EntregadaFechaHora = DATE_TIME_FIELD_NULL_VALUE
                    .Retira = ""
                End If
                .Asiento = -1
            Else
                .Retira = ""
            End If
            
            .IDPersonaCuentaCorriente = Val(txtPersonaCuentaCorriente.Tag)
            .Facturar = (chkFacturar.value = vbChecked)
            .FacturarNotas = txtFacturarNotas.Text
            .FacturaNumero = txtFacturaNumero.Text
            .Notas = txtNotas.Text
            .ReservadoPor = txtReservadoPor.Text
            
            .RefreshListSkip = True
            If Not .Update() Then
                Exit Sub
            End If
        End With
        
        Set ViajeDetalle_Conexion = New ViajeDetalle_Conexion
        ViajeDetalle_Conexion.IDViajeDetalle = mViajeDetalle.IDViajeDetalle
        ViajeDetalle_Conexion.Conexion_IDViajeDetalle = ViajeDetalle.IDViajeDetalle
        If Not ViajeDetalle_Conexion.Update() Then
            Exit Sub
        End If
        Set ViajeDetalle_Conexion = Nothing
    End If
    
    If picPagos.Visible Then
        If Not mViajeDetalle.UpdatePagos(mCPagosToAdd, mCPagosToUpdate, mCPagosToDelete) Then
            Exit Sub
        End If
    End If
    
    mViajeDetalle.RefreshListSkip = False
    Call mViajeDetalle.RefreshList
    Call RefreshList_Module.RefreshList_RefreshViaje(DATE_TIME_FIELD_NULL_VALUE, "")
    
    Set RutaConexion = Nothing
    Set Viaje = Nothing
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    'CONFIRMACION
    If cmdOK.Visible Then
        If MsgBox("¿Desea descartar los cambios realizados a esta Reserva?", vbExclamation + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Function VerificarAsientoMensaje(ByVal AsientoAsignado As Long, ByVal ViajeDescripcion As String) As String
    Select Case AsientoAsignado
        Case 0
            VerificarAsientoMensaje = ViajeDescripcion & "Hay lugar parado. No quedan Asientos disponibles."
        Case -1
            VerificarAsientoMensaje = ViajeDescripcion & "No hay más lugar en este Viaje."
        Case -2
            VerificarAsientoMensaje = ViajeDescripcion & "No se ha especificado el Vehículo del Viaje, por lo tanto no se puede determinar si hay lugar o no."
        Case -3
            VerificarAsientoMensaje = "CONEXION: No se encontró ningún viaje en conexión con el de origen."
        Case Else
            VerificarAsientoMensaje = ViajeDescripcion & "Hay lugar disponible para esta Reserva."
    End Select
End Function

Private Sub SetCaption()
    If mNew Then
        If optTipoComision.value Then
            Caption = "Propiedades de Nueva Comisión" & IIf(txtPersona.Text <> "", " de " & txtPersona.Text, "")
        Else
            Caption = "Propiedades de Nuevo Pasajero" & IIf(txtPersona.Text <> "", ": " & txtPersona.Text, "")
        End If
    Else
        If optTipoComision.value Then
            Caption = "Propiedades de la Comisión de " & txtPersona.Text
        Else
            Caption = "Propiedades del Pasajero " & txtPersona.Text
        End If
    End If
End Sub

Public Sub PersonaSelected(ByVal IDPersona As Long, ByVal Tag As String)
    Dim Viaje As Viaje
    Dim ViajeDetalle As ViajeDetalle
    Dim Persona As Persona
    Dim SaldoActual As Currency
    Dim Feriado As Feriado

    If IDPersona = 0 Then
        Exit Sub
    End If
    
    Select Case Tag
        Case "PP"
            Set Persona = New Persona
            Persona.IDPersona = IDPersona
            If Not Persona.Load() Then
                Set Persona = Nothing
                Exit Sub
            End If
            If Not Persona.HabilitadoViajar Then
                If Not pCPermiso.GotPermission(PERMISO_PERSONA_HABILITACION_VIAJAR_IGNORAR, False) Then
                    MsgBox "Esta persona ha sido inhabilitada para viajar en esta empresa.", vbExclamation, App.Title
                    Set Persona = Nothing
                    Exit Sub
                End If
            End If
            
            txtPersona.Tag = IDPersona
            
            If txtReservadoPor.Text = txtPersona.Text Then
                txtReservadoPor.Text = frmMDI.cboPersona.Text
            End If
            
            txtPersona.Text = frmMDI.cboPersona.Text
            
            If datcboRuta.BoundText <> "" And datcboRuta.BoundText <> pParametro.Ruta_ID_Otra And datcboRuta.BoundText <> pParametro.Ruta_Paquete_ID Then
                Set ViajeDetalle = New ViajeDetalle
                ViajeDetalle.FechaHora = CDate(dtpFecha.value & " " & datcboHora.BoundText)
                ViajeDetalle.IDRuta = datcboRuta.BoundText
                ViajeDetalle.IDPersona = IDPersona
                ViajeDetalle.IDRuta = datcboRuta.BoundText
                ViajeDetalle.OcupanteTipo = IIf(optTipoComision.value, OCUPANTE_TIPO_COMISION, OCUPANTE_TIPO_PASAJERO)
                Set Viaje = New Viaje
                Viaje.FechaHora = CDate(dtpFecha.value & " " & datcboHora.BoundText)
                Viaje.IDRuta = datcboRuta.BoundText
                If Not Viaje.Load() Then
                    Set Persona = Nothing
                    Set ViajeDetalle = Nothing
                    Set Viaje = Nothing
                    Exit Sub
                End If
                If Not ViajeDetalle.GetNewValues(Viaje.DiaSemanaBase) Then
                    Set Persona = Nothing
                    Set ViajeDetalle = Nothing
                    Set Viaje = Nothing
                    Exit Sub
                End If
                Set Viaje = Nothing
                'txtPersonaCuentaCorriente.Tag = ViajeDetalle.IDPersonaCuentaCorriente
                datcboListaPrecio.BoundText = ViajeDetalle.IDListaPrecio
                datcboOrigen.BoundText = ViajeDetalle.IDOrigen
                txtSube.Text = ViajeDetalle.Sube
                datcboDestino.BoundText = ViajeDetalle.IDDestino
                txtBaja.Text = ViajeDetalle.Baja
                If ViajeDetalle.Importe > -1 Then
                    txtImporte.Text = Format(ViajeDetalle.Importe, "Currency")
                Else
                    txtImporte.Text = ""
                End If
                Set ViajeDetalle = Nothing
            End If
            
            txtPersonaCuentaCorriente.Tag = Persona.IDPersonaCuentaCorriente
            
            'VERIFICO SI TIENE INASISTENCIAS
            Call Persona.VerificarInasistencias
            
            'VERIFICO SI TIENE RESERVAS PARA UN FERIADO
            Set Feriado = New Feriado
            Feriado.VerificarReservasDelPasajero IDPersona
            Set Feriado = Nothing
            
            'SALDO ACTUAL
            Persona.ViajeActual_FechaHora = mViajeDetalle.FechaHora
            Persona.ViajeActual_IDRuta = mViajeDetalle.IDRuta
            Persona.ViajeActual_Indice = mViajeDetalle.Indice
            If Val(txtPersonaCuentaCorriente.Tag) <> 0 Then
                Persona.IDPersona = Val(txtPersonaCuentaCorriente.Tag)
                If Persona.Load() Then
                    txtPersonaCuentaCorriente.Text = Persona.ApellidoNombre
                End If
            End If
            Persona.LoadSaldoActual
            SaldoActual = Persona.SaldoActual
            
            'AVISA QUE DEBE
            If Persona.SaldoActual < 0 And Not Persona.PermiteViajarSinPagar Then
                Load frmPersonaSaldo
                frmPersonaSaldo.txtPersona.Tag = Persona.IDPersona
                frmPersonaSaldo.txtPersona.Text = " " & Persona.ApellidoNombre
                frmPersonaSaldo.txtSaldo.Text = Persona.SaldoActual_Formatted
                frmPersonaSaldo.FillListView
                frmPersonaSaldo.Show
            End If
            
            Set Persona = Nothing
            
            On Error Resume Next
            txtImporte.SetFocus
            
        Case "PR"   'RECIBE
            txtPersonaRecibe.Tag = IDPersona
            Set Persona = New Persona
            Persona.IDPersona = IDPersona
            If Persona.Load() Then
                txtPersonaRecibe.Text = Persona.ApellidoNombre
            End If
            
            Call CambioPersona
                        
        Case "PF"   'FACTURA
            If IDPersona = Val(txtPersona.Tag) Then
                MsgBox "No se puede seleccionar la misma Persona para Facturar.", vbExclamation, App.Title
            Else
                Set Persona = New Persona
                Persona.IDPersona = IDPersona
                If Persona.Load() Then
                    If Persona.IDPersonaCuentaCorriente > 0 Then
                        MsgBox "No se puede seleccionar a una Persona que ya tiene especificado debitar a otra Persona.", vbExclamation, App.Title
                    Else
                        txtPersonaCuentaCorriente.Tag = IDPersona
                        txtPersonaCuentaCorriente.Text = Persona.ApellidoNombre
                    
                        Persona.ViajeActual_FechaHora = mViajeDetalle.FechaHora
                        Persona.ViajeActual_IDRuta = mViajeDetalle.IDRuta
                        Persona.ViajeActual_Indice = mViajeDetalle.Indice
                        Persona.LoadSaldoActual
                        SaldoActual = Persona.SaldoActual
                        
                        'AVISA QUE DEBE
                        If Persona.SaldoActual < 0 And Not Persona.PermiteViajarSinPagar Then
                            Load frmPersonaSaldo
                            frmPersonaSaldo.txtPersona.Tag = Persona.IDPersona
                            frmPersonaSaldo.txtPersona.Text = " " & Persona.ApellidoNombre
                            frmPersonaSaldo.txtSaldo.Text = Persona.SaldoActual_Formatted
                            frmPersonaSaldo.FillListView
                            frmPersonaSaldo.Show
                        End If
                    End If
                End If
                Set Persona = Nothing
            End If
            
            On Error Resume Next
    End Select
    txtSaldoActual.Tag = SaldoActual
    CalcularImporteCuentaCorriente
End Sub

Private Sub ShowControls()
    lblPersona.Caption = IIf(optTipoComision.value, "&Envía:", "&Pasajero:")

    lblPersonaRecibe.Visible = optTipoComision.value
    txtPersonaRecibe.Visible = optTipoComision.value
    cmdPersonaRecibe.Visible = optTipoComision.value
    cmdPersonaRecibeUltimo.Visible = optTipoComision.value
    
    optPagaEnvia.Visible = optTipoComision.value
    optPagaRecibe.Visible = optTipoComision.value
    
    lblDescripcion.Visible = optTipoComision.value
    txtDescripcion.Visible = optTipoComision.value
    lblDomicilio.Visible = optTipoComision.value
    txtDomicilio.Visible = optTipoComision.value
    lblHorario.Visible = optTipoComision.value
    txtHorario.Visible = optTipoComision.value
    lblTelefono.Visible = optTipoComision.value
    txtTelefono.Visible = optTipoComision.value
    cboDejarTraer.Visible = optTipoComision.value
    
    lblRealizado.Visible = optTipoPasajero.value
    cboRealizado.Visible = optTipoPasajero.value
    lblForzarDebito.Visible = (optTipoPasajero.value And cboRealizado.ListIndex = VIAJE_DETALLE_REALIZADO_NO)
    chkForzarDebito.Visible = (optTipoPasajero.value And cboRealizado.ListIndex = VIAJE_DETALLE_REALIZADO_NO)
    chkEntregada.Visible = optTipoComision.value
    dtpEntregadaFecha.Visible = (optTipoComision.value And chkEntregada.value = vbChecked)
    dtpEntregadaHora.Visible = (optTipoComision.value And chkEntregada.value = vbChecked)
    lblRetira.Visible = (optTipoComision.value And chkEntregada.value = vbChecked)
    txtRetira.Visible = (optTipoComision.value And chkEntregada.value = vbChecked)
    
    lblListaPrecio.Visible = Not (mEsRutaEspecial Or mEsRutaPaquete)
    datcboListaPrecio.Visible = Not (mEsRutaEspecial Or mEsRutaPaquete)
    cmdListaPrecio.Visible = Not (mEsRutaEspecial Or mEsRutaPaquete)
    
    lblValorDeclarado.Visible = (optTipoComision.value And pParametro.Comision_Seguro_Habilitar)
    txtValorDeclarado.Visible = (optTipoComision.value And pParametro.Comision_Seguro_Habilitar)
    lblImporteSeguro.Visible = (optTipoComision.value And pParametro.Comision_Seguro_Habilitar)
    txtImporteSeguro.Visible = (optTipoComision.value And pParametro.Comision_Seguro_Habilitar)
    lblImporteTotal.Visible = (optTipoComision.value And pParametro.Comision_Seguro_Habilitar)
    txtImporteTotal.Visible = (optTipoComision.value And pParametro.Comision_Seguro_Habilitar)
    
    lblImporte.Visible = Not mEsRutaEspecial
    txtImporte.Visible = Not mEsRutaEspecial
    txtImporte.TabStop = mEsRutaPaquete
    
    picImporte.Visible = Not (pParametro.ViajeDetalle_Paquete_Permite_Multiples_Pagos And mEsRutaPaquete)
    picPagos.Visible = (pParametro.ViajeDetalle_Paquete_Permite_Multiples_Pagos And mEsRutaPaquete)
    
    lblImporteContado.Visible = Not mEsRutaEspecial
    txtImporteContado.Visible = Not mEsRutaEspecial
    
    datcboMedioPago.Visible = (Not mEsRutaEspecial)
    lblOperacion.Visible = (Not mEsRutaEspecial)
    txtOperacion.Visible = (Not mEsRutaEspecial)
    
    lblImporteCuentaCorriente.Visible = Not mEsRutaEspecial
    txtImporteCuentaCorriente.Visible = Not mEsRutaEspecial
    chkImprimirSaldo.Visible = Not mEsRutaEspecial
    txtSaldoActual.Visible = Not mEsRutaEspecial
    cmdSaldoActual.Visible = Not mEsRutaEspecial
    
    If (optTipoComision.value And pParametro.ViajeDetalle_Comision_Precio_PermitirModificar And pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_COMISION_IMPORTE_ALLOWMODIFY, False)) Or (optTipoPasajero.value And pParametro.ViajeDetalle_Pasajero_Precio_PermitirModificar And pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_PASAJERO_IMPORTE_ALLOWMODIFY, False)) Or (mEsRutaPaquete And pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_PAQUETE_PASAJERO_IMPORTE_ALLOWMODIFY, False)) Then
        txtImporte.Locked = False
        txtImporte.TabStop = True
        txtImporte.BackColor = vbWindowBackground
    Else
        txtImporte.Locked = True
        txtImporte.TabStop = False
        txtImporte.BackColor = vbButtonFace
    End If
    cmdVerificarAsiento.Visible = optTipoPasajero.value
End Sub

Private Sub EnableControls(ByVal value As Boolean)
    cmdAnterior.Visible = value
    dtpFecha.Enabled = value
    cmdSiguiente.Visible = value
    cmdHoy.Visible = value
    datcboHora.Enabled = value
    datcboRuta.Enabled = value
    cmdRuta.Visible = value
    
    optPagaEnvia.Enabled = value
    cmdPersona.Visible = value
    cmdPersonaUltimo.Visible = value
    
    optPagaRecibe.Enabled = value
    cmdPersonaRecibe.Visible = value
    cmdPersonaRecibeUltimo.Visible = value
    txtDescripcion.Enabled = value
    txtDomicilio.Enabled = value
    txtHorario.Enabled = value
    txtTelefono.Enabled = value
    cboDejarTraer.Enabled = value
    
    datcboOrigen.Enabled = value
    cmdLugarOrigen.Visible = value
    txtSube.Enabled = value
    datcboDestino.Enabled = value
    cmdLugarDestino.Visible = value
    txtBaja.Enabled = value
    
    cmdVerificarAsiento.Visible = value
    
    datcboListaPrecio.Enabled = value
    cmdListaPrecio.Visible = value
    
    txtImporte.Enabled = value
    txtImporteContado.Enabled = value
    datcboMedioPago.Enabled = value
    cboCuotas.Enabled = value
    txtOperacion.Enabled = value
    
    chkImprimirSaldo.Enabled = value
    datcboCuentaCorrienteCaja.Enabled = value
    cmdCuentaCorrienteCaja.Visible = value
    
    cmdPersonaCuentaCorriente.Visible = value
    cmdPersonaCuentaCorrienteClear.Visible = value
    
    txtReservadoPor.Enabled = value
    cboRealizado.Enabled = value
    chkForzarDebito.Enabled = value
    chkEntregada.Enabled = value
    dtpEntregadaFecha.Enabled = value
    dtpEntregadaHora.Enabled = value
    
    chkFacturar.Enabled = value
    txtFacturaNumero.Enabled = value
    txtFacturarNotas.Enabled = value
    
    txtNotas.Enabled = value
    
    cmdOK.Visible = value
    If value Then
        cmdCancel.Caption = "Cancelar"
    Else
        cmdCancel.Caption = "Cerrar"
    End If
End Sub

Private Sub CambioPersona()
    Dim Persona As Persona
    Dim SaldoActual As Currency
    
    Set Persona = New Persona
    If optPagaEnvia.value Then
        Persona.IDPersona = Val(txtPersona.Tag)
    Else
        Persona.IDPersona = Val(txtPersonaRecibe.Tag)
    End If
    If Persona.IDPersona = 0 Then
        Set Persona = Nothing
        Exit Sub
    End If
    If Persona.Load() Then
        If Persona.IDPersonaCuentaCorriente = 0 Then
            txtPersonaCuentaCorriente.Tag = 0
            txtPersonaCuentaCorriente.Text = ""
        Else
            txtPersonaCuentaCorriente.Tag = Persona.IDPersonaCuentaCorriente
            txtPersonaCuentaCorriente.Text = Persona.ApellidoNombre
        End If
    End If
    If Val(txtPersonaCuentaCorriente.Tag) <> 0 Then
        Persona.IDPersona = Val(txtPersonaCuentaCorriente.Tag)
        If Persona.Load() Then
            txtPersonaCuentaCorriente.Text = Persona.ApellidoNombre
        End If
    End If
    
    'SALDO ACTUAL
    Persona.ViajeActual_FechaHora = mViajeDetalle.FechaHora
    Persona.ViajeActual_IDRuta = mViajeDetalle.IDRuta
    Persona.ViajeActual_Indice = mViajeDetalle.Indice
    Persona.LoadSaldoActual
    SaldoActual = Persona.SaldoActual
    
    'AVISA QUE DEBE
    If Persona.SaldoActual < 0 And Not Persona.PermiteViajarSinPagar Then
        Load frmPersonaSaldo
        frmPersonaSaldo.txtPersona.Tag = Persona.IDPersona
        frmPersonaSaldo.txtPersona.Text = " " & Persona.ApellidoNombre
        frmPersonaSaldo.txtSaldo.Text = Persona.SaldoActual_Formatted
        frmPersonaSaldo.FillListView
        frmPersonaSaldo.Show
    End If
    
    Set Persona = Nothing
    
    txtSaldoActual.Tag = SaldoActual
    CalcularImporteCuentaCorriente
End Sub

Private Function CalcularImporte_Function(ByVal IDListaPrecio As Long, ByVal IDRuta As String, ByVal IDOrigen As Long, ByVal IDDestino As Long) As Currency
    Dim ListaPrecioDetalle As ListaPrecioDetalle
    
    Set ListaPrecioDetalle = New ListaPrecioDetalle
    ListaPrecioDetalle.IDListaPrecio = IDListaPrecio
    ListaPrecioDetalle.OcupanteTipo = IIf(optTipoComision.value, OCUPANTE_TIPO_COMISION, OCUPANTE_TIPO_PASAJERO)
    ListaPrecioDetalle.IDRuta = IDRuta
    ListaPrecioDetalle.IDOrigen = IDOrigen
    ListaPrecioDetalle.IDDestino = IDDestino
    If ListaPrecioDetalle.GetImporteByLugar() Then
        If mViajeDetalle.IDUsuarioCreacion = pParametro.ReservaWebIdUsuario Then
            If ListaPrecioDetalle.ImporteWeb = -1 Then
                CalcularImporte_Function = ListaPrecioDetalle.Importe
            Else
                CalcularImporte_Function = ListaPrecioDetalle.ImporteWeb
            End If
        Else
            CalcularImporte_Function = ListaPrecioDetalle.Importe
        End If
    Else
        CalcularImporte_Function = -1
    End If
    Set ListaPrecioDetalle = Nothing
End Function

Private Sub CalcularImporte()
    Dim Importe As Currency
    
    If Val(datcboListaPrecio.BoundText) <> 0 And Val(datcboOrigen.BoundText) <> 0 And Val(datcboDestino.BoundText) <> 0 Then
        Importe = CalcularImporte_Function(Val(datcboListaPrecio.BoundText), datcboRuta.BoundText, Val(datcboOrigen.BoundText), IIf(chkRutaConexion.Visible And chkRutaConexion.value = vbChecked, mTramo1_Tramo2_IDLugar, Val(datcboDestino.BoundText)))
        If Importe > -1 Then
            txtImporte.Text = Format(Importe, "Currency")
        Else
            txtImporte.Text = ""
        End If
    Else
        txtImporte.Text = ""
    End If
End Sub

Private Sub CalcularImporteTotal()
    If IsNumeric(txtImporteSeguro.Text) Then
        If IsNumeric(txtImporte.Text) Then
            txtImporteTotal.Text = Format(CCur(txtImporteSeguro.Text) + CCur(txtImporte.Text), "Currency")
        Else
            txtImporteTotal.Text = txtImporteSeguro.Text
        End If
    Else
        If IsNumeric(txtImporte.Text) Then
            txtImporteTotal.Text = txtImporte.Text
        Else
            txtImporteTotal.Text = ""
        End If
    End If
End Sub

Private Sub FillComboBoxViajes()
    Dim RutaDetalle As RutaDetalle
    Dim RutaConexion As RutaConexion
    
    Dim RutaDetalle_Duracion As Integer
    Dim Tramo1_Tramo2_IDLugar As Long
    
    Dim SQLStatement As String
        
    If datcboHora.Text = "" Or Val(datcboOrigen.BoundText) = 0 Or Val(datcboDestino.BoundText) = 0 Then
        Set datcboViajeConexion.DataSource = Nothing
        Exit Sub
    End If
    
    If datcboRutaConexion.BoundText <> "" Then
        '//////////////////////////////////
        'COMBINACION DE RUTAS
        '//////////////////////////////////
        
        Set RutaConexion = New RutaConexion
        RutaConexion.IDRutaConexion = Val(datcboRutaConexion.BoundText)
        If RutaConexion.Load() Then
            Tramo1_Tramo2_IDLugar = RutaConexion.Tramo1_Tramo2_IDLugar
        End If
        Set RutaConexion = Nothing
        
        'BUSCO LAS RUTAS COMBINADAS POSIBLES Y BUSCO SI HAY LUGAR EN ALGUNA
        Set RutaDetalle = New RutaDetalle
        RutaDetalle.IDRuta = datcboRuta.BoundText
        RutaDetalle.IDLugar = Tramo1_Tramo2_IDLugar
        If RutaDetalle.Load() Then
            RutaDetalle_Duracion = RutaDetalle.Duracion
        End If
        Set RutaDetalle = Nothing
        
        SQLStatement = "SELECT DISTINCT Viaje.FechaHora, CONVERT(CHAR(10), Viaje.FechaHora, 111) + ' ' + CONVERT(CHAR(5), Viaje.FechaHora, 108) + '|' + RTrim(Viaje.IDRuta) AS BoundField, CONVERT(CHAR(5), Viaje.FechaHora, 108) + ' - ' + RTrim(Viaje.IDRuta) AS ListField" & vbCr
        SQLStatement = SQLStatement & "FROM (Viaje INNER JOIN RutaConexionDetalle ON Viaje.IDRuta = RutaConexionDetalle.Tramo2_IDRuta) INNER JOIN RutaDetalle ON Viaje.IDRuta = RutaDetalle.IDRuta" & vbCr
        SQLStatement = SQLStatement & "WHERE RutaDetalle.IDLugar = " & Tramo1_Tramo2_IDLugar & " AND RutaConexionDetalle.IDRutaConexion = " & Val(datcboRutaConexion.BoundText) & " AND dateadd(minute, RutaDetalle.Duracion + RutaDetalle.Espera, Viaje.FechaHora) BETWEEN '" & Format(DateAdd("n", RutaDetalle_Duracion, CDate(dtpFecha.value & " " & datcboHora.Text)), "yyyy/mm/dd hh:nn") & "' AND '" & Format(DateAdd("n", RutaDetalle_Duracion + pParametro.Viaje_RutaConexion_TiempoEsperaMaximo_Minutos, CDate(dtpFecha.value & " " & datcboHora.Text)), "yyyy/mm/dd hh:nn") & "' AND (Viaje.Estado = 'AC' OR Viaje.Estado = 'EP')" & vbCr
        SQLStatement = SQLStatement & "ORDER BY Viaje.FechaHora"
        
        Call CSM_Control_DataCombo.FillFromSQL(datcboViajeConexion, SQLStatement, "BoundField", "ListField", "Conexiones de Viajes", cscpCurrentOrFirst)
    End If
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

Public Sub FillComboBoxLugar()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboOrigen.BoundText)
    Set recData = datcboOrigen.RowSource
    recData.Requery
    Set recData = Nothing
    datcboOrigen.BoundText = KeySave

    KeySave = Val(datcboDestino.BoundText)
    Set recData = datcboDestino.RowSource
    recData.Requery
    Set recData = Nothing
    datcboDestino.BoundText = KeySave
End Sub

Public Sub FillComboBoxListaPrecio()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboListaPrecio.BoundText)
    Set recData = datcboListaPrecio.RowSource
    recData.Requery
    Set recData = Nothing
    datcboListaPrecio.BoundText = KeySave
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

Public Sub FillComboBoxCuentaCorrienteCaja()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboCuentaCorrienteCaja.BoundText)
    Set recData = datcboCuentaCorrienteCaja.RowSource
    recData.Requery
    Set recData = Nothing
    datcboCuentaCorrienteCaja.BoundText = KeySave
End Sub

Private Sub Pago_Add(ByVal Pago As CuentaCorriente)
    mCPagosToAdd.Add Pago
End Sub

Private Sub Pago_Update(ByVal Pago As CuentaCorriente)
    If Pago.IDMovimiento > 0 Then
        mCPagosToUpdate.Add Pago
    End If
End Sub

Private Sub Pago_Delete(ByVal Pago As CuentaCorriente)
    If Pago.IDMovimiento > 0 Then
        mCPagosToDelete.Add Pago
    End If
End Sub

Private Function VerificarLugarDisponiblePorHorario(ByVal IDRuta As String, ByVal IDLugar As Integer, ByVal DiaSemana As Byte, ByVal Hora As Date) As Boolean
    Dim cmdData As ADODB.command
    
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrorHandler
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "usp_RutaDetalle_HorarioDisponible"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta", adChar, adParamInput, 20, IDRuta)
    cmdData.Parameters.Append cmdData.CreateParameter("IDLugar", adInteger, adParamInput, , IDLugar)
    cmdData.Parameters.Append cmdData.CreateParameter("DiaSemana", adTinyInt, adParamInput, , DiaSemana - 1)
    cmdData.Parameters.Append cmdData.CreateParameter("Hora", adDBTime, adParamInput, , Hora)
    cmdData.Parameters.Append cmdData.CreateParameter("Disponible", adBoolean, adParamOutput)
    cmdData.Execute
    
    VerificarLugarDisponiblePorHorario = cmdData.Parameters("Disponible").value

    Set cmdData = Nothing

    Screen.MousePointer = vbDefault
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Forms.ViajeDetallePropiedad.VerificarLugarDisponiblePorHorario", "Error al verificar la disponibilidad el Lugar."
End Function
