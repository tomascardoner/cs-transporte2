VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmViajeConductor 
   Caption         =   "Viajes por Conductor"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12720
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ViajeConductor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6195
   ScaleWidth      =   12720
   Begin VB.CommandButton cmdShowReport 
      Caption         =   "Reporte"
      Height          =   975
      Left            =   10980
      Picture         =   "ViajeConductor.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   360
      Width           =   975
   End
   Begin VB.Frame fraTramo 
      Caption         =   "Tramos:"
      Height          =   1095
      Left            =   9420
      TabIndex        =   20
      Top             =   240
      Width           =   1395
      Begin VB.CheckBox chkTramo_Tramo2 
         Caption         =   "Tramo 2"
         Height          =   210
         Left            =   180
         TabIndex        =   23
         Top             =   780
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkTramo_Tramo1 
         Caption         =   "Tramo 1"
         Height          =   210
         Left            =   180
         TabIndex        =   22
         Top             =   540
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkTramo_Completo 
         Caption         =   "Completo"
         Height          =   210
         Left            =   180
         TabIndex        =   21
         Top             =   300
         Value           =   1  'Checked
         Width           =   1035
      End
   End
   Begin VB.Frame fraEstado 
      Caption         =   "Estado:"
      Height          =   1095
      Left            =   6240
      TabIndex        =   15
      Top             =   240
      Width           =   3075
      Begin VB.CheckBox chkEstadoCancelado 
         Caption         =   "Cancelado"
         Height          =   210
         Left            =   1740
         TabIndex        =   19
         Top             =   720
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkEstadoFinalizado 
         Caption         =   "Finalizado"
         Height          =   210
         Left            =   1740
         TabIndex        =   18
         Top             =   360
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkEstadoEnProgreso 
         Caption         =   "En Progreso"
         Height          =   210
         Left            =   180
         TabIndex        =   17
         Top             =   720
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkEstadoActivo 
         Caption         =   "Activo"
         Height          =   210
         Left            =   180
         TabIndex        =   16
         Top             =   360
         Value           =   1  'Checked
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRuta 
      Caption         =   "..."
      Height          =   315
      Left            =   4440
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Rutas"
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton cmdConductor 
      Caption         =   "..."
      Height          =   315
      Left            =   4440
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Personas"
      Top             =   540
      Width           =   255
   End
   Begin VB.CommandButton cmdAnteriorDesde 
      Height          =   315
      Left            =   1080
      Picture         =   "ViajeConductor.frx":0636
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Anterior"
      Top             =   120
      Width           =   300
   End
   Begin VB.CommandButton cmdSiguienteDesde 
      Height          =   315
      Left            =   2820
      Picture         =   "ViajeConductor.frx":0BC0
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Siguiente"
      Top             =   120
      Width           =   300
   End
   Begin VB.CommandButton cmdHoyDesde 
      Height          =   315
      Left            =   3120
      Picture         =   "ViajeConductor.frx":114A
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Hoy"
      Top             =   120
      Width           =   315
   End
   Begin VB.CommandButton cmdAnteriorHasta 
      Height          =   315
      Left            =   3720
      Picture         =   "ViajeConductor.frx":1294
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Anterior"
      Top             =   120
      Width           =   300
   End
   Begin VB.CommandButton cmdSiguienteHasta 
      Height          =   315
      Left            =   5460
      Picture         =   "ViajeConductor.frx":181E
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Siguiente"
      Top             =   120
      Width           =   300
   End
   Begin VB.CommandButton cmdHoyHasta 
      Height          =   315
      Left            =   5760
      Picture         =   "ViajeConductor.frx":1DA8
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Hoy"
      Top             =   120
      Width           =   315
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   25
      Top             =   5835
      Width           =   12720
      _ExtentX        =   22437
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21898
            MinWidth        =   176
            Key             =   "TEXT"
         EndProperty
      EndProperty
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
   Begin MSComCtl2.DTPicker dtpFechaDesde 
      Height          =   315
      Left            =   1380
      TabIndex        =   2
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
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
   Begin MSComCtl2.DTPicker dtpFechaHasta 
      Height          =   315
      Left            =   4020
      TabIndex        =   6
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
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
   Begin MSDataListLib.DataCombo datcboConductor 
      Height          =   330
      Left            =   1080
      TabIndex        =   10
      Top             =   540
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
   Begin MSComctlLib.ListView lvwData 
      Height          =   3615
      Left            =   120
      TabIndex        =   24
      Top             =   2040
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "DiaSemana"
         Text            =   "Día"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Fecha"
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Hora"
         Text            =   "Hora"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "IDRuta"
         Text            =   "Ruta"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Vehiculo"
         Text            =   "Vehículo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "Estado"
         Text            =   "Estado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Key             =   "Tramo"
         Text            =   "Tramo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Key             =   "Importe"
         Text            =   "Importe"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSDataListLib.DataCombo datcboRuta 
      Height          =   330
      Left            =   1080
      TabIndex        =   13
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
   Begin VB.Label lblRuta 
      AutoSize        =   -1  'True
      Caption         =   "&Ruta:"
      Height          =   210
      Left            =   120
      TabIndex        =   12
      Top             =   1020
      Width           =   375
   End
   Begin VB.Label lblConductor 
      AutoSize        =   -1  'True
      Caption         =   "&Conductor:"
      Height          =   210
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   795
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   495
   End
   Begin VB.Label lblFechaA 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   210
      Left            =   3540
      TabIndex        =   26
      Top             =   180
      Width           =   90
   End
End
Attribute VB_Name = "frmViajeConductor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLoading As Boolean

Public Function FillListView() As Boolean
    Dim ListItem As MSComctlLib.ListItem
    Dim recData As ADODB.Recordset
    Dim SQL_Where As String
    Dim SQL_OrderBy As String
    Dim Viaje As Viaje
    Dim ImporteTotal As Currency
    
    If mLoading Or Val(datcboConductor.BoundText) = 0 Then
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    SQL_Where = ""
    
    'FILTRO SI PERSONAL ESTÁ ACTIVADO
    If pPersonal Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", "WHERE ", " AND ") & "Viaje.Personal = 0"
    End If
    
    'FILTRO DE FECHAS
    SQL_Where = SQL_Where & IIf(SQL_Where = "", "WHERE ", " AND ") & "Viaje.FechaHora BETWEEN '" & Format(dtpFechaDesde.value, "yyyy/mm/dd") & " 00:00:00' AND '" & Format(dtpFechaHasta.value, "yyyy/mm/dd") & " 23:59:00'"
    
    'FILTRO DE CONDUCTOR
    SQL_Where = SQL_Where & IIf(SQL_Where = "", "WHERE ", " AND ") & "(Viaje.IDConductor = " & Val(datcboConductor.BoundText) & " OR Viaje.IDConductor2 = " & Val(datcboConductor.BoundText) & ")"
    
    'FILTRO DE RUTAS
    If datcboRuta.BoundText <> CSM_Constant.ITEM_ALL_FEMALE Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", "WHERE ", " AND ") & "Viaje.IDRuta = '" & ReplaceQuote(datcboRuta.BoundText) & "'"
    Else
        If pCPermiso.RutaWhere <> "" Then
            SQL_Where = SQL_Where & IIf(SQL_Where = "", "WHERE ", " AND ") & Replace(pCPermiso.RutaWhere, "%TABLENAME%", "Viaje")
        End If
    End If
    
    'FILTROS DE ESTADOS
    If chkEstadoActivo.value = vbUnchecked Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", "WHERE ", " AND ") & "Viaje.Estado <> '" & VIAJE_ESTADO_ACTIVO & "'"
    End If
    If chkEstadoEnProgreso.value = vbUnchecked Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", "WHERE ", " AND ") & "Viaje.Estado <> '" & VIAJE_ESTADO_EN_PROGRESO & "'"
    End If
    If chkEstadoFinalizado.value = vbUnchecked Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", "WHERE ", " AND ") & "Viaje.Estado <> '" & VIAJE_ESTADO_FINALIZADO & "'"
    End If
    If chkEstadoCancelado.value = vbUnchecked Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", "WHERE ", " AND ") & "Viaje.Estado <> '" & VIAJE_ESTADO_CANCELADO & "'"
    End If
    
    'FILTROS DE TRAMOS
    If pParametro.Viaje_Permite_2_Conductores Then
        If chkTramo_Completo.value = vbUnchecked Then
            SQL_Where = SQL_Where & IIf(SQL_Where = "", "WHERE ", " AND ") & "NOT (ISNULL(Viaje.IDConductor2, 0) = 0)"
        End If
        If chkTramo_Tramo1.value = vbUnchecked Then
            SQL_Where = SQL_Where & IIf(SQL_Where = "", "WHERE ", " AND ") & "NOT (ISNULL(Viaje.IDConductor2, 0) <> 0 AND Viaje.IDConductor = " & Val(datcboConductor.BoundText) & ")"
        End If
        If chkTramo_Tramo2.value = vbUnchecked Then
            SQL_Where = SQL_Where & IIf(SQL_Where = "", "WHERE ", " AND ") & "NOT (ISNULL(Viaje.IDConductor2, 0) = " & Val(datcboConductor.BoundText) & ")"
        End If
    End If
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    'ORDEN
    Select Case lvwData.SortKey
        Case 0  'DIA SEMANA
            SQL_OrderBy = " ORDER BY datepart(weekday, Viaje.FechaHora)" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Viaje.FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Viaje.IDRuta + ': ' + Viaje.RutaOtra" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 1  'FECHA
            SQL_OrderBy = " ORDER BY Viaje.FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Viaje.IDRuta + ': ' + Viaje.RutaOtra" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 2  'HORA
            SQL_OrderBy = " ORDER BY convert(char(8), Viaje.FechaHora, 108)" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", datepart(weekday, Viaje.FechaHora)" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Viaje.FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Viaje.IDRuta + ': ' + Viaje.RutaOtra" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 3  'RUTA
            SQL_OrderBy = " ORDER BY Viaje.IDRuta + ISNULL(': ' + Viaje.RutaOtra, '')" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", datepart(weekday, Viaje.FechaHora)" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Viaje.FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 4  'VEHICULO
            SQL_OrderBy = " ORDER BY Vehiculo.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", datepart(weekday, Viaje.FechaHora)" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Viaje.FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Viaje.IDRuta + ': ' + Viaje.RutaOtra" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 5  'ESTADO
            SQL_OrderBy = " ORDER BY Viaje.Estado" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", datepart(weekday, Viaje.FechaHora)" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Viaje.FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Viaje.IDRuta + ': ' + Viaje.RutaOtra" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 6  'TRAMO
            SQL_OrderBy = " ORDER BY Tramo" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", datepart(weekday, Viaje.FechaHora)" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Viaje.FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Viaje.IDRuta + ': ' + Viaje.RutaOtra" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 7  'IMPORTE
            SQL_OrderBy = " ORDER BY Importe" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", datepart(weekday, Viaje.FechaHora)" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Viaje.FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Viaje.IDRuta + ': ' + Viaje.RutaOtra" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
    End Select
    
    Set Viaje = New Viaje
    
    lvwData.ListItems.Clear
    
    Set recData = New ADODB.Recordset
    recData.Source = "SELECT Viaje.FechaHora, Viaje.IDRuta, Viaje.RutaOtra, Viaje.Estado, Vehiculo.Nombre AS Vehiculo, "
    If pParametro.Viaje_Permite_2_Conductores Then
        recData.Source = recData.Source & "dbo.udf_GetViajeTramoNombre(" & Val(datcboConductor.BoundText) & ", Viaje.IDConductor, Viaje.IDConductor2) AS Tramo, "
        recData.Source = recData.Source & "(CASE dbo.udf_GetViajeTramoNumero(" & Val(datcboConductor.BoundText) & ", Viaje.IDConductor, Viaje.IDConductor2) WHEN 0 THEN dbo.udf_GetViajeTramoImporte(ConductorRuta1.ConductorImporteTramoCompleto, Horario.ConductorImporteTramoCompleto, Ruta.ConductorImporteTramoCompleto) WHEN 1 THEN dbo.udf_GetViajeTramoImporte(ConductorRuta1.ConductorImporteTramo1, Horario.ConductorImporteTramo1, Ruta.ConductorImporteTramo1) WHEN 2 THEN dbo.udf_GetViajeTramoImporte(ConductorRuta2.ConductorImporteTramo2, Horario.ConductorImporteTramo2, Ruta.ConductorImporteTramo2) END) AS Importe" & vbCr
        recData.Source = recData.Source & "FROM ("
    Else
        recData.Source = recData.Source & "'Completo' AS Tramo, "
        recData.Source = recData.Source & "dbo.udf_GetViajeTramoImporte(ConductorRuta1.ConductorImporteTramoCompleto, Horario.ConductorImporteTramoCompleto, Ruta.ConductorImporteTramoCompleto) AS Importe" & vbCr
        recData.Source = recData.Source & "FROM "
    End If
    recData.Source = recData.Source & "(((Viaje LEFT JOIN Vehiculo ON Viaje.IDVehiculo = Vehiculo.IDVehiculo) INNER JOIN Ruta ON Viaje.IDRuta = Ruta.IDRuta) LEFT JOIN Horario ON Viaje.DiaSemanaBase = Horario.DiaSemana AND CONVERT(CHAR(8), Viaje.FechaHora, 108) = CONVERT(CHAR(8), Horario.Hora, 108) AND Viaje.IDRuta = Horario.IDRuta) LEFT JOIN ConductorRuta AS ConductorRuta1 ON Viaje.IDConductor = ConductorRuta1.IDPersona AND Viaje.IDRuta = ConductorRuta1.IDRuta"
    If pParametro.Viaje_Permite_2_Conductores Then
        recData.Source = recData.Source & ") LEFT JOIN ConductorRuta AS ConductorRuta2 ON Viaje.IDConductor2 = ConductorRuta2.IDPersona AND Viaje.IDRuta = ConductorRuta2.IDRuta" & vbCr
    Else
        recData.Source = recData.Source & vbCr
    End If
    
    recData.Source = recData.Source & SQL_Where & SQL_OrderBy
    recData.MaxRecords = pParametro.Recordset_MaxRecords
    recData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With recData
        If Not .EOF Then
            Do While Not .EOF
                Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & .Fields("FechaHora").value & KEY_DELIMITER & RTrim(.Fields("IDRuta").value), WeekdayName(Weekday(.Fields("FechaHora").value)))
                ListItem.SubItems(1) = Format(.Fields("FechaHora").value, "Short Date")
                ListItem.SubItems(2) = Format(.Fields("FechaHora").value, "Short Time")
                ListItem.SubItems(3) = RTrim(.Fields("IDRuta").value) & IIf(RTrim(.Fields("IDRuta").value) = pParametro.Ruta_ID_Otra, ": " & .Fields("RutaOtra").value, "")
                ListItem.SubItems(4) = .Fields("Vehiculo").value & ""
                Viaje.Estado = .Fields("Estado").value
                ListItem.SubItems(5) = Viaje.Estado_ToString
                ListItem.SubItems(6) = .Fields("Tramo").value
                If .Fields("Importe").value = 0 Then
                    ListItem.SubItems(7) = " "
                Else
                    ListItem.SubItems(7) = Format(.Fields("Importe").value, "Currency")
                    ImporteTotal = ImporteTotal + .Fields("Importe").value
                End If
                ListItem.ForeColor = Viaje.Estado_ToColor
                ListItem.Bold = Viaje.Estado_ToBold
                .MoveNext
            Loop
            
            stbMain.Panels("TEXT").Text = .RecordCount & " items" & IIf(.RecordCount = .MaxRecords, " (Limitados)", "") & "  -  Importe Total: " & Format(ImporteTotal, "Currency")
        Else
            stbMain.Panels("TEXT").Text = "No hay items."
        End If
        .Close
    End With
    Set recData = Nothing
    
    Set Viaje = Nothing
    
    On Error Resume Next

    If frmMDI.ActiveForm.Name = Me.Name And GetForegroundWindow() = frmMDI.hwnd And frmMDI.WindowState <> vbMinimized Then
        lvwData.SetFocus
    End If
    
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Forms.ViajeConductor.FillListView", "Error al leer la lista de Viajes."
End Function

Private Sub cmdAnteriorDesde_Click()
    dtpFechaDesde.value = DateAdd("d", -1, dtpFechaDesde.value)
    dtpFechaDesde.SetFocus
    dtpFechaDesde_Change
End Sub

Private Sub cmdShowReport_Click()
    Dim Reporte As New Reporte
    
    If pCPermiso.GotPermission(PERMISO_REPORTE_REPORTE & "Viaje_Listado_Conductor_Importe") Then
        Set Reporte = New Reporte
        Reporte.IDReporte = "Viaje_Listado_Conductor_Importe"
        If Reporte.Load() Then
            Reporte.Titulo = "Viajes del Conductor: " & datcboConductor.Text & " - Desde el " & dtpFechaDesde.value & " al " & dtpFechaHasta.value
            Reporte.Parametros("FechaDesde").Valor = dtpFechaDesde.value & " 00:00"
            Reporte.Parametros("FechaHasta").Valor = dtpFechaHasta.value & " 23:59"
            Reporte.Parametros("IDConductor").Valor = Val(datcboConductor.BoundText)
            If datcboRuta.Text <> CSM_Constant.ITEM_ALL_FEMALE Then
                Reporte.Parametros("IDRuta").Valor = datcboRuta.Text
            End If
            
            'FILTROS DE ESTADOS
            Reporte.Parametros("EstadoActivo").Valor = chkEstadoActivo.value
            Reporte.Parametros("EstadoEnProgreso").Valor = chkEstadoEnProgreso.value
            Reporte.Parametros("EstadoFinalizado").Valor = chkEstadoFinalizado.value
            Reporte.Parametros("EstadoCancelado").Valor = chkEstadoCancelado.value
            
            'FILTROS DE TRAMOS
            If pParametro.Viaje_Permite_2_Conductores Then
                Reporte.Parametros("TramoCompleto").Valor = chkTramo_Completo.value
                Reporte.Parametros("Tramo1").Valor = chkTramo_Tramo1.value
                Reporte.Parametros("Tramo2").Valor = chkTramo_Tramo2.value
            End If
        
            If Reporte.OpenReport() Then
                Reporte.PrintReport True
            End If
            
            Set Reporte = Nothing
        End If
    End If
End Sub

Private Sub dtpFechaDesde_Change()
    FillListView
End Sub

Private Sub cmdSiguienteDesde_Click()
    dtpFechaDesde.value = DateAdd("d", 1, dtpFechaDesde.value)
    dtpFechaDesde.SetFocus
    dtpFechaDesde_Change
End Sub

Private Sub cmdHoyDesde_Click()
    Dim OldValue As Date
    
    OldValue = dtpFechaDesde.value
    dtpFechaDesde.value = Date
    dtpFechaDesde.SetFocus
    If OldValue <> dtpFechaDesde.value Then
        dtpFechaDesde_Change
    End If
End Sub

Private Sub cmdAnteriorHasta_Click()
    dtpFechaHasta.value = DateAdd("d", -1, dtpFechaHasta.value)
    dtpFechaHasta.SetFocus
    dtpFechaHasta_Change
End Sub

Private Sub dtpFechaHasta_Change()
    FillListView
End Sub

Private Sub cmdSiguienteHasta_Click()
    dtpFechaHasta.value = DateAdd("d", 1, dtpFechaHasta.value)
    dtpFechaHasta.SetFocus
    dtpFechaHasta_Change
End Sub

Private Sub cmdHoyHasta_Click()
    Dim OldValue As Date
    
    OldValue = dtpFechaHasta.value
    dtpFechaHasta.value = Date
    dtpFechaHasta.SetFocus
    If OldValue <> dtpFechaHasta.value Then
        dtpFechaHasta_Change
    End If
End Sub

Private Sub datcboConductor_Change()
    FillListView
End Sub

Private Sub cmdConductor_Click()
    If pCPermiso.GotPermission(PERMISO_PERSONA) Then
        Call frmPersona.FindAndShowItem(Val(datcboConductor.BoundText), UCase(Left(datcboConductor.Text, 1)), Me.Name, ENTIDAD_TIPO_PERSONA_CONDUCTOR, "")
    End If
End Sub

Private Sub datcboRuta_Change()
    FillListView
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

Private Sub chkEstadoActivo_Click()
    FillListView
End Sub

Private Sub chkEstadoEnProgreso_Click()
    FillListView
End Sub

Private Sub chkEstadoFinalizado_Click()
    FillListView
End Sub

Private Sub chkEstadoCancelado_Click()
    FillListView
End Sub

Private Sub chkTramo_Completo_Click()
    FillListView
End Sub

Private Sub chkTramo_Tramo1_Click()
    FillListView
End Sub

Private Sub chkTramo_Tramo2_Click()
    FillListView
End Sub

Private Sub lvwData_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwData.ColumnHeaders(lvwData.SortKey + 1).Icon = 0
    If ColumnHeader.Index - 1 = lvwData.SortKey Then
        lvwData.SortOrder = IIf(lvwData.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lvwData.SortKey = ColumnHeader.Index - 1
        lvwData.SortOrder = lvwAscending
    End If
    ColumnHeader.Icon = lvwData.SortOrder + 1
    FillListView
End Sub

Private Sub Form_Load()
    mLoading = True
    
    lvwData.GridLines = pParametro.ListView_GridLines
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL LISTVIEW
    Set lvwData.ColumnHeaderIcons = frmMDI.ilsFormSortColumn
    '//////////////////////////////////////////////////////////
    
    dtpFechaDesde.value = Date
    dtpFechaHasta.value = Date
    
    Call CSM_Control_DataCombo.FillFromSQL(datcboConductor, "SELECT IDPersona, dbo.udf_GetEntidadApellidoYNombre(Apellido, Nombre) AS ApellidoNombre FROM Persona WHERE Activo = 1 AND EntidadTipo = '" & ENTIDAD_TIPO_PERSONA_CONDUCTOR & "' ORDER BY ApellidoNombre", "IDPersona", "ApellidoNombre", "Conductores", cscpFirst)
    
    Call CSM_Control_DataCombo.FillFromSQL(datcboRuta, "(SELECT '" & CSM_Constant.ITEM_ALL_FEMALE & "' AS IDRuta) UNION (SELECT RTRIM(IDRuta) AS IDRuta FROM Ruta WHERE Activo = 1" & IIf(pCPermiso.RutaWhere <> "", " AND " & Replace(pCPermiso.RutaWhere, "%TABLENAME%", "Ruta"), "") & ") ORDER BY IDRuta", "IDRuta", "IDRuta", "Rutas", cscpFirst)
        
    CSM_Forms.ResizeAndPosition frmMDI, Me
    pParametro.GetListViewSettings "ViajeConductor", lvwData
    lvwData.ColumnHeaders(lvwData.SortKey + 1).Icon = lvwData.SortOrder + 1
    
    fraTramo.Visible = pParametro.Viaje_Permite_2_Conductores

    mLoading = False
End Sub

Private Sub Form_Resize()
    ResizeControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visible = False
    WindowState = vbNormal
    pParametro.SaveListViewSettings "ViajeConductor", lvwData
End Sub

Private Sub ResizeControls()
    Const CONTROL_SPACE = 60
    
    On Error Resume Next
    
    lvwData.Top = datcboRuta.Top + datcboRuta.Height + (CONTROL_SPACE * 2)
    lvwData.Left = CONTROL_SPACE
    lvwData.Width = ScaleWidth - (CONTROL_SPACE * 2)
    lvwData.Height = ScaleHeight - lvwData.Top - CONTROL_SPACE - stbMain.Height
End Sub
