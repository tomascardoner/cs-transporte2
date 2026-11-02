VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmViaje 
   Caption         =   "Viajes"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10800
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Viaje.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5565
   ScaleWidth      =   10800
   Begin MSComctlLib.Toolbar tlbPin 
      Height          =   330
      Left            =   60
      TabIndex        =   21
      Top             =   5220
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PIN"
            ImageIndex      =   1
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrMain 
      Align           =   1  'Align Top
      Height          =   1020
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   1799
      BandCount       =   6
      FixedOrder      =   -1  'True
      _CBWidth        =   10800
      _CBHeight       =   1020
      _Version        =   "6.7.9816"
      Child1          =   "tlbMain"
      MinHeight1      =   570
      Width1          =   3000
      FixedBackground1=   0   'False
      Key1            =   "Toolbar"
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Child2          =   "picFilterDiaSemana"
      MinWidth2       =   1695
      MinHeight2      =   360
      Width2          =   1695
      Key2            =   "FilterDiaSemana"
      NewRow2         =   0   'False
      AllowVertical2  =   0   'False
      Child3          =   "picFilterFecha"
      MinWidth3       =   6705
      MinHeight3      =   360
      Width3          =   6705
      Key3            =   "FilterFecha"
      NewRow3         =   0   'False
      AllowVertical3  =   0   'False
      Child4          =   "picFilterRutaGrupo"
      MinHeight4      =   360
      Width4          =   795
      Key4            =   "FilterRutaGrupo"
      NewRow4         =   0   'False
      AllowVertical4  =   0   'False
      Child5          =   "picFilterRuta"
      MinWidth5       =   3015
      MinHeight5      =   360
      Width5          =   3015
      Key5            =   "FilterRuta"
      NewRow5         =   0   'False
      AllowVertical5  =   0   'False
      Child6          =   "picFilterEstado"
      MinWidth6       =   2625
      MinHeight6      =   330
      Width6          =   2625
      Key6            =   "FilterEstado"
      NewRow6         =   0   'False
      AllowVertical6  =   0   'False
      Begin VB.PictureBox picFilterRutaGrupo 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   10770
         ScaleHeight     =   360
         ScaleWidth      =   15
         TabIndex        =   27
         Top             =   135
         Width           =   15
         Begin VB.OptionButton optRutaGrupo 
            Height          =   315
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   15
            Width           =   795
         End
      End
      Begin VB.PictureBox picFilterEstado 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8085
         ScaleHeight     =   330
         ScaleWidth      =   2625
         TabIndex        =   24
         Top             =   645
         Width           =   2625
         Begin VB.ComboBox cboFilterEstado 
            Height          =   330
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   0
            Width           =   2010
         End
         Begin VB.Label lblFilterEstado 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   210
            Left            =   0
            TabIndex        =   26
            Top             =   60
            Width           =   540
         End
      End
      Begin VB.PictureBox picFilterRuta 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   165
         ScaleHeight     =   360
         ScaleWidth      =   7695
         TabIndex        =   13
         Top             =   630
         Width           =   7695
         Begin VB.ComboBox cboRuta 
            Height          =   330
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   0
            Width           =   2550
         End
         Begin VB.Label lblRuta 
            AutoSize        =   -1  'True
            Caption         =   "Ruta:"
            Height          =   210
            Left            =   0
            TabIndex        =   14
            Top             =   60
            Width           =   375
         End
      End
      Begin VB.PictureBox picFilterFecha 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3840
         ScaleHeight     =   360
         ScaleWidth      =   6705
         TabIndex        =   7
         Top             =   135
         Width           =   6705
         Begin VB.TextBox txtDiaSemana 
            Height          =   315
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.CommandButton cmdHoyHasta 
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
            Left            =   6360
            Picture         =   "Viaje.frx":058A
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            ToolTipText     =   "Hoy"
            Top             =   0
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton cmdSiguienteHasta 
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
            Left            =   6060
            Picture         =   "Viaje.frx":06D4
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   "Siguiente"
            Top             =   0
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.CommandButton cmdAnteriorHasta 
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
            Picture         =   "Viaje.frx":0C5E
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "Anterior"
            Top             =   0
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.CommandButton cmdHoyDesde 
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
            Left            =   3720
            Picture         =   "Viaje.frx":11E8
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Hoy"
            Top             =   0
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton cmdSiguienteDesde 
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
            Left            =   3420
            Picture         =   "Viaje.frx":1332
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Siguiente"
            Top             =   0
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.CommandButton cmdAnteriorDesde 
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
            Left            =   1680
            Picture         =   "Viaje.frx":18BC
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Anterior"
            Top             =   0
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.ComboBox cboFecha 
            Height          =   330
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   0
            Width           =   1035
         End
         Begin MSComCtl2.DTPicker dtpFechaDesde 
            Height          =   315
            Left            =   1980
            TabIndex        =   9
            Top             =   0
            Visible         =   0   'False
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
            Left            =   4620
            TabIndex        =   10
            Top             =   0
            Visible         =   0   'False
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
         Begin VB.Label lblFechaAnd 
            AutoSize        =   -1  'True
            Caption         =   "y"
            Height          =   210
            Left            =   4140
            TabIndex        =   12
            Top             =   60
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Label lblFecha 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   210
            Left            =   0
            TabIndex        =   11
            Top             =   60
            Width           =   495
         End
      End
      Begin VB.PictureBox picFilterDiaSemana 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         ScaleHeight     =   360
         ScaleWidth      =   1695
         TabIndex        =   4
         Top             =   135
         Width           =   1695
         Begin VB.ComboBox cboDiaSemana 
            Height          =   330
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   0
            Width           =   1350
         End
         Begin VB.Label lblDiaSemana 
            AutoSize        =   -1  'True
            Caption         =   "Día:"
            Height          =   210
            Left            =   0
            TabIndex        =   6
            Top             =   60
            Width           =   270
         End
      End
      Begin MSComctlLib.Toolbar tlbMain 
         Height          =   570
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   1005
         ButtonWidth     =   2381
         ButtonHeight    =   1005
         ToolTips        =   0   'False
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Generar"
               Key             =   "GENERATE"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Nuevo"
               Key             =   "NEW"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Propiedades"
               Key             =   "PROPERTIES"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Eliminar"
               Key             =   "DELETE"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Seleccionar"
               Key             =   "SELECT"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Detalle"
               Key             =   "DETAIL"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cambiar Estado"
               Key             =   "CHANGE_STATUS"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Enviar E-mail"
               Key             =   "EMAIL"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "EMAIL_OBSERVACIONES"
                     Text            =   "Planilla con Observaciones"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "EMAIL_DOCUMENTO"
                     Text            =   "Planilla con Documento"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   5205
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   661
            MinWidth        =   661
            Key             =   "PIN"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17833
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
   Begin MSComctlLib.ListView lvwData 
      Height          =   3615
      Left            =   60
      TabIndex        =   0
      Top             =   1440
      Width           =   5955
      _ExtentX        =   10504
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
      NumItems        =   10
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
         Key             =   "Conductor"
         Text            =   "Conductor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "Conductor2"
         Text            =   "Conductor 2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "Estado"
         Text            =   "Estado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Key             =   "AsientoLibre"
         Text            =   "Asientos"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Key             =   "Notas"
         Text            =   "Notas"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmViaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mLoading As Boolean

Public FormWaitingForSelect As String
Public AllowMultipleSelect As Boolean
Public AllowMultipleRuta As Boolean
Public CSelectEstadosFilter As Collection
Public SelectTag As String

Public Sub FillListView(ByVal FechaHora As Date, ByVal IDRuta As String)
    Dim KeySave As Variant
    Dim CKeySave As Collection
    Dim Viaje As Viaje
    
    Dim SQL_Where As String
    Dim SQL_OrderBy As String
    Dim recData As ADODB.Recordset
    Dim ListItem As MSComctlLib.ListItem
    
    Dim Index As Long
    Dim IDRutaGrupo As Long
    
    If mLoading Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    If IDRuta = "" Then
        If Not lvwData.SelectedItem Is Nothing Then
            KeySave = lvwData.SelectedItem.Key
            Set CKeySave = New Collection
            For Each ListItem In lvwData.ListItems
                If ListItem.Selected Then
                    CKeySave.Add ListItem.Key
                End If
            Next ListItem
        End If
    Else
        KeySave = KEY_STRINGER & FechaHora & KEY_DELIMITER & IDRuta
    End If
        
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
        
    'PERSONAL
    If pPersonal Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Viaje.Personal = 0"
    End If
    
    'DATE FILTER
    Select Case cboFecha.ListIndex
        Case 0  'ALL
        Case 1  'EQUAL
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Viaje.FechaHora BETWEEN '" & Format(dtpFechaDesde.value, "yyyy/mm/dd") & " 00:00:00' AND '" & Format(dtpFechaDesde.value, "yyyy/mm/dd") & " 23:59:00'"
        Case 2  'GREATER
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Viaje.FechaHora > '" & Format(dtpFechaDesde.value, "yyyy/mm/dd") & " 23:59:00'"
        Case 3  'GREATER OR EQUAL
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Viaje.FechaHora >= '" & Format(dtpFechaDesde.value, "yyyy/mm/dd") & " 00:00:00'"
        Case 4  'MINOR
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Viaje.FechaHora < '" & Format(dtpFechaDesde.value, "yyyy/mm/dd") & " 00:00:00'"
        Case 5  'MINOR OR EQUAL
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Viaje.FechaHora <= '" & Format(dtpFechaDesde.value, "yyyy/mm/dd") & " 23:59:00'"
        Case 6  'NOT EQUAL
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Viaje.FechaHora NOT BETWEEN '" & Format(dtpFechaDesde.value, "yyyy/mm/dd") & " 00:00:00' AND '" & Format(dtpFechaDesde.value, "yyyy/mm/dd") & " 23:59:00'"
        Case 7  'BETWEEN
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Viaje.FechaHora BETWEEN '" & Format(dtpFechaDesde.value, "yyyy/mm/dd") & " 00:00:00' AND '" & Format(dtpFechaHasta.value, "yyyy/mm/dd") & " 23:59:00'"
    End Select
    
    'RUTA Y GRUPO DE RUTAS
    If cboRuta.ListIndex = 0 Then
        For Index = 0 To optRutaGrupo.Count - 1
            If optRutaGrupo(Index).value Then
                IDRutaGrupo = Val(optRutaGrupo(Index).Tag)
                Exit For
            End If
        Next Index
        If IDRutaGrupo > 0 Then
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Ruta.IDRutaGrupo = " & IDRutaGrupo
        End If
        If pCPermiso.RutaWhere <> "" Then
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & Replace(pCPermiso.RutaWhere, "%TABLENAME%", "Viaje")
        End If
    Else
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "(RTrim(Viaje.IDRuta) = '" & ReplaceQuote(cboRuta.Text) & "'" & IIf(pParametro.Viaje_Especial_MostrarEnTodasLasRutas, " OR Viaje.IDRuta = '" & pParametro.Ruta_ID_Otra & "' OR Viaje.IDRuta = '" & pParametro.Ruta_Paquete_ID & "'", "") & ")"
    End If
    
    'ESTADO
    Select Case cboFilterEstado.ListIndex
        Case 0
            '<Todos>
        Case 1
            'Activo
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Viaje.Estado = '" & VIAJE_ESTADO_ACTIVO & "'"
        Case 2
            'En Progreso
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Viaje.Estado = '" & VIAJE_ESTADO_EN_PROGRESO & "'"
        Case 3
            'Finalizado
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Viaje.Estado = '" & VIAJE_ESTADO_FINALIZADO & "'"
        Case 4
            'Cancelado
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Viaje.Estado = '" & VIAJE_ESTADO_CANCELADO & "'"
    End Select
    
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
        Case 5  'CONDUCTOR
            SQL_OrderBy = " ORDER BY dbo.udf_GetEntidadApellidoYNombre(Conductor.Apellido, Conductor.Nombre)" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", datepart(weekday, Viaje.FechaHora)" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Viaje.FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Viaje.IDRuta + ': ' + Viaje.RutaOtra" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 6  'CONDUCTOR 2
            SQL_OrderBy = " ORDER BY dbo.udf_GetEntidadApellidoYNombre(Conductor2.Apellido, Conductor2.Nombre)" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", datepart(weekday, Viaje.FechaHora)" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Viaje.FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Viaje.IDRuta + ': ' + Viaje.RutaOtra" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 7  'ESTADO
            SQL_OrderBy = " ORDER BY Viaje.Estado" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", datepart(weekday, Viaje.FechaHora)" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Viaje.FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Viaje.IDRuta + ': ' + Viaje.RutaOtra" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 8  'ASIENTO LIBRE
            SQL_OrderBy = " ORDER BY AsientoLibre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", datepart(weekday, Viaje.FechaHora)" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Viaje.FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Viaje.IDRuta + ': ' + Viaje.RutaOtra" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 9  'NOTAS + DIA SEMANA + FECHA-HORA + RUTA
            SQL_OrderBy = " ORDER BY Viaje.Notas" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", datepart(weekday, Viaje.FechaHora)" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Viaje.FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Viaje.IDRuta + ': ' + Viaje.RutaOtra" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
    End Select
    
    Set Viaje = New Viaje
    
    lvwData.ListItems.Clear
    
    Set recData = New ADODB.Recordset
    recData.Source = "SELECT Viaje.FechaHora, Viaje.IDRuta, Viaje.RutaOtra, Viaje.Estado, Vehiculo.Asiento - Viaje.AsientoOcupado AS AsientoLibre, Vehiculo.Nombre AS Vehiculo, dbo.udf_GetEntidadApellidoYNombre(Conductor.Apellido, Conductor.Nombre) AS Conductor, dbo.udf_GetEntidadApellidoYNombre(Conductor2.Apellido, Conductor2.Nombre) AS Conductor2, Viaje.Notas" & vbCr
    recData.Source = recData.Source & "FROM (((Viaje INNER JOIN Ruta ON Viaje.IDRuta = Ruta.IDRuta) LEFT JOIN Vehiculo ON Viaje.IDVehiculo = Vehiculo.IDVehiculo) LEFT JOIN Persona AS Conductor ON Viaje.IDConductor = Conductor.IDPersona) LEFT JOIN Persona AS Conductor2 ON Viaje.IDConductor2 = Conductor2.IDPersona" & SQL_Where & SQL_OrderBy
    recData.MaxRecords = pParametro.Recordset_MaxRecords
    recData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With recData
        If Not .EOF Then
            Do While Not .EOF
                Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & .Fields("FechaHora").value & KEY_DELIMITER & RTrim(.Fields("IDRuta").value), WeekdayName(Weekday(.Fields("FechaHora").value)))
                ListItem.SubItems(1) = Format(.Fields("FechaHora").value, "Short Date")
                ListItem.SubItems(2) = Format(.Fields("FechaHora").value, "Short Time")
                ListItem.SubItems(3) = RTrim(.Fields("IDRuta").value) & IIf(RTrim(.Fields("IDRuta").value) = pParametro.Ruta_ID_Otra Or RTrim(.Fields("IDRuta").value) = pParametro.Ruta_Paquete_ID, ": " & .Fields("RutaOtra").value, "")
                ListItem.SubItems(4) = .Fields("Vehiculo").value & ""
                ListItem.SubItems(5) = .Fields("Conductor").value & ""
                Viaje.Estado = .Fields("Estado").value
                If pParametro.Viaje_Permite_2_Conductores Then
                    ListItem.SubItems(6) = .Fields("Conductor2").value & ""
                    ListItem.SubItems(7) = Viaje.Estado_ToString
                    ListItem.SubItems(8) = .Fields("AsientoLibre").value & ""
                    ListItem.SubItems(9) = .Fields("Notas").value & ""
                Else
                    ListItem.SubItems(6) = Viaje.Estado_ToString
                    ListItem.SubItems(7) = .Fields("AsientoLibre").value & ""
                    ListItem.SubItems(8) = .Fields("Notas").value & ""
                End If
                
                If RTrim(.Fields("IDRuta").value) = pParametro.Ruta_ID_Otra Or RTrim(.Fields("IDRuta").value) = pParametro.Ruta_Paquete_ID Then
                    ListItem.Bold = pParametro.Viaje_Especial_Bold
                    ListItem.ForeColor = pParametro.Viaje_Especial_Color
                Else
                    ListItem.Bold = Viaje.Estado_ToBold
                    ListItem.ForeColor = Viaje.Estado_ToColor
                End If
                .MoveNext
            Loop
            
            stbMain.Panels("TEXT").Text = .RecordCount & " items" & IIf(.RecordCount = .MaxRecords, " (Limitados)", "")
        Else
            stbMain.Panels("TEXT").Text = "No hay items."
        End If
        .Close
    End With
    Set recData = Nothing
    
    Set Viaje = Nothing
    
    On Error Resume Next
    lvwData.SelectedItem.Selected = False
    Set lvwData.SelectedItem = lvwData.ListItems(KeySave)
    lvwData.SelectedItem.EnsureVisible
    
    If Not CKeySave Is Nothing Then
        If CKeySave.Count > 1 Then
            For Each KeySave In CKeySave
                lvwData.ListItems(KeySave).Selected = True
            Next KeySave
        End If
    End If

    If frmMDI.ActiveForm.Name = Me.Name And GetForegroundWindow() = frmMDI.hwnd And frmMDI.WindowState <> vbMinimized Then
        lvwData.SetFocus
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.Viaje.FillListView", "Error al leer la lista de Viajes."
End Sub

Public Sub FillComboBoxRuta()
    Dim recRuta As ADODB.Recordset
    Dim KeySave As String
    Dim Index As Long
    Dim IDRutaGrupo As Long
    
    KeySave = cboRuta.Text

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    'GRUPO DE RUTAS
    For Index = 0 To optRutaGrupo.Count - 1
        If optRutaGrupo(Index).value Then
            IDRutaGrupo = Val(optRutaGrupo(Index).Tag)
            Exit For
        End If
    Next Index
    
    Set recRuta = New ADODB.Recordset
    If IDRutaGrupo = 0 Then
        recRuta.Source = "SELECT IDRuta FROM Ruta WHERE Activo = 1" & IIf(pCPermiso.RutaWhere <> "", " AND " & Replace(pCPermiso.RutaWhere, "%TABLENAME%", "Ruta"), "") & " ORDER BY IDRuta"
    Else
        recRuta.Source = "SELECT IDRuta FROM Ruta WHERE IDRutaGrupo = " & IDRutaGrupo & " AND Activo = 1" & IIf(pCPermiso.RutaWhere <> "", " AND " & Replace(pCPermiso.RutaWhere, "%TABLENAME%", "Ruta"), "") & " ORDER BY IDRuta"
    End If
    recRuta.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    cboRuta.Clear
    cboRuta.AddItem ITEM_ALL_FEMALE
    Do While Not recRuta.EOF
        cboRuta.AddItem RTrim(recRuta("IDRuta").value)
        recRuta.MoveNext
    Loop
    recRuta.Close
    Set recRuta = Nothing

    cboRuta.ListIndex = CSM_Control_ComboBox.GetListIndexByText(cboRuta, KeySave, cscpItemOrFirst)
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.Viaje.FillComboBoxRuta", "Error al leer la lista de Rutas."
End Sub

Private Sub cboFecha_Click()
    txtDiaSemana.Visible = (cboFecha.ListIndex > 0 And cboFecha.ListIndex < 7)
    cmdAnteriorDesde.Visible = (cboFecha.ListIndex > 0)
    dtpFechaDesde.Visible = (cboFecha.ListIndex > 0)
    cmdSiguienteDesde.Visible = (cboFecha.ListIndex > 0)
    cmdHoyDesde.Visible = (cboFecha.ListIndex > 0)
    
    lblFechaAnd.Visible = (cboFecha.ListIndex = 7)
    
    cmdAnteriorHasta.Visible = (cboFecha.ListIndex = 7)
    dtpFechaHasta.Visible = (cboFecha.ListIndex = 7)
    cmdSiguienteHasta.Visible = (cboFecha.ListIndex = 7)
    cmdHoyHasta.Visible = (cboFecha.ListIndex = 7)
    
    cmdAnteriorDesde.Left = 1680
    dtpFechaDesde.Left = 1980
    cmdSiguienteDesde.Left = 3420
    cmdHoyDesde.Left = 3720
    
    If cboFecha.ListIndex > 0 And cboFecha.ListIndex < 7 Then
        cmdAnteriorDesde.Left = cmdAnteriorDesde.Left + txtDiaSemana.Width
        dtpFechaDesde.Left = dtpFechaDesde.Left + txtDiaSemana.Width
        cmdSiguienteDesde.Left = cmdSiguienteDesde.Left + txtDiaSemana.Width
        cmdHoyDesde.Left = cmdHoyDesde.Left + txtDiaSemana.Width
    End If
    
    FillListView Now, ""
End Sub

Private Sub cboFilterEstado_Click()
    FillListView Now, ""
End Sub

Private Sub cboRuta_Click()
    FillListView Now, ""
End Sub

Private Sub cbrMain_HeightChanged(ByVal NewHeight As Single)
    ResizeControls NewHeight
End Sub

Private Sub cboDiaSemana_Click()
    FillListView Now, ""
End Sub

Private Sub dtpFechaDesde_Change()
    txtDiaSemana.Text = WeekdayName(Weekday(dtpFechaDesde.value))
    FillListView Now, ""
End Sub

Private Sub cmdAnteriorDesde_Click()
    dtpFechaDesde.value = DateAdd("d", -1, dtpFechaDesde.value)
    dtpFechaDesde.SetFocus
    dtpFechaDesde_Change
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

Private Sub dtpFechaHasta_Change()
    FillListView Now, ""
End Sub

Private Sub cmdAnteriorHasta_Click()
    dtpFechaHasta.value = DateAdd("d", -1, dtpFechaHasta.value)
    dtpFechaHasta.SetFocus
    dtpFechaHasta_Change
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

Private Sub Form_Initialize()
    Set CSelectEstadosFilter = New Collection
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
            Case vbKeyN
                tlbMain_ButtonClick tlbMain.Buttons.Item("NEW")
            Case vbKeyP
                tlbMain_ButtonClick tlbMain.Buttons.Item("PROPERTIES")
            Case vbKeyE
                tlbMain_ButtonClick tlbMain.Buttons.Item("DELETE")
            Case vbKeyD
                tlbMain_ButtonClick tlbMain.Buttons.Item("DETAIL")
            Case vbKeyS
                tlbMain_ButtonClick tlbMain.Buttons.Item("SELECT")
        End Select
    End If
End Sub

Private Sub Form_Load()
    Dim DiaSemana As Byte
    Dim RutaGrupo As RutaGrupo
    Dim Index  As Long
    
    lvwData.GridLines = pParametro.ListView_GridLines
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL TOOLBAR
    Set tlbMain.ImageList = frmMDI.ilsFormToolbar
    Set tlbMain.HotImageList = frmMDI.ilsFormToolbarHot
    tlbMain.Buttons("GENERATE").Image = "VIAJE_GENERATE"
    tlbMain.Buttons("NEW").Image = "NEW"
    tlbMain.Buttons("PROPERTIES").Image = "PROPERTIES"
    tlbMain.Buttons("DELETE").Image = "DELETE"
    tlbMain.Buttons("SELECT").Image = "SELECT"
    tlbMain.Buttons("DETAIL").Image = "DETAIL"
    tlbMain.Buttons("CHANGE_STATUS").Image = "CHANGE_STATUS"
    tlbMain.Buttons("EMAIL").Image = "EMAIL"
    
    cbrMain.Bands("Toolbar").MinWidth = CSM_Control_Toolbar.GetTotalWidth(tlbMain)
    '//////////////////////////////////////////////////////////
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL LISTVIEW
    Set lvwData.ColumnHeaderIcons = frmMDI.ilsFormSortColumn
    '//////////////////////////////////////////////////////////
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL PIN
    Set tlbPin.ImageList = frmMDI.ilsFormPin
    '//////////////////////////////////////////////////////////
    
    mLoading = True
    
    cboDiaSemana.AddItem ITEM_ALL_MALE
    For DiaSemana = 1 To 7
        cboDiaSemana.AddItem WeekdayName(DiaSemana)
    Next DiaSemana
    cboDiaSemana.ListIndex = 0
    
    cboFecha.AddItem "<Todas>"
    cboFecha.AddItem "="
    cboFecha.AddItem ">"
    cboFecha.AddItem ">="
    cboFecha.AddItem "<"
    cboFecha.AddItem "<="
    cboFecha.AddItem "<>"
    cboFecha.AddItem "Entre"
    cboFecha.ListIndex = 1
    
    dtpFechaDesde.value = Date
    txtDiaSemana.Text = WeekdayName(Weekday(dtpFechaDesde.value))
    dtpFechaHasta.value = Date
    
    'GRUPOS DE RUTAS
    Set RutaGrupo = New RutaGrupo
    Set RutaGrupo.Database = pDatabase
    If RutaGrupo.LoadFirst() Then
        optRutaGrupo(0).Caption = RutaGrupo.Nombre
        optRutaGrupo(0).Tag = RutaGrupo.IDRutaGrupo
        Do While RutaGrupo.LoadNext
            Index = Index + 1
            Load optRutaGrupo(Index)
            With optRutaGrupo(Index)
                .Top = optRutaGrupo(0).Top
                .Left = optRutaGrupo(Index - 1).Left + optRutaGrupo(Index - 1).Width + 60
                .Width = optRutaGrupo(0).Width
                .Caption = RutaGrupo.Nombre
                .Tag = RutaGrupo.IDRutaGrupo
                .Visible = True
            End With
        Loop
        cbrMain.Bands("FilterRutaGrupo").MinWidth = optRutaGrupo(Index).Left + optRutaGrupo(Index).Width
    End If
    Set RutaGrupo = Nothing
    
    FillComboBoxRuta
    cboRuta.ListIndex = 0
    
    cboFilterEstado.AddItem ITEM_ALL_MALE
    cboFilterEstado.AddItem VIAJE_ESTADO_ACTIVO_NOMBRE
    cboFilterEstado.AddItem VIAJE_ESTADO_EN_PROGRESO_NOMBRE
    cboFilterEstado.AddItem VIAJE_ESTADO_FINALIZADO_NOMBRE
    cboFilterEstado.AddItem VIAJE_ESTADO_CANCELADO_NOMBRE
    cboFilterEstado.ListIndex = 0
        
    CSM_Forms.ResizeAndPosition frmMDI, Me
    pParametro.GetCoolBarSettings "Viaje", cbrMain
    pParametro.GetListViewSettings "Viaje", lvwData
    lvwData.ColumnHeaders(lvwData.SortKey + 1).Icon = lvwData.SortOrder + 1
    tlbPin.Buttons("PIN").value = pParametro.Usuario_LeerNumero("Viaje_Pin", tlbPin.Buttons("PIN").value)
    If tlbPin.Buttons("PIN").value = tbrUnpressed Then
        tlbPin.Buttons("PIN").Image = 1
    Else
        tlbPin.Buttons("PIN").Image = 2
    End If
    
    If Not pParametro.Viaje_Permite_2_Conductores Then
        lvwData.ColumnHeaders.Remove ("Conductor2")
    End If

    mLoading = False

    FillListView Now, ""
End Sub

Private Sub Form_Resize()
    ResizeControls cbrMain.Height
End Sub

Private Sub Form_Terminate()
    Set CSelectEstadosFilter = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visible = False
    WindowState = vbNormal
    pParametro.SaveCoolBarSettings "Viaje", cbrMain
    pParametro.SaveListViewSettings "Viaje", lvwData
    pParametro.Usuario_GuardarNumero "Viaje_Pin", tlbPin.Buttons("PIN").value
    Set frmViaje = Nothing
End Sub

Private Sub lvwData_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    '//////////////////////////////////////////////////////////////////
    'Debido a que no se puede ordenar por Disponibles, ignoro el click en la columna
    If ColumnHeader.Key = "Disponibles" Then
        Exit Sub
    End If
    lvwData.ColumnHeaders(lvwData.SortKey + 1).Icon = 0
    If ColumnHeader.Index - 1 = lvwData.SortKey Then
        lvwData.SortOrder = IIf(lvwData.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lvwData.SortKey = ColumnHeader.Index - 1
        lvwData.SortOrder = lvwAscending
    End If
    ColumnHeader.Icon = lvwData.SortOrder + 1
    FillListView Date, ""
End Sub

Private Sub lvwData_DblClick()
    If GetFormIndex(FormWaitingForSelect) > 0 Then
        tlbMain_ButtonClick tlbMain.Buttons.Item("SELECT")
    Else
        tlbMain_ButtonClick tlbMain.Buttons.Item("DETAIL")
    End If
End Sub

Private Sub lvwData_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        lvwData_DblClick
    End If
End Sub

Private Sub optRutaGrupo_Click(Index As Integer)
    Call FillComboBoxRuta
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim FormIndex As Long
    Dim ItemIndex As Long
    Dim SelectedItemCount As Long
    
    Dim Viaje As Viaje
    Dim CFechaHora As Collection
    Dim CIDRuta As Collection
    Dim IDRuta As String
    
    Dim SelectedItems As Collection
    Dim EstadoFilter As Variant
    Dim EstadoFilterFounded As Boolean
    
    Select Case Button.Key
        Case "GENERATE"
            If pCPermiso.GotPermission(PERMISO_VIAJE_GENERATE) Then
                Screen.MousePointer = vbHourglass
                frmViajeGenerar.LoadDataAndShow Me, IIf(cboFecha.ListIndex > 0, dtpFechaDesde.value, Date)
                Screen.MousePointer = vbDefault
            End If
        Case "NEW"
            If pCPermiso.GotPermission(PERMISO_VIAJE_ADD) Then
                Screen.MousePointer = vbHourglass
                Set Viaje = New Viaje
                Viaje.FechaHora = IIf(cboFecha.ListIndex > 0, dtpFechaDesde.value, Date)
                frmViajePropiedad.LoadDataAndShow Me, Viaje
                Set Viaje = Nothing
                Screen.MousePointer = vbDefault
            End If
        Case "PROPERTIES"
            If pCPermiso.GotPermission(PERMISO_VIAJE_MODIFY) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                SelectedItemCount = 0
                For ItemIndex = 1 To lvwData.ListItems.Count
                    If lvwData.ListItems(ItemIndex).Selected Then
                        SelectedItemCount = SelectedItemCount + 1
                    End If
                Next ItemIndex
                
                Select Case SelectedItemCount
                    Case 0
                        MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                        lvwData.SetFocus
                        Exit Sub
                    Case 1
                        Screen.MousePointer = vbHourglass
                        Set Viaje = New Viaje
                        Viaje.FechaHora = CDate(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                        Viaje.IDRuta = CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER)
                        If Viaje.Load() Then
                            frmViajePropiedad.LoadDataAndShow Me, Viaje
                        End If
                        Set Viaje = Nothing
                        Screen.MousePointer = vbDefault
                    Case Else
                        Set CFechaHora = New Collection
                        Set CIDRuta = New Collection
                        For ItemIndex = 1 To lvwData.ListItems.Count
                            If lvwData.ListItems(ItemIndex).Selected Then
                                CFechaHora.Add CDate(GetSubString(Mid(lvwData.ListItems(ItemIndex).Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                                CIDRuta.Add ReplaceQuote(GetSubString(Mid(lvwData.ListItems(ItemIndex).Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER))
                            End If
                        Next ItemIndex
                        Screen.MousePointer = vbHourglass
                        frmViajePropiedadMultiple.LoadDataAndShow Me, CFechaHora, CIDRuta
                        Screen.MousePointer = vbDefault
                        Set CFechaHora = Nothing
                        Set CIDRuta = Nothing
                End Select
            End If
        Case "DELETE"
            If pCPermiso.GotPermission(PERMISO_VIAJE_DELETE) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                SelectedItemCount = 0
                For ItemIndex = 1 To lvwData.ListItems.Count
                    If lvwData.ListItems(ItemIndex).Selected Then
                        SelectedItemCount = SelectedItemCount + 1
                    End If
                Next ItemIndex
                
                Set Viaje = New Viaje
                If MsgBox(IIf(SelectedItemCount = 1, "¿Desea eliminar el Viaje seleccionado?" & vbCr & vbCr & "Se eliminarán todas las Reservas del Viaje.", "¿Desea eliminar los " & SelectedItemCount & " Viajes seleccionados?" & vbCr & vbCr & "Se eliminarán todas las Reservas de los Viajes."), vbQuestion + vbYesNo, App.Title) = vbYes Then
                    Screen.MousePointer = vbHourglass
                    If SelectedItemCount = 1 Then
                        Viaje.FechaHora = CDate(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                        Viaje.IDRuta = ReplaceQuote(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER))
                        If Viaje.Load() Then
                            Viaje.Delete
                        End If
                    Else
                        Viaje.RefreshListSkip = True
                        For ItemIndex = 1 To lvwData.ListItems.Count
                            If lvwData.ListItems(ItemIndex).Selected Then
                                Viaje.FechaHora = CDate(GetSubString(Mid(lvwData.ListItems(ItemIndex).Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                                Viaje.IDRuta = ReplaceQuote(GetSubString(Mid(lvwData.ListItems(ItemIndex).Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER))
                                If Viaje.Load() Then
                                    Viaje.Delete
                                End If
                            End If
                        Next ItemIndex
                        RefreshList_RefreshCuentaCorriente 0
                        RefreshList_RefreshViaje Now, ""
                    End If
                    Screen.MousePointer = vbDefault
                End If
                Set Viaje = Nothing
            End If
        Case "SELECT"
            FormIndex = GetFormIndex(FormWaitingForSelect)
            If FormIndex >= 0 Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                If Not AllowMultipleSelect Then
                    SelectedItemCount = 0
                    For ItemIndex = 1 To lvwData.ListItems.Count
                        If lvwData.ListItems(ItemIndex).Selected Then
                            SelectedItemCount = SelectedItemCount + 1
                            If SelectedItemCount > 1 Then
                                MsgBox "No se puede Seleccionar más de un Viaje a la vez.", vbExclamation, App.Title
                                Exit Sub
                            End If
                        End If
                    Next ItemIndex
                
                    Set Viaje = New Viaje
                    Viaje.FechaHora = CDate(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                    Viaje.IDRuta = ReplaceQuote(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER))
                    If Not Viaje.Load() Then
                        Set Viaje = Nothing
                        Exit Sub
                    End If
                    
                    EstadoFilterFounded = False
                    If Not CSelectEstadosFilter Is Nothing Then
                        For Each EstadoFilter In CSelectEstadosFilter
                            If EstadoFilter = Viaje.Estado Then
                                EstadoFilterFounded = True
                            End If
                        Next EstadoFilter
                        If Not EstadoFilterFounded Then
                            MsgBox "No se permite seleccionar Viajes con Estado " & Viaje.Estado_ToString & ".", vbInformation, App.Title
                            Set Viaje = Nothing
                            lvwData.SetFocus
                            Exit Sub
                        End If
                    End If
                    Set CSelectEstadosFilter = New Collection
                    
                    Screen.MousePointer = vbHourglass
                    Forms(FormIndex).ViajeSelected Viaje, SelectTag
                    Set Viaje = Nothing
                Else
                    Screen.MousePointer = vbHourglass
                    Set SelectedItems = New Collection
                    For ItemIndex = 1 To lvwData.ListItems.Count
                        If lvwData.ListItems(ItemIndex).Selected Then
                            Set Viaje = New Viaje
                            Viaje.FechaHora = CDate(GetSubString(Mid(lvwData.ListItems(ItemIndex).Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                            Viaje.IDRuta = ReplaceQuote(GetSubString(Mid(lvwData.ListItems(ItemIndex).Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER))
                            If Not Viaje.Load() Then
                                Set Viaje = Nothing
                                Exit Sub
                            End If
                        
                            EstadoFilterFounded = False
                            If Not CSelectEstadosFilter Is Nothing Then
                                For Each EstadoFilter In CSelectEstadosFilter
                                    If EstadoFilter = Viaje.Estado Then
                                        EstadoFilterFounded = True
                                    End If
                                Next EstadoFilter
                                If Not EstadoFilterFounded Then
                                    MsgBox "No se permite seleccionar Viajes con Estado " & Viaje.Estado_ToString & ".", vbInformation, App.Title
                                    Set Viaje = Nothing
                                    lvwData.SetFocus
                                    Exit Sub
                                End If
                            End If
                    
                            If AllowMultipleRuta = False And IDRuta <> Viaje.IDRuta Then
                                If IDRuta <> "" Then
                                    Screen.MousePointer = vbDefault
                                    MsgBox "No se pueden seleccionar Viajes de Distintas Rutas.", vbInformation, App.Title
                                    Set SelectedItems = Nothing
                                    Exit Sub
                                End If
                                IDRuta = Viaje.IDRuta
                            End If
                            
                            SelectedItems.Add lvwData.ListItems(ItemIndex).Key
                        End If
                    Next ItemIndex
                    Set CSelectEstadosFilter = New Collection
                    Forms(FormIndex).MultipleViajeSelected SelectedItems
                    Set SelectedItems = Nothing
                End If
                
                'Forms(FormIndex).SetFocus
                If tlbPin.Buttons("PIN").value = tbrUnpressed Then
                    Unload Me
                End If
                Screen.MousePointer = vbDefault
            End If
            FormWaitingForSelect = ""
        Case "DETAIL"
            If pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                SelectedItemCount = 0
                For ItemIndex = 1 To lvwData.ListItems.Count
                    If lvwData.ListItems(ItemIndex).Selected Then
                        SelectedItemCount = SelectedItemCount + 1
                        If SelectedItemCount > 1 Then
                            MsgBox "No se puede mostrar el Detalle de más de un Viaje a la vez.", vbExclamation, App.Title
                            lvwData.SetFocus
                            Exit Sub
                        End If
                    End If
                Next ItemIndex
                
                Screen.MousePointer = vbHourglass
                
                Set Viaje = New Viaje
                Viaje.FechaHora = CDate(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                Viaje.IDRuta = ReplaceQuote(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER))
                If Viaje.Load() Then
                    frmViajeDetalle.LoadDataAndShow Viaje
                End If
                Set Viaje = Nothing
                
                Screen.MousePointer = vbDefault
            End If
        Case "CHANGE_STATUS"
            If pCPermiso.GotPermission(PERMISO_VIAJE_CHANGE_STATUS) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                SelectedItemCount = 0
                For ItemIndex = 1 To lvwData.ListItems.Count
                    If lvwData.ListItems(ItemIndex).Selected Then
                        SelectedItemCount = SelectedItemCount + 1
                        If SelectedItemCount > 1 Then
                            MsgBox "No se puede Cambiar el Estado de más de un Viaje a la vez.", vbExclamation, App.Title
                            lvwData.SetFocus
                            Exit Sub
                        End If
                    End If
                Next ItemIndex
                
                Screen.MousePointer = vbHourglass
                frmViajeCambiarEstado.LoadDataAndShow Me, CDate(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER)), CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER)
                Screen.MousePointer = vbDefault
            End If
        Case "EMAIL"
            Call tlbMain_ButtonMenuClick(Button.ButtonMenus("EMAIL_OBSERVACIONES"))
    End Select
End Sub

Private Sub tlbMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim SelectedItemCount As Long
    Dim ItemIndex As Long
    Dim Viaje As Viaje
    
    If pCPermiso.GotPermission(PERMISO_VIAJE_SEND_EMAIL) Then
        If lvwData.SelectedItem Is Nothing Then
            MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
            lvwData.SetFocus
            Exit Sub
        End If
        
        SelectedItemCount = 0
        For ItemIndex = 1 To lvwData.ListItems.Count
            If lvwData.ListItems(ItemIndex).Selected Then
                SelectedItemCount = SelectedItemCount + 1
                If SelectedItemCount > 1 Then
                    MsgBox "No se puede Enviar más de un Viaje a la vez.", vbExclamation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
            End If
        Next ItemIndex
        
        Screen.MousePointer = vbHourglass
        
        Set Viaje = New Viaje
        Viaje.FechaHora = CDate(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
        Viaje.IDRuta = ReplaceQuote(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER))
        If Not Viaje.Load() Then
            Set Viaje = Nothing
            Exit Sub
        End If
        
        Select Case Viaje.Estado
            Case VIAJE_ESTADO_ACTIVO, VIAJE_ESTADO_FINALIZADO
                Screen.MousePointer = vbDefault
                If MsgBox("Este Viaje no está En Progreso." & vbCr & vbCr & "¿Desea enviarlo de todos modos?", vbExclamation + vbYesNo, App.Title) = vbNo Then
                    Set Viaje = Nothing
                    Exit Sub
                End If
            Case VIAJE_ESTADO_CANCELADO
                Screen.MousePointer = vbDefault
                MsgBox "Este Viaje está Cancelado, no se puede enviar el E-mail.", vbExclamation, App.Title
                Set Viaje = Nothing
                Exit Sub
            Case VIAJE_ESTADO_EN_PROGRESO
                Screen.MousePointer = vbDefault
                If MsgBox("¿Desea enviar la Planilla de este Viaje por E-mail?", vbQuestion + vbYesNo, App.Title) = vbNo Then
                    Set Viaje = Nothing
                    Exit Sub
                End If
        End Select
    
        If pCPermiso.GotPermission(PERMISO_REPORTE_REPORTE & "Viaje_Planilla_" & Switch(ButtonMenu.Key = "EMAIL_DOCUMENTO", "Documento", ButtonMenu.Key = "EMAIL_OBSERVACIONES", "Observaciones")) Then
            Call PlanillaViajeSendEmail(Viaje, Switch(ButtonMenu.Key = "EMAIL_DOCUMENTO", True, ButtonMenu.Key = "EMAIL_OBSERVACIONES", False))
        End If
    
        Set Viaje = Nothing
        lvwData.SetFocus
    End If
End Sub

Private Sub tlbPin_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.value = tbrUnpressed Then
        Button.Image = 1
    Else
        Button.Image = 2
    End If
End Sub

Private Sub ResizeControls(ByVal CoolBarHeight As Single)
    Const CONTROL_SPACE = 60
    
    On Error Resume Next
    
    lvwData.Top = CoolBarHeight + CONTROL_SPACE
    lvwData.Left = CONTROL_SPACE
    lvwData.Width = ScaleWidth - (CONTROL_SPACE * 2)
    lvwData.Height = ScaleHeight - lvwData.Top - CONTROL_SPACE - stbMain.Height
    
    tlbPin.Top = ScaleHeight - 330
    tlbPin.Left = 15
End Sub

Private Function PlanillaViajeSendEmail(ByRef Viaje As Viaje, ByVal EnviarConDocumento As Boolean) As Boolean
    Dim Reporte As Reporte
    Dim EmailMessage As EmailMessage
    Dim EmailAttachment As EmailAttachment
    
    Dim SucursalNombres As String
    Dim SucursalEmails As String
    
    Const FILENAME_REPORT As String = "PlanillaViaje"
    
    Const FILENAME_EXCEL_DOCUMENTO As String = "PlanillaViaje_Documento.xls"
    Const FILENAME_EXCEL_OBSERVACIONES As String = "PlanillaViaje_Observaciones.xls"
    
    'SELECCIONAR LAS SUCURSALES A ENVIAR
    frmSucursalSelect.Show vbModal, frmMDI
    If frmSucursalSelect.Tag <> "CANCEL" Then
        SucursalNombres = frmSucursalSelect.SucursalNombres
        SucursalEmails = frmSucursalSelect.SucursalEmails
        Unload frmSucursalSelect
        Set frmSucursalSelect = Nothing
    Else
        Unload frmSucursalSelect
        Set frmSucursalSelect = Nothing
        FileDelete pSpecialFolders.Temp & FILENAME_REPORT
        Exit Function
    End If


    'REPORTE
    If pParametro.PlanillaViajeEmail_SendReportFormat > 0 Then
        FileDelete pSpecialFolders.Temp & FILENAME_REPORT
        Set Reporte = New Reporte
        If Not PlanillaViajeGenerateReport(Viaje, Reporte, FILENAME_REPORT, EnviarConDocumento) Then
            Set Reporte = Nothing
            Exit Function
        End If
    End If
    
    'EXCEL DOCUMENTO
    If pParametro.PlanillaViajeEmail_SendExcel And EnviarConDocumento Then
        FileDelete pSpecialFolders.Temp & FILENAME_EXCEL_DOCUMENTO
        If Not PlanillaViajeGenerateExcel(Viaje, True, FILENAME_EXCEL_DOCUMENTO) Then
            If pParametro.PlanillaViajeEmail_SendReportFormat > 0 Then
                Set Reporte = Nothing
                FileDelete pSpecialFolders.Temp & FILENAME_REPORT
            End If
            FileDelete pSpecialFolders.Temp & FILENAME_EXCEL_DOCUMENTO
            Exit Function
        End If
    End If
    
    'EXCEL OBSERVACIONES
    If pParametro.PlanillaViajeEmail_SendExcel Then
        FileDelete pSpecialFolders.Temp & FILENAME_EXCEL_OBSERVACIONES
        If Not PlanillaViajeGenerateExcel(Viaje, False, FILENAME_EXCEL_OBSERVACIONES) Then
            If pParametro.PlanillaViajeEmail_SendReportFormat > 0 Then
                Set Reporte = Nothing
                FileDelete pSpecialFolders.Temp & FILENAME_REPORT
            End If
            FileDelete pSpecialFolders.Temp & FILENAME_EXCEL_OBSERVACIONES
            Exit Function
        End If
    End If
    
    'GUARDO EL EMAIL EN LA BASE DE DATOS
    Set EmailMessage = New EmailMessage
    With EmailMessage
        .DateTime = Now
        .SenderDisplayName = "Sucursal " & pSucursal.Nombre
        .SenderAddress = pSucursal.Email
        .RecipientToDisplayName = SucursalNombres
        .RecipientToAddress = SucursalEmails
        .Subject = "Planilla del Viaje: " & Viaje.FechaHora_Formatted & " - " & Viaje.Ruta_DisplayName
        .SMTPHost = pParametro.PlanillaViajeEmail_SMTPHost
        .SMTPUserName = pParametro.PlanillaViajeEmail_SMTPUserName
        .SMTPPassword = pParametro.PlanillaViajeEmail_SMTPPassword
        If Not .Add() Then
            Set EmailMessage = Nothing
            If pParametro.PlanillaViajeEmail_SendReportFormat > 0 Then
                Set Reporte = Nothing
                FileDelete pSpecialFolders.Temp & FILENAME_REPORT
            End If
            If pParametro.PlanillaViajeEmail_SendExcel Then
                If EnviarConDocumento Then
                    FileDelete pSpecialFolders.Temp & FILENAME_EXCEL_DOCUMENTO
                End If
                FileDelete pSpecialFolders.Temp & FILENAME_EXCEL_OBSERVACIONES
            End If
            Exit Function
        End If
    End With
    
    'GUARDO EL ADJUNTO CON EL REPORTE
    If pParametro.PlanillaViajeEmail_SendReportFormat > 0 Then
        Set EmailAttachment = New EmailAttachment
        EmailAttachment.MessageID = EmailMessage.MessageID
        EmailAttachment.FileName = "PlanillaViaje_" & Format(Viaje.FechaHora, "yyyymmdd-hhnn") & "_" & Viaje.IDRuta
        EmailAttachment.FileExtension = pParametro.PlanillaViajeEmail_SendReportFormat_Extension
        EmailAttachment.FileSourcePath = Reporte.ExportOptions.DiskFileName
        If Not EmailAttachment.Add() Then
            Set EmailAttachment = Nothing
            Set EmailMessage = Nothing
            Set Reporte = Nothing
            FileDelete pSpecialFolders.Temp & FILENAME_REPORT
            If pParametro.PlanillaViajeEmail_SendExcel Then
                If EnviarConDocumento Then
                    FileDelete pSpecialFolders.Temp & FILENAME_EXCEL_DOCUMENTO
                End If
                FileDelete pSpecialFolders.Temp & FILENAME_EXCEL_OBSERVACIONES
            End If
            Exit Function
        End If
        Set Reporte = Nothing
        If pIsCompiled Then
            FileDelete pSpecialFolders.Temp & FILENAME_REPORT
        End If
    End If
    
    'GUARDO EL ADJUNTO CON LA PLANILLA EXCEL CON DOCUMENTO
    If pParametro.PlanillaViajeEmail_SendExcel And EnviarConDocumento Then
        Set EmailAttachment = New EmailAttachment
        EmailAttachment.MessageID = EmailMessage.MessageID
        EmailAttachment.FileName = "PlanillaViaje_" & Format(Viaje.FechaHora, "yyyymmdd-hhnn") & "_" & Viaje.IDRuta & "_Documento"
        EmailAttachment.FileExtension = "xls"
        EmailAttachment.FileSourcePath = pSpecialFolders.Temp & FILENAME_EXCEL_DOCUMENTO
        If Not EmailAttachment.Add() Then
            Set EmailAttachment = Nothing
            Set EmailMessage = Nothing
            If pParametro.PlanillaViajeEmail_SendReportFormat > 0 Then
                Set Reporte = Nothing
                FileDelete pSpecialFolders.Temp & FILENAME_REPORT
            End If
            FileDelete pSpecialFolders.Temp & FILENAME_EXCEL_DOCUMENTO
            FileDelete pSpecialFolders.Temp & FILENAME_EXCEL_OBSERVACIONES
            Exit Function
        End If
        If pIsCompiled Then
            FileDelete pSpecialFolders.Temp & FILENAME_EXCEL_DOCUMENTO
        End If
    End If
            
    'GUARDO EL ADJUNTO CON LA PLANILLA EXCEL CON DOCUMENTO
    If pParametro.PlanillaViajeEmail_SendExcel Then
        Set EmailAttachment = New EmailAttachment
        EmailAttachment.MessageID = EmailMessage.MessageID
        EmailAttachment.FileName = "PlanillaViaje_" & Format(Viaje.FechaHora, "yyyymmdd-hhnn") & "_" & Viaje.IDRuta & "_Observaciones"
        EmailAttachment.FileExtension = "xls"
        EmailAttachment.FileSourcePath = pSpecialFolders.Temp & FILENAME_EXCEL_OBSERVACIONES
        If Not EmailAttachment.Add() Then
            Set EmailAttachment = Nothing
            Set EmailMessage = Nothing
            If pParametro.PlanillaViajeEmail_SendReportFormat > 0 Then
                Set Reporte = Nothing
                FileDelete pSpecialFolders.Temp & FILENAME_REPORT
            End If
            If EnviarConDocumento Then
                FileDelete pSpecialFolders.Temp & FILENAME_EXCEL_DOCUMENTO
            End If
            FileDelete pSpecialFolders.Temp & FILENAME_EXCEL_OBSERVACIONES
            Exit Function
        End If
        If pIsCompiled Then
            FileDelete pSpecialFolders.Temp & FILENAME_EXCEL_OBSERVACIONES
        End If
    End If
            
    'ACTUALIZO EL EMAIL
    EmailMessage.ReadyToSend = True
    EmailMessage.Update
            
    Set EmailAttachment = Nothing
    Set EmailMessage = Nothing
            
    MsgBox "En instantes, se enviará el E-mail de este Viaje.", vbInformation, App.Title
End Function

Private Function PlanillaViajeGenerateReport(ByRef Viaje As Viaje, ByRef Reporte As Reporte, ByVal FileName As String, ByVal EnviarConDocumento As Boolean) As Boolean
    With Reporte
        If EnviarConDocumento Then
            .IDReporte = "Viaje_Planilla_Documento"
        Else
            .IDReporte = "Viaje_Planilla_Observaciones"
        End If
        If .Load() Then
            .Titulo = "Detalle del Viaje: " & Viaje.FechaHora_Formatted & " - " & Viaje.Ruta_DisplayName
            .Parametros("FechaHora_FILTER").Valor = Viaje.FechaHora
            .Parametros("IDRuta_FILTER").Valor = Viaje.IDRuta

            If .OpenReport() Then
                .ExportOptions.DestinationType = crEDTDiskFile
                .ExportOptions.DiskFileName = pSpecialFolders.Temp & FileName
                .ExportOptions.formattype = pParametro.PlanillaViajeEmail_SendReportFormat
                PlanillaViajeGenerateReport = .ExportReport()
            End If
        End If
    End With
End Function

Private Function PlanillaViajeGenerateExcel(ByRef Viaje As Viaje, ByVal MostrarDocumento As Boolean, ByVal FileName As String) As Boolean
    Dim ExcelApplication As Object
    Dim ExcelWorkbook As Object
    Dim ExcelWorksheet As Object
'    Dim ExcelApplication As Excel.Application
'    Dim ExcelWorkbook As Excel.Workbook
'    Dim ExcelWorksheet As Excel.Worksheet
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    Dim RowNumber As Integer
    Dim RowNumberDetalleStart As Integer
    Dim RowNumberDetalleEnd As Integer
'    Dim Index As Long
    Dim Ruta As Ruta
    Dim LugarOrigen As Lugar
    Dim LugarDestino As Lugar
    
    Screen.MousePointer = vbHourglass
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    'INICIO UNA SESION DE EXCEL
    Set ExcelApplication = CreateObject("Excel.Application")
    ExcelApplication.Visible = Not pIsCompiled
    
    'ABRO EL ARCHIVO ESPECIFICADO
    Set ExcelWorkbook = ExcelApplication.Workbooks.Add
    Set ExcelWorksheet = ExcelWorkbook.Worksheets(1)
    
    '//////////////////////////////////////////////////////////
    'ANCHO DE LAS COLUMNAS
    ExcelWorksheet.Range("A1").ColumnWidth = pParametro.PlanillaViajeEmail_ColumnWidth_A
    ExcelWorksheet.Range("B1").ColumnWidth = pParametro.PlanillaViajeEmail_ColumnWidth_B
    ExcelWorksheet.Range("C1").ColumnWidth = pParametro.PlanillaViajeEmail_ColumnWidth_C
    ExcelWorksheet.Range("D1").ColumnWidth = pParametro.PlanillaViajeEmail_ColumnWidth_D
    ExcelWorksheet.Range("E1").ColumnWidth = pParametro.PlanillaViajeEmail_ColumnWidth_E
    ExcelWorksheet.Range("F1").ColumnWidth = pParametro.PlanillaViajeEmail_ColumnWidth_F
    ExcelWorksheet.Range("G1").ColumnWidth = pParametro.PlanillaViajeEmail_ColumnWidth_G
    ExcelWorksheet.Range("H1").ColumnWidth = pParametro.PlanillaViajeEmail_ColumnWidth_H
    ExcelWorksheet.Range("I1").ColumnWidth = pParametro.PlanillaViajeEmail_ColumnWidth_I
    ExcelWorksheet.Range("J1").ColumnWidth = pParametro.PlanillaViajeEmail_ColumnWidth_J
    ExcelWorksheet.Range("K1").ColumnWidth = pParametro.PlanillaViajeEmail_ColumnWidth_K
    
    '//////////////////////////////////////////////////////////
    'ALTO DE LAS FILAS
    ExcelWorksheet.Range("A2").RowHeight = pParametro.PlanillaViajeEmail_RowHeight_2
    ExcelWorksheet.Range("A4").RowHeight = pParametro.PlanillaViajeEmail_RowHeight_4
    ExcelWorksheet.Range("A6").RowHeight = pParametro.PlanillaViajeEmail_RowHeight_6
    ExcelWorksheet.Range("A8").RowHeight = pParametro.PlanillaViajeEmail_RowHeight_8
    ExcelWorksheet.Range("A35").RowHeight = pParametro.PlanillaViajeEmail_RowHeight_35
    ExcelWorksheet.Range("A37").RowHeight = pParametro.PlanillaViajeEmail_RowHeight_37
    'ExcelWorksheet.Range("A45").RowHeight = pParametro.PlanillaViajeEmail_RowHeight_45
    ExcelWorksheet.Range("A9").RowHeight = pParametro.PlanillaViajeEmail_RowHeight_ColumnHeader
    ExcelWorksheet.Range("A10", "A34").RowHeight = pParametro.PlanillaViajeEmail_RowHeight_Detalle
    ExcelWorksheet.Range("A38").RowHeight = pParametro.PlanillaViajeEmail_RowHeight_ColumnHeader
    ExcelWorksheet.Range("A39", "A44").RowHeight = pParametro.PlanillaViajeEmail_RowHeight_Detalle
    
    '//////////////////////////////////////////////////////////
    'MARGENES
    ExcelWorksheet.PageSetup.TopMargin = pParametro.PlanillaViajeEmail_Margin_Top
    ExcelWorksheet.PageSetup.LeftMargin = pParametro.PlanillaViajeEmail_Margin_Left
    ExcelWorksheet.PageSetup.RightMargin = pParametro.PlanillaViajeEmail_Margin_Right
    ExcelWorksheet.PageSetup.BottomMargin = pParametro.PlanillaViajeEmail_Margin_Bottom
    
    '//////////////////////////////////////////////////////////
    'APARIENCIA DE LA PLANILLA
    ExcelWorksheet.Name = "Planilla del Viaje"
'    For Index = 2 To ExcelWorkbook.Worksheets.Count
'        ExcelWorkbook.Worksheets(2).Delete
'    Next Index
    ExcelWorksheet.Range("L1", "IV1").Columns.Hidden = True
    ExcelWorksheet.Range("A47", "A65536").Rows.Hidden = True
    ExcelWorksheet.Range("A1", "K46").Interior.Color = vbWhite
    
    '//////////////////////////////////////////////////////////
    'HEADER
    RowNumber = 1
    With ExcelWorksheet.Range("A" & RowNumber)
        .value = pParametro.CompanyName
        .Font.Name = "Arial"
        .Font.Size = 10
        .Font.Bold = True
    End With
    
    With ExcelWorksheet.Range("F" & RowNumber)
        .value = Now
        .Font.Name = "Arial"
        .Font.Size = 10
    End With
    ExcelWorksheet.Range("F" & RowNumber, "K" & RowNumber).Merge
    ExcelWorksheet.Range("F" & RowNumber).HorizontalAlignment = xlRight
    
    RowNumber = RowNumber + 2
    With ExcelWorksheet.Range("A" & RowNumber)
        .value = "Detalle del Viaje: " & Viaje.FechaHora_Formatted & " - " & Viaje.Ruta_DisplayName
        .HorizontalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 12
        .Font.Bold = True
    End With
    ExcelWorksheet.Range("A" & RowNumber, "K" & RowNumber).Merge
        
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_Report_Viaje_Planilla"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("FechaHora_FILTER", adDate, adParamInput, , Viaje.FechaHora)
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta_FILTER", adChar, adParamInput, 20, Viaje.IDRuta)
    Set recData = New ADODB.Recordset
    recData.CursorType = adOpenForwardOnly
    recData.LockType = adLockReadOnly
    recData.Open cmdData
    Set cmdData = Nothing
    
    If Not recData.EOF Then
        RowNumber = RowNumber + 2
        With ExcelWorksheet.Range("A" & RowNumber)
            .value = "Conductor: " & recData("Conductor").value & ""
            .Font.Name = "Arial"
            .Font.Size = 10
            .Font.Bold = True
        End With
        
        With ExcelWorksheet.Range("F" & RowNumber)
            .value = "Vehículo: " & recData("Vehiculo").value & ""
            .Font.Name = "Arial"
            .Font.Size = 10
            .Font.Bold = True
        End With
    End If
    
    'RUTA
    Set Ruta = New Ruta
    Ruta.IDRuta = Viaje.IDRuta
    Call Ruta.Load
    Set LugarOrigen = New Lugar
    LugarOrigen.IDLugar = Ruta.IDOrigen
    Call LugarOrigen.Load
    Set LugarDestino = New Lugar
    LugarDestino.IDLugar = Ruta.IDDestino
    Call LugarDestino.Load
    
    '//////////////////////////////////////////////////////////
    'PASAJEROS
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_Report_Viaje_Planilla_Pasajero"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("FechaHora_FILTER", adDate, adParamInput, , Viaje.FechaHora)
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta_FILTER", adChar, adParamInput, 20, Viaje.IDRuta)
    Set recData = New ADODB.Recordset
    recData.CursorType = adOpenForwardOnly
    recData.LockType = adLockReadOnly
    recData.Open cmdData
    Set cmdData = Nothing
        
    RowNumber = RowNumber + 2
    With ExcelWorksheet.Range("A" & RowNumber)
        .value = "PASAJEROS"
        .HorizontalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 14
        .Font.Bold = True
        .Interior.Color = &HC6C3C6
    End With
    ExcelWorksheet.Range("A" & RowNumber, "K" & RowNumber).Merge
    
    'ENCABEZADOS DE LAS COLUMNAS
    RowNumber = RowNumber + 2
    ExcelWorksheet.Range("A" & RowNumber).value = "#"
    ExcelWorksheet.Range("B" & RowNumber).value = "Apellido y Nombre"
    ExcelWorksheet.Range("B" & RowNumber, "C" & RowNumber).Merge
    ExcelWorksheet.Range("D" & RowNumber).value = "Sube"
    ExcelWorksheet.Range("E" & RowNumber).value = "Baja"
    ExcelWorksheet.Range("F" & RowNumber).value = "Importe"
    ExcelWorksheet.Range("G" & RowNumber).value = "Pagado"
    ExcelWorksheet.Range("H" & RowNumber).value = "Saldo"
    ExcelWorksheet.Range("I" & RowNumber).value = "R"
    If MostrarDocumento Then
        ExcelWorksheet.Range("J" & RowNumber).value = "Documento"
    Else
        ExcelWorksheet.Range("J" & RowNumber).value = "Observaciones"
    End If
    ExcelWorksheet.Range("A" & RowNumber, "J" & RowNumber).VerticalAlignment = xlCenter
    
    With ExcelWorksheet.Range("A" & RowNumber, "J" & RowNumber)
        .HorizontalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Bold = True
    End With
            
    'BORDES
    With ExcelWorksheet.Range("A" & RowNumber, "J" & RowNumber)
        .Borders(xlInsideVertical).Weight = xlThin
        .BorderAround 1, xlMedium
    End With
    
    RowNumberDetalleStart = RowNumber + 1
    Do While Not recData.EOF
        RowNumber = RowNumber + 1
        ExcelWorksheet.Range("A" & RowNumber).value = recData.AbsolutePosition
        ExcelWorksheet.Range("B" & RowNumber).value = recData("Pasajero").value
        If Not (recData("IDOrigen").value = Ruta.IDOrigen And recData("Origen").value = LugarOrigen.Nombre) Then
            ExcelWorksheet.Range("D" & RowNumber).value = recData("Origen").value
        End If
        If Not (recData("IDDestino").value = Ruta.IDDestino And recData("Destino").value = LugarDestino.Nombre) Then
            ExcelWorksheet.Range("E" & RowNumber).value = recData("Destino").value
        End If
        ExcelWorksheet.Range("F" & RowNumber).value = recData("Importe").value
        If recData("ImportePagado").value = 0 Then
            ExcelWorksheet.Range("G" & RowNumber).Locked = False
            ExcelWorksheet.Range("G" & RowNumber).Font.Color = vbRed
        Else
            ExcelWorksheet.Range("G" & RowNumber).value = recData("ImportePagado").value
        End If
        If recData("ImprimirSaldo").value And Not IsNull(recData("SaldoActual").value) Then
            ExcelWorksheet.Range("H" & RowNumber).value = recData("SaldoActual").value
        End If
        If IsNull(recData("Realizado").value) Then
            ExcelWorksheet.Range("I" & RowNumber).Locked = False
            ExcelWorksheet.Range("I" & RowNumber).Font.Color = vbRed
        Else
            If recData("Realizado").value Then
                ExcelWorksheet.Range("I" & RowNumber).value = "x"
            End If
        End If
        If MostrarDocumento Then
            If IsNull(recData("Documento").value) Then
                ExcelWorksheet.Range("J" & RowNumber).Locked = False
                ExcelWorksheet.Range("J" & RowNumber).Font.Color = vbRed
            Else
                ExcelWorksheet.Range("J" & RowNumber).value = recData("Documento").value
            End If
        Else
            If IsNull(recData("Notas").value) Then
                ExcelWorksheet.Range("J" & RowNumber).Locked = False
                ExcelWorksheet.Range("J" & RowNumber).Font.Color = vbRed
            Else
                ExcelWorksheet.Range("J" & RowNumber).value = recData("Notas").value
            End If
        End If
        
        recData.MoveNext
    Loop
    RowNumberDetalleEnd = RowNumberDetalleStart + 24
    
    'DESPROTEJO LAS CELDAS VACIAS Y CAMBIO EL COLOR DE LA TIPOGRAFIA EN LAS CELDAS VACIAS
    If RowNumber < RowNumberDetalleEnd Then
        With ExcelWorksheet.Range("A" & RowNumber + 1, "J" & RowNumberDetalleEnd)
            .Locked = False
            .Font.Color = vbRed
        End With
    End If
        
    'ALINEACION
    ExcelWorksheet.Range("A" & RowNumberDetalleStart, "A" & RowNumberDetalleEnd).HorizontalAlignment = xlCenter
    ExcelWorksheet.Range("B" & RowNumberDetalleStart, "B" & RowNumberDetalleEnd).HorizontalAlignment = xlLeft
    ExcelWorksheet.Range("D" & RowNumberDetalleStart, "D" & RowNumberDetalleEnd).HorizontalAlignment = xlLeft
    ExcelWorksheet.Range("E" & RowNumberDetalleStart, "E" & RowNumberDetalleEnd).HorizontalAlignment = xlLeft
    ExcelWorksheet.Range("F" & RowNumberDetalleStart, "F" & RowNumberDetalleEnd).HorizontalAlignment = xlRight
    ExcelWorksheet.Range("G" & RowNumberDetalleStart, "G" & RowNumberDetalleEnd).HorizontalAlignment = xlRight
    ExcelWorksheet.Range("H" & RowNumberDetalleStart, "H" & RowNumberDetalleEnd).HorizontalAlignment = xlRight
    ExcelWorksheet.Range("I" & RowNumberDetalleStart, "I" & RowNumberDetalleEnd).HorizontalAlignment = xlCenter
    ExcelWorksheet.Range("J" & RowNumberDetalleStart, "J" & RowNumberDetalleEnd).HorizontalAlignment = xlLeft
    ExcelWorksheet.Range("A" & RowNumberDetalleStart, "J" & RowNumberDetalleEnd).VerticalAlignment = xlCenter
    
    'HAGO MERGE DE LOS ROWS
    For RowNumber = RowNumberDetalleStart To RowNumberDetalleEnd
        ExcelWorksheet.Range("B" & RowNumber, "C" & RowNumber).Merge
    Next RowNumber
    
    'TIPOGRAFIAS
    With ExcelWorksheet.Range("A" & RowNumberDetalleStart, "H" & RowNumberDetalleEnd)
        .Font.Name = "Arial"
        .Font.Size = 9
    End With
    With ExcelWorksheet.Range("I" & RowNumberDetalleStart, "I" & RowNumberDetalleEnd)
        .Font.Name = "Wingdings"
        .Font.Size = 10
    End With
    With ExcelWorksheet.Range("J" & RowNumberDetalleStart, "J" & RowNumberDetalleEnd)
        .Font.Name = "Arial"
        .Font.Size = 8
    End With
    
    'FORMATOS
    ExcelWorksheet.Range("F" & RowNumberDetalleStart, "H" & RowNumberDetalleEnd).NumberFormat = "$ #,##0.00;$ -#,##0.00"
    
    'BORDES
    With ExcelWorksheet.Range("A" & RowNumberDetalleStart, "J" & RowNumberDetalleEnd)
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideHorizontal).Weight = xlThin
        .BorderAround 1, xlMedium
    End With
    
    '//////////////////////////////////////////////////////////
    'COMISIONES
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_Report_Viaje_Planilla_Comision"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("FechaHora_FILTER", adDate, adParamInput, , Viaje.FechaHora)
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta_FILTER", adChar, adParamInput, 20, Viaje.IDRuta)
    Set recData = New ADODB.Recordset
    recData.CursorType = adOpenForwardOnly
    recData.LockType = adLockReadOnly
    recData.Open cmdData
    Set cmdData = Nothing
    
    RowNumber = RowNumber + 1
    With ExcelWorksheet.Range("A" & RowNumber)
        .value = "COMISIONES"
        .HorizontalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 14
        .Font.Bold = True
        .Interior.Color = &HC6C3C6
    End With
    ExcelWorksheet.Range("A" & RowNumber, "K" & RowNumber).Merge
        
    'ENCABEZADOS DE LAS COLUMNAS
    RowNumber = RowNumber + 2
    ExcelWorksheet.Range("A" & RowNumber).value = "#"
    ExcelWorksheet.Range("B" & RowNumber).value = "Envía"
    ExcelWorksheet.Range("C" & RowNumber).value = "Recibe"
    ExcelWorksheet.Range("D" & RowNumber).value = "Sube"
    ExcelWorksheet.Range("E" & RowNumber).value = "Baja"
    ExcelWorksheet.Range("F" & RowNumber).value = "Importe"
    ExcelWorksheet.Range("G" & RowNumber).value = "Pagado"
    ExcelWorksheet.Range("H" & RowNumber).value = "Saldo"
    ExcelWorksheet.Range("I" & RowNumber).value = "Observaciones"
    ExcelWorksheet.Range("A" & RowNumber, "J" & RowNumber).VerticalAlignment = xlCenter
    ExcelWorksheet.Range("I" & RowNumber, "J" & RowNumber).Merge
    
    With ExcelWorksheet.Range("A" & RowNumber, "J" & RowNumber)
        .HorizontalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Bold = True
    End With
            
    'BORDES
    With ExcelWorksheet.Range("A" & RowNumber, "J" & RowNumber)
        .Borders(xlInsideVertical).Weight = xlThin
        .BorderAround 1, xlMedium
    End With
    
    RowNumberDetalleStart = RowNumber + 1
    Do While Not recData.EOF
        RowNumber = RowNumber + 1
        ExcelWorksheet.Range("A" & RowNumber).value = recData.AbsolutePosition
        ExcelWorksheet.Range("B" & RowNumber).value = recData("Envia").value
        ExcelWorksheet.Range("C" & RowNumber).value = recData("Recibe").value
        If Not (recData("IDOrigen").value = Ruta.IDOrigen And recData("Origen").value = LugarOrigen.Nombre) Then
            ExcelWorksheet.Range("D" & RowNumber).value = recData("Origen").value
        End If
        If Not (recData("IDDestino").value = Ruta.IDDestino And recData("Destino").value = LugarDestino.Nombre) Then
            ExcelWorksheet.Range("E" & RowNumber).value = recData("Destino").value
        End If
        ExcelWorksheet.Range("F" & RowNumber).value = recData("Importe").value
        If recData("ImportePagado").value = 0 Then
            ExcelWorksheet.Range("G" & RowNumber).Locked = False
            ExcelWorksheet.Range("G" & RowNumber).Font.Color = vbRed
        Else
            ExcelWorksheet.Range("G" & RowNumber).value = recData("ImportePagado").value
        End If
        If recData("ImprimirSaldo").value And Not IsNull(recData("SaldoActual").value) Then
            ExcelWorksheet.Range("H" & RowNumber).value = recData("SaldoActual").value
        End If
        If IsNull(recData("Notas").value) Then
            ExcelWorksheet.Range("I" & RowNumber).Locked = False
            ExcelWorksheet.Range("I" & RowNumber).Font.Color = vbRed
        Else
            ExcelWorksheet.Range("I" & RowNumber).value = recData("Notas").value & ""
        End If
        recData.MoveNext
    Loop
    RowNumberDetalleEnd = RowNumberDetalleStart + 5
    
    'DESPROTEJO LAS CELDAS VACIAS y CAMBIO EL COLOR DE LA TIPOGRAFIA EN LAS CELDAS VACIAS
    If RowNumber < RowNumberDetalleEnd Then
        With ExcelWorksheet.Range("A" & RowNumber + 1, "J" & RowNumberDetalleEnd)
            .Locked = False
            .Font.Color = vbRed
        End With
    End If
    
    'ALINEACION
    ExcelWorksheet.Range("A" & RowNumberDetalleStart, "A" & RowNumberDetalleEnd).HorizontalAlignment = xlCenter
    ExcelWorksheet.Range("B" & RowNumberDetalleStart, "B" & RowNumberDetalleEnd).HorizontalAlignment = xlLeft
    ExcelWorksheet.Range("C" & RowNumberDetalleStart, "C" & RowNumberDetalleEnd).HorizontalAlignment = xlLeft
    ExcelWorksheet.Range("D" & RowNumberDetalleStart, "D" & RowNumberDetalleEnd).HorizontalAlignment = xlLeft
    ExcelWorksheet.Range("E" & RowNumberDetalleStart, "E" & RowNumberDetalleEnd).HorizontalAlignment = xlLeft
    ExcelWorksheet.Range("F" & RowNumberDetalleStart, "F" & RowNumberDetalleEnd).HorizontalAlignment = xlRight
    ExcelWorksheet.Range("G" & RowNumberDetalleStart, "G" & RowNumberDetalleEnd).HorizontalAlignment = xlRight
    ExcelWorksheet.Range("H" & RowNumberDetalleStart, "H" & RowNumberDetalleEnd).HorizontalAlignment = xlRight
    ExcelWorksheet.Range("I" & RowNumberDetalleStart, "I" & RowNumberDetalleEnd).HorizontalAlignment = xlLeft
    ExcelWorksheet.Range("A" & RowNumberDetalleStart, "I" & RowNumberDetalleEnd).VerticalAlignment = xlCenter
    
    'HAGO MERGE DE LOS ROWS
    For RowNumber = RowNumberDetalleStart To RowNumberDetalleEnd
        ExcelWorksheet.Range("I" & RowNumber, "J" & RowNumber).Merge
    Next RowNumber
            
    'TIPOGRAFIAS
    With ExcelWorksheet.Range("A" & RowNumberDetalleStart, "H" & RowNumberDetalleEnd)
        .Font.Name = "Arial"
        .Font.Size = 9
    End With
    With ExcelWorksheet.Range("I" & RowNumberDetalleStart, "I" & RowNumberDetalleEnd)
        .Font.Name = "Arial"
        .Font.Size = 8
    End With
    
    'FORMATOS
    ExcelWorksheet.Range("F" & RowNumberDetalleStart, "H" & RowNumberDetalleEnd).NumberFormat = "$ #,##0.00;$ -#,##0.00"
    
    'BORDES
    With ExcelWorksheet.Range("A" & RowNumberDetalleStart, "J" & RowNumberDetalleEnd)
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideHorizontal).Weight = xlThin
        .BorderAround 1, xlMedium
    End With
    
    'COPYRIGHT
    ExcelWorksheet.Range("A45", "J46").Borders(xlInsideHorizontal).Weight = xlMedium
    With ExcelWorksheet.Range("A46")
        .value = "iNet Soluciones Informáticas"
        .Font.Name = "Times New Roman"
        .Font.Size = 10
    End With

    If pIsCompiled Then
        'PROTEJO LA PLANILLA
        ExcelWorksheet.Protect LCase(pParametro.CompanyName)
    End If
    
    recData.Close
    Set recData = Nothing
    
    'GUARDO LA PLANILLA
    ExcelWorkbook.SaveAs pSpecialFolders.Temp & FileName
    
    'LIBERO LOS OBJETOS
    Set ExcelWorksheet = Nothing
    ExcelWorkbook.Close
    Set ExcelWorkbook = Nothing
    ExcelApplication.Quit
    Set ExcelApplication = Nothing
    Set Ruta = Nothing
    Set LugarOrigen = Nothing
    Set LugarDestino = Nothing
    
    Screen.MousePointer = vbDefault
    PlanillaViajeGenerateExcel = True
    Exit Function
    
ErrorHandler:
    Select Case Err.Number
        Case 429
            Screen.MousePointer = vbDefault
            MsgBox "No se puede iniciar una sesión de Microsoft Excel." & vbCr & "Reinstale Microsoft Excel.", vbCritical, App.Title
        Case Else
            ShowErrorMessage "Forms.Viaje.PlanillaViajeGenerateExcel", "Error al Generar la Planilla del Viaje en Excel."
    End Select
    If Not Ruta Is Nothing Then
        Set Ruta = Nothing
    End If
    If Not LugarOrigen Is Nothing Then
        Set LugarOrigen = Nothing
    End If
    If Not LugarDestino Is Nothing Then
        Set LugarDestino = Nothing
    End If
    If Not ExcelWorksheet Is Nothing Then
        Set ExcelWorksheet = Nothing
    End If
    If Not ExcelWorkbook Is Nothing Then
        ExcelWorkbook.Close
        Set ExcelWorkbook = Nothing
    End If
    If Not ExcelApplication Is Nothing Then
        ExcelApplication.Quit
        Set ExcelApplication = Nothing
    End If
    If Not cmdData Is Nothing Then
        Set cmdData = Nothing
    End If
    If Not recData Is Nothing Then
        If recData.State = adStateOpen Then
            recData.Close
        End If
        Set recData = Nothing
    End If
End Function

Private Function FileDelete(ByVal PathAndFileName As String) As Boolean
    On Error GoTo ErrorHandler
    
    FileSystem.Kill PathAndFileName
    FileDelete = True
    Exit Function
    
ErrorHandler:
End Function
