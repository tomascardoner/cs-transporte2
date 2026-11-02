VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmViajeAsistencia 
   Caption         =   "Asistencia de Viajes"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11010
   Icon            =   "ViajeAsistencia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6345
   ScaleWidth      =   11010
   Begin VB.CommandButton cmdSelectViaje 
      Caption         =   "Seleccionar Viajes"
      Height          =   1335
      Left            =   4020
      TabIndex        =   14
      Top             =   720
      Width           =   1335
   End
   Begin MSComctlLib.ImageList ilsData 
      Left            =   3660
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViajeAsistencia.frx":062A
            Key             =   "PASAJERO_CONFIRMADO"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViajeAsistencia.frx":0BC4
            Key             =   "PASAJERO_CONDICIONAL"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViajeAsistencia.frx":115E
            Key             =   "PASAJERO_CANCELADO"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViajeAsistencia.frx":16F8
            Key             =   "COMISION_CONFIRMADO"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViajeAsistencia.frx":1C92
            Key             =   "COMISION_CANCELADO"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbPin 
      Height          =   330
      Left            =   60
      TabIndex        =   4
      Top             =   6480
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
      Height          =   630
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   1111
      BandCount       =   4
      FixedOrder      =   -1  'True
      _CBWidth        =   11010
      _CBHeight       =   630
      _Version        =   "6.7.9816"
      Child1          =   "tlbMain"
      MinWidth1       =   4995
      MinHeight1      =   570
      Width1          =   4995
      FixedBackground1=   0   'False
      Key1            =   "Toolbar"
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Child2          =   "picFilterEstado"
      MinWidth2       =   1845
      MinHeight2      =   330
      Width2          =   1845
      Key2            =   "FilterEstado"
      NewRow2         =   0   'False
      AllowVertical2  =   0   'False
      Child3          =   "picFilterRealizado"
      MinWidth3       =   1800
      MinHeight3      =   330
      Width3          =   1800
      Key3            =   "FilterRealizado"
      NewRow3         =   0   'False
      AllowVertical3  =   0   'False
      Child4          =   "picMostrarSaldo"
      MinWidth4       =   1410
      MinHeight4      =   330
      Width4          =   1410
      Key4            =   "MostrarSaldo"
      NewRow4         =   0   'False
      AllowVertical4  =   0   'False
      Begin VB.PictureBox picMostrarSaldo 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   9510
         ScaleHeight     =   330
         ScaleWidth      =   1410
         TabIndex        =   11
         Top             =   150
         Width           =   1410
         Begin VB.CheckBox chkMostrarSaldo 
            Caption         =   "Mostrar Saldo"
            Height          =   195
            Left            =   0
            TabIndex        =   12
            Top             =   60
            Width           =   1455
         End
      End
      Begin VB.PictureBox picFilterRealizado 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   7485
         ScaleHeight     =   330
         ScaleWidth      =   1800
         TabIndex        =   8
         Top             =   150
         Width           =   1800
         Begin VB.ComboBox cboFilterRealizado 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   0
            Width           =   990
         End
         Begin VB.Label lblFilterRealizado 
            AutoSize        =   -1  'True
            Caption         =   "Realizado:"
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
            Left            =   0
            TabIndex        =   10
            Top             =   60
            Width           =   750
         End
      End
      Begin VB.PictureBox picFilterEstado 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   5415
         ScaleHeight     =   330
         ScaleWidth      =   1845
         TabIndex        =   5
         Top             =   150
         Width           =   1845
         Begin VB.ComboBox cboFilterEstado 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   0
            Width           =   1230
         End
         Begin VB.Label lblFilterEstado 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
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
            Left            =   0
            TabIndex        =   7
            Top             =   60
            Width           =   540
         End
      End
      Begin MSComctlLib.Toolbar tlbMain 
         Height          =   570
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   1005
         ButtonWidth     =   2381
         ButtonHeight    =   1005
         ToolTips        =   0   'False
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Propiedades"
               Key             =   "PROPERTIES"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Asistencia"
               Key             =   "ASISTENCIA"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cambiar Estado"
               Key             =   "CHANGE_STATUS"
               ImageIndex      =   7
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   5985
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   661
            MinWidth        =   661
            Key             =   "PIN"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "TEXT"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15637
            Key             =   "INFO"
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
   Begin MSComctlLib.ListView lvwViajeDetalle 
      Height          =   3315
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5847
      View            =   3
      LabelEdit       =   1
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
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Orden"
         Text            =   "Orden"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Persona"
         Text            =   "Persona"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Documento"
         Text            =   "Documento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   "Importe"
         Text            =   "Importe"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Key             =   "Pagado"
         Text            =   "Pagado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Key             =   "Debe"
         Text            =   "Debe"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Key             =   "SaldoActual"
         Text            =   "Saldo"
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
         Key             =   "Asiento"
         Text            =   "Asiento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Key             =   "Realizado"
         Text            =   "Realizado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Key             =   "Origen"
         Text            =   "Origen"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Key             =   "Destino"
         Text            =   "Destino"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Key             =   "Tipo"
         Text            =   "Tipo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Key             =   "Facturar"
         Text            =   "Facturar"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Key             =   "Notas"
         Text            =   "Observaciones"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Key             =   "ListaPasajero"
         Text            =   "Lista de Pasajeros"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwViaje 
      Height          =   1335
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   2355
      View            =   3
      LabelEdit       =   1
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
      NumItems        =   9
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
         Key             =   "Estado"
         Text            =   "Estado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "AsientoLibre"
         Text            =   "Asientos"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Key             =   "Notas"
         Text            =   "Notas"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   10920
      Y1              =   2400
      Y2              =   2400
   End
End
Attribute VB_Name = "frmViajeAsistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLoading As Boolean

Private mCViaje As Collection
    
Public FormWaitingForSelect As String

Public Function FillListViewViaje(Optional RefreshViajeDetalle As Boolean = True)
    Dim ListItem As MSComctlLib.ListItem
    Dim recData As ADODB.Recordset
    Dim KeySave As Variant
    Dim SQL_Where As String
    Dim SQL_OrderBy As String
    Dim Viaje As Viaje
    Dim ViajeKey As Variant
    
    If mLoading Then
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    If Not lvwViaje.SelectedItem Is Nothing Then
        KeySave = lvwViaje.SelectedItem.Key
    End If
    
    SQL_Where = ""
    
    'VIAJES
    If mCViaje.Count = 0 Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "0 = 1"
    Else
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "("
        For Each ViajeKey In mCViaje
            SQL_Where = SQL_Where & "(Viaje.FechaHora = '" & Format(CDate(GetSubString(Mid(ViajeKey, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER)), "yyyy/mm/dd hh:nn:ss") & "' AND Viaje.IDRuta = '" & ReplaceQuote(GetSubString(Mid(ViajeKey, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER)) & "') OR "
        Next ViajeKey
        SQL_Where = Left(SQL_Where, Len(SQL_Where) - 4)
        SQL_Where = SQL_Where & ")"
    End If
    
    'PERSONAL
    If pPersonal Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Viaje.Personal = 0"
    End If
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Select Case lvwViaje.SortKey
        Case 0  'DIA SEMANA
            SQL_OrderBy = " ORDER BY datepart(weekday, Viaje.FechaHora)" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC") & ", Viaje.FechaHora" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC") & ", Viaje.IDRuta + ': ' + Viaje.RutaOtra" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC")
        Case 1  'FECHA
            SQL_OrderBy = " ORDER BY Viaje.FechaHora" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC") & ", Viaje.IDRuta + ': ' + Viaje.RutaOtra" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC")
        Case 2  'HORA
            SQL_OrderBy = " ORDER BY convert(char(8), Viaje.FechaHora, 108)" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC") & ", datepart(weekday, Viaje.FechaHora)" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC") & ", Viaje.FechaHora" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC") & ", Viaje.IDRuta + ': ' + Viaje.RutaOtra" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC")
        Case 3  'RUTA
            SQL_OrderBy = " ORDER BY Viaje.IDRuta + ISNULL(': ' + Viaje.RutaOtra, '')" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC") & ", datepart(weekday, Viaje.FechaHora)" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC") & ", Viaje.FechaHora" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC")
        Case 4  'VEHICULO
            SQL_OrderBy = " ORDER BY Vehiculo.Nombre" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC") & ", datepart(weekday, Viaje.FechaHora)" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC") & ", Viaje.FechaHora" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC") & ", Viaje.IDRuta + ': ' + Viaje.RutaOtra" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC")
        Case 5  'CONDUCTOR
            SQL_OrderBy = " ORDER BY dbo.udf_GetEntidadApellidoYNombre(Persona.Apellido, Persona.Nombre)" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC") & ", datepart(weekday, Viaje.FechaHora)" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC") & ", Viaje.FechaHora" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC") & ", Viaje.IDRuta + ': ' + Viaje.RutaOtra" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC")
        Case 6  'ESTADO
            SQL_OrderBy = " ORDER BY Viaje.Estado" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC") & ", datepart(weekday, Viaje.FechaHora)" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC") & ", Viaje.FechaHora" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC") & ", Viaje.IDRuta + ': ' + Viaje.RutaOtra" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC")
        Case 7  'ASIENTO LIBRE
            SQL_OrderBy = " ORDER BY AsientoLibre" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC") & ", datepart(weekday, Viaje.FechaHora)" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC") & ", Viaje.FechaHora" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC") & ", Viaje.IDRuta + ': ' + Viaje.RutaOtra" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC")
        Case 8  'NOTAS + DIA SEMANA + FECHA-HORA + RUTA
            SQL_OrderBy = " ORDER BY Viaje.Notas" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC") & ", datepart(weekday, Viaje.FechaHora)" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC") & ", Viaje.FechaHora" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC") & ", Viaje.IDRuta + ': ' + Viaje.RutaOtra" & IIf(lvwViaje.SortOrder = lvwAscending, "", " DESC")
    End Select
    
    Set Viaje = New Viaje
    
    lvwViaje.ListItems.Clear
    
    Set recData = New ADODB.Recordset
    recData.Source = "SELECT Viaje.FechaHora, Viaje.IDRuta, Viaje.RutaOtra, Viaje.Estado, Vehiculo.Asiento - Viaje.AsientoOcupado AS AsientoLibre, Vehiculo.Nombre AS Vehiculo, dbo.udf_GetEntidadApellidoYNombre(Persona.Apellido, Persona.Nombre) AS Conductor, Viaje.Notas FROM (Viaje LEFT JOIN Vehiculo ON Viaje.IDVehiculo = Vehiculo.IDVehiculo) LEFT JOIN Persona ON Viaje.IDConductor = Persona.IDPersona" & SQL_Where & SQL_OrderBy
    recData.MaxRecords = pParametro.Recordset_MaxRecords
    recData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With recData
        If Not .EOF Then
            Do While Not .EOF
                Set ListItem = lvwViaje.ListItems.Add(, KEY_STRINGER & .Fields("FechaHora").value & KEY_DELIMITER & RTrim(.Fields("IDRuta").value), WeekdayName(Weekday(.Fields("FechaHora").value)))
                ListItem.SubItems(1) = Format(.Fields("FechaHora").value, "Short Date")
                ListItem.SubItems(2) = Format(.Fields("FechaHora").value, "Short Time")
                ListItem.SubItems(3) = RTrim(.Fields("IDRuta").value) & IIf(RTrim(.Fields("IDRuta").value) = pParametro.Ruta_ID_Otra, ": " & .Fields("RutaOtra").value, "")
                ListItem.SubItems(4) = .Fields("Vehiculo").value & ""
                ListItem.SubItems(5) = .Fields("Conductor").value & ""
                Viaje.Estado = .Fields("Estado").value
                ListItem.SubItems(6) = Viaje.Estado_ToString
                ListItem.SubItems(7) = .Fields("AsientoLibre").value & ""
                ListItem.SubItems(8) = .Fields("Notas").value & ""
                
                If RTrim(.Fields("IDRuta").value) = pParametro.Ruta_ID_Otra Then
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
    lvwViaje.SelectedItem.Selected = False
    Set lvwViaje.SelectedItem = lvwViaje.ListItems(KeySave)
    lvwViaje.SelectedItem.EnsureVisible
    
    If frmMDI.ActiveForm.Name = Me.Name And GetForegroundWindow() = frmMDI.hwnd And frmMDI.WindowState <> vbMinimized Then
        lvwViaje.SetFocus
    End If
    
    Call FillListViewViajeDetalle
    
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Forms.Viaje.FillListView", "Error al leer la lista de Viajes."
End Function

Public Function FillListViewViajeDetalle() As Boolean
    Dim MousePointerSave As Integer
    Dim ListItem As MSComctlLib.ListItem
    Dim recData As ADODB.Recordset
    Dim KeySave As String
    Dim SQL_Where As String
    Dim SQL_OrderBy As String
    Dim EstadoKey As String
    Dim UltimoTipo As String
    Dim UltimoEstadoPersona As String
    Dim UltimoEstadoComision As String
    Dim ViajeKey As Variant
    Dim ViajeDetalle As ViajeDetalle
    
    If mLoading Then
        Exit Function
    End If
    
RESTART:
    
    MousePointerSave = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    If Not lvwViajeDetalle.SelectedItem Is Nothing Then
        KeySave = lvwViajeDetalle.SelectedItem.Key
    End If
    
    SQL_Where = ""
    
    'VIAJES
    If mCViaje.Count = 0 Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "0 = 1"
    Else
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "("
        For Each ViajeKey In mCViaje
            SQL_Where = SQL_Where & "(ViajeDetalle.FechaHora = '" & Format(CDate(GetSubString(Mid(ViajeKey, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER)), "yyyy/mm/dd hh:nn:ss") & "' AND ViajeDetalle.IDRuta = '" & ReplaceQuote(GetSubString(Mid(ViajeKey, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER)) & "') OR "
        Next ViajeKey
        SQL_Where = Left(SQL_Where, Len(SQL_Where) - 4)
        SQL_Where = SQL_Where & ")"
    End If
        
    Select Case cboFilterEstado.ListIndex
        Case 0
            '<Todos>
        Case 1
            'Confirmado
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "ViajeDetalle.Estado = '" & VIAJE_DETALLE_ESTADO_CONFIRMADO & "'"
        Case 2
            'Condicional
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "ViajeDetalle.Estado = '" & VIAJE_DETALLE_ESTADO_CONDICIONAL & "'"
        Case 3
            'Cancelado
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "ViajeDetalle.Estado = '" & VIAJE_DETALLE_ESTADO_CANCELADO & "'"
    End Select
    
    If cboFilterRealizado.ListIndex > 0 Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Realizado " & IIf(cboFilterRealizado.ListIndex = 1, "IS NULL", IIf(cboFilterRealizado.ListIndex, "= 1", "= 0"))
    End If
    
    Select Case lvwViajeDetalle.SortKey
        Case 0  'ORDEN
            SQL_OrderBy = " ORDER BY ViajeDetalle.OcupanteTipo DESC, ViajeDetalle.Estado, ViajeDetalle.Orden" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC")
        Case 1  'PERSONA
            SQL_OrderBy = " ORDER BY ViajeDetalle.OcupanteTipo DESC, ViajeDetalle.Estado, dbo.udf_GetEntidadApellidoYNombre(Persona.Apellido, Persona.Nombre)" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC") & ", ViajeDetalle.Orden" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC")
        Case 2  'DOCUMENTO
            SQL_OrderBy = " ORDER BY ViajeDetalle.OcupanteTipo DESC, Persona.IDDocumentoTipo + Persona.DocumentoNumero" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC") & ", dbo.udf_GetEntidadApellidoYNombre(Persona.Apellido, Persona.Nombre)" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC") & ", ViajeDetalle.Orden" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC")
        Case 3  'IMPORTE
            SQL_OrderBy = " ORDER BY ViajeDetalle.OcupanteTipo DESC, ViajeDetalle.Importe" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC") & ", ViajeDetalle.Orden" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC")
        Case 4  'IMPORTE PAGADO
            SQL_OrderBy = " ORDER BY ViajeDetalle.OcupanteTipo DESC, ViajeDetalle.ImporteContado + ViajeDetalle.ImporteCuentaCorriente" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC") & ", ViajeDetalle.Orden" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC")
        Case 5  'DEBE
            SQL_OrderBy = " ORDER BY ViajeDetalle.OcupanteTipo DESC, ViajeDetalle.Importe - ViajeDetalle.ImporteContado - ViajeDetalle.ImporteCuentaCorriente" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC") & ", ViajeDetalle.Orden" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC")
        Case 6  'SALDO
            SQL_OrderBy = " ORDER BY ViajeDetalle.OcupanteTipo DESC, ViajeDetalle.Estado, ViajeDetalle.Orden" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC")
        Case 7  'ESTADO
            SQL_OrderBy = " ORDER BY ViajeDetalle.OcupanteTipo DESC, ViajeDetalle.Estado" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC") & ", ViajeDetalle.Orden" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC")
        Case 8  'ASIENTO
            SQL_OrderBy = " ORDER BY ViajeDetalle.OcupanteTipo DESC, ViajeDetalle.Estado" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC") & ", ViajeDetalle.Asiento" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC")
        Case 9  'REALIZADO
            SQL_OrderBy = " ORDER BY ViajeDetalle.OcupanteTipo DESC, ViajeDetalle.Estado" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC") & ", ViajeDetalle.Realizado" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC")
        Case 10  'ORIGEN
            SQL_OrderBy = " ORDER BY ViajeDetalle.OcupanteTipo DESC, ViajeDetalle.Sube + Lugar_Origen.Nombre" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC") & ", ViajeDetalle.Orden" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC")
        Case 11  'DESTINO
            SQL_OrderBy = " ORDER BY ViajeDetalle.OcupanteTipo DESC, ViajeDetalle.Baja + Lugar_Destino.Nombre" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC") & ", ViajeDetalle.Orden" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC")
        Case 12  'RESERVA TIPO
            SQL_OrderBy = " ORDER BY ViajeDetalle.OcupanteTipo DESC, ViajeDetalle.ReservaTipo" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC") & ", ViajeDetalle.Orden" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC")
        Case 13  'FACTURAR
            SQL_OrderBy = " ORDER BY ViajeDetalle.OcupanteTipo DESC, ViajeDetalle.Facturar" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC") & ", ViajeDetalle.Orden" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC")
        Case 14  'NOTAS
            SQL_OrderBy = " ORDER BY ViajeDetalle.OcupanteTipo DESC, ViajeDetalle.Notas" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC") & ", ViajeDetalle.Orden" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC")
        Case 15  'LISTA PASAJEROS
            SQL_OrderBy = " ORDER BY ViajeDetalle.OcupanteTipo DESC, ViajeDetalle.Estado" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC") & ", Persona.ListaPasajero" & IIf(lvwViajeDetalle.SortOrder = lvwAscending, "", " DESC")
    End Select
       
    lvwViajeDetalle.ListItems.Clear
    Set recData = New ADODB.Recordset
    recData.Source = "SELECT ViajeDetalle.FechaHora, ViajeDetalle.IDRuta, ViajeDetalle.Indice, ViajeDetalle.OcupanteTipo, ViajeDetalle.Estado, ViajeDetalle.Asiento, ViajeDetalle.Realizado, ViajeDetalle.Orden, Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Persona, Persona.IDDocumentoTipo, Persona.DocumentoNumero, ViajeDetalle.Importe, ViajeDetalle.ImporteContado + ViajeDetalle.ImporteCuentaCorriente AS ImportePagado, ViajeDetalle.Importe - ViajeDetalle.ImporteContado - ViajeDetalle.ImporteCuentaCorriente AS Debe, Lugar_Origen.Nombre AS Origen, ViajeDetalle.Sube, Lugar_Destino.Nombre AS Destino, ViajeDetalle.Baja, ViajeDetalle.ReservaTipo, ViajeDetalle.Facturar, ViajeDetalle.Notas, Persona.ListaPasajero, ViajeDetalle.CreadoEnProgreso, ViajeDetalle.ModificadoEnProgreso"
    recData.Source = recData.Source & IIf(chkMostrarSaldo.value = vbChecked, ", (SELECT Sum(Importe) AS SaldoActual FROM CuentaCorriente WHERE CuentaCorriente.IDPersona = (CASE isnull(ViajeDetalle.IDPersonaCuentaCorriente, 0) WHEN 0 THEN ViajeDetalle.IDPersona ELSE ViajeDetalle.IDPersonaCuentaCorriente END)) AS SaldoActual", "") & " "
    recData.Source = recData.Source & "FROM ((ViajeDetalle INNER JOIN Persona ON ViajeDetalle.IDPersona = Persona.IDPersona) INNER JOIN Lugar AS Lugar_Origen ON ViajeDetalle.IDOrigen = Lugar_Origen.IDLugar) INNER JOIN Lugar AS Lugar_Destino ON ViajeDetalle.IDDestino = Lugar_Destino.IDLugar" & SQL_Where & SQL_OrderBy
    recData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Set ViajeDetalle = New ViajeDetalle
    
    With recData
        If Not .EOF Then
            Do While Not .EOF
                Select Case .Fields("Estado").value
                    Case VIAJE_DETALLE_ESTADO_CONFIRMADO
                        EstadoKey = "CONFIRMADO"
                    Case VIAJE_DETALLE_ESTADO_CONDICIONAL
                        EstadoKey = "CONDICIONAL"
                    Case VIAJE_DETALLE_ESTADO_CANCELADO
                        EstadoKey = "CANCELADO"
                    Case Else
                        EstadoKey = ""
                End Select
                
                '//////////////////////////////////////////////////
                'ROWS SEPARATOR BY TYPE
                If pParametro.ViajeDetalle_SeparateRowsByType Then
                    If UltimoTipo <> .Fields("OcupanteTipo").value Then
                        If UltimoTipo <> "" Then
                            Set ListItem = lvwViajeDetalle.ListItems.Add(, KEY_STRINGER & (.Fields("Indice").value * -1), "")
                        End If
                        UltimoTipo = .Fields("OcupanteTipo").value
                    End If
                End If
                '//////////////////////////////////////////////////
                'ROWS SEPARATOR BY STATUS
                If .Fields("OcupanteTipo").value = OCUPANTE_TIPO_PASAJERO Then
                    If pParametro.ViajeDetalle_SeparateRowsByStatus Then
                        If UltimoEstadoPersona <> .Fields("Estado").value And (lvwViajeDetalle.SortKey = 0 Or lvwViajeDetalle.SortKey = 1 Or lvwViajeDetalle.SortKey = 6) Then
                            If UltimoEstadoPersona <> "" Then
                                Set ListItem = lvwViajeDetalle.ListItems.Add(, KEY_STRINGER & (.Fields("Indice").value * -1), "")
                            End If
                            UltimoEstadoPersona = .Fields("Estado").value
                        End If
                    End If
                Else
                    If pParametro.ViajeDetalle_SeparateRowsByStatus Then
                        If UltimoEstadoComision <> .Fields("Estado").value And (lvwViajeDetalle.SortKey = 0 Or lvwViajeDetalle.SortKey = 6) Then
                            If UltimoEstadoComision <> "" Then
                                Set ListItem = lvwViajeDetalle.ListItems.Add(, KEY_STRINGER & (.Fields("Indice").value * -1), "")
                            End If
                            UltimoEstadoComision = .Fields("Estado").value
                        End If
                    End If
                End If
                
                Select Case .Fields("OcupanteTipo").value
                    Case OCUPANTE_TIPO_PASAJERO
                        'Pasajero
                        Set ListItem = lvwViajeDetalle.ListItems.Add(, KEY_STRINGER & .Fields("FechaHora").value & KEY_DELIMITER & RTrim(.Fields("IDRuta").value) & KEY_DELIMITER & .Fields("Indice").value, .Fields("Orden").value, , "PASAJERO_" & EstadoKey)
                        If .Fields("Estado").value = VIAJE_DETALLE_ESTADO_CONFIRMADO Then
                            ListItem.SubItems(9) = IIf(IsNull(.Fields("Realizado").value), "", IIf(.Fields("Realizado").value, "Sí", "No"))
                        End If
                    Case OCUPANTE_TIPO_COMISION
                        'Comisión
                        Set ListItem = lvwViajeDetalle.ListItems.Add(, KEY_STRINGER & .Fields("FechaHora").value & KEY_DELIMITER & RTrim(.Fields("IDRuta").value) & KEY_DELIMITER & .Fields("Indice").value, .Fields("Orden").value, , "COMISION_" & EstadoKey)
                End Select
                ListItem.SubItems(1) = .Fields("Persona").value
                ListItem.SubItems(2) = IIf(IsNull(.Fields("DocumentoNumero").value), "", IIf(IsNull(.Fields("IDDocumentoTipo").value), .Fields("DocumentoNumero").value, RTrim(.Fields("IDDocumentoTipo").value) & ": " & .Fields("DocumentoNumero").value))
                ListItem.SubItems(3) = Format(.Fields("Importe").value, "Currency")
                ListItem.SubItems(4) = Format(.Fields("ImportePagado").value, "Currency")
                ListItem.SubItems(5) = Format(.Fields("Debe").value, "Currency")
                If chkMostrarSaldo.value = vbChecked Then
                    ListItem.SubItems(6) = IIf(IsNull(.Fields("SaldoActual").value), " ", Format(.Fields("SaldoActual").value, "Currency"))
                Else
                    ListItem.SubItems(6) = " "
                End If
                ViajeDetalle.Estado = .Fields("Estado").value & ""
                ListItem.SubItems(7) = ViajeDetalle.Estado_ToString
                ListItem.SubItems(8) = .Fields("Asiento").value & ""
                ListItem.SubItems(10) = IIf(IsNull(.Fields("Sube").value), .Fields("Origen").value, .Fields("Sube").value)
                ListItem.SubItems(11) = IIf(IsNull(.Fields("Baja").value), .Fields("Destino").value, .Fields("Baja").value)
                ViajeDetalle.ReservaTipo = .Fields("ReservaTipo").value
                ListItem.SubItems(12) = ViajeDetalle.ReservaTipo_ToString
                ListItem.SubItems(13) = IIf(.Fields("Facturar").value, "Sí", "No")
                ListItem.SubItems(14) = .Fields("Notas").value & ""
                ListItem.SubItems(15) = IIf(.Fields("ListaPasajero").value, "Sí", "")
                
                If .Fields("CreadoEnProgreso").value Then
                   ListItem.ForeColor = pParametro.ViajeDetalle_CreadoEnProgreso_Color
                   ListItem.Bold = True
                End If
                If .Fields("ModificadoEnProgreso").value Then
                   ListItem.ForeColor = pParametro.ViajeDetalle_ModificadoEnProgreso_Color
                   ListItem.Bold = True
                End If
                .MoveNext
            Loop
            
            stbMain.Panels("TEXT").Text = .RecordCount & " items"
        Else
            stbMain.Panels("TEXT").Text = "No hay items."
        End If
        .Close
    End With
    Set recData = Nothing
    
    Set ViajeDetalle = Nothing
    
    On Error Resume Next
    Set lvwViajeDetalle.SelectedItem = lvwViajeDetalle.ListItems(KeySave)
    lvwViajeDetalle.SelectedItem.EnsureVisible
    
    If frmMDI.ActiveForm.Name = Me.Name And GetForegroundWindow() = frmMDI.hwnd And frmMDI.WindowState <> vbMinimized Then
        lvwViajeDetalle.SetFocus
    End If
    
    FillListViewViajeDetalle = True
    Screen.MousePointer = MousePointerSave
    Exit Function
    
ErrorHandler:
    If Err.Number = ERROR_TYPE_MISMATCH Or Err.Number = ERROR_ELEMENT_NOT_FOUND Then
        'mViaje.Asiento_Asignar
        Resume RESTART
    Else
        ShowErrorMessage "Forms.ViajeDetalle.FillListView", "Error al obtener el Detalle del Viaje."
    End If
End Function

Private Sub cboFilterRealizado_Click()
    FillListViewViajeDetalle
End Sub

Private Sub cbrMain_HeightChanged(ByVal NewHeight As Single)
    ResizeControls NewHeight
End Sub

Private Sub cboFilterEstado_Click()
    FillListViewViajeDetalle
End Sub

Private Sub chkMostrarSaldo_Click()
    FillListViewViajeDetalle
End Sub

Private Sub cmdSelectViaje_Click()
    Screen.MousePointer = vbHourglass
    frmViaje.Show
    If frmViaje.WindowState = vbMinimized Then
        frmViaje.WindowState = vbNormal
    End If
    frmViaje.FormWaitingForSelect = Me.Name
    frmViaje.AllowMultipleSelect = True
    frmViaje.AllowMultipleRuta = False
    frmViaje.CSelectEstadosFilter.Add VIAJE_ESTADO_ACTIVO
    frmViaje.CSelectEstadosFilter.Add VIAJE_ESTADO_EN_PROGRESO
    frmViaje.SetFocus
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
            Case vbKeyP
                tlbMain_ButtonClick tlbMain.Buttons.Item("PROPERTIES")
        End Select
    End If
End Sub

Private Sub Form_Load()
    mLoading = True
    
    Set mCViaje = New Collection
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL TOOLBAR
    Set tlbMain.ImageList = frmMDI.ilsFormToolbar
    Set tlbMain.HotImageList = frmMDI.ilsFormToolbarHot
    tlbMain.Buttons("PROPERTIES").Image = "PROPERTIES"
    tlbMain.Buttons("ASISTENCIA").Image = "ASISTENCIA"
    tlbMain.Buttons("CHANGE_STATUS").Image = "CHANGE_STATUS"
    '//////////////////////////////////////////////////////////
    
    '//////////////////////////////////////////////////////////
    'PREPARO LOS LISTVIEW
    lvwViaje.GridLines = pParametro.ListView_GridLines
    lvwViajeDetalle.GridLines = pParametro.ListView_GridLines
    
    Set lvwViaje.ColumnHeaderIcons = frmMDI.ilsFormSortColumn
    Set lvwViajeDetalle.ColumnHeaderIcons = frmMDI.ilsFormSortColumn
    Set lvwViajeDetalle.SmallIcons = ilsData
    '//////////////////////////////////////////////////////////
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL PIN
    Set tlbPin.ImageList = frmMDI.ilsFormPin
    '//////////////////////////////////////////////////////////
    
    cboFilterEstado.AddItem ITEM_ALL_MALE
    cboFilterEstado.AddItem VIAJE_DETALLE_ESTADO_CONFIRMADO_NOMBRE
    cboFilterEstado.AddItem VIAJE_DETALLE_ESTADO_CONDICIONAL_NOMBRE
    cboFilterEstado.AddItem VIAJE_DETALLE_ESTADO_CANCELADO_NOMBRE
    cboFilterEstado.ListIndex = 0
    
    cboFilterRealizado.AddItem ITEM_ALL_MALE
    cboFilterRealizado.AddItem "--"
    cboFilterRealizado.AddItem "Sí"
    cboFilterRealizado.AddItem "No"
    cboFilterRealizado.ListIndex = 1
    
    CSM_Forms.ResizeAndPosition frmMDI, Me
    pParametro.GetCoolBarSettings "ViajeAsistencia", cbrMain
    
    pParametro.GetListViewSettings "Viaje", lvwViaje
    lvwViaje.ColumnHeaders(lvwViaje.SortKey + 1).Icon = lvwViaje.SortOrder + 1
    
    pParametro.GetListViewSettings "ViajeDetalle", lvwViajeDetalle
    lvwViajeDetalle.ColumnHeaders(lvwViajeDetalle.SortKey + 1).Icon = lvwViajeDetalle.SortOrder + 1
    
    tlbPin.Buttons("PIN").value = pParametro.Usuario_LeerNumero("ViajeAsistencia_Pin", tlbPin.Buttons("PIN").value)
    If tlbPin.Buttons("PIN").value = tbrUnpressed Then
        tlbPin.Buttons("PIN").Image = 1
    Else
        tlbPin.Buttons("PIN").Image = 2
    End If

    mLoading = False
End Sub

Private Sub Form_Resize()
    ResizeControls cbrMain.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visible = False
    WindowState = vbNormal
    pParametro.SaveCoolBarSettings "ViajeAsistencia", cbrMain
    
    pParametro.SaveListViewSettings "Viaje", lvwViaje
    pParametro.SaveListViewSettings "ViajeDetalle", lvwViajeDetalle
    
    pParametro.Usuario_GuardarNumero "ViajeAsistencia_Pin", tlbPin.Buttons("PIN").value
    Set mCViaje = Nothing
    Set frmViajeDetalle = Nothing
End Sub

Private Sub lvwViaje_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwViaje.ColumnHeaders(lvwViaje.SortKey + 1).Icon = 0
    If ColumnHeader.Index - 1 = lvwViaje.SortKey Then
        lvwViaje.SortOrder = IIf(lvwViaje.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lvwViaje.SortKey = ColumnHeader.Index - 1
        lvwViaje.SortOrder = lvwAscending
    End If
    ColumnHeader.Icon = lvwViaje.SortOrder + 1
    
    FillListViewViaje False
End Sub

Private Sub lvwViajeDetalle_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwViajeDetalle.ColumnHeaders(lvwViajeDetalle.SortKey + 1).Icon = 0
    If ColumnHeader.Index - 1 = lvwViajeDetalle.SortKey Then
        lvwViajeDetalle.SortOrder = IIf(lvwViajeDetalle.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lvwViajeDetalle.SortKey = ColumnHeader.Index - 1
        lvwViajeDetalle.SortOrder = lvwAscending
    End If
    ColumnHeader.Icon = lvwViajeDetalle.SortOrder + 1
    FillListViewViajeDetalle
End Sub

Private Sub lvwViajeDetalle_DblClick()
    If GetFormIndex(FormWaitingForSelect) > 0 Then
        tlbMain_ButtonClick tlbMain.Buttons.Item("SELECT")
    Else
        tlbMain_ButtonClick tlbMain.Buttons.Item("PROPERTIES")
    End If
End Sub

Private Sub lvwViajeDetalle_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Val(Mid(Item.Key, Len(KEY_STRINGER) + 1)) < 0 Then
        Set lvwViaje.SelectedItem = Nothing
    Else
        Set lvwViaje.SelectedItem = lvwViaje.ListItems(KEY_STRINGER & CSM_String.GetSubString(Mid(Item.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER) & KEY_DELIMITER & CSM_String.GetSubString(Mid(Item.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER))
        lvwViaje.SelectedItem.EnsureVisible
    End If
End Sub

Private Sub lvwViajeDetalle_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        lvwViajeDetalle_DblClick
    End If
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim Viaje As Viaje
    Dim ViajeDetalle As ViajeDetalle
    Dim Ruta As Ruta
    Dim RutaDetalleLimite As RutaDetalle
    Dim RutaDetalleOrigen As RutaDetalle
    
    Select Case Button.Key
        Case "PROPERTIES"
            If Button.Enabled Then
                If pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_MODIFY) Then
                    If lvwViajeDetalle.SelectedItem Is Nothing Then
                        MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                        lvwViajeDetalle.SetFocus
                        Exit Sub
                    End If
                    If Val(Mid(lvwViajeDetalle.SelectedItem.Key, 2)) < 0 Then
                        MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                        lvwViajeDetalle.SetFocus
                        Exit Sub
                    End If
                    
                    Screen.MousePointer = vbHourglass
                    
                    Set ViajeDetalle = New ViajeDetalle
                    ViajeDetalle.FechaHora = CDate(GetSubString(Mid(lvwViajeDetalle.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                    ViajeDetalle.IDRuta = CSM_String.GetSubString(Mid(lvwViajeDetalle.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER)
                    ViajeDetalle.Indice = Val(GetSubString(Mid(lvwViajeDetalle.SelectedItem.Key, Len(KEY_STRINGER) + 1), 3, KEY_DELIMITER))
                    If ViajeDetalle.Load() Then
                        frmViajeDetallePropiedad.LoadDataAndShow Me, ViajeDetalle
                    Else
                        lvwViajeDetalle.SetFocus
                    End If
                    Set ViajeDetalle = Nothing
                    Screen.MousePointer = vbDefault
                End If
            End If
        Case "ASISTENCIA"
            If pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_MODIFY) Then
                If lvwViajeDetalle.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwViajeDetalle.SetFocus
                    Exit Sub
                End If
                If Val(Mid(lvwViajeDetalle.SelectedItem.Key, 2)) < 0 Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwViajeDetalle.SetFocus
                    Exit Sub
                End If
                
                Screen.MousePointer = vbHourglass
                
                Set ViajeDetalle = New ViajeDetalle
                With ViajeDetalle
                    .FechaHora = CDate(GetSubString(Mid(lvwViajeDetalle.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                    .IDRuta = CSM_String.GetSubString(Mid(lvwViajeDetalle.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER)
                    .Indice = Val(GetSubString(Mid(lvwViajeDetalle.SelectedItem.Key, Len(KEY_STRINGER) + 1), 3, KEY_DELIMITER))
                    If Not .Load() Then
                        Set ViajeDetalle = Nothing
                        Exit Sub
                    End If
                    If .Estado <> VIAJE_DETALLE_ESTADO_CONFIRMADO Then
                        MsgBox "Sólo se puede cargar la Asistencia de los Pasajeros o Comisiones con Estado Confirmado.", vbInformation, App.Title
                        lvwViajeDetalle.SetFocus
                        Set ViajeDetalle = Nothing
                        Exit Sub
                    End If
                    If Not pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_FINALIZADO_MODIFY, False) Then
                        Set Viaje = New Viaje
                        Viaje.FechaHora = CDate(GetSubString(Mid(lvwViaje.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                        Viaje.IDRuta = CSM_String.GetSubString(Mid(lvwViaje.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER)
                        If Not Viaje.Load() Then
                            lvwViajeDetalle.SetFocus
                            Set ViajeDetalle = Nothing
                            Set Viaje = Nothing
                            Exit Sub
                        End If
                        
                        If ViajeDetalle.OcupanteTipo = OCUPANTE_TIPO_PASAJERO Then
                            If Viaje.Estado = VIAJE_ESTADO_FINALIZADO And (ViajeDetalle.Realizado = 2 Or (ViajeDetalle.Realizado = 1 And ViajeDetalle.ImporteContado = ViajeDetalle.Importe)) Then
                                MsgBox "Ya se le dió Asistencia a esta Reserva.", vbInformation, App.Title
                                lvwViajeDetalle.SetFocus
                                Set ViajeDetalle = Nothing
                                Set Viaje = Nothing
                                Exit Sub
                            End If
                        Else
                            If Viaje.Estado = VIAJE_ESTADO_FINALIZADO And ViajeDetalle.ImporteContado = ViajeDetalle.Importe And ViajeDetalle.Entregada Then
                                MsgBox "Ya se le dió Asistencia a esta Comisión.", vbInformation, App.Title
                                lvwViajeDetalle.SetFocus
                                Set ViajeDetalle = Nothing
                                Set Viaje = Nothing
                                Exit Sub
                            End If
                        End If
                        
                        Set Viaje = Nothing
                    End If
                    frmViajeDetalleAsistencia.LoadDataAndShow Me, ViajeDetalle
                End With
                Set ViajeDetalle = Nothing
                Screen.MousePointer = vbDefault
            End If
        Case "CHANGE_STATUS"
            If Button.Enabled Then
                If pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_CHANGE_STATUS) Then
                    If lvwViajeDetalle.SelectedItem Is Nothing Then
                        MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                        lvwViajeDetalle.SetFocus
                        Exit Sub
                    End If
                    If Val(Mid(lvwViajeDetalle.SelectedItem.Key, 2)) < 0 Then
                        MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                        lvwViajeDetalle.SetFocus
                        Exit Sub
                    End If
                    
                    Set ViajeDetalle = New ViajeDetalle
                    With ViajeDetalle
                        .FechaHora = CDate(GetSubString(Mid(lvwViajeDetalle.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                        .IDRuta = CSM_String.GetSubString(Mid(lvwViajeDetalle.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER)
                        .Indice = Val(GetSubString(Mid(lvwViajeDetalle.SelectedItem.Key, Len(KEY_STRINGER) + 1), 3, KEY_DELIMITER))
                        If Not .Load() Then
                            lvwViajeDetalle.SetFocus
                            Set ViajeDetalle = Nothing
                            Exit Sub
                        End If
                        
                        'VERIFICO QUE NO HAYA PASADO EL TIEMPO LIMITE
                        Set Viaje = New Viaje
                        Viaje.FechaHora = CDate(GetSubString(Mid(lvwViaje.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                        Viaje.IDRuta = CSM_String.GetSubString(Mid(lvwViaje.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER)
                        If Not Viaje.Load() Then
                            lvwViajeDetalle.SetFocus
                            Set ViajeDetalle = Nothing
                            Set Viaje = Nothing
                            Exit Sub
                        End If
                        
                        Set Ruta = New Ruta
                        Ruta.IDRuta = Viaje.IDRuta
                        If Ruta.Load() Then
                            If Ruta.LimiteCancelacionDuracion > 0 And Ruta.LimiteCancelacionIDLugar > 0 Then
                                Set RutaDetalleLimite = New RutaDetalle
                                RutaDetalleLimite.IDRuta = Viaje.IDRuta
                                RutaDetalleLimite.IDLugar = Ruta.LimiteCancelacionIDLugar
                                If RutaDetalleLimite.Load() Then
                                    Set RutaDetalleOrigen = New RutaDetalle
                                    RutaDetalleOrigen.IDRuta = Viaje.IDRuta
                                    RutaDetalleOrigen.IDLugar = ViajeDetalle.IDOrigen
                                    If RutaDetalleOrigen.Load() Then
                                        If RutaDetalleOrigen.Indice <= RutaDetalleLimite.Indice Then
                                            If DateDiff("n", ViajeDetalle.FechaHora, Now) > Ruta.LimiteCancelacionDuracion Then
                                                'Tiempo Vencido, habilito según Permiso
                                                If pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_CHANGE_STATUS_AFTER_LIMIT, False) Then
                                                    'Permitido por Permiso
                                                    Select Case .Estado
                                                        Case VIAJE_DETALLE_ESTADO_CONFIRMADO, VIAJE_DETALLE_ESTADO_CONDICIONAL
                                                            frmViajeDetalleCancelar.LoadDataAndShow Me, ViajeDetalle
                                                        Case VIAJE_DETALLE_ESTADO_CANCELADO
                                                            If .OcupanteTipo = OCUPANTE_TIPO_PASAJERO Then
                                                                If MsgBox("Esta Reserva está Cancelada." & vbCr & vbCr & "Orden: " & .Orden & vbCr & "Pasajero: " & lvwViajeDetalle.SelectedItem.SubItems(1) & ", " & lvwViajeDetalle.SelectedItem.SubItems(2) & vbCr & vbCr & "¿Desea Activar esta Reserva?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                                                                    .Estado = ""
                                                                    Call .CambiarEstado(pParametro.Viaje_Permite_RutaConexion)
                                                                End If
                                                            Else
                                                                If MsgBox("Esta Comisión está Cancelada." & vbCr & vbCr & "Orden: " & .Orden & vbCr & "Remitente: " & lvwViajeDetalle.SelectedItem.SubItems(1) & ", " & lvwViajeDetalle.SelectedItem.SubItems(2) & vbCr & vbCr & "¿Desea Activar esta Comisión?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                                                                    .Estado = VIAJE_DETALLE_ESTADO_CONFIRMADO
                                                                    Call .CambiarEstado(pParametro.Viaje_Permite_RutaConexion)
                                                                End If
                                                            End If
                                                        Case Else
                                                            MsgBox "Estado Incorrecto.", vbCritical, App.Title
                                                    End Select
                                                Else
                                                    MsgBox "Este Viaje ya ha cumplido el tiempo límite para realizar cambios de estado de las Reservas.", vbInformation, App.Title
                                                End If
                                            Else
                                                'Permitido porque aún no pasado el tiempo límite
                                                Select Case .Estado
                                                    Case VIAJE_DETALLE_ESTADO_CONFIRMADO, VIAJE_DETALLE_ESTADO_CONDICIONAL
                                                        frmViajeDetalleCancelar.LoadDataAndShow Me, ViajeDetalle
                                                    Case VIAJE_DETALLE_ESTADO_CANCELADO
                                                        If .OcupanteTipo = OCUPANTE_TIPO_PASAJERO Then
                                                            If MsgBox("Esta Reserva está Cancelada." & vbCr & vbCr & "Orden: " & .Orden & vbCr & "Pasajero: " & lvwViajeDetalle.SelectedItem.SubItems(1) & ", " & lvwViajeDetalle.SelectedItem.SubItems(2) & vbCr & vbCr & "¿Desea Activar esta Reserva?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                                                                .Estado = ""
                                                                Call .CambiarEstado(pParametro.Viaje_Permite_RutaConexion)
                                                            End If
                                                        Else
                                                            If MsgBox("Esta Comisión está Cancelada." & vbCr & vbCr & "Orden: " & .Orden & vbCr & "Remitente: " & lvwViajeDetalle.SelectedItem.SubItems(1) & ", " & lvwViajeDetalle.SelectedItem.SubItems(2) & vbCr & vbCr & "¿Desea Activar esta Comisión?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                                                                .Estado = VIAJE_DETALLE_ESTADO_CONFIRMADO
                                                                Call .CambiarEstado(pParametro.Viaje_Permite_RutaConexion)
                                                            End If
                                                        End If
                                                    Case Else
                                                        MsgBox "Estado Incorrecto.", vbCritical, App.Title
                                                End Select
                                            End If
                                        Else
                                            'Permitido porque el Origen está antes que el Límite
                                            Select Case .Estado
                                                Case VIAJE_DETALLE_ESTADO_CONFIRMADO, VIAJE_DETALLE_ESTADO_CONDICIONAL
                                                    frmViajeDetalleCancelar.LoadDataAndShow Me, ViajeDetalle
                                                Case VIAJE_DETALLE_ESTADO_CANCELADO
                                                    If .OcupanteTipo = OCUPANTE_TIPO_PASAJERO Then
                                                        If MsgBox("Esta Reserva está Cancelada." & vbCr & vbCr & "Orden: " & .Orden & vbCr & "Pasajero: " & lvwViajeDetalle.SelectedItem.SubItems(1) & ", " & lvwViajeDetalle.SelectedItem.SubItems(2) & vbCr & vbCr & "¿Desea Activar esta Reserva?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                                                            .Estado = ""
                                                            Call .CambiarEstado(pParametro.Viaje_Permite_RutaConexion)
                                                        End If
                                                    Else
                                                        If MsgBox("Esta Comisión está Cancelada." & vbCr & vbCr & "Orden: " & .Orden & vbCr & "Remitente: " & lvwViajeDetalle.SelectedItem.SubItems(1) & ", " & lvwViajeDetalle.SelectedItem.SubItems(2) & vbCr & vbCr & "¿Desea Activar esta Comisión?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                                                            .Estado = VIAJE_DETALLE_ESTADO_CONFIRMADO
                                                            Call .CambiarEstado(pParametro.Viaje_Permite_RutaConexion)
                                                        End If
                                                    End If
                                                Case Else
                                                    MsgBox "Estado Incorrecto.", vbCritical, App.Title
                                            End Select
                                        End If
                                    Else
                                        lvwViajeDetalle.SetFocus
                                        Set ViajeDetalle = Nothing
                                        Set Viaje = Nothing
                                        Set Ruta = Nothing
                                        Set RutaDetalleLimite = Nothing
                                        Set RutaDetalleOrigen = Nothing
                                        Exit Sub
                                    End If
                                    Set RutaDetalleOrigen = Nothing
                                Else
                                    lvwViajeDetalle.SetFocus
                                    Set ViajeDetalle = Nothing
                                    Set Viaje = Nothing
                                    Set Ruta = Nothing
                                    Set RutaDetalleLimite = Nothing
                                    Exit Sub
                                End If
                                Set RutaDetalleLimite = Nothing
                            Else
                                'Permitido porque la Ruta no tiene Límite
                                Select Case .Estado
                                    Case VIAJE_DETALLE_ESTADO_CONFIRMADO, VIAJE_DETALLE_ESTADO_CONDICIONAL
                                        frmViajeDetalleCancelar.LoadDataAndShow Me, ViajeDetalle
                                    Case VIAJE_DETALLE_ESTADO_CANCELADO
                                        If .OcupanteTipo = OCUPANTE_TIPO_PASAJERO Then
                                            If MsgBox("Esta Reserva está Cancelada." & vbCr & vbCr & "Orden: " & .Orden & vbCr & "Pasajero: " & lvwViajeDetalle.SelectedItem.SubItems(1) & ", " & lvwViajeDetalle.SelectedItem.SubItems(2) & vbCr & vbCr & "¿Desea Activar esta Reserva?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                                                .Estado = ""
                                                Call .CambiarEstado(pParametro.Viaje_Permite_RutaConexion)
                                            End If
                                        Else
                                            If MsgBox("Esta Comisión está Cancelada." & vbCr & vbCr & "Orden: " & .Orden & vbCr & "Remitente: " & lvwViajeDetalle.SelectedItem.SubItems(1) & ", " & lvwViajeDetalle.SelectedItem.SubItems(2) & vbCr & vbCr & "¿Desea Activar esta Comisión?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                                                .Estado = VIAJE_DETALLE_ESTADO_CONFIRMADO
                                                Call .CambiarEstado(pParametro.Viaje_Permite_RutaConexion)
                                            End If
                                        End If
                                    Case Else
                                        MsgBox "Estado Incorrecto.", vbCritical, App.Title
                                End Select
                            End If
                        Else
                            lvwViajeDetalle.SetFocus
                            Set ViajeDetalle = Nothing
                            Set Viaje = Nothing
                            Set Ruta = Nothing
                            Exit Sub
                        End If
                        
                        Set Viaje = Nothing
                        Set Ruta = Nothing
                        
                        SetLastPersona ViajeDetalle.IDPersona
                    End With
                    Set ViajeDetalle = Nothing
                End If
            End If
    End Select
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
    
    lvwViaje.Top = CoolBarHeight + CONTROL_SPACE
    lvwViaje.Left = CONTROL_SPACE
    lvwViaje.Width = ScaleWidth - (CONTROL_SPACE * 3) - cmdSelectViaje.Width
    
    cmdSelectViaje.Top = lvwViaje.Top
    cmdSelectViaje.Left = lvwViaje.Left + lvwViaje.Width + CONTROL_SPACE
    
    Line1.Y1 = lvwViaje.Top + lvwViaje.Height + (CONTROL_SPACE * 2)
    Line1.X1 = CONTROL_SPACE
    Line1.Y2 = Line1.Y1
    Line1.X2 = ScaleWidth - (CONTROL_SPACE * 2)
    
    lvwViajeDetalle.Top = Line1.Y1 + Line1.BorderWidth + (CONTROL_SPACE * 2)
    lvwViajeDetalle.Left = CONTROL_SPACE
    lvwViajeDetalle.Width = ScaleWidth - (CONTROL_SPACE * 2)
    lvwViajeDetalle.Height = ScaleHeight - lvwViajeDetalle.Top - CONTROL_SPACE - stbMain.Height
    
    tlbPin.Top = ScaleHeight - 330
    tlbPin.Left = 15
End Sub

Public Sub MultipleViajeSelected(ByVal Viajes As Collection)
    Set mCViaje = Viajes
    
    Call FillListViewViaje
End Sub
