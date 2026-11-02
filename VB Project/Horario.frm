VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHorario 
   Caption         =   "Horarios"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11895
   Icon            =   "Horario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5730
   ScaleWidth      =   11895
   Begin MSComctlLib.Toolbar tlbPin 
      Height          =   330
      Left            =   15
      TabIndex        =   15
      Top             =   6645
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
      Height          =   1380
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   2434
      BandCount       =   5
      FixedOrder      =   -1  'True
      _CBWidth        =   11895
      _CBHeight       =   1380
      _Version        =   "6.7.9816"
      Child1          =   "tlbMain"
      MinWidth1       =   4395
      MinHeight1      =   570
      Width1          =   4395
      FixedBackground1=   0   'False
      Key1            =   "Toolbar"
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Child2          =   "picFilterDiaSemana"
      MinWidth2       =   1695
      MinHeight2      =   360
      Width2          =   1695
      Key2            =   "FilterDiaSemana"
      NewRow2         =   -1  'True
      AllowVertical2  =   0   'False
      Child3          =   "picFilterHora"
      MinWidth3       =   4665
      MinHeight3      =   360
      Width3          =   4665
      Key3            =   "FilterHora"
      NewRow3         =   0   'False
      AllowVertical3  =   0   'False
      Child4          =   "picFilterRuta"
      MinWidth4       =   3015
      MinHeight4      =   360
      Width4          =   1095
      Key4            =   "FilterRuta"
      NewRow4         =   0   'False
      AllowVertical4  =   0   'False
      Child5          =   "picFilterActivo"
      MinWidth5       =   1605
      MinHeight5      =   330
      Width5          =   1605
      Key5            =   "FilterActivo"
      NewRow5         =   0   'False
      AllowVertical5  =   0   'False
      Begin VB.PictureBox picFilterActivo 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   165
         ScaleHeight     =   330
         ScaleWidth      =   11640
         TabIndex        =   17
         Top             =   1020
         Width           =   11640
         Begin VB.ComboBox cboFilterActivo 
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
            TabIndex        =   18
            Top             =   0
            Width           =   990
         End
         Begin VB.Label lblFilterActivo 
            AutoSize        =   -1  'True
            Caption         =   "Activo:"
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
            TabIndex        =   19
            Top             =   60
            Width           =   510
         End
      End
      Begin VB.PictureBox picFilterHora 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   2085
         ScaleHeight     =   360
         ScaleWidth      =   4665
         TabIndex        =   9
         Top             =   630
         Width           =   4665
         Begin VB.ComboBox cboHora 
            Height          =   315
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   0
            Width           =   1035
         End
         Begin MSComCtl2.DTPicker dtpHoraDesde 
            Height          =   315
            Left            =   1560
            TabIndex        =   11
            Top             =   0
            Visible         =   0   'False
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
            CurrentDate     =   36950
         End
         Begin MSComCtl2.DTPicker dtpHoraHasta 
            Height          =   315
            Left            =   3240
            TabIndex        =   12
            Top             =   0
            Visible         =   0   'False
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
            CurrentDate     =   36950
         End
         Begin VB.Label lblHoraAnd 
            AutoSize        =   -1  'True
            Caption         =   "y"
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
            Left            =   3060
            TabIndex        =   14
            Top             =   60
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Label lblHora 
            AutoSize        =   -1  'True
            Caption         =   "Hora:"
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
            TabIndex        =   13
            Top             =   60
            Width           =   390
         End
      End
      Begin VB.PictureBox picFilterRuta 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   6975
         ScaleHeight     =   360
         ScaleWidth      =   4830
         TabIndex        =   7
         Top             =   630
         Width           =   4830
         Begin VB.ComboBox cboRuta 
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
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   0
            Width           =   2550
         End
         Begin VB.Label lblRuta 
            AutoSize        =   -1  'True
            Caption         =   "Ruta:"
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
            TabIndex        =   8
            Top             =   60
            Width           =   375
         End
      End
      Begin VB.PictureBox picFilterDiaSemana 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   165
         ScaleHeight     =   360
         ScaleWidth      =   1695
         TabIndex        =   4
         Top             =   630
         Width           =   1695
         Begin VB.ComboBox cboDiaSemana 
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
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   0
            Width           =   1350
         End
         Begin VB.Label lblWeekday 
            AutoSize        =   -1  'True
            Caption         =   "Día:"
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
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   1005
         ButtonWidth     =   2170
         ButtonHeight    =   1005
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Nuevo"
               Key             =   "NEW"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Propiedades"
               Key             =   "PROPERTIES"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Eliminar"
               Key             =   "DELETE"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Seleccionar"
               Key             =   "SELECT"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   5370
      Width           =   11895
      _ExtentX        =   20981
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
            Object.Width           =   19764
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
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   6165
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "DiaSemana"
         Text            =   "Día"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Hora"
         Text            =   "Hora"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "IDRuta"
         Text            =   "Ruta"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Vehiculo"
         Text            =   "Vehículo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Conductor"
         Text            =   "Conductor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "Conductor2"
         Text            =   "Conductor 2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "Activo"
         Text            =   "Activo"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmHorario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLoading As Boolean

Public FormWaitingForSelect As String
Public AllowMultipleSelect As Boolean
Public AllowMultipleRuta As Boolean

Public Sub FillListView(ByVal DiaSemana As Byte, ByVal Hora As Date, ByVal IDRuta As String)
    Dim ListItem As MSComctlLib.ListItem
    Dim recData As ADODB.Recordset
    Dim KeySave As Variant
    Dim CKeySave As Collection
    Dim SQL_Where As String
    Dim SQL_OrderBy As String
    
    If mLoading Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    If DiaSemana = 0 Then
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
        KeySave = KEY_STRINGER & DiaSemana & KEY_DELIMITER & Hora & KEY_DELIMITER & IDRuta
    End If
    
    SQL_Where = ""
    
    If pPersonal Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Horario.Personal = 0"
    End If
    
    If cboDiaSemana.ListIndex > 0 Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Horario.DiaSemana = " & cboDiaSemana.ListIndex
    End If
    
    If cboHora.ListIndex > 0 Then
        If cboHora.ListIndex < 7 Then
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "convert(char(8), Horario.Hora, 108) " & cboHora.Text & " '" & Format(dtpHoraDesde.value, "hh:nn:ss") & "'"
        Else
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "(convert(char(8), Horario.Hora, 108) >= '" & Format(dtpHoraDesde.value, "hh:nn:ss") & "' AND convert(char(8), Horario.Hora, 108) <= '" & Format(dtpHoraHasta.value, "hh:nn:ss") & "')"
        End If
    End If
    
    If cboRuta.ListIndex > 0 Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Horario.IDRuta = '" & ReplaceQuote(cboRuta.Text) & "'"
    Else
        If pCPermiso.RutaWhere <> "" Then
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & Replace(pCPermiso.RutaWhere, "%TABLENAME%", "Horario")
        End If
    End If
    
    If cboFilterActivo.ListIndex > 0 Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Horario.Activo = " & IIf(cboFilterActivo.ListIndex = 1, 1, 0)
    End If
    
    lvwData.ListItems.Clear
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Select Case lvwData.SortKey
        Case 0  'DIA SEMANA
            SQL_OrderBy = " ORDER BY Horario.DiaSemana" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Horario.Hora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Horario.IDRuta" & IIf(lvwData.SortOrder = lvwDescending, "", " DESC")
        Case 1  'HORA
            SQL_OrderBy = " ORDER BY Horario.Hora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Horario.DiaSemana" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Horario.IDRuta" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 2  'RUTA
            SQL_OrderBy = " ORDER BY Horario.DiaSemana" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Horario.IDRuta" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Horario.Hora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 3  'VEHICULO
            SQL_OrderBy = " ORDER BY Vehiculo.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Horario.DiaSemana" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Horario.IDRuta" & IIf(lvwData.SortOrder = lvwDescending, "", " DESC") & ", Horario.Hora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 4  'CONDUCTOR
            SQL_OrderBy = " ORDER BY Conductor" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Horario.DiaSemana" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Horario.IDRuta" & IIf(lvwData.SortOrder = lvwDescending, "", " DESC") & ", Horario.Hora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 5  'CONDUCTOR 2 o ACTIVO
            If pParametro.Viaje_Permite_2_Conductores Then
                SQL_OrderBy = " ORDER BY Conductor2" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Horario.DiaSemana" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Horario.IDRuta" & IIf(lvwData.SortOrder = lvwDescending, "", " DESC") & ", Horario.Hora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
            Else
                SQL_OrderBy = " ORDER BY Horario.Activo" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Horario.DiaSemana" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Horario.IDRuta" & IIf(lvwData.SortOrder = lvwDescending, "", " DESC") & ", Horario.Hora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
            End If
        Case 6  'ACTIVO
            SQL_OrderBy = " ORDER BY Horario.Activo" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Horario.DiaSemana" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Horario.IDRuta" & IIf(lvwData.SortOrder = lvwDescending, "", " DESC") & ", Horario.Hora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
    End Select
    
    Set recData = New ADODB.Recordset
    recData.Source = "SELECT Horario.DiaSemana, Horario.Hora, Horario.IDRuta, Vehiculo.Nombre AS Vehiculo, dbo.udf_GetEntidadApellidoYNombre(Conductor.Apellido, Conductor.Nombre) AS Conductor, dbo.udf_GetEntidadApellidoYNombre(Conductor2.Apellido, Conductor2.Nombre) AS Conductor2, Horario.Activo" & vbCr
    recData.Source = recData.Source & "FROM ((Horario LEFT JOIN Vehiculo ON Horario.IDVehiculo = Vehiculo.IDVehiculo) LEFT JOIN Persona AS Conductor ON Horario.IDConductor = Conductor.IDPersona) LEFT JOIN Persona AS Conductor2 ON Horario.IDConductor2 = Conductor2.IDPersona" & SQL_Where & SQL_OrderBy
    recData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With recData
        If Not .EOF Then
            Do While Not .EOF
                Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & .Fields("DiaSemana").value & KEY_DELIMITER & Format(.Fields("Hora").value, "Long Time") & KEY_DELIMITER & RTrim(.Fields("IDRuta").value), WeekdayName(.Fields("DiaSemana").value))
                ListItem.SubItems(1) = Format(.Fields("Hora").value, "Short Time")
                ListItem.SubItems(2) = RTrim(.Fields("IDRuta").value)
                ListItem.SubItems(3) = .Fields("Vehiculo").value & ""
                ListItem.SubItems(4) = .Fields("Conductor").value & ""
                If pParametro.Viaje_Permite_2_Conductores Then
                    ListItem.SubItems(5) = .Fields("Conductor2").value & ""
                    ListItem.SubItems(6) = GetBooleanString(.Fields("Activo").value)
                Else
                    ListItem.SubItems(5) = GetBooleanString(.Fields("Activo").value)
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
    ShowErrorMessage "Forms.Horario.FillListView", "Error al obtener la lista de Horarios."
End Sub

Public Sub FillComboBoxRuta()
    Dim recRuta As ADODB.Recordset
    Dim IDRutaSave As String
    
    IDRutaSave = cboRuta.Text

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Set recRuta = New ADODB.Recordset
    recRuta.Source = "SELECT IDRuta FROM Ruta WHERE Activo = 1" & IIf(pCPermiso.RutaWhere <> "", " AND " & Replace(pCPermiso.RutaWhere, "%TABLENAME%", "Ruta"), "") & " ORDER BY IDRuta"
    recRuta.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    cboRuta.Clear
    cboRuta.AddItem "<Todas>"
    Do While Not recRuta.EOF
        cboRuta.AddItem RTrim(recRuta("IDRuta").value)
        recRuta.MoveNext
    Loop
    recRuta.Close
    Set recRuta = Nothing

    cboRuta.ListIndex = CSM_Control_ComboBox.GetListIndexByText(cboRuta, IDRutaSave, cscpItemOrFirst)
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.Horario.FillComboBoxRuta", "Error al leer la lista de Rutas."
End Sub

Private Sub cboHora_Click()
    dtpHoraDesde.Visible = (cboHora.ListIndex > 0)
    lblHoraAnd.Visible = (cboHora.ListIndex = 7)
    dtpHoraHasta.Visible = (cboHora.ListIndex = 7)
    FillListView 0, Date, ""
End Sub

Private Sub cboRuta_Click()
    FillListView 0, Date, ""
End Sub

Private Sub cboFilterActivo_Click()
    FillListView 0, Date, ""
End Sub

Private Sub cboDiaSemana_Click()
    FillListView 0, Date, ""
End Sub

Private Sub dtpHoraDesde_Change()
    FillListView 0, Date, ""
End Sub

Private Sub dtpHoraHasta_Change()
    FillListView 0, Date, ""
End Sub

Private Sub cbrMain_HeightChanged(ByVal NewHeight As Single)
    ResizeControls NewHeight
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
            Case vbKeyS
                tlbMain_ButtonClick tlbMain.Buttons.Item("SELECT")
        End Select
    End If
End Sub

Private Sub Form_Load()
    Dim Weekday As Byte
    
    mLoading = True
    
    lvwData.GridLines = pParametro.ListView_GridLines
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL TOOLBAR
    Set tlbMain.ImageList = frmMDI.ilsFormToolbar
    Set tlbMain.HotImageList = frmMDI.ilsFormToolbarHot
    tlbMain.Buttons("NEW").Image = "NEW"
    tlbMain.Buttons("PROPERTIES").Image = "PROPERTIES"
    tlbMain.Buttons("DELETE").Image = "DELETE"
    tlbMain.Buttons("SELECT").Image = "SELECT"
    '//////////////////////////////////////////////////////////
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL LISTVIEW
    Set lvwData.ColumnHeaderIcons = frmMDI.ilsFormSortColumn
    '//////////////////////////////////////////////////////////
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL PIN
    Set tlbPin.ImageList = frmMDI.ilsFormPin
    '//////////////////////////////////////////////////////////
    
    cboDiaSemana.AddItem ITEM_ALL_MALE
    For Weekday = 1 To 7
        cboDiaSemana.AddItem WeekdayName(Weekday)
    Next Weekday
    cboDiaSemana.ListIndex = 0
        
    cboHora.AddItem "<Todas>"
    cboHora.AddItem "="
    cboHora.AddItem ">"
    cboHora.AddItem ">="
    cboHora.AddItem "<"
    cboHora.AddItem "<="
    cboHora.AddItem "<>"
    cboHora.AddItem "Entre"
    cboHora.ListIndex = 0
    
    dtpHoraDesde.value = CDate("00:00:00")
    dtpHoraHasta.value = CDate("23:59:00")
    
    FillComboBoxRuta
    cboRuta.ListIndex = 0
    
    cboFilterActivo.AddItem ITEM_ALL_MALE
    cboFilterActivo.AddItem "Sí"
    cboFilterActivo.AddItem "No"
    cboFilterActivo.ListIndex = FILTER_ACTIVO_LIST_INDEX
    
    CSM_Forms.ResizeAndPosition frmMDI, Me
    pParametro.GetCoolBarSettings "Horario", cbrMain
    pParametro.GetListViewSettings "Horario", lvwData
    lvwData.ColumnHeaders(lvwData.SortKey + 1).Icon = lvwData.SortOrder + 1
    tlbPin.Buttons("PIN").value = pParametro.Usuario_LeerNumero("Horario_Pin", tlbPin.Buttons("PIN").value)
    If tlbPin.Buttons("PIN").value = tbrUnpressed Then
        tlbPin.Buttons("PIN").Image = 1
    Else
        tlbPin.Buttons("PIN").Image = 2
    End If

    If Not pParametro.Viaje_Permite_2_Conductores Then
        lvwData.ColumnHeaders.Remove ("Conductor2")
    End If

    mLoading = False

    FillListView 0, Date, ""
End Sub

Private Sub Form_Resize()
    ResizeControls cbrMain.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visible = False
    WindowState = vbNormal
    pParametro.SaveCoolBarSettings "Horario", cbrMain
    pParametro.SaveListViewSettings "Horario", lvwData
    pParametro.Usuario_GuardarNumero "Horario_Pin", tlbPin.Buttons("PIN").value
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
    FillListView 0, Date, ""
End Sub

Private Sub lvwData_DblClick()
    If GetFormIndex(FormWaitingForSelect) > 0 Then
        tlbMain_ButtonClick tlbMain.Buttons.Item("SELECT")
    Else
        tlbMain_ButtonClick tlbMain.Buttons.Item("PROPERTIES")
    End If
End Sub

Private Sub lvwData_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        lvwData_DblClick
    End If
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim FormIndex As Long
    Dim ItemIndex As Long
    Dim SelectedItemCount As Long
    
    Dim Horario As Horario
    Dim CDiaSemana As Collection
    Dim CHora As Collection
    Dim CIDRuta As Collection
    Dim IDRuta As String
    
    Dim SelectedItems As Collection
    
    Select Case Button.Key
        Case "NEW"
            If pCPermiso.GotPermission(PERMISO_HORARIO_ADD) Then
                Screen.MousePointer = vbHourglass
                
                Set Horario = New Horario
                frmHorarioPropiedad.LoadDataAndShow Me, Horario
                Set Horario = Nothing
                
                Screen.MousePointer = vbDefault
            End If
        Case "PROPERTIES"
            If pCPermiso.GotPermission(PERMISO_HORARIO_MODIFY) Then
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
                
                If SelectedItemCount = 1 Then
                    Screen.MousePointer = vbHourglass
                    Set Horario = New Horario
                    Horario.DiaSemana = Val(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                    Horario.Hora = CDate(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER))
                    Horario.IDRuta = CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 3, KEY_DELIMITER)
                    If Horario.Load() Then
                        frmHorarioPropiedad.LoadDataAndShow Me, Horario
                    End If
                    Set Horario = Nothing
                    Screen.MousePointer = vbDefault
                Else
                    Set CDiaSemana = New Collection
                    Set CHora = New Collection
                    Set CIDRuta = New Collection
                    For ItemIndex = 1 To lvwData.ListItems.Count
                        If lvwData.ListItems(ItemIndex).Selected Then
                            CDiaSemana.Add Val(GetSubString(Mid(lvwData.ListItems(ItemIndex).Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                            CHora.Add CDate(GetSubString(Mid(lvwData.ListItems(ItemIndex).Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER))
                            CIDRuta.Add ReplaceQuote(GetSubString(Mid(lvwData.ListItems(ItemIndex).Key, Len(KEY_STRINGER) + 1), 3, KEY_DELIMITER))
                        End If
                    Next ItemIndex
                    Screen.MousePointer = vbHourglass
                    frmHorarioPropiedadMultiple.LoadDataAndShow Me, CDiaSemana, CHora, CIDRuta
                    Screen.MousePointer = vbDefault
                    Set CDiaSemana = Nothing
                    Set CHora = Nothing
                    Set CIDRuta = Nothing
                End If
            End If
        Case "DELETE"
            If pCPermiso.GotPermission(PERMISO_HORARIO_DELETE) Then
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
                
                If MsgBox(IIf(SelectedItemCount = 1, "¿Desea eliminar el Horario seleccionado?", "¿Desea eliminar los " & SelectedItemCount & " Horarios seleccionados?"), vbQuestion + vbYesNo, App.Title) = vbYes Then
                    Set Horario = New Horario
                    If SelectedItemCount = 1 Then
                        Horario.DiaSemana = Val(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                        Horario.Hora = CDate(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER))
                        Horario.IDRuta = CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 3, KEY_DELIMITER)
                        If Horario.Load() Then
                            Call Horario.Delete
                        End If
                    Else
                        Horario.RefreshListSkip = True
                        For ItemIndex = 1 To lvwData.ListItems.Count
                            If lvwData.ListItems(ItemIndex).Selected Then
                                Horario.DiaSemana = Val(GetSubString(Mid(lvwData.ListItems(ItemIndex).Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                                Horario.Hora = CDate(GetSubString(Mid(lvwData.ListItems(ItemIndex).Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER))
                                Horario.IDRuta = CSM_String.GetSubString(Mid(lvwData.ListItems(ItemIndex).Key, Len(KEY_STRINGER) + 1), 3, KEY_DELIMITER)
                                If Horario.Load() Then
                                    Call Horario.Delete
                                End If
                            End If
                        Next ItemIndex
                        RefreshList_RefreshHorario 1, Time, ""
                    End If
                    Set Horario = Nothing
                    
                    lvwData.SetFocus
                End If
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
                                MsgBox "No se puede Seleccionar más de un Horario a la vez.", vbExclamation, App.Title
                                Exit Sub
                            End If
                        End If
                    Next ItemIndex
                
                    Screen.MousePointer = vbHourglass
                    Forms(FormIndex).HorarioSelected Val(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER)), CDate(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER)), CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 3, KEY_DELIMITER)
                Else
                    Screen.MousePointer = vbHourglass
                    Set SelectedItems = New Collection
                    For ItemIndex = 1 To lvwData.ListItems.Count
                        If lvwData.ListItems(ItemIndex).Selected Then
                            If AllowMultipleRuta = False And IDRuta <> CSM_String.GetSubString(Mid(lvwData.ListItems(ItemIndex).Key, Len(KEY_STRINGER) + 1), 3, KEY_DELIMITER) Then
                                If IDRuta <> "" Then
                                    Screen.MousePointer = vbDefault
                                    MsgBox "No se pueden seleccionar Horarios de Distintas Rutas.", vbInformation, App.Title
                                    Set SelectedItems = Nothing
                                    Exit Sub
                                End If
                                IDRuta = CSM_String.GetSubString(Mid(lvwData.ListItems(ItemIndex).Key, Len(KEY_STRINGER) + 1), 3, KEY_DELIMITER)
                            End If
                            SelectedItems.Add lvwData.ListItems(ItemIndex).Key
                        End If
                    Next ItemIndex
                    Forms(FormIndex).MultipleHorarioSelected SelectedItems
                    Set SelectedItems = Nothing
                End If
                
                'Forms(FormIndex).SetFocus
                If tlbPin.Buttons("PIN").value = tbrUnpressed Then
                    Unload Me
                End If
                Screen.MousePointer = vbDefault
            End If
            FormWaitingForSelect = ""
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
    
    lvwData.Top = CoolBarHeight + CONTROL_SPACE
    lvwData.Left = CONTROL_SPACE
    lvwData.Width = ScaleWidth - (CONTROL_SPACE * 2)
    lvwData.Height = ScaleHeight - lvwData.Top - CONTROL_SPACE - stbMain.Height
    
    tlbPin.Top = ScaleHeight - 330
    tlbPin.Left = 15
End Sub
