VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFranco 
   Caption         =   "Francos"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Franco.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   9420
   Begin MSComctlLib.Toolbar tlbPin 
      Height          =   330
      Left            =   15
      TabIndex        =   4
      Top             =   5640
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
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   1799
      FixedOrder      =   -1  'True
      _CBWidth        =   9420
      _CBHeight       =   1020
      _Version        =   "6.7.9816"
      Child1          =   "tlbMain"
      MinWidth1       =   4410
      MinHeight1      =   570
      Width1          =   4410
      FixedBackground1=   0   'False
      Key1            =   "Toolbar"
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Child2          =   "picConductor"
      MinWidth2       =   4545
      MinHeight2      =   330
      Width2          =   4545
      Key2            =   "Conductor"
      NewRow2         =   0   'False
      AllowVertical2  =   0   'False
      Child3          =   "picFecha"
      MinWidth3       =   6675
      MinHeight3      =   360
      Width3          =   6675
      Key3            =   "Fecha"
      NewRow3         =   0   'False
      AllowVertical3  =   0   'False
      Begin VB.PictureBox picFecha 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   165
         ScaleHeight     =   360
         ScaleWidth      =   9165
         TabIndex        =   8
         Top             =   630
         Width           =   9165
         Begin VB.ComboBox cboFecha 
            Height          =   330
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   0
            Width           =   1035
         End
         Begin VB.CommandButton cmdAnteriorDesde 
            Height          =   315
            Left            =   1680
            Picture         =   "Franco.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Anterior"
            Top             =   0
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.CommandButton cmdSiguienteDesde 
            Height          =   315
            Left            =   3420
            Picture         =   "Franco.frx":0596
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Siguiente"
            Top             =   0
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.CommandButton cmdHoyDesde 
            Height          =   315
            Left            =   3720
            Picture         =   "Franco.frx":0B20
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Hoy"
            Top             =   0
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton cmdAnteriorHasta 
            Height          =   315
            Left            =   4320
            Picture         =   "Franco.frx":0C6A
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   "Anterior"
            Top             =   0
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.CommandButton cmdSiguienteHasta 
            Height          =   315
            Left            =   6060
            Picture         =   "Franco.frx":11F4
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Siguiente"
            Top             =   0
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.CommandButton cmdHoyHasta 
            Height          =   315
            Left            =   6360
            Picture         =   "Franco.frx":177E
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Hoy"
            Top             =   0
            Visible         =   0   'False
            Width           =   315
         End
         Begin MSComCtl2.DTPicker dtpFechaDesde 
            Height          =   315
            Left            =   1980
            TabIndex        =   16
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
            TabIndex        =   17
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
         Begin VB.Label lblFecha 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   210
            Left            =   0
            TabIndex        =   19
            Top             =   60
            Width           =   495
         End
         Begin VB.Label lblFechaAnd 
            AutoSize        =   -1  'True
            Caption         =   "y"
            Height          =   210
            Left            =   4140
            TabIndex        =   18
            Top             =   60
            Visible         =   0   'False
            Width           =   90
         End
      End
      Begin VB.PictureBox picConductor 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   4665
         ScaleHeight     =   330
         ScaleWidth      =   4665
         TabIndex        =   5
         Top             =   150
         Width           =   4665
         Begin VB.ComboBox cboConductor 
            Height          =   330
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   0
            Width           =   3570
         End
         Begin VB.Label lblConductor 
            AutoSize        =   -1  'True
            Caption         =   "Conductor:"
            Height          =   210
            Left            =   0
            TabIndex        =   7
            Top             =   60
            Width           =   795
         End
      End
      Begin MSComctlLib.Toolbar tlbMain 
         Height          =   570
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   1005
         ButtonWidth     =   2170
         ButtonHeight    =   1005
         ToolTips        =   0   'False
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
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
      Top             =   5520
      Width           =   9420
      _ExtentX        =   16616
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
            Object.Width           =   15399
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
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   1140
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   7435
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Fecha"
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Conductor"
         Text            =   "Conductor"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   "Importe"
         Text            =   "Importe"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmFranco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLoading As Boolean

Public FormWaitingForSelect As String

Public Sub FillListView(ByVal Fecha As Date, ByVal IDPersona As Long)
    Dim ListItem As MSComctlLib.ListItem
    Dim recData As ADODB.Recordset
    Dim KeySave As String
    
    Dim SQL_Where As String
    Dim SQL_OrderBy As String
    
    If mLoading Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    If Fecha = DATE_TIME_FIELD_NULL_VALUE Then
        If Not lvwData.SelectedItem Is Nothing Then
            KeySave = lvwData.SelectedItem.Key
        End If
    Else
        KeySave = KEY_STRINGER & Fecha & KEY_DELIMITER & IDPersona
    End If
    
    lvwData.ListItems.Clear
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    'CONDUCTOR
    If cboConductor.ListIndex > 0 Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Franco.IDPersona = " & cboConductor.ItemData(cboConductor.ListIndex)
    End If
    
    'DATE FILTER
    Select Case cboFecha.ListIndex
        Case 0  'ALL
        Case 1  'EQUAL
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Franco.Fecha BETWEEN '" & Format(dtpFechaDesde.value, "yyyy/mm/dd") & " 00:00:00' AND '" & Format(dtpFechaDesde.value, "yyyy/mm/dd") & " 23:59:00'"
        Case 2  'GREATER
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Franco.Fecha > '" & Format(dtpFechaDesde.value, "yyyy/mm/dd") & " 23:59:00'"
        Case 3  'GREATER OR EQUAL
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Franco.Fecha >= '" & Format(dtpFechaDesde.value, "yyyy/mm/dd") & " 00:00:00'"
        Case 4  'MINOR
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Franco.Fecha < '" & Format(dtpFechaDesde.value, "yyyy/mm/dd") & " 00:00:00'"
        Case 5  'MINOR OR EQUAL
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Franco.Fecha <= '" & Format(dtpFechaDesde.value, "yyyy/mm/dd") & " 23:59:00'"
        Case 6  'NOT EQUAL
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Franco.Fecha NOT BETWEEN '" & Format(dtpFechaDesde.value, "yyyy/mm/dd") & " 00:00:00' AND '" & Format(dtpFechaDesde.value, "yyyy/mm/dd") & " 23:59:00'"
        Case 7  'BETWEEN
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Franco.Fecha BETWEEN '" & Format(dtpFechaDesde.value, "yyyy/mm/dd") & " 00:00:00' AND '" & Format(dtpFechaHasta.value, "yyyy/mm/dd") & " 23:59:00'"
    End Select
    
    Select Case lvwData.SortKey
        Case 0  'FECHA + CONDUCTOR
            SQL_OrderBy = " ORDER BY Franco.Fecha" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", dbo.udf_GetEntidadApellidoYNombre(Persona.Apellido, Persona.Nombre)" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 1  'CONDUCTOR + FECHA
            SQL_OrderBy = " ORDER BY dbo.udf_GetEntidadApellidoYNombre(Persona.Apellido, Persona.Nombre)" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Franco.Fecha" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 2  'IMPORTE + FECHA + CONDUCTOR
            SQL_OrderBy = " ORDER BY Franco.Importe" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Franco.Fecha" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", dbo.udf_GetEntidadApellidoYNombre(Persona.Apellido, Persona.Nombre)" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
    End Select
    
    lvwData.ListItems.Clear
    
    Set recData = New ADODB.Recordset
    recData.Source = "SELECT Franco.IDPersona, Franco.Fecha, dbo.udf_GetEntidadApellidoYNombre(Persona.Apellido, Persona.Nombre) AS ApellidoNombre, Franco.Importe FROM Franco INNER JOIN Persona ON Franco.IDPersona = Persona.IDPersona" & SQL_Where & SQL_OrderBy
    recData.MaxRecords = pParametro.Recordset_MaxRecords
    recData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With recData
        If Not .EOF Then
            Do While Not .EOF
                Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & .Fields("Fecha").value & KEY_DELIMITER & .Fields("IDPersona").value, Format(.Fields("Fecha").value, "Short Date"))
                ListItem.SubItems(1) = .Fields("ApellidoNombre").value
                If IsNull(.Fields("Importe").value) Then
                    ListItem.SubItems(2) = " "
                Else
                    ListItem.SubItems(2) = Format(.Fields("Importe").value, "Currency")
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
    Set lvwData.SelectedItem = lvwData.ListItems(KeySave)
    lvwData.SelectedItem.EnsureVisible

    If frmMDI.ActiveForm.Name = Me.Name And GetForegroundWindow() = frmMDI.hwnd And frmMDI.WindowState <> vbMinimized Then
        lvwData.SetFocus
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.Franco.FillListView", "Error al obtener la lista de Francos"
End Sub

Private Sub cboFecha_Click()
    cmdAnteriorDesde.Visible = (cboFecha.ListIndex > 0)
    dtpFechaDesde.Visible = (cboFecha.ListIndex > 0)
    cmdSiguienteDesde.Visible = (cboFecha.ListIndex > 0)
    cmdHoyDesde.Visible = (cboFecha.ListIndex > 0)
    
    lblFechaAnd.Visible = (cboFecha.ListIndex = 7)
    
    cmdAnteriorHasta.Visible = (cboFecha.ListIndex = 7)
    dtpFechaHasta.Visible = (cboFecha.ListIndex = 7)
    cmdSiguienteHasta.Visible = (cboFecha.ListIndex = 7)
    cmdHoyHasta.Visible = (cboFecha.ListIndex = 7)
    
    FillListView DATE_TIME_FIELD_NULL_VALUE, 0
End Sub

Private Sub dtpFechaDesde_Change()
    FillListView DATE_TIME_FIELD_NULL_VALUE, 0
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
    FillListView DATE_TIME_FIELD_NULL_VALUE, 0
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

Private Sub cboConductor_Click()
    FillListView DATE_TIME_FIELD_NULL_VALUE, 0
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
    
    FillComboBoxConductor
    
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
    dtpFechaHasta.value = Date
        
    CSM_Forms.ResizeAndPosition frmMDI, Me
    pParametro.GetCoolBarSettings "Franco", cbrMain
    pParametro.GetListViewSettings "Franco", lvwData
    tlbPin.Buttons("PIN").value = pParametro.Usuario_LeerNumero("Franco_Pin", tlbPin.Buttons("PIN").value)
    If tlbPin.Buttons("PIN").value = tbrUnpressed Then
        tlbPin.Buttons("PIN").Image = 1
    Else
        tlbPin.Buttons("PIN").Image = 2
    End If

    mLoading = False

    FillListView DATE_TIME_FIELD_NULL_VALUE, 0
End Sub

Private Sub Form_Resize()
    ResizeControls cbrMain.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visible = False
    WindowState = vbNormal
    pParametro.SaveCoolBarSettings "Franco", cbrMain
    pParametro.SaveListViewSettings "Franco", lvwData
    pParametro.Usuario_GuardarNumero "Franco_Pin", tlbPin.Buttons("PIN").value
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
    Dim Franco As Franco
    
    Select Case Button.Key
        Case "NEW"
            If pCPermiso.GotPermission(PERMISO_FRANCO_ADD) Then
                Screen.MousePointer = vbHourglass
                
                Set Franco = New Franco
                frmFrancoPropiedad.LoadDataAndShow Me, Franco
                Set Franco = Nothing
                
                Screen.MousePointer = vbDefault
            End If
        Case "PROPERTIES"
            If pCPermiso.GotPermission(PERMISO_FRANCO_MODIFY) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                Screen.MousePointer = vbHourglass
                
                Set Franco = New Franco
                Franco.Fecha = CDate(CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                Franco.IDPersona = Val(CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER))
                If Franco.Load() Then
                    frmFrancoPropiedad.LoadDataAndShow Me, Franco
                End If
                Set Franco = Nothing
                
                Screen.MousePointer = vbDefault
            End If
        Case "DELETE"
            If pCPermiso.GotPermission(PERMISO_FRANCO_DELETE) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                If MsgBox("¿Desea eliminar el Franco seleccionado?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                    Set Franco = New Franco
                    Franco.Fecha = CDate(CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                    Franco.IDPersona = Val(CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER))
                    If Franco.Load() Then
                        Call Franco.Delete
                    End If
                    Set Franco = Nothing
                    
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
                Screen.MousePointer = vbHourglass
                Forms(FormIndex).FrancoSelected CDate(CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER)), Val(CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER))

                Forms(FormIndex).SetFocus
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

Public Sub FillComboBoxConductor()
    Dim KeySave As Long
    
    If cboConductor.ListCount > 0 Then
        KeySave = cboConductor.ItemData(cboConductor.ListIndex)
    End If
    Call CSM_Control_ComboBox.FillFromSQL(cboConductor, "(SELECT -1 AS IDPersona, '<Todos>' AS ApellidoNombre, 1 AS Orden FROM Persona) UNION (SELECT IDPersona, dbo.udf_GetEntidadApellidoYNombre(Apellido, Nombre) AS ApellidoNombre, 2 AS Orden FROM Persona WHERE Activo = 1 AND EntidadTipo = '" & ENTIDAD_TIPO_PERSONA_CONDUCTOR & "') ORDER BY Orden, ApellidoNombre", "IDPersona", "ApellidoNombre", "Conductores", cscpItemOrFirst, KeySave)
End Sub
