VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmVehiculoMantenimiento 
   Caption         =   "Mantenimiento de Veh�culos"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   Icon            =   "VehiculoMantenimiento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   9195
   Begin MSComctlLib.Toolbar tlbPin 
      Height          =   330
      Left            =   0
      TabIndex        =   7
      Top             =   5520
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
      Height          =   600
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   1058
      BandCount       =   2
      FixedOrder      =   -1  'True
      _CBWidth        =   9195
      _CBHeight       =   600
      _Version        =   "6.7.9782"
      Child1          =   "tlbMain"
      MinWidth1       =   5505
      MinHeight1      =   540
      Width1          =   5505
      FixedBackground1=   0   'False
      Key1            =   "Toolbar"
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Child2          =   "picFilterActivo"
      MinWidth2       =   1605
      MinHeight2      =   330
      Width2          =   1605
      Key2            =   "FilterActivo"
      NewRow2         =   0   'False
      AllowVertical2  =   0   'False
      Begin VB.PictureBox picFilterActivo 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   7500
         ScaleHeight     =   330
         ScaleWidth      =   1605
         TabIndex        =   4
         Top             =   135
         Width           =   1605
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
            TabIndex        =   5
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
            TabIndex        =   6
            Top             =   60
            Width           =   510
         End
      End
      Begin MSComctlLib.Toolbar tlbMain 
         Height          =   540
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   953
         ButtonWidth     =   1931
         ButtonHeight    =   953
         ToolTips        =   0   'False
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
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
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Copiar"
               Key             =   "COPY"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   5475
      Width           =   9195
      _ExtentX        =   16219
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
            Object.Width           =   15002
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
      Left            =   60
      TabIndex        =   0
      Top             =   1200
      Width           =   4815
      _ExtentX        =   8493
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Vehiculo"
         Text            =   "Veh�culo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Grupo"
         Text            =   "Grupo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Tipo"
         Text            =   "Tipo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Activo"
         Text            =   "Activo"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmVehiculoMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLoading As Boolean

Public FormWaitingForSelect As String

Public Sub FillListView(ByVal IDVehiculo As Long, ByVal IDVehiculoMantenimientoGrupo As Long)
    Dim ListItem As MSComctlLib.ListItem
    Dim recData As ADODB.Recordset
    Dim KeySave As String
    Dim SQL_Where As String
    Dim SQL_OrderBy As String
    Dim VehiculoMantenimiento As VehiculoMantenimiento
    
    If mLoading Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    If IDVehiculo = 0 Then
        If Not lvwData.SelectedItem Is Nothing Then
            KeySave = lvwData.SelectedItem.Key
        End If
    Else
        KeySave = KEY_STRINGER & IDVehiculo & KEY_DELIMITER & IDVehiculoMantenimientoGrupo
    End If
    
    SQL_Where = ""
    
    If cboFilterActivo.ListIndex > 0 Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "VehiculoMantenimiento.Activo = " & IIf(cboFilterActivo.ListIndex = 1, 1, 0)
    End If
    
    Select Case lvwData.SortKey
        Case 0  'VEHICULO + GRUPO
            SQL_OrderBy = " ORDER BY Vehiculo.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", VehiculoMantenimientoGrupo.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 1  'GRUPO + VEHICULO
            SQL_OrderBy = " ORDER BY VehiculoMantenimientoGrupo.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Vehiculo.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 2  'TIPO + VEHICULO + GRUPO
            SQL_OrderBy = " ORDER BY VehiculoMantenimiento.Tipo" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Vehiculo.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", VehiculoMantenimientoGrupo.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 3  'ACTIVO + VEHICULO + GRUPO
            SQL_OrderBy = " ORDER BY VehiculoMantenimiento.Activo" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Vehiculo.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", VehiculoMantenimientoGrupo.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
    End Select
    
    lvwData.ListItems.Clear
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Set recData = New ADODB.Recordset
    recData.Source = "SELECT VehiculoMantenimiento.IDVehiculo, Vehiculo.Nombre AS Vehiculo, VehiculoMantenimiento.IDVehiculoMantenimientoGrupo, VehiculoMantenimientoGrupo.Nombre AS VehiculoMantenimientoGrupo, VehiculoMantenimiento.Tipo, VehiculoMantenimiento.Activo FROM (VehiculoMantenimiento INNER JOIN Vehiculo ON VehiculoMantenimiento.IDVehiculo = Vehiculo.IDVehiculo) INNER JOIN VehiculoMantenimientoGrupo ON VehiculoMantenimiento.IDVehiculoMantenimientoGrupo = VehiculoMantenimientoGrupo.IDVehiculoMantenimientoGrupo" & SQL_Where & SQL_OrderBy
    recData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With recData
        If Not .EOF Then
            Set VehiculoMantenimiento = New VehiculoMantenimiento
            
            Do While Not .EOF
                Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & .Fields("IDVehiculo").Value & KEY_DELIMITER & .Fields("IDVehiculoMantenimientoGrupo").Value, .Fields("Vehiculo").Value)
                ListItem.SubItems(1) = .Fields("VehiculoMantenimientoGrupo").Value
                VehiculoMantenimiento.Tipo = .Fields("Tipo").Value
                ListItem.SubItems(2) = VehiculoMantenimiento.Tipo_ToString
                ListItem.SubItems(3) = GetBooleanString(.Fields("Activo").Value)
                .MoveNext
            Loop
            
            Set VehiculoMantenimiento = Nothing
            
            stbMain.Panels("TEXT").Text = .RecordCount & " items"
        Else
            stbMain.Panels("TEXT").Text = "No hay items."
        End If
        .Close
    End With
    Set recData = Nothing
    
    On Error Resume Next
    Set lvwData.SelectedItem = lvwData.ListItems(KeySave)

    If frmMDI.ActiveForm.Name = Me.Name And GetForegroundWindow() = frmMDI.hwnd And frmMDI.WindowState <> vbMinimized Then
        lvwData.SetFocus
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.VehiculoMantenimiento.FillListView", "Error al obtener la lista de los Mantenimientos de los Veh�culos."
End Sub

Private Sub cboFilterActivo_Click()
    FillListView 0, 0
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
    tlbMain.Buttons("COPY").Image = "COPY"
    '//////////////////////////////////////////////////////////
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL LISTVIEW
    Set lvwData.ColumnHeaderIcons = frmMDI.ilsFormSortColumn
    '//////////////////////////////////////////////////////////
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL PIN
    Set tlbPin.ImageList = frmMDI.ilsFormPin
    '//////////////////////////////////////////////////////////
    
    cboFilterActivo.AddItem ITEM_ALL_MALE
    cboFilterActivo.AddItem "S�"
    cboFilterActivo.AddItem "No"
    cboFilterActivo.ListIndex = FILTER_ACTIVO_LIST_INDEX
    
    CSM_Forms.ResizeAndPosition frmMDI, Me
    pParametro.GetCoolBarSettings "VehiculoMantenimiento", cbrMain
    pParametro.GetListViewSettings "VehiculoMantenimiento", lvwData
    lvwData.ColumnHeaders(lvwData.SortKey + 1).Icon = lvwData.SortOrder + 1
    tlbPin.Buttons("PIN").Value = pParametro.Usuario_LeerNumero("VehiculoMantenimiento_Pin", tlbPin.Buttons("PIN").Value)
    If tlbPin.Buttons("PIN").Value = tbrUnpressed Then
        tlbPin.Buttons("PIN").Image = 1
    Else
        tlbPin.Buttons("PIN").Image = 2
    End If
    
    mLoading = False

    FillListView 0, 0
End Sub

Private Sub Form_Resize()
    ResizeControls cbrMain.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visible = False
    WindowState = vbNormal
    pParametro.SaveCoolBarSettings "VehiculoMantenimiento", cbrMain
    pParametro.SaveListViewSettings "VehiculoMantenimiento", lvwData
    pParametro.Usuario_GuardarNumero "VehiculoMantenimiento_Pin", tlbPin.Buttons("PIN").Value
    Set frmVehiculoMantenimiento = Nothing
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
    FillListView 0, 0
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
    Dim VehiculoMantenimiento As VehiculoMantenimiento
    
    Set VehiculoMantenimiento = New VehiculoMantenimiento
    
    Select Case Button.Key
        Case "NEW"
            If pCPermiso.GotPermission(PERMISO_VEHICULO_MANTENIMIENTO_ADD) Then
                Screen.MousePointer = vbHourglass
                frmVehiculoMantenimientoPropiedad.LoadDataAndShow Me, VehiculoMantenimiento
                Screen.MousePointer = vbDefault
            End If
        Case "PROPERTIES"
            If pCPermiso.GotPermission(PERMISO_VEHICULO_MANTENIMIENTO_MODIFY) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ning�n Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                VehiculoMantenimiento.IDVehiculo = Val(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                VehiculoMantenimiento.IDVehiculoMantenimientoGrupo = Val(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER))
                If Not VehiculoMantenimiento.Load() Then
                    Set VehiculoMantenimiento = Nothing
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                Screen.MousePointer = vbHourglass
                frmVehiculoMantenimientoPropiedad.LoadDataAndShow Me, VehiculoMantenimiento
                Screen.MousePointer = vbDefault
            End If
        Case "DELETE"
            If pCPermiso.GotPermission(PERMISO_VEHICULO_MANTENIMIENTO_DELETE) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ning�n Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                If MsgBox("�Desea eliminar el Mantenimiento del Veh�culo seleccionado?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                    
                    VehiculoMantenimiento.IDVehiculo = Val(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                    VehiculoMantenimiento.IDVehiculoMantenimientoGrupo = Val(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER))
                    If Not VehiculoMantenimiento.Load() Then
                        Set VehiculoMantenimiento = Nothing
                        lvwData.SetFocus
                        Exit Sub
                    End If
                    If Not VehiculoMantenimiento.Delete() Then
                        Set VehiculoMantenimiento = Nothing
                        lvwData.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        Case "SELECT"
            FormIndex = GetFormIndex(FormWaitingForSelect)
            If FormIndex >= 0 Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ning�n Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                If Not GetBooleanValueFromString(lvwData.SelectedItem.SubItems(2)) Then
                    MsgBox "No puede seleccionar este Item ya que est� inactivo.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                Screen.MousePointer = vbHourglass
                Forms(FormIndex).VehiculoMantenimientoSelected Val(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER)), Val(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER))
                Forms(FormIndex).SetFocus
                If tlbPin.Buttons("PIN").Value = tbrUnpressed Then
                    Unload Me
                End If
                Screen.MousePointer = vbDefault
            End If
            FormWaitingForSelect = ""
        Case "COPY"
            frmVehiculoMantenimientoCopy.Show
    End Select
    
    Set VehiculoMantenimiento = Nothing
End Sub

Private Sub tlbPin_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Value = tbrUnpressed Then
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
