VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmUsuario 
   Caption         =   "Usuarios"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7455
   Icon            =   "Usuario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   7455
   Begin MSComctlLib.Toolbar tlbPin 
      Height          =   330
      Left            =   15
      TabIndex        =   7
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
      Height          =   990
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   1746
      FixedOrder      =   -1  'True
      _CBWidth        =   7455
      _CBHeight       =   990
      _Version        =   "6.7.9782"
      Child1          =   "tlbMain"
      MinWidth1       =   4410
      MinHeight1      =   570
      Width1          =   4410
      FixedBackground1=   0   'False
      Key1            =   "Toolbar"
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Child2          =   "picUsuarioGrupo"
      MinWidth2       =   3405
      MinHeight2      =   330
      Width2          =   3405
      Key2            =   "UsuarioGrupo"
      NewRow2         =   0   'False
      AllowVertical2  =   0   'False
      Child3          =   "picActivo"
      MinWidth3       =   1605
      MinHeight3      =   330
      Width3          =   1605
      Key3            =   "Activo"
      NewRow3         =   0   'False
      AllowVertical3  =   0   'False
      Begin VB.PictureBox picUsuarioGrupo 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   165
         ScaleHeight     =   330
         ScaleWidth      =   5370
         TabIndex        =   8
         Top             =   630
         Width           =   5370
         Begin VB.ComboBox cboUsuarioGrupo 
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
            Left            =   660
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   0
            Width           =   2730
         End
         Begin VB.Label lblUsuarioGrupo 
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
            Left            =   0
            TabIndex        =   10
            Top             =   60
            Width           =   495
         End
      End
      Begin VB.PictureBox picActivo 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   5760
         ScaleHeight     =   330
         ScaleWidth      =   1605
         TabIndex        =   4
         Top             =   630
         Width           =   1605
         Begin VB.ComboBox cboActivo 
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
         Begin VB.Label lblActivo 
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
         Height          =   570
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   7335
         _ExtentX        =   12938
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
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Propiedades"
               Key             =   "PROPERTIES"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Eliminar"
               Key             =   "DELETE"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Seleccionar"
               Key             =   "SELECT"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   5595
      Width           =   7455
      _ExtentX        =   13150
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
            Object.Width           =   11933
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
      Height          =   3915
      Left            =   300
      TabIndex        =   0
      Top             =   1140
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   6906
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "ID"
         Text            =   "Login Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Nombre"
         Text            =   "Nombre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Descripcion"
         Text            =   "Descripci�n"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Grupo"
         Text            =   "Grupo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Empresa"
         Text            =   "Empresa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "Activo"
         Text            =   "Activo"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLoading As Boolean

Public FormWaitingForSelect As String

Public Sub FillListView(ByVal IDUsuario As String)
    Dim ListItem As MSComctlLib.ListItem
    Dim recData As ADODB.Recordset
    Dim KeySave As String
    Dim SQL_Where As String
    
    If mLoading Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    If IDUsuario = "" Then
        If Not lvwData.SelectedItem Is Nothing Then
            KeySave = lvwData.SelectedItem.Key
        End If
    Else
        KeySave = KEY_STRINGER & IDUsuario
    End If
    
    SQL_Where = " WHERE Usuario.IDUsuario <> '" & USUARIO_ID_ADMINISTRATOR & "' AND Usuario.LoginName <> '" & USUARIO_LOGINNAME_INTERNET & "'"
    
    If cboUsuarioGrupo.ListIndex > 0 Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Usuario.IDUsuarioGrupo = " & cboUsuarioGrupo.ItemData(cboUsuarioGrupo.ListIndex)
    End If
    
    If cboActivo.ListIndex > 0 Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Usuario.Activo = " & IIf(cboActivo.ListIndex = 1, 1, 0)
    End If
    
    lvwData.ListItems.Clear
    
    On Error GoTo ErrorHandler
    
    Set recData = New ADODB.Recordset
    recData.Source = "SELECT Usuario.IDUsuario, Usuario.LoginName, Usuario.Nombre, Usuario.Descripcion, UsuarioGrupo.Nombre AS UsuarioGrupo, Empresa.Nombre AS Empresa, Usuario.Activo FROM (Usuario INNER JOIN UsuarioGrupo ON Usuario.IDUsuarioGrupo = UsuarioGrupo.IDUsuarioGrupo) INNER JOIN Empresa ON Usuario.IDEmpresa = Empresa.IDEmpresa" & SQL_Where
    recData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With recData
        If Not .EOF Then
            Do While Not .EOF
                Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & .Fields("IDUsuario").Value, .Fields("LoginName").Value)
                ListItem.SubItems(1) = .Fields("Nombre").Value
                ListItem.SubItems(2) = .Fields("Descripcion").Value & ""
                ListItem.SubItems(3) = .Fields("UsuarioGrupo").Value
                ListItem.SubItems(4) = .Fields("Empresa").Value
                ListItem.SubItems(5) = GetBooleanString(.Fields("Activo").Value)
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

    If frmMDI.ActiveForm.Name = Me.Name And GetForegroundWindow() = frmMDI.hwnd And frmMDI.WindowState <> vbMinimized Then
        lvwData.SetFocus
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.Usuario.FillListView", "Error al obtener la Lista de Usuarios."
End Sub

Private Sub cboUsuarioGrupo_Click()
    FillListView ""
End Sub

Private Sub cboActivo_Click()
    FillListView ""
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
    
    Call FillComboBoxUsuarioGrupo
    
    cboActivo.AddItem ITEM_ALL_MALE
    cboActivo.AddItem "S�"
    cboActivo.AddItem "No"
    cboActivo.ListIndex = FILTER_ACTIVO_LIST_INDEX
    
    CSM_Forms.ResizeAndPosition frmMDI, Me
    pParametro.GetCoolBarSettings "Usuario", cbrMain
    pParametro.GetListViewSettings "Usuario", lvwData
    lvwData.ColumnHeaders(lvwData.SortKey + 1).Icon = lvwData.SortOrder + 1
    tlbPin.Buttons("PIN").Value = pParametro.Usuario_LeerNumero("Usuario_Pin", tlbPin.Buttons("PIN").Value)
    If tlbPin.Buttons("PIN").Value = tbrUnpressed Then
        tlbPin.Buttons("PIN").Image = 1
    Else
        tlbPin.Buttons("PIN").Image = 2
    End If

    mLoading = False

    FillListView ""
End Sub

Private Sub Form_Resize()
    ResizeControls cbrMain.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visible = False
    WindowState = vbNormal
    pParametro.SaveCoolBarSettings "Usuario", cbrMain
    pParametro.SaveListViewSettings "Usuario", lvwData
    pParametro.Usuario_GuardarNumero "Usuario_Pin", tlbPin.Buttons("PIN").Value
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
    Dim Usuario As Usuario
    
    Select Case Button.Key
        Case "NEW"
            If pCPermiso.GotPermission(PERMISO_USUARIO_ADD) Then
                Screen.MousePointer = vbHourglass
                
                Set Usuario = New Usuario
                frmUsuarioPropiedad.LoadDataAndShow Me, Usuario
                Set Usuario = Nothing
                
                Screen.MousePointer = vbDefault
            End If
        Case "PROPERTIES"
            If pCPermiso.GotPermission(PERMISO_USUARIO_MODIFY) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ning�n Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                Screen.MousePointer = vbHourglass
                
                Set Usuario = New Usuario
                Usuario.IDUsuario = Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1)
                If Usuario.Load() Then
                    frmUsuarioPropiedad.LoadDataAndShow Me, Usuario
                Else
                    lvwData.SetFocus
                End If
                Set Usuario = Nothing
                
                Screen.MousePointer = vbDefault
            End If
        Case "DELETE"
            If pCPermiso.GotPermission(PERMISO_USUARIO_DELETE) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ning�n Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                If MsgBox("�Desea eliminar el Usuario seleccionado?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                    Set Usuario = New Usuario
                    Usuario.IDUsuario = Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1)
                    If Usuario.Load() Then
                        Call Usuario.Delete
                    End If
                    Set Usuario = Nothing
                    lvwData.SetFocus
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
                If Not GetBooleanValueFromString(lvwData.SelectedItem.SubItems(3)) Then
                    MsgBox "No puede seleccionar este Item ya que est� inactivo.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                Screen.MousePointer = vbHourglass
                Forms(FormIndex).UsuarioSelected Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1)
                Forms(FormIndex).SetFocus
                If tlbPin.Buttons("PIN").Value = tbrUnpressed Then
                    Unload Me
                End If
                Screen.MousePointer = vbDefault
            End If
            FormWaitingForSelect = ""
    End Select
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
    
    lvwData.Top = CoolBarHeight + (CONTROL_SPACE * 2)
    lvwData.Left = CONTROL_SPACE
    lvwData.Width = ScaleWidth - (CONTROL_SPACE * 2)
    lvwData.Height = ScaleHeight - lvwData.Top - CONTROL_SPACE - stbMain.Height
    
    tlbPin.Top = ScaleHeight - 330
    tlbPin.Left = 15
End Sub

Public Sub FillComboBoxUsuarioGrupo()
    Dim KeySave As Long
    
    If cboUsuarioGrupo.ListCount > 0 Then
        KeySave = cboUsuarioGrupo.ItemData(cboUsuarioGrupo.ListIndex)
    End If
    Call CSM_Control_ComboBox.FillFromSQL(cboUsuarioGrupo, "(SELECT -1 AS IDUsuarioGrupo, '<Todos>' AS Nombre, 1 AS Orden FROM UsuarioGrupo) UNION (SELECT IDUsuarioGrupo, Nombre, 2 AS Orden FROM UsuarioGrupo WHERE Activo = 1) ORDER BY Orden, Nombre", "IDUsuarioGrupo", "Nombre", "Grupos de Usuario", cscpItemOrfirst, KeySave)
End Sub
