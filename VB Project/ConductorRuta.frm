VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmConductorRuta 
   Caption         =   "Rutas del Conductor"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10170
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ConductorRuta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5745
   ScaleWidth      =   10170
   Begin MSComctlLib.Toolbar tlbPin 
      Height          =   330
      Left            =   15
      TabIndex        =   4
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
      Height          =   600
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   1058
      BandCount       =   2
      FixedOrder      =   -1  'True
      _CBWidth        =   10170
      _CBHeight       =   600
      _Version        =   "6.7.9816"
      Child1          =   "tlbMain"
      MinWidth1       =   4410
      MinHeight1      =   540
      Width1          =   4410
      FixedBackground1=   0   'False
      Key1            =   "Toolbar"
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Child2          =   "picConductor"
      MinWidth2       =   3885
      MinHeight2      =   360
      Width2          =   3885
      NewRow2         =   0   'False
      AllowVertical2  =   0   'False
      Begin VB.PictureBox picConductor 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   6195
         ScaleHeight     =   360
         ScaleWidth      =   3885
         TabIndex        =   5
         Top             =   120
         Width           =   3885
         Begin VB.ComboBox cboConductor 
            Height          =   330
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   0
            Width           =   2970
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
         Height          =   540
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   953
         ButtonWidth     =   1931
         ButtonHeight    =   953
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
      Top             =   5385
      Width           =   10170
      _ExtentX        =   17939
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
            Object.Width           =   16722
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
      Height          =   4095
      Left            =   180
      TabIndex        =   0
      Top             =   780
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   7223
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Conductor"
         Text            =   "Conductor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Ruta"
         Text            =   "Ruta"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   "ImporteTramoCompleto"
         Text            =   "Importe Tramo Completo"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   "ImporteTramo1"
         Text            =   "Importe Tramo 1"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Key             =   "ImporteTramo2"
         Text            =   "Importe Tramo 2"
         Object.Width           =   2822
      EndProperty
   End
End
Attribute VB_Name = "frmConductorRuta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FormWaitingForSelect As String

Private mLoading As Boolean

Public Sub ForceRefresh()
    FillListView cboConductor.ItemData(cboConductor.ListIndex), ""
End Sub

Public Function FillListView(ByVal IDPersona As Long, ByVal IDRuta As String) As Boolean
    Dim ListItem As MSComctlLib.ListItem
    Dim recData As ADODB.Recordset
    Dim KeySave As Variant
    Dim SQL_Where As String
    Dim SQL_OrderBy As String
    
    If cboConductor.ItemData(cboConductor.ListIndex) <> IDPersona And cboConductor.ListIndex > 0 Then
        Exit Function
    End If
    
    If mLoading Then
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    If IDRuta = "" Then
        If Not lvwData.SelectedItem Is Nothing Then
            KeySave = lvwData.SelectedItem.Key
        End If
    Else
        KeySave = KEY_STRINGER & IDPersona & KEY_DELIMITER & IDRuta
    End If
    
    If cboConductor.ListIndex > 0 Then
        SQL_Where = " WHERE ConductorRuta.IDPersona = " & IDPersona
    End If
    
    If pCPermiso.RutaWhere <> "" Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & Replace(pCPermiso.RutaWhere, "%TABLENAME%", "ConductorRuta")
    End If
    
    lvwData.ListItems.Clear
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Select Case lvwData.SortKey
        Case 0  'CONDUCTOR
            SQL_OrderBy = " ORDER BY Conductor" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", ConductorRuta.IDRuta" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 1  'RUTA
            SQL_OrderBy = " ORDER BY ConductorRuta.IDRuta" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Conductor" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 2  'IMPORTE TRAMO COMPLETO
            SQL_OrderBy = " ORDER BY ConductorRuta.ConductorImporteTramoCompleto" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Conductor" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", ConductorRuta.IDRuta" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 3  'IMPORTE TRAMO 1
            SQL_OrderBy = " ORDER BY ConductorRuta.ConductorImporteTramo1" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Conductor" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", ConductorRuta.IDRuta" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 4  'IMPORTE TRAMO 2
            SQL_OrderBy = " ORDER BY ConductorRuta.ConductorImporteTramo2" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Conductor" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", ConductorRuta.IDRuta" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
    End Select
    
    Set recData = New ADODB.Recordset
    recData.Source = "SELECT Persona.IDPersona, Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Conductor, ConductorRuta.IDRuta, Ruta.Permite2Conductores, ConductorRuta.ConductorImporteTramoCompleto, ConductorRuta.ConductorImporteTramo1, ConductorRuta.ConductorImporteTramo2 FROM (ConductorRuta INNER JOIN Persona ON ConductorRuta.IDPersona = Persona.IDPersona) INNER JOIN Ruta ON ConductorRuta.IDRuta = Ruta.IDRuta" & SQL_Where & SQL_OrderBy
    recData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With recData
        If Not .EOF Then
            Do While Not .EOF
                Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & .Fields("IDPersona").value & KEY_DELIMITER & RTrim(.Fields("IDRuta").value), .Fields("Conductor").value)
                ListItem.SubItems(1) = RTrim(.Fields("IDRuta").value)
                If IsNull(.Fields("ConductorImporteTramoCompleto").value) Then
                    ListItem.SubItems(2) = " "
                Else
                    ListItem.SubItems(2) = Format(.Fields("ConductorImporteTramoCompleto").value, "Currency")
                End If
                If pParametro.Viaje_Permite_2_Conductores Then
                    If .Fields("Permite2Conductores").value = False Or IsNull(.Fields("ConductorImporteTramo1").value) Then
                        ListItem.SubItems(3) = " "
                    Else
                        ListItem.SubItems(3) = Format(.Fields("ConductorImporteTramo1").value, "Currency")
                    End If
                    If .Fields("Permite2Conductores").value = False Or IsNull(.Fields("ConductorImporteTramo2").value) Then
                        ListItem.SubItems(4) = " "
                    Else
                        ListItem.SubItems(4) = Format(.Fields("ConductorImporteTramo2").value, "Currency")
                    End If
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
    
    If frmMDI.ActiveForm.Name = Me.Name And GetForegroundWindow() = frmMDI.hwnd And frmMDI.WindowState <> vbMinimized Then
        lvwData.SetFocus
    End If
    
    Screen.MousePointer = vbDefault
    FillListView = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Forms.ConductorRuta.FillListView", "Error al obtener la Lista de Rutas del Conductor." & vbCr & vbCr & "IDPersona: " & IDPersona
End Function

Private Sub cboConductor_Click()
    ForceRefresh
End Sub

Private Sub cbrMain_HeightChanged(ByVal NewHeight As Single)
    ResizeControls NewHeight
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
            Case vbKeyN
                tlbMain_ButtonClick tlbMain.Buttons.Item("NEW")
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
    
    CSM_Forms.ResizeAndPosition frmMDI, Me
    pParametro.GetCoolBarSettings "ConductorRuta", cbrMain
    pParametro.GetListViewSettings "ConductorRuta", lvwData
    lvwData.ColumnHeaders(lvwData.SortKey + 1).Icon = lvwData.SortOrder + 1
    tlbPin.Buttons("PIN").value = pParametro.Usuario_LeerNumero("ConductorRuta_Pin", tlbPin.Buttons("PIN").value)
    If tlbPin.Buttons("PIN").value = tbrUnpressed Then
        tlbPin.Buttons("PIN").Image = 1
    Else
        tlbPin.Buttons("PIN").Image = 2
    End If

    If Not pParametro.Viaje_Permite_2_Conductores Then
        lvwData.ColumnHeaders.Remove ("ImporteTramo1")
        lvwData.ColumnHeaders.Remove ("ImporteTramo2")
    End If
    
    mLoading = False
    
    FillListView 0, ""
End Sub

Private Sub Form_Resize()
    ResizeControls cbrMain.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visible = False
    WindowState = vbNormal
    pParametro.SaveCoolBarSettings "ConductorRuta", cbrMain
    pParametro.SaveListViewSettings "ConductorRuta", lvwData
    pParametro.Usuario_GuardarNumero "ConductorRuta_Pin", tlbPin.Buttons("PIN").value
    Set frmConductorRuta = Nothing
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
    FillListView cboConductor.ItemData(cboConductor.ListIndex), ""
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
    Dim ConductorRuta As ConductorRuta
    
    Select Case Button.Key
        Case "NEW"
            If pCPermiso.GotPermission(PERMISO_CONDUCTOR_RUTA_ADD) Then
                Screen.MousePointer = vbHourglass
                
                Set ConductorRuta = New ConductorRuta
                ConductorRuta.IDPersona = cboConductor.ItemData(cboConductor.ListIndex)
                frmConductorRutaPropiedad.LoadDataAndShow Me, ConductorRuta
                Set ConductorRuta = Nothing
                
                Screen.MousePointer = vbDefault
            End If
        Case "PROPERTIES"
            If pCPermiso.GotPermission(PERMISO_CONDUCTOR_RUTA_MODIFY) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                Screen.MousePointer = vbHourglass
                
                Set ConductorRuta = New ConductorRuta
                ConductorRuta.IDPersona = Val(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                ConductorRuta.IDRuta = CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER)
                If ConductorRuta.Load() Then
                    frmConductorRutaPropiedad.LoadDataAndShow Me, ConductorRuta
                End If
                Set ConductorRuta = Nothing
                
                Screen.MousePointer = vbDefault
            End If
        Case "DELETE"
            If pCPermiso.GotPermission(PERMISO_CONDUCTOR_RUTA_DELETE) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                If MsgBox("¿Desea eliminar la Ruta seleccionada?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                    Set ConductorRuta = New ConductorRuta
                    ConductorRuta.IDPersona = Val(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                    ConductorRuta.IDRuta = CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER)
                    If ConductorRuta.Load() Then
                        Call ConductorRuta.Delete
                    End If
                    Set ConductorRuta = Nothing
                End If
                lvwData.SetFocus
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
                Forms(FormIndex).ConductorRutaSelected Val(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER)), CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER)
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
    Call CSM_Control_ComboBox.FillFromSQL(cboConductor, "(SELECT 0 AS IDPersona, '<Todos>' AS ApellidoNombre, 1 AS Orden) UNION (SELECT IDPersona, dbo.udf_GetEntidadApellidoYNombre(Apellido, Nombre) AS ApellidoNombre, 2 AS Orden FROM Persona WHERE Activo = 1 AND EntidadTipo = '" & ENTIDAD_TIPO_PERSONA_CONDUCTOR & "') ORDER BY Orden, ApellidoNombre", "IDPersona", "ApellidoNombre", "Conductores", cscpItemOrFirst, KeySave)
End Sub
