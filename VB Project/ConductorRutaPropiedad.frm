VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmConductorRutaPropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ConductorRutaPropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4440
   ScaleWidth      =   5205
   Begin VB.Frame fraConductorImporte 
      Caption         =   "Importe a acreditar al conductor:"
      Height          =   1695
      Left            =   180
      TabIndex        =   6
      Top             =   1980
      Width           =   4875
      Begin VB.TextBox txtConductorImporteTramoCompleto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1620
         MaxLength       =   20
         TabIndex        =   8
         Tag             =   "CURRENCY|EMPTY|ZERO|POSITIVE"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtConductorImporteTramo1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1620
         MaxLength       =   20
         TabIndex        =   10
         Tag             =   "CURRENCY|EMPTY|ZERO|POSITIVE"
         Top             =   780
         Width           =   1455
      End
      Begin VB.TextBox txtConductorImporteTramo2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1620
         MaxLength       =   20
         TabIndex        =   12
         Tag             =   "CURRENCY|EMPTY|ZERO|POSITIVE"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblConductorImporteTramoCompleto 
         AutoSize        =   -1  'True
         Caption         =   "Tramo completo:"
         Height          =   210
         Left            =   180
         TabIndex        =   7
         Top             =   420
         Width           =   1185
      End
      Begin VB.Label lblConductorImporteTramo1 
         AutoSize        =   -1  'True
         Caption         =   "Tramo 1:"
         Height          =   210
         Left            =   180
         TabIndex        =   9
         Top             =   840
         Width           =   630
      End
      Begin VB.Label lblConductorImporteTramo2 
         AutoSize        =   -1  'True
         Caption         =   "Tramo 2:"
         Height          =   210
         Left            =   180
         TabIndex        =   11
         Top             =   1260
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdAuditoria 
      Caption         =   "Auditoría"
      Height          =   615
      Left            =   4080
      Picture         =   "ConductorRutaPropiedad.frx":054A
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdConductor 
      Caption         =   "..."
      Height          =   315
      Left            =   4800
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Personas"
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   3900
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3840
      TabIndex        =   14
      Top             =   3900
      Width           =   1215
   End
   Begin VB.CommandButton cmdRuta 
      Caption         =   "..."
      Height          =   315
      Left            =   4800
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Rutas"
      Top             =   1440
      Width           =   255
   End
   Begin VB.Frame fraLine 
      Height          =   75
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   4935
   End
   Begin MSDataListLib.DataCombo datcboRuta 
      Height          =   330
      Left            =   1440
      TabIndex        =   4
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
   Begin MSDataListLib.DataCombo datcboConductor 
      Height          =   330
      Left            =   1440
      TabIndex        =   1
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
   Begin VB.Label lblConductor 
      AutoSize        =   -1  'True
      Caption         =   "Conductor:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   1020
      Width           =   795
   End
   Begin VB.Label lblRuta 
      AutoSize        =   -1  'True
      Caption         =   "&Ruta:"
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   1500
      Width           =   375
   End
   Begin VB.Image imgIcon2 
      Height          =   480
      Left            =   480
      Picture         =   "ConductorRutaPropiedad.frx":0B74
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "ConductorRutaPropiedad.frx":0E7E
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Datos de la Ruta del Conductor"
      Height          =   210
      Left            =   1140
      TabIndex        =   16
      Top             =   240
      Width           =   2235
   End
End
Attribute VB_Name = "frmConductorRutaPropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mConductorRuta As ConductorRuta
Private mNew As Boolean
Private mKeyDecimal As Boolean
Private mPermite2Conductores As Boolean

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef ConductorRuta As ConductorRuta)
    Set mConductorRuta = ConductorRuta
    mNew = (mConductorRuta.IDRuta = "")
    
    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    With mConductorRuta
        If Not CSM_Control_DataCombo.FillFromSQL(datcboConductor, "SELECT IDPersona, dbo.udf_GetEntidadApellidoYNombre(Apellido, Nombre) AS ApellidoNombre FROM Persona WHERE (Activo = 1 OR IDPersona = " & .IDPersona & ") AND EntidadTipo = '" & ENTIDAD_TIPO_PERSONA_CONDUCTOR & "' ORDER BY ApellidoNombre", "IDPersona", "ApellidoNombre", "Conductores", cscpItemOrNone, .IDPersona) Then
            Unload Me
            Exit Sub
        End If
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboRuta, "SELECT RTRIM(IDRuta) AS IDRuta FROM Ruta WHERE IDRuta <> '" & ReplaceQuote(pParametro.Ruta_ID_Otra) & "' AND (Activo = 1 OR IDRuta = '" & ReplaceQuote(.IDRuta) & "')" & IIf(pCPermiso.RutaWhere <> "", " AND " & Replace(pCPermiso.RutaWhere, "%TABLENAME%", "Ruta"), "") & " ORDER BY IDRuta", "IDRuta", "IDRuta", "Rutas", cscpItemOrFirst, .IDRuta) Then
            Unload Me
            Exit Sub
        End If
        
        txtConductorImporteTramoCompleto.Text = .ConductorImporteTramoCompleto_FormattedAsString
        If mPermite2Conductores Then
            txtConductorImporteTramo1.Text = .ConductorImporteTramo1_FormattedAsString
            txtConductorImporteTramo2.Text = .ConductorImporteTramo2_FormattedAsString
        End If
    End With
    
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Private Sub Form_Load()
    lblConductorImporteTramo1.Visible = pParametro.Viaje_Permite_2_Conductores
    txtConductorImporteTramo1.Visible = pParametro.Viaje_Permite_2_Conductores
    lblConductorImporteTramo2.Visible = pParametro.Viaje_Permite_2_Conductores
    txtConductorImporteTramo2.Visible = pParametro.Viaje_Permite_2_Conductores
End Sub

Private Sub cmdAuditoria_Click()
    frmAuditoriaGenerico.LoadDataAndShow mConductorRuta
End Sub

Private Sub cmdConductor_Click()
    If pCPermiso.GotPermission(PERMISO_PERSONA) Then
        Call frmPersona.FindAndShowItem(Val(datcboConductor.BoundText), UCase(Left(datcboConductor.Text, 1)), Me.Name, ENTIDAD_TIPO_PERSONA_CONDUCTOR, "")
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

Private Sub datcboRuta_Change()
    ShowControls
End Sub

Private Sub txtConductorImporteTramoCompleto_GotFocus()
    CSM_Control_TextBox.SelAllText txtConductorImporteTramoCompleto
End Sub

Private Sub txtConductorImporteTramoCompleto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDecimal Then
        mKeyDecimal = True
    End If
End Sub

Private Sub txtConductorImporteTramoCompleto_KeyPress(KeyAscii As Integer)
    Call CSM_Control_TextBox.CheckKeyPress(txtConductorImporteTramoCompleto, KeyAscii, mKeyDecimal)
End Sub

Private Sub txtConductorImporteTramoCompleto_LostFocus()
    Call CSM_Control_TextBox.FormatValue_ByTag(txtConductorImporteTramoCompleto)
End Sub

Private Sub txtConductorImporteTramo1_GotFocus()
    CSM_Control_TextBox.SelAllText txtConductorImporteTramo1
End Sub

Private Sub txtConductorImporteTramo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDecimal Then
        mKeyDecimal = True
    End If
End Sub

Private Sub txtConductorImporteTramo1_KeyPress(KeyAscii As Integer)
    Call CSM_Control_TextBox.CheckKeyPress(txtConductorImporteTramo1, KeyAscii, mKeyDecimal)
End Sub

Private Sub txtConductorImporteTramo1_LostFocus()
    Call CSM_Control_TextBox.FormatValue_ByTag(txtConductorImporteTramo1)
End Sub

Private Sub txtConductorImporteTramo2_GotFocus()
    CSM_Control_TextBox.SelAllText txtConductorImporteTramo2
End Sub

Private Sub txtConductorImporteTramo2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDecimal Then
        mKeyDecimal = True
    End If
End Sub

Private Sub txtConductorImporteTramo2_KeyPress(KeyAscii As Integer)
    Call CSM_Control_TextBox.CheckKeyPress(txtConductorImporteTramo2, KeyAscii, mKeyDecimal)
End Sub

Private Sub txtConductorImporteTramo2_LostFocus()
    Call CSM_Control_TextBox.FormatValue_ByTag(txtConductorImporteTramo2)
End Sub

Private Sub cmdOK_Click()
    If datcboConductor.BoundText = "" Then
        MsgBox "Debe seleccionar el Conductor.", vbInformation, App.Title
        datcboConductor.SetFocus
        Exit Sub
    End If
    If datcboRuta.BoundText = "" Then
        MsgBox "Debe seleccionar la Ruta.", vbInformation, App.Title
        datcboRuta.SetFocus
        Exit Sub
    End If
    
    Call txtConductorImporteTramoCompleto_LostFocus
    If mPermite2Conductores Then
        Call txtConductorImporteTramo1_LostFocus
        Call txtConductorImporteTramo2_LostFocus
    End If
        
    With mConductorRuta
        .IDPersona = Val(datcboConductor.BoundText)
        .IDRuta = datcboRuta.BoundText
        .ConductorImporteTramoCompleto_FormattedAsString = txtConductorImporteTramoCompleto.Text
        If mPermite2Conductores Then
            .ConductorImporteTramo1_FormattedAsString = txtConductorImporteTramo1.Text
            .ConductorImporteTramo2_FormattedAsString = txtConductorImporteTramo2.Text
        Else
            .ConductorImporteTramo1 = -1
            .ConductorImporteTramo2 = -1
        End If
        If Not .Update() Then
            Exit Sub
        End If
    End With
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mConductorRuta = Nothing
    Set frmConductorRutaPropiedad = Nothing
End Sub

Public Sub FillComboBoxConductor()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboConductor.BoundText)
    Set recData = datcboConductor.RowSource
    recData.Requery
    Set recData = Nothing
    datcboConductor.BoundText = KeySave
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

Private Sub ShowControls()
    Dim Ruta As Ruta
    
    If datcboRuta.BoundText <> "" Then
        Set Ruta = New Ruta
        Ruta.IDRuta = datcboRuta.BoundText
        If Ruta.Load() Then
            mPermite2Conductores = (pParametro.Viaje_Permite_2_Conductores And Ruta.Permite2Conductores)
            lblConductorImporteTramo1.Visible = mPermite2Conductores
            txtConductorImporteTramo1.Visible = mPermite2Conductores
            lblConductorImporteTramo2.Visible = mPermite2Conductores
            txtConductorImporteTramo2.Visible = mPermite2Conductores
        End If
        Set Ruta = Nothing
    End If
End Sub
