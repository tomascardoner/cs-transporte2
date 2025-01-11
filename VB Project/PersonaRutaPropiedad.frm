VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPersonaRutaPropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PersonaRutaPropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   5175
   Begin VB.CommandButton cmdAuditoria 
      Caption         =   "Auditoría"
      Height          =   615
      Left            =   4080
      Picture         =   "PersonaRutaPropiedad.frx":054A
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   60
      Width           =   975
   End
   Begin VB.TextBox txtSube 
      Height          =   315
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2400
      Width           =   3615
   End
   Begin VB.TextBox txtBaja 
      Height          =   315
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   9
      Top             =   3300
      Width           =   3615
   End
   Begin VB.TextBox txtPersona 
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   960
      Width           =   3555
   End
   Begin VB.CommandButton cmdListaPrecio 
      Caption         =   "..."
      Height          =   315
      Left            =   4800
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Listas de Precios"
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3840
      TabIndex        =   13
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdLugarDestino 
      Caption         =   "..."
      Height          =   315
      Left            =   4800
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Lugares"
      Top             =   2940
      Width           =   255
   End
   Begin VB.CommandButton cmdLugarOrigen 
      Caption         =   "..."
      Height          =   315
      Left            =   4800
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Lugares"
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton cmdRuta 
      Caption         =   "..."
      Height          =   315
      Left            =   4800
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Rutas"
      Top             =   1500
      Width           =   255
   End
   Begin VB.Frame fraLine 
      Height          =   75
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Width           =   4935
   End
   Begin MSDataListLib.DataCombo datcboRuta 
      Height          =   330
      Left            =   1440
      TabIndex        =   1
      Top             =   1500
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
      Left            =   1440
      TabIndex        =   3
      Top             =   2040
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
      Left            =   1440
      TabIndex        =   7
      Top             =   2940
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
      Left            =   1440
      TabIndex        =   11
      Top             =   3840
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
   Begin VB.Label lblSube 
      AutoSize        =   -1  'True
      Caption         =   "Sube:"
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   2460
      Width           =   420
   End
   Begin VB.Label lblBaja 
      AutoSize        =   -1  'True
      Caption         =   "Baja:"
      Height          =   210
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   360
   End
   Begin VB.Label lblPersona 
      AutoSize        =   -1  'True
      Caption         =   "Pasajero:"
      Height          =   210
      Left            =   120
      TabIndex        =   21
      Top             =   1020
      Width           =   675
   End
   Begin VB.Label lblListaPrecio 
      AutoSize        =   -1  'True
      Caption         =   "Lista de Precios:"
      Height          =   210
      Left            =   120
      TabIndex        =   10
      Top             =   3900
      Width           =   1200
   End
   Begin VB.Label lblDestino 
      AutoSize        =   -1  'True
      Caption         =   "&Destino:"
      Height          =   210
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   585
   End
   Begin VB.Label lblOrigen 
      AutoSize        =   -1  'True
      Caption         =   "&Origen:"
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   2100
      Width           =   525
   End
   Begin VB.Label lblRuta 
      AutoSize        =   -1  'True
      Caption         =   "&Ruta:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   375
   End
   Begin VB.Image imgIcon2 
      Height          =   480
      Left            =   480
      Picture         =   "PersonaRutaPropiedad.frx":0B74
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "PersonaRutaPropiedad.frx":0E7E
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Datos de la Ruta de la Persona"
      Height          =   210
      Left            =   1140
      TabIndex        =   19
      Top             =   240
      Width           =   2220
   End
End
Attribute VB_Name = "frmPersonaRutaPropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mPersonaRuta As PersonaRuta
Private mNew As Boolean

Private mIDDestino As Long
Private mRutaDetalleIndiceMinimo As Long
Private mRutaDetallaIndiceMaximo As Long

Private mListaPrecio_PrepagoVencimiento As String

Private Sub cmdAuditoria_Click()
    frmAuditoriaGenerico.LoadDataAndShow mPersonaRuta
End Sub

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef PersonaRuta As PersonaRuta)
    Set mPersonaRuta = PersonaRuta
    mNew = (mPersonaRuta.IDRuta = "")
    
    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    Dim SubeBaja As Boolean
    SubeBaja = pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_SUBEBAJA_MOSTRAR, False)
    lblSube.Visible = SubeBaja
    txtSube.Visible = SubeBaja
    lblBaja.Visible = SubeBaja
    txtBaja.Visible = SubeBaja
    
    txtPersona.Text = mPersonaRuta.Persona.ApellidoNombre
    
    With mPersonaRuta
        If Not CSM_Control_DataCombo.FillFromSQL(datcboRuta, "SELECT RTRIM(IDRuta) AS IDRuta FROM Ruta WHERE IDRuta <> '" & ReplaceQuote(pParametro.Ruta_ID_Otra) & "' AND (Activo = 1 OR IDRuta = '" & ReplaceQuote(.IDRuta) & "')" & IIf(pCPermiso.RutaWhere <> "", " AND " & Replace(pCPermiso.RutaWhere, "%TABLENAME%", "Ruta"), "") & " ORDER BY IDRuta", "IDRuta", "IDRuta", "Rutas") Then
            Unload Me
            Exit Sub
        End If
        datcboRuta.BoundText = .IDRuta
        
        datcboOrigen.BoundText = .IDOrigen
        txtSube.Text = .Sube
        datcboDestino.BoundText = .IDDestino
        txtBaja.Text = .Baja
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboListaPrecio, "SELECT IDListaPrecio, Nombre FROM ListaPrecio WHERE Activo = 1 OR IDListaPrecio = " & .IDListaPrecio & IIf(pCPermiso.ListaPrecioWhere <> "", " AND " & Replace(pCPermiso.ListaPrecioWhere, "%TABLENAME%", "ListaPrecio"), "") & " ORDER BY Nombre", "IDListaPrecio", "Nombre", "Listas de Precios") Then
            Unload Me
            Exit Sub
        End If
        datcboListaPrecio.BoundText = .IDListaPrecio
        If Val(datcboListaPrecio.BoundText) = 0 Then
            datcboListaPrecio.BoundText = pParametro.ListaPrecio_ID_Predeterminada
        End If
    End With
    
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
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

Public Sub FillComboBoxRuta()
    Dim KeySave As String
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboRuta.BoundText)
    Set recData = datcboRuta.RowSource
    recData.Requery
    Set recData = Nothing
    datcboRuta.BoundText = KeySave
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

Private Sub datcboRuta_Change()
    Dim IDOrigenSave As Long
    Dim Ruta As Ruta
    
    If datcboRuta.BoundText <> "" Then
        Set Ruta = New Ruta
        Ruta.IDRuta = datcboRuta.BoundText
        If Not Ruta.GetStatistics(0, mRutaDetalleIndiceMinimo, mRutaDetallaIndiceMaximo) Then
            Set Ruta = Nothing
            Exit Sub
        End If
        Set Ruta = Nothing
        
        IDOrigenSave = Val(datcboOrigen.BoundText)
        Call CSM_Control_DataCombo.FillFromSQL(datcboOrigen, "SELECT RutaDetalle.IDLugar, Lugar.Nombre FROM RutaDetalle INNER JOIN Lugar ON RutaDetalle.IDLugar = Lugar.IDLugar WHERE RutaDetalle.IDRuta = '" & ReplaceQuote(datcboRuta.Text) & "' AND RutaDetalle.Indice < " & mRutaDetallaIndiceMaximo & " AND RutaDetalle.IDLugar <> " & pParametro.Lugar_ID_Otro & " AND (Lugar.Activo = 1 OR Lugar.IDLugar = " & mPersonaRuta.IDOrigen & ") ORDER BY RutaDetalle.Indice, Lugar.Nombre", "IDLugar", "Nombre", "Orígenes", cscpItemOrFirst, IDOrigenSave)
        
        datcboOrigen_Change
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

Private Sub datcboOrigen_Change()
    Dim RutaDetalle As RutaDetalle
    Dim RutaDetalleIndice As Long
    
    Dim IDDestinoSave As Long
    
    If Val(datcboOrigen.BoundText) = 0 Then
        Exit Sub
    End If
        
    'Busco el Detalle de la Ruta para filtrar el ComboBox de Destino a partir del Origen
    Set RutaDetalle = New RutaDetalle
    RutaDetalle.IDRuta = datcboRuta.Text
    RutaDetalle.IDLugar = Val(datcboOrigen.BoundText)
    If RutaDetalle.Load() Then
        RutaDetalleIndice = RutaDetalle.Indice
    End If
    Set RutaDetalle = Nothing

    IDDestinoSave = Val(datcboDestino.BoundText)
    Call CSM_Control_DataCombo.FillFromSQL(datcboDestino, "SELECT RutaDetalle.IDLugar, Lugar.Nombre FROM RutaDetalle INNER JOIN Lugar ON RutaDetalle.IDLugar = Lugar.IDLugar WHERE RutaDetalle.IDRuta = '" & ReplaceQuote(datcboRuta.BoundText) & "' AND RutaDetalle.Indice > " & RutaDetalleIndice & " AND RutaDetalle.IDLugar <> " & pParametro.Lugar_ID_Otro & " AND (Lugar.Activo = 1 OR Lugar.IDLugar = " & mIDDestino & ") ORDER BY RutaDetalle.Indice, Lugar.Nombre", "IDLugar", "Nombre", "Destinos", cscpItemOrLast, IDDestinoSave)
End Sub

Private Sub cmdLugarOrigen_Click()
    If pCPermiso.GotPermission(PERMISO_LUGAR) Then
        Screen.MousePointer = vbHourglass
        frmLugar.Show
        On Error Resume Next
        Set frmLugar.lvwData.SelectedItem = frmLugar.lvwData.ListItems(KEY_STRINGER & Val(datcboOrigen.BoundText))
        frmLugar.lvwData.SelectedItem.EnsureVisible
        If frmLugar.WindowState = vbMinimized Then
            frmLugar.WindowState = vbNormal
        End If
        frmLugar.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub txtSube_GotFocus()
    CSM_Control_TextBox.SelAllText txtSube
End Sub

Private Sub cmdLugarDestino_Click()
    If pCPermiso.GotPermission(PERMISO_LUGAR) Then
        Screen.MousePointer = vbHourglass
        frmLugar.Show
        On Error Resume Next
        Set frmLugar.lvwData.SelectedItem = frmLugar.lvwData.ListItems(KEY_STRINGER & Val(datcboDestino.BoundText))
        frmLugar.lvwData.SelectedItem.EnsureVisible
        If frmLugar.WindowState = vbMinimized Then
            frmLugar.WindowState = vbNormal
        End If
        frmLugar.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub txtBaja_GotFocus()
    CSM_Control_TextBox.SelAllText txtBaja
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

Private Sub cmdOK_Click()
    If datcboRuta.BoundText = "" Then
        MsgBox "Debe seleccionar la Ruta.", vbInformation, App.Title
        datcboRuta.SetFocus
        Exit Sub
    End If
    If Val(datcboListaPrecio.BoundText) = 0 Then
        MsgBox "Debe seleccionar la Lista de Precios.", vbInformation, App.Title
        datcboListaPrecio.SetFocus
        Exit Sub
    End If
    
    With mPersonaRuta
        .IDRuta = datcboRuta.BoundText
        .IDOrigen = Val(datcboOrigen.BoundText)
        .Sube = txtSube.Text
        .IDDestino = Val(datcboDestino.BoundText)
        .Baja = txtBaja.Text
        .IDListaPrecio = Val(datcboListaPrecio.BoundText)
        
        If mNew Then
            If Not .AddNew() Then
                Exit Sub
            End If
        Else
            If Not .Update() Then
                Exit Sub
            End If
        End If
    End With
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mPersonaRuta = Nothing
    Set frmPersonaRutaPropiedad = Nothing
End Sub

