VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmViajeDetalleTransferencia 
   Caption         =   "Transferencia de Pasajeros entre Viajes"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10455
   Icon            =   "ViajeDetalleTransferencia.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5130
   ScaleWidth      =   10455
   Begin VB.CommandButton cmdBoth 
      Height          =   615
      Left            =   4920
      Picture         =   "ViajeDetalleTransferencia.frx":062A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1620
      Width           =   615
   End
   Begin VB.Timer tmrListViewSettingsUpdate 
      Interval        =   3000
      Left            =   4980
      Top             =   4560
   End
   Begin MSComctlLib.ImageList ilsData 
      Left            =   4920
      Top             =   3780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViajeDetalleTransferencia.frx":0EF4
            Key             =   "PASAJERO_CONFIRMADO"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViajeDetalleTransferencia.frx":148E
            Key             =   "PASAJERO_CONDICIONAL"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViajeDetalleTransferencia.frx":1A28
            Key             =   "PASAJERO_CANCELADO"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViajeDetalleTransferencia.frx":1FC2
            Key             =   "COMISION"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRightToLeft 
      Height          =   615
      Left            =   4920
      Picture         =   "ViajeDetalleTransferencia.frx":255C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3060
      Width           =   615
   End
   Begin VB.CommandButton cmdLeftToRight 
      Height          =   615
      Left            =   4920
      Picture         =   "ViajeDetalleTransferencia.frx":2E26
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2340
      Width           =   615
   End
   Begin VB.Frame fraViajeRight 
      Height          =   1815
      Left            =   5640
      TabIndex        =   15
      Top             =   60
      Width           =   4695
      Begin VB.CommandButton cmdSelectRight 
         Caption         =   "Seleccionar"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   180
         Width           =   4455
      End
      Begin VB.TextBox txtFechaDiaSemanaRight 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   660
         Width           =   1170
      End
      Begin VB.TextBox txtFechaRight 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   2220
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   660
         Width           =   1410
      End
      Begin VB.TextBox txtHoraRight 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1170
      End
      Begin VB.TextBox txtRutaRight 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1380
         Width           =   3570
      End
      Begin VB.Label lblFechaRight 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   210
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblHoraRight 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
         Height          =   210
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label lblRutaRight 
         AutoSize        =   -1  'True
         Caption         =   "Ruta:"
         Height          =   210
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   375
      End
   End
   Begin VB.Frame fraViajeLeft 
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   60
      Width           =   4695
      Begin VB.CommandButton cmdSelectLeft 
         Caption         =   "Seleccionar"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   180
         Width           =   4455
      End
      Begin VB.TextBox txtRutaLeft 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1380
         Width           =   3570
      End
      Begin VB.TextBox txtHoraLeft 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1170
      End
      Begin VB.TextBox txtFechaLeft 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   2220
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   660
         Width           =   1410
      End
      Begin VB.TextBox txtFechaDiaSemanaLeft 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   660
         Width           =   1170
      End
      Begin VB.Label lblRutaLeft 
         AutoSize        =   -1  'True
         Caption         =   "Ruta:"
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label lblHoraLeft 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label lblFechaLeft 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   495
      End
   End
   Begin MSComctlLib.ListView lvwDataLeft 
      Height          =   2955
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5212
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
      NumItems        =   8
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
         Key             =   "Estado"
         Text            =   "Estado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Asiento"
         Text            =   "Asiento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Origen"
         Text            =   "Origen"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "Destino"
         Text            =   "Destino"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Key             =   "Tipo"
         Text            =   "Tipo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "ListaPasajero"
         Text            =   "Lista de Pasajeros"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwDataRight 
      Height          =   2955
      Left            =   5640
      TabIndex        =   4
      Top             =   2040
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5212
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
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
         Key             =   "Estado"
         Text            =   "Estado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Asiento"
         Text            =   "Asiento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Origen"
         Text            =   "Origen"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "Destino"
         Text            =   "Destino"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Key             =   "Tipo"
         Text            =   "Tipo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "ListaPasajero"
         Text            =   "Lista de Pasajeros"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmViajeDetalleTransferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mIDViajeLeft As Long
Private mFechaHoraLeft As Date
Private mIDRutaLeft As String

Private mIDViajeRight As Long
Private mFechaHoraRight As Date
Private mIDRutaRight As String

Public FormWaitingForSelect As String

Public Function FillListViewLeft(ByVal FechaHora As Date, ByVal IDRuta As String, ByVal Indice As Long) As Boolean
    Dim ListItem As MSComctlLib.ListItem
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    Dim KeySave As String
    Dim ViajeDetalle As ViajeDetalle
    Dim EstadoKey As String
    Dim UltimoEstado As String
    
    If FechaHora <> mFechaHoraLeft Or IDRuta <> mIDRutaLeft Then
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    If Indice = 0 Then
        If Not lvwDataLeft.SelectedItem Is Nothing Then
            KeySave = lvwDataLeft.SelectedItem.Key
        End If
    Else
        KeySave = KEY_STRINGER & Indice
    End If
        
    lvwDataLeft.ListItems.Clear
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_ViajeDetalle_ListPasajero"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("FechaHora", adDate, adParamInput, , FechaHora)
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta", adChar, adParamInput, 20, IDRuta)
    Set recData = New ADODB.Recordset
    Set recData.ActiveConnection = pDatabase.Connection
    recData.Open cmdData, , adOpenForwardOnly, adLockReadOnly
    Set cmdData = Nothing
        
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
                End Select
                
                '//////////////////////////////////////////////////
                'ROWS SEPARATOR BY STATUS
                If pParametro.ViajeDetalle_SeparateRowsByStatus Then
                    If UltimoEstado <> .Fields("Estado").value Then
                        If UltimoEstado <> "" Then
                            Set ListItem = lvwDataLeft.ListItems.Add(, KEY_STRINGER & (.Fields("Indice").value * -1), "")
                        End If
                        UltimoEstado = .Fields("Estado").value
                    End If
                End If
                
                'Pasajero
                Set ListItem = lvwDataLeft.ListItems.Add(, KEY_STRINGER & .Fields("Indice").value, Val(.Fields("Orden").value & ""), , "PASAJERO_" & EstadoKey)
                ListItem.SubItems(1) = .Fields("Persona").value
                ViajeDetalle.Estado = .Fields("Estado").value & ""
                ListItem.SubItems(2) = ViajeDetalle.Estado_ToString
                ListItem.SubItems(3) = .Fields("Asiento").value & ""
                ListItem.SubItems(4) = IIf(IsNull(.Fields("Sube").value), .Fields("Origen").value, .Fields("Sube").value)
                ListItem.SubItems(5) = IIf(IsNull(.Fields("Baja").value), .Fields("Destino").value, .Fields("Baja").value)
                ViajeDetalle.ReservaTipo = .Fields("ReservaTipo").value
                ListItem.SubItems(6) = ViajeDetalle.ReservaTipo_ToString
                ListItem.SubItems(7) = IIf(.Fields("ListaPasajero").value, "Sí", "")
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set recData = Nothing
    
    Set ViajeDetalle = Nothing
    
    On Error Resume Next
    Set lvwDataLeft.SelectedItem = lvwDataLeft.ListItems(KeySave)
    
    If frmMDI.ActiveForm.Name = Me.Name And GetForegroundWindow() = frmMDI.hwnd And frmMDI.WindowState <> vbMinimized Then
        lvwDataLeft.SetFocus
    End If
    
    FillListViewLeft = True
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Forms.ViajeDetalleTransferencia.FillListViewLeft", "Error al obtener el Detalle del Viaje."
End Function

Public Function FillListViewRight(ByVal FechaHora As Date, ByVal IDRuta As String, ByVal Indice As Long) As Boolean
    Dim ListItem As MSComctlLib.ListItem
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    Dim KeySave As String
    Dim ViajeDetalle As ViajeDetalle
    Dim EstadoKey As String
    Dim UltimoEstado As String
    
    If FechaHora <> mFechaHoraRight Or IDRuta <> mIDRutaRight Then
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    If Indice = 0 Then
        If Not lvwDataRight.SelectedItem Is Nothing Then
            KeySave = lvwDataRight.SelectedItem.Key
        End If
    Else
        KeySave = KEY_STRINGER & Indice
    End If
    
    lvwDataRight.ListItems.Clear
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_ViajeDetalle_ListPasajero"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("FechaHora", adDate, adParamInput, , FechaHora)
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta", adChar, adParamInput, 20, IDRuta)
    Set recData = New ADODB.Recordset
    Set recData.ActiveConnection = pDatabase.Connection
    recData.Open cmdData, , adOpenForwardOnly, adLockReadOnly
    Set cmdData = Nothing
    
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
                End Select
                
                '//////////////////////////////////////////////////
                'ROWS SEPARATOR BY STATUS
                If pParametro.ViajeDetalle_SeparateRowsByStatus Then
                    If UltimoEstado <> .Fields("Estado").value Then
                        If UltimoEstado <> "" Then
                            Set ListItem = lvwDataRight.ListItems.Add(, KEY_STRINGER & (.Fields("Indice").value * -1), "")
                        End If
                        UltimoEstado = .Fields("Estado").value
                    End If
                End If
                
                'Pasajero
                Set ListItem = lvwDataRight.ListItems.Add(, KEY_STRINGER & .Fields("Indice").value, Val(.Fields("Orden").value & ""), , "PASAJERO_" & EstadoKey)
                ListItem.SubItems(1) = .Fields("Persona").value
                ViajeDetalle.Estado = .Fields("Estado").value & ""
                ListItem.SubItems(2) = ViajeDetalle.Estado_ToString
                ListItem.SubItems(3) = .Fields("Asiento").value & ""
                ListItem.SubItems(4) = IIf(IsNull(.Fields("Sube").value), .Fields("Origen").value, .Fields("Sube").value)
                ListItem.SubItems(5) = IIf(IsNull(.Fields("Baja").value), .Fields("Destino").value, .Fields("Baja").value)
                ViajeDetalle.ReservaTipo = .Fields("ReservaTipo").value
                ListItem.SubItems(6) = ViajeDetalle.ReservaTipo_ToString
                ListItem.SubItems(7) = IIf(.Fields("ListaPasajero").value, "Sí", "")
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set recData = Nothing
    
    Set ViajeDetalle = Nothing
    
    On Error Resume Next
    Set lvwDataRight.SelectedItem = lvwDataRight.ListItems(KeySave)
    
    If frmMDI.ActiveForm.Name = Me.Name And GetForegroundWindow() = frmMDI.hwnd And frmMDI.WindowState <> vbMinimized Then
        lvwDataRight.SetFocus
    End If
    
    FillListViewRight = True
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Forms.ViajeDetalleTransferencia.FillListViewRight", "Error al obtener el Detalle del Viaje."
End Function

Private Sub cmdBoth_Click()
    Dim ViajeDetalleLeft As ViajeDetalle
    Dim ViajeDetalleRight As ViajeDetalle
    Dim RutaDetalle As RutaDetalle
    
    If mIDRutaLeft = "" Then
        MsgBox "Debe seleccionar el Viaje de la Izquierda.", vbInformation, App.Title
        cmdSelectLeft.SetFocus
        Exit Sub
    End If
    If mIDRutaRight = "" Then
        MsgBox "Debe seleccionar el Viaje de la Derecha.", vbInformation, App.Title
        cmdSelectRight.SetFocus
        Exit Sub
    End If
    
    If lvwDataLeft.SelectedItem Is Nothing Then
        MsgBox "No hay ninguna Reserva seleccionada en el Viaje de la Izquierda.", vbInformation, App.Title
        lvwDataLeft.SetFocus
        Exit Sub
    End If
    If Val(Mid(lvwDataLeft.SelectedItem.Key, 2)) < 0 Then
        MsgBox "No hay ninguna Reserva seleccionada en el Viaje de la Izquierda.", vbInformation, App.Title
        lvwDataLeft.SetFocus
        Exit Sub
    End If
    If lvwDataRight.SelectedItem Is Nothing Then
        MsgBox "No hay ninguna Reserva seleccionada en el Viaje de la Derecha.", vbInformation, App.Title
        lvwDataRight.SetFocus
        Exit Sub
    End If
    If Val(Mid(lvwDataRight.SelectedItem.Key, 2)) < 0 Then
        MsgBox "No hay ninguna Reserva seleccionada en el Viaje de la Derecha.", vbInformation, App.Title
        lvwDataRight.SetFocus
        Exit Sub
    End If
    
    If MsgBox("¿Desea intercambiar los Pasajeros seleccionados?", vbExclamation + vbYesNo, App.Title) = vbYes Then
    
        'PRIMERO CARGO LOS OBJETOS PARA VER SI NO HAY ALGUNA RESERVA ASISTIDA
        Set ViajeDetalleLeft = New ViajeDetalle
        With ViajeDetalleLeft
            .RefreshListSkip = True
            .FechaHora = mFechaHoraLeft
            .IDRuta = mIDRutaLeft
            .Indice = Val(Mid(lvwDataLeft.SelectedItem.Key, Len(KEY_STRINGER) + 1))
            If Not .Load() Then
                lvwDataLeft.SetFocus
                Exit Sub
            End If
        End With
        
        Set ViajeDetalleRight = New ViajeDetalle
        With ViajeDetalleRight
            .RefreshListSkip = True
            .FechaHora = mFechaHoraRight
            .IDRuta = mIDRutaRight
            .Indice = Val(Mid(lvwDataRight.SelectedItem.Key, Len(KEY_STRINGER) + 1))
            If Not .Load() Then
                lvwDataRight.SetFocus
                Exit Sub
            End If
        End With
        
        If ViajeDetalleLeft.Realizado <> VIAJE_DETALLE_REALIZADO_UNKNOWN Then
            MsgBox "La Reserva de la izquierda ya está asistida, por lo tanto, no se puede transferir.", vbExclamation, App.Title
        ElseIf ViajeDetalleRight.Realizado <> VIAJE_DETALLE_REALIZADO_UNKNOWN Then
            MsgBox "La Reserva de la derecha ya está asistida, por lo tanto, no se puede transferir.", vbExclamation, App.Title
        Else
            If mIDRutaLeft <> mIDRutaRight Then
                'VERIFICO SI EXISTE LA SUBIDA
                'DEL VIAJE DE LA IZQUIERDA EN LA RUTA DE LA DERECHA
                Set RutaDetalle = New RutaDetalle
                With RutaDetalle
                    .NoMatchRaiseError = False
                    .IDRuta = ViajeDetalleRight.IDRuta
                    .IDLugar = ViajeDetalleLeft.IDOrigen
                    If Not .Load() Then
                        Set RutaDetalle = Nothing
                        MsgBox "No se Transfirieron las Reservas porque no se pudieron comprobar las paradas.", vbCritical, App.Title
                        Exit Sub
                    End If
                    If .NoMatch Then
                        Set RutaDetalle = Nothing
                        MsgBox "No se pueden Transferir las Reservas porque el Lugar adonde Sube el Pasajero de la Izquierda no existe en la Ruta de la Derecha.", vbExclamation, App.Title
                        Exit Sub
                    End If
                End With
                Set RutaDetalle = Nothing
            
                'VERIFICO SI EXISTE LA BAJADA
                'DEL VIAJE DE LA IZQUIERDA EN LA RUTA DE LA DERECHA
                Set RutaDetalle = New RutaDetalle
                With RutaDetalle
                    .NoMatchRaiseError = False
                    .IDRuta = ViajeDetalleRight.IDRuta
                    .IDLugar = ViajeDetalleLeft.IDDestino
                    If Not .Load() Then
                        Set RutaDetalle = Nothing
                        MsgBox "No se Transfirieron las Reservas porque no se pudieron comprobar las paradas.", vbCritical, App.Title
                        Exit Sub
                    End If
                    If .NoMatch Then
                        Set RutaDetalle = Nothing
                        MsgBox "No se pueden Transferir las Reservas porque el Lugar adonde Baja el Pasajero de la Izquierda no existe en la Ruta de la Derecha.", vbExclamation, App.Title
                        Exit Sub
                    End If
                End With
                Set RutaDetalle = Nothing
            
                'VERIFICO SI EXISTE LA SUBIDA
                'DEL VIAJE DE LA DERECHA EN LA RUTA DE LA IZQUIERDA
                Set RutaDetalle = New RutaDetalle
                With RutaDetalle
                    .NoMatchRaiseError = False
                    .IDRuta = ViajeDetalleLeft.IDRuta
                    .IDLugar = ViajeDetalleRight.IDOrigen
                    If Not .Load() Then
                        Set RutaDetalle = Nothing
                        MsgBox "No se Transfirieron las Reservas porque no se pudieron comprobar las paradas.", vbCritical, App.Title
                        Exit Sub
                    End If
                    If .NoMatch Then
                        Set RutaDetalle = Nothing
                        MsgBox "No se pueden Transferir las Reservas porque el Lugar adonde Sube el Pasajero de la Derecha no existe en la Ruta de la Izquierda.", vbExclamation, App.Title
                        Exit Sub
                    End If
                End With
                Set RutaDetalle = Nothing
            
                'VERIFICO SI EXISTE LA BAJADA
                'DEL VIAJE DE LA DERECHA EN LA RUTA DE LA IZQUIERDA
                Set RutaDetalle = New RutaDetalle
                With RutaDetalle
                    .NoMatchRaiseError = False
                    .IDRuta = ViajeDetalleLeft.IDRuta
                    .IDLugar = ViajeDetalleRight.IDDestino
                    If Not .Load() Then
                        Set RutaDetalle = Nothing
                        MsgBox "No se Transfirieron las Reservas porque no se pudieron comprobar las paradas.", vbCritical, App.Title
                        Exit Sub
                    End If
                    If .NoMatch Then
                        Set RutaDetalle = Nothing
                        MsgBox "No se pueden Transferir las Reservas porque el Lugar adonde Baja el Pasajero de la Derecha no existe en la Ruta de la Izquierda.", vbExclamation, App.Title
                        Exit Sub
                    End If
                End With
                Set RutaDetalle = Nothing
            End If
            
            With ViajeDetalleLeft
                LogAccionAdd ENTIDAD_TIPO_VIAJE_DETALLE, "Transferencia Reserva: " & .Persona.ApellidoNombre & " // Desde: " & .FechaHora_Formatted & " - " & .IDRuta & " - " & " // Hasta: " & Format(mFechaHoraRight, "Short Date") & " " & Format(mFechaHoraRight, "Short Time") & " - " & mIDRutaRight
                .FechaHora = mFechaHoraRight
                .IDRuta = mIDRutaRight
                .Update
            End With
            
            With ViajeDetalleRight
                LogAccionAdd ENTIDAD_TIPO_VIAJE_DETALLE, "Transferencia Reserva: " & .Persona.ApellidoNombre & " // Desde: " & .FechaHora_Formatted & " - " & .IDRuta & " - " & " // Hasta: " & Format(mFechaHoraLeft, "Short Date") & " " & Format(mFechaHoraLeft, "Short Time") & " - " & mIDRutaLeft
                .FechaHora = mFechaHoraLeft
                .IDRuta = mIDRutaLeft
                .Update
            End With
        
            RefreshList_RefreshCuentaCorriente 0
            RefreshList_RefreshViajeDetalle mFechaHoraRight, mIDRutaRight, ViajeDetalleLeft.Indice
            RefreshList_RefreshViajeDetalle mFechaHoraLeft, mIDRutaLeft, ViajeDetalleRight.Indice
        End If
        
        Set ViajeDetalleLeft = Nothing
        Set ViajeDetalleRight = Nothing
    End If
End Sub

Private Sub cmdLeftToRight_Click()
    Dim ViajeDetalle As ViajeDetalle
    Dim RutaDetalle As RutaDetalle
    
    If mIDRutaLeft = "" Then
        MsgBox "Debe seleccionar el Viaje de Origen.", vbInformation, App.Title
        cmdSelectLeft.SetFocus
        Exit Sub
    End If
    If mIDRutaRight = "" Then
        MsgBox "Debe seleccionar el Viaje de Destino.", vbInformation, App.Title
        cmdSelectRight.SetFocus
        Exit Sub
    End If
    
    If lvwDataLeft.SelectedItem Is Nothing Then
        MsgBox "No hay ninguna Reserva seleccionada para Transferir.", vbInformation, App.Title
        lvwDataLeft.SetFocus
        Exit Sub
    End If
    If Val(Mid(lvwDataLeft.SelectedItem.Key, 2)) < 0 Then
        MsgBox "No hay ninguna Reserva seleccionada para Transferir.", vbInformation, App.Title
        lvwDataLeft.SetFocus
        Exit Sub
    End If
    
    If MsgBox("¿Desea Transferir la Reserva seleccionada en el Viaje de la Izquierda al Viaje de la Derecha?", vbExclamation + vbYesNo, App.Title) = vbYes Then
        Set ViajeDetalle = New ViajeDetalle
        ViajeDetalle.RefreshListSkip = True
        ViajeDetalle.FechaHora = mFechaHoraLeft
        ViajeDetalle.IDRuta = mIDRutaLeft
        ViajeDetalle.Indice = Val(Mid(lvwDataLeft.SelectedItem.Key, Len(KEY_STRINGER) + 1))
        If ViajeDetalle.Load() Then
            If ViajeDetalle.Realizado <> VIAJE_DETALLE_REALIZADO_UNKNOWN Then
                MsgBox "Esta Reserva ya está asistida, por lo tanto, no se puede transferir.", vbExclamation, App.Title
            Else
                If mIDRutaLeft <> mIDRutaRight Then
                    'VERIFICO SI EXISTE LA SUBIDA
                    'DEL VIAJE DE LA IZQUIERDA EN LA RUTA DE LA DERECHA
                    Set RutaDetalle = New RutaDetalle
                    With RutaDetalle
                        .NoMatchRaiseError = False
                        .IDRuta = mIDRutaRight
                        .IDLugar = ViajeDetalle.IDOrigen
                        If Not .Load() Then
                            Set RutaDetalle = Nothing
                            MsgBox "No se Transfirió la Reserva porque no se pudieron comprobar las paradas.", vbCritical, App.Title
                            Exit Sub
                        End If
                        If .NoMatch Then
                            Set RutaDetalle = Nothing
                            MsgBox "No se puede Transferir la Reserva porque el Lugar adonde Sube el Pasajero de la Izquierda no existe en la Ruta de la Derecha.", vbExclamation, App.Title
                            Exit Sub
                        End If
                    End With
                    Set RutaDetalle = Nothing
                
                    'VERIFICO SI EXISTE LA BAJADA
                    'DEL VIAJE DE LA IZQUIERDA EN LA RUTA DE LA DERECHA
                    Set RutaDetalle = New RutaDetalle
                    With RutaDetalle
                        .NoMatchRaiseError = False
                        .IDRuta = mIDRutaRight
                        .IDLugar = ViajeDetalle.IDDestino
                        If Not .Load() Then
                            Set RutaDetalle = Nothing
                            MsgBox "No se Transfirió la Reserva porque no se pudieron comprobar las paradas.", vbCritical, App.Title
                            Exit Sub
                        End If
                        If .NoMatch Then
                            Set RutaDetalle = Nothing
                            MsgBox "No se puede Transferir la Reserva porque el Lugar adonde Baja el Pasajero de la Izquierda no existe en la Ruta de la Derecha.", vbExclamation, App.Title
                            Exit Sub
                        End If
                    End With
                    Set RutaDetalle = Nothing
                End If
            
                With ViajeDetalle
                    LogAccionAdd ENTIDAD_TIPO_VIAJE_DETALLE, "Transferencia Reserva: " & .Persona.ApellidoNombre & " // Desde: " & .FechaHora_Formatted & " - " & .IDRuta & " - " & " // Hasta: " & Format(mFechaHoraRight, "Short Date") & " " & Format(mFechaHoraRight, "Short Time") & " - " & mIDRutaRight
                    .FechaHora = mFechaHoraRight
                    .IDRuta = mIDRutaRight
                    .Update
                End With
                
                RefreshList_RefreshCuentaCorriente 0
                RefreshList_RefreshViajeDetalle mFechaHoraLeft, mIDRutaLeft, 0
                RefreshList_RefreshViajeDetalle mFechaHoraRight, mIDRutaRight, ViajeDetalle.Indice
            End If
            lvwDataRight.SetFocus
        End If
        Set ViajeDetalle = Nothing
    End If
End Sub

Private Sub cmdRightToLeft_Click()
    Dim ViajeDetalle As ViajeDetalle
    Dim RutaDetalle As RutaDetalle

    If mIDRutaRight = "" Then
        MsgBox "Debe seleccionar el Viaje de Origen.", vbInformation, App.Title
        cmdSelectRight.SetFocus
        Exit Sub
    End If
    If mIDRutaLeft = "" Then
        MsgBox "Debe seleccionar el Viaje de Destino.", vbInformation, App.Title
        cmdSelectLeft.SetFocus
        Exit Sub
    End If
    
    If lvwDataRight.SelectedItem Is Nothing Then
        MsgBox "No hay ninguna Reserva seleccionada para Transferir.", vbInformation, App.Title
        lvwDataRight.SetFocus
        Exit Sub
    End If
    If Val(Mid(lvwDataRight.SelectedItem.Key, 2)) < 0 Then
        MsgBox "No hay ninguna Reserva seleccionada para Transferir.", vbInformation, App.Title
        lvwDataRight.SetFocus
        Exit Sub
    End If
    
    If MsgBox("¿Desea Transferir la Reserva seleccionada en el Viaje de la Derecha al Viaje de la Izquierda?", vbExclamation + vbYesNo, App.Title) = vbYes Then
        Set ViajeDetalle = New ViajeDetalle
        ViajeDetalle.RefreshListSkip = True
        ViajeDetalle.FechaHora = mFechaHoraRight
        ViajeDetalle.IDRuta = mIDRutaRight
        ViajeDetalle.Indice = Val(Mid(lvwDataRight.SelectedItem.Key, Len(KEY_STRINGER) + 1))
        If ViajeDetalle.Load() Then
            If ViajeDetalle.Realizado <> VIAJE_DETALLE_REALIZADO_UNKNOWN Then
                MsgBox "Esta Reserva ya está asistida, por lo tanto, no se puede transferir.", vbExclamation, App.Title
            Else
                If mIDRutaLeft <> mIDRutaRight Then
                    'VERIFICO SI EXISTE LA SUBIDA
                    'DEL VIAJE DE LA DERECHA EN LA RUTA DE LA IZQUIERDA
                    Set RutaDetalle = New RutaDetalle
                    With RutaDetalle
                        .NoMatchRaiseError = False
                        .IDRuta = mIDRutaLeft
                        .IDLugar = ViajeDetalle.IDOrigen
                        If Not .Load() Then
                            Set RutaDetalle = Nothing
                            MsgBox "No se Transfirió la Reserva porque no se pudieron comprobar las paradas.", vbCritical, App.Title
                            Exit Sub
                        End If
                        If .NoMatch Then
                            Set RutaDetalle = Nothing
                            MsgBox "No se puede Transferir la Reserva porque el Lugar adonde Sube el Pasajero de la Derecha no existe en la Ruta de la Izquierda.", vbExclamation, App.Title
                            Exit Sub
                        End If
                    End With
                    Set RutaDetalle = Nothing
                
                    'VERIFICO SI EXISTE LA BAJADA
                    'DEL VIAJE DE LA DERECHA EN LA RUTA DE LA IZQUIERDA
                    Set RutaDetalle = New RutaDetalle
                    With RutaDetalle
                        .NoMatchRaiseError = False
                        .IDRuta = mIDRutaLeft
                        .IDLugar = ViajeDetalle.IDDestino
                        If Not .Load() Then
                            Set RutaDetalle = Nothing
                            MsgBox "No se Transfirió la Reserva porque no se pudieron comprobar las paradas.", vbCritical, App.Title
                            Exit Sub
                        End If
                        If .NoMatch Then
                            Set RutaDetalle = Nothing
                            MsgBox "No se puede Transferir la Reserva porque el Lugar adonde Baja el Pasajero de la Derecha no existe en la Ruta de la Izquierda.", vbExclamation, App.Title
                            Exit Sub
                        End If
                    End With
                    Set RutaDetalle = Nothing
                End If
                
                With ViajeDetalle
                    LogAccionAdd ENTIDAD_TIPO_VIAJE_DETALLE, "Transferencia Reserva: " & .Persona.ApellidoNombre & " // Desde: " & .FechaHora_Formatted & " - " & .IDRuta & " - " & " // Hasta: " & Format(mFechaHoraLeft, "Short Date") & " " & Format(mFechaHoraLeft, "Short Time") & " - " & mIDRutaLeft
                    .FechaHora = mFechaHoraLeft
                    .IDRuta = mIDRutaLeft
                    .Update
                End With
                
                RefreshList_RefreshCuentaCorriente 0
                RefreshList_RefreshViajeDetalle mFechaHoraLeft, mIDRutaLeft, ViajeDetalle.Indice
                RefreshList_RefreshViajeDetalle mFechaHoraRight, mIDRutaRight, 0
            End If
            lvwDataRight.SetFocus
        End If
        Set ViajeDetalle = Nothing
    End If
End Sub

Private Sub cmdSelectLeft_Click()
    Screen.MousePointer = vbHourglass
    frmViaje.Show
    If frmViaje.WindowState = vbMinimized Then
        frmViaje.WindowState = vbNormal
    End If
    frmViaje.FormWaitingForSelect = Me.Name
    frmViaje.SelectTag = "LEFT"
    frmViaje.CSelectEstadosFilter.Add VIAJE_ESTADO_ACTIVO
    frmViaje.CSelectEstadosFilter.Add VIAJE_ESTADO_EN_PROGRESO
    frmViaje.SetFocus
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSelectRight_Click()
    Screen.MousePointer = vbHourglass
    frmViaje.Show
    If frmViaje.WindowState = vbMinimized Then
        frmViaje.WindowState = vbNormal
    End If
    frmViaje.FormWaitingForSelect = Me.Name
    frmViaje.SelectTag = "RIGHT"
    frmViaje.CSelectEstadosFilter.Add VIAJE_ESTADO_ACTIVO
    frmViaje.CSelectEstadosFilter.Add VIAJE_ESTADO_EN_PROGRESO
    frmViaje.SetFocus
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    lvwDataLeft.GridLines = pParametro.ListView_GridLines
    lvwDataRight.GridLines = pParametro.ListView_GridLines

    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL LISTVIEW
    Set lvwDataLeft.SmallIcons = ilsData
    Set lvwDataRight.SmallIcons = ilsData
    '//////////////////////////////////////////////////////////
    
    tmrListViewSettingsUpdate.Enabled = pIsCompiled
    
    CSM_Forms.ResizeAndPosition frmMDI, Me
    pParametro.GetListViewSettings "ViajeDetalleTransferencia", lvwDataLeft
    Call pParametro.CopyListViewSettings(lvwDataLeft, lvwDataRight)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visible = False
    WindowState = vbNormal
    pParametro.SaveListViewSettings "ViajeDetalleTransferencia", lvwDataLeft
    Set frmViajeDetalleTransferencia = Nothing
End Sub

Private Sub Form_Resize()
    Const CONTROL_SPACE = 60
    
    On Error Resume Next
        
    fraViajeLeft.Top = CONTROL_SPACE
    fraViajeLeft.Left = CONTROL_SPACE
    
    lvwDataLeft.Top = fraViajeLeft.Top + fraViajeLeft.Height + CONTROL_SPACE
    lvwDataLeft.Left = fraViajeLeft.Left
    lvwDataLeft.Height = ScaleHeight - lvwDataLeft.Top - CONTROL_SPACE
    lvwDataLeft.Width = (ScaleWidth - cmdLeftToRight.Width - (CONTROL_SPACE * 4)) / 2
    
    cmdBoth.Top = lvwDataLeft.Top + ((lvwDataLeft.Height - cmdBoth.Height - CONTROL_SPACE - cmdLeftToRight.Height - CONTROL_SPACE - cmdRightToLeft.Height) / 2)
    cmdBoth.Left = lvwDataLeft.Left + lvwDataLeft.Width + CONTROL_SPACE
    
    cmdLeftToRight.Top = cmdBoth.Top + cmdBoth.Height + CONTROL_SPACE
    cmdLeftToRight.Left = cmdBoth.Left

    cmdRightToLeft.Top = cmdLeftToRight.Top + cmdLeftToRight.Height + CONTROL_SPACE
    cmdRightToLeft.Left = cmdBoth.Left

    fraViajeRight.Top = CONTROL_SPACE
    fraViajeRight.Left = ScaleWidth - lvwDataLeft.Width - CONTROL_SPACE
    
    lvwDataRight.Top = lvwDataLeft.Top
    lvwDataRight.Left = fraViajeRight.Left
    lvwDataRight.Height = lvwDataLeft.Height
    lvwDataRight.Width = lvwDataLeft.Width
End Sub

Public Sub ViajeSelected(ByRef Viaje As Viaje, ByVal SelectTag As String)
    If SelectTag = "LEFT" Then
        'CHEQUEO LOS VIAJES
        If mIDRutaRight <> "" Then
            If Viaje.IDRuta <> mIDRutaRight Then
                If pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_TRANSFER_ALLOWDIFFERENTROUTE, False) Then
                    MsgBox "Está intentando Transferir Reservas entre Viajes con Rutas diferentes. Utilice esta opción sólo en casos realmente necesarios.", vbExclamation, App.Title
                Else
                    MsgBox "La Ruta del Viaje de Origen debe ser la misma que la del Viaje de Destino.", vbExclamation, App.Title
                    cmdSelectLeft.SetFocus
                    Exit Sub
                End If
            End If
            If Viaje.IDViaje = mIDViajeRight Then
                MsgBox "El Viaje de Origen seleccionado es el mismo que el Viaje de Destino.", vbInformation, App.Title
                cmdSelectLeft.SetFocus
                Exit Sub
            End If
            If Not pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_TRANSFER_ALLOWOUTOFRANGE, False) Then
                If pParametro.ViajeDetalle_Transferencia_RangoMaximoMinutos > 0 Then
                    If Abs(DateDiff("n", Viaje.FechaHora, mFechaHoraRight)) > pParametro.ViajeDetalle_Transferencia_RangoMaximoMinutos Then
                        MsgBox "La diferencia horaria entre el Viaje de Origen y el Viaje de Destino debe ser de " & pParametro.ViajeDetalle_Transferencia_RangoMaximoMinutos & " minutos como máximo.", vbInformation, App.Title
                        cmdSelectLeft.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        End If
        
        mIDViajeLeft = Viaje.IDViaje
        mFechaHoraLeft = Viaje.FechaHora
        mIDRutaLeft = Viaje.IDRuta
        
        txtFechaDiaSemanaLeft.Text = Viaje.FechaHora_WeekdayName
        txtFechaLeft.Text = Viaje.FechaHora_FormattedAsDate
        txtHoraLeft.Text = Viaje.FechaHora_FormattedAsTime
        txtRutaLeft.Text = Viaje.Ruta_DisplayName
        
        FillListViewLeft mFechaHoraLeft, mIDRutaLeft, 0
    Else
        'CHEQUEO LOS VIAJES
        If mIDRutaLeft <> "" Then
            If Viaje.IDRuta <> mIDRutaLeft Then
                If pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_TRANSFER_ALLOWDIFFERENTROUTE, False) Then
                    MsgBox "Está intentando Transferir Reservas entre Viajes con Rutas diferentes. Utilice esta opción sólo en casos realmente necesarios.", vbExclamation, App.Title
                Else
                    MsgBox "La Ruta del Viaje de Origen debe ser la misma que la del Viaje de Destino.", vbExclamation, App.Title
                    cmdSelectRight.SetFocus
                    Exit Sub
                End If
            End If
            If Viaje.IDViaje = mIDViajeLeft Then
                MsgBox "El Viaje de Origen seleccionado es el mismo que el Viaje de Destino.", vbInformation, App.Title
                cmdSelectRight.SetFocus
                Exit Sub
            End If
            If Not pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_TRANSFER_ALLOWOUTOFRANGE, False) Then
                If pParametro.ViajeDetalle_Transferencia_RangoMaximoMinutos > 0 Then
                    If Abs(DateDiff("n", Viaje.FechaHora, mFechaHoraLeft)) > pParametro.ViajeDetalle_Transferencia_RangoMaximoMinutos Then
                        MsgBox "La diferencia horaria entre el Viaje de Origen y el Viaje de Destino debe ser de " & pParametro.ViajeDetalle_Transferencia_RangoMaximoMinutos & " minutos como máximo.", vbInformation, App.Title
                        cmdSelectRight.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        End If
        
        mIDViajeRight = Viaje.IDViaje
        mFechaHoraRight = Viaje.FechaHora
        mIDRutaRight = Viaje.IDRuta
        
        txtFechaDiaSemanaRight.Text = Viaje.FechaHora_WeekdayName
        txtFechaRight.Text = Viaje.FechaHora_FormattedAsDate
        txtHoraRight.Text = Viaje.FechaHora_FormattedAsTime
        txtRutaRight.Text = Viaje.Ruta_DisplayName
        
        FillListViewRight mFechaHoraRight, mIDRutaRight, 0
    End If
    
    Set Viaje = Nothing
End Sub

Private Sub tmrListViewSettingsUpdate_Timer()
    Call pParametro.CopyListViewSettings(lvwDataLeft, lvwDataRight)
End Sub
