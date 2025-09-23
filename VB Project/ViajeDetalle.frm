VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmViajeDetalle 
   Caption         =   "Detalle del Viaje"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   Icon            =   "ViajeDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5295
   ScaleWidth      =   9360
   Begin MSComctlLib.ImageList ilsData 
      Left            =   3660
      Top             =   2640
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
            Picture         =   "ViajeDetalle.frx":058A
            Key             =   "PASAJERO_CONFIRMADO"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViajeDetalle.frx":0B24
            Key             =   "PASAJERO_CONDICIONAL"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViajeDetalle.frx":10BE
            Key             =   "PASAJERO_CANCELADO"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViajeDetalle.frx":1658
            Key             =   "COMISION_CONFIRMADO"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViajeDetalle.frx":1BF2
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
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   1111
      BandCount       =   4
      FixedOrder      =   -1  'True
      _CBWidth        =   9360
      _CBHeight       =   630
      _Version        =   "6.7.9816"
      Child1          =   "tlbMain"
      MinHeight1      =   570
      Width1          =   3000
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
         Left            =   7290
         ScaleHeight     =   330
         ScaleWidth      =   1980
         TabIndex        =   11
         Top             =   150
         Width           =   1980
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
         Left            =   5265
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
         Left            =   3195
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
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   1005
         ButtonWidth     =   2381
         ButtonHeight    =   1005
         ToolTips        =   0   'False
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
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
               Caption         =   "Imprimir"
               Key             =   "PRINT"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   13
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PRINT_PLANILLA_OBSERVACIONES"
                     Text            =   "Planilla Completa (con Observaciones)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PRINT_PLANILLA_DOCUMENTO"
                     Text            =   "Planilla Completa (con Documento)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PRINT_PLANILLA_PASAJERO_OBSERVACIONES"
                     Text            =   "Planilla de Pasajeros (con Observaciones)"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PRINT_PLANILLA_PASAJERO_DOCUMENTO"
                     Text            =   "Planilla de Pasajeros (con Documento)"
                  EndProperty
                  BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PRINT_PLANILLA_PASAJERO_DOMICILIO"
                     Text            =   "Planilla de Pasajeros (con Domicilio)"
                  EndProperty
                  BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PRINT_PLANILLA_COMISION"
                     Text            =   "Planilla de Comisiones"
                  EndProperty
                  BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PRINT_COMISION_REMITO"
                     Text            =   "Remito de Comisión"
                  EndProperty
                  BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PRINT_COMISION_LISTADO"
                     Text            =   "Listado de Comisiones"
                  EndProperty
                  BeginProperty ButtonMenu12 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu13 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PRINT_FACTURA"
                     Text            =   "Imprimir Factura"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Asistencia"
               Key             =   "ASISTENCIA"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "ASISTENCIA_SIMPLE"
                     Text            =   "Al Pasajero o Comisión Seleccionado"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "ASISTENCIA_MULTIPLE"
                     Text            =   "A Todos los Pasajeros"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      Top             =   4935
      Width           =   9360
      _ExtentX        =   16510
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
            Object.Width           =   12726
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
   Begin MSComctlLib.ListView lvwData 
      Height          =   3195
      Left            =   60
      TabIndex        =   0
      Top             =   1320
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5636
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
         Alignment       =   2
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
End
Attribute VB_Name = "frmViajeDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLoading As Boolean

Private mViaje As Viaje

Public FormWaitingForSelect As String

Public Sub LoadDataAndShow(ByRef Viaje As Viaje)
    Set mViaje = Viaje
    
    Load frmViajeDetalle
    
    If Not FillListView(mViaje.FechaHora, mViaje.IDRuta, 0) Then
        Unload Me
        Exit Sub
    End If

    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Public Sub ForceRefresh()
    FillListView mViaje.FechaHora, mViaje.IDRuta, 0
End Sub

Public Function FillListView(ByVal FechaHora As Date, ByVal IDRuta As String, ByVal Indice As Long) As Boolean
    Dim MousePointerSave As Integer
    
    Dim KeySave As String
    
    Dim EstadoKey As String
    Dim UltimoTipo As String
    Dim UltimoEstadoPersona As String
    Dim UltimoEstadoComision As String
    Dim ViajeDetalle As ViajeDetalle
    Dim Vehiculo As Vehiculo
    Dim VehiculoNombre As String
    Dim Conductor As Persona
    Dim ConductorNombre As String
    
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    Dim OrderBy As String
    Dim ListItem As MSComctlLib.ListItem
    
    If mLoading Or FechaHora <> mViaje.FechaHora Or IDRuta <> mViaje.IDRuta Then
        Exit Function
    End If
    
RESTART:
    
    MousePointerSave = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    mViaje.Load
    'ESTADO DEL VIAJE
    tlbMain.Buttons("NEW").Enabled = True
    tlbMain.Buttons("PROPERTIES").Enabled = True
    tlbMain.Buttons("DELETE").Enabled = True
    tlbMain.Buttons("SELECT").Enabled = True
    tlbMain.Buttons("PRINT").Enabled = True
    tlbMain.Buttons("ASISTENCIA").Enabled = True
    tlbMain.Buttons("CHANGE_STATUS").Enabled = True
    
    Select Case mViaje.Estado
        Case VIAJE_ESTADO_FINALIZADO
            If Not pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_FINALIZADO_MODIFY, False) Then
                tlbMain.Buttons("NEW").Enabled = False
                tlbMain.Buttons("DELETE").Enabled = False
                tlbMain.Buttons("CHANGE_STATUS").Enabled = False
            End If
        Case VIAJE_ESTADO_CANCELADO
            tlbMain.Buttons("NEW").Enabled = False
            tlbMain.Buttons("PROPERTIES").Enabled = False
            tlbMain.Buttons("DELETE").Enabled = False
            tlbMain.Buttons("SELECT").Enabled = False
            tlbMain.Buttons("PRINT").Enabled = False
            tlbMain.Buttons("ASISTENCIA").Enabled = False
            tlbMain.Buttons("CHANGE_STATUS").Enabled = False
    End Select
        
    If Indice = 0 Then
        If Not lvwData.SelectedItem Is Nothing Then
            KeySave = lvwData.SelectedItem.Key
        End If
    Else
        KeySave = KEY_STRINGER & Indice
    End If
    
    'ASIENTOS LIBRES
    If mViaje.IDVehiculo = 0 Then
        stbMain.Panels("INFO").Text = ""
        VehiculoNombre = ""
    Else
        Set Vehiculo = New Vehiculo
        Vehiculo.IDVehiculo = mViaje.IDVehiculo
        If Vehiculo.Load() Then
            stbMain.Panels("INFO").Text = "Asientos Libres: " & (Vehiculo.Asiento - mViaje.AsientoOcupado)
            VehiculoNombre = " | " & Vehiculo.Nombre
        Else
            stbMain.Panels("INFO").Text = ""
            VehiculoNombre = ""
        End If
        Set Vehiculo = Nothing
    End If
    
    'CONDUCTOR
    If mViaje.IDConductor = 0 Then
        ConductorNombre = ""
    Else
        Set Conductor = New Persona
        Conductor.IDPersona = mViaje.IDConductor
        If Conductor.Load() Then
            ConductorNombre = " | " & Conductor.ApellidoNombre
        Else
            ConductorNombre = ""
        End If
        Set Conductor = Nothing
    End If
    
    Caption = mViaje.FechaHora_WeekdayName & " " & mViaje.FechaHora_Formatted & " | " & mViaje.Ruta_DisplayName & VehiculoNombre & ConductorNombre
    
    lvwData.ListItems.Clear
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    If pParametro.ViajeDetalle_Paquete_Permite_Multiples_Pagos And mViaje.IDRuta = pParametro.Ruta_Paquete_ID Then
        cmdData.CommandText = "sp_ViajeDetalle_ListGrid_Paquete_MultiplesPagos" & IIf(chkMostrarSaldo.value = vbChecked, "_WithSaldo", "")
    Else
        cmdData.CommandText = "sp_ViajeDetalle_ListGrid" & IIf(chkMostrarSaldo.value = vbChecked, "_WithSaldo", "")
    End If
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("FechaHora", adDate, adParamInput, , mViaje.FechaHora)
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta", adChar, adParamInput, 20, mViaje.IDRuta)
    cmdData.Parameters.Append cmdData.CreateParameter("FilterEstado", adTinyInt, adParamInput, , cboFilterEstado.ListIndex)
    cmdData.Parameters.Append cmdData.CreateParameter("FilterRealizado", adTinyInt, adParamInput, , cboFilterRealizado.ListIndex)
    Set recData = New ADODB.Recordset
    Set recData.ActiveConnection = pDatabase.Connection
    recData.Open cmdData, , adOpenForwardOnly, adLockReadOnly
    Set cmdData = Nothing
    
    Select Case lvwData.SortKey
        Case 0  'ORDEN
            OrderBy = "OcupanteTipo DESC, Estado, Orden" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 1  'PERSONA
            OrderBy = "OcupanteTipo DESC, Persona" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Orden" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 2  'DOCUMENTO
            OrderBy = "OcupanteTipo DESC, DocumentoTipoNombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", DocumentoNumero" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Persona" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Orden" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 3  'IMPORTE
            OrderBy = "OcupanteTipo DESC, Importe" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Orden" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 4  'IMPORTE PAGADO
            OrderBy = "OcupanteTipo DESC, ImportePagado" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Orden" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 5  'DEBE
            OrderBy = "OcupanteTipo DESC, Debe" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Orden" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 6  'SALDO
            OrderBy = "OcupanteTipo DESC, Estado, Orden" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 7  'ESTADO
            OrderBy = "OcupanteTipo DESC, Estado" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Orden" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 8  'ASIENTO
            OrderBy = "OcupanteTipo DESC, Estado" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", AsientoIdentificacion" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 9  'REALIZADO
            OrderBy = "OcupanteTipo DESC, Estado" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Realizado" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 10  'ORIGEN
            OrderBy = "OcupanteTipo DESC, Origen" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Orden" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 11  'DESTINO
            OrderBy = "OcupanteTipo DESC, Destino" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Orden" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 12  'RESERVA TIPO
            OrderBy = "OcupanteTipo DESC, ReservaTipo" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Orden" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 13  'FACTURAR
            OrderBy = "OcupanteTipo DESC, Facturar" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Orden" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 14  'NOTAS
            OrderBy = "OcupanteTipo DESC, Notas" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Orden" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 15  'LISTA PASAJEROS
            OrderBy = "OcupanteTipo DESC, Estado" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", ListaPasajero" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
    End Select
        
    Set ViajeDetalle = New ViajeDetalle
    With recData
        .Sort = OrderBy
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
                            Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & (.Fields("Indice").value * -1), "")
                        End If
                        UltimoTipo = .Fields("OcupanteTipo").value
                    End If
                End If
                '//////////////////////////////////////////////////
                'ROWS SEPARATOR BY STATUS
                If .Fields("OcupanteTipo").value = OCUPANTE_TIPO_PASAJERO Then
                    If pParametro.ViajeDetalle_SeparateRowsByStatus Then
                        If UltimoEstadoPersona <> .Fields("Estado").value And (lvwData.SortKey = 0 Or lvwData.SortKey = 6) Then
                            If UltimoEstadoPersona <> "" Then
                                Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & (.Fields("Indice").value * -1), "")
                            End If
                            UltimoEstadoPersona = .Fields("Estado").value
                        End If
                    End If
                Else
                    If pParametro.ViajeDetalle_SeparateRowsByStatus Then
                        If UltimoEstadoComision <> .Fields("Estado").value And (lvwData.SortKey = 0 Or lvwData.SortKey = 6) Then
                            If UltimoEstadoComision <> "" Then
                                Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & (.Fields("Indice").value * -1), "")
                            End If
                            UltimoEstadoComision = .Fields("Estado").value
                        End If
                    End If
                End If
                
                Select Case .Fields("OcupanteTipo").value
                    Case OCUPANTE_TIPO_PASAJERO
                        'Pasajero
                        If EstadoKey = "" Then
                            Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & .Fields("Indice").value, .Fields("Orden").value)
                        Else
                            Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & .Fields("Indice").value, .Fields("Orden").value, , "PASAJERO_" & EstadoKey)
                            If .Fields("Estado").value = VIAJE_DETALLE_ESTADO_CONFIRMADO Then
                                ListItem.SubItems(9) = IIf(IsNull(.Fields("Realizado").value), "", IIf(.Fields("Realizado").value, "Sí", "No"))
                            End If
                        End If
                    Case OCUPANTE_TIPO_COMISION
                        'Comisión
                        Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & .Fields("Indice").value, Val(.Fields("Orden").value & ""), , "COMISION_" & EstadoKey)
                End Select
                ListItem.SubItems(1) = .Fields("Persona").value
                ListItem.SubItems(2) = IIf(IsNull(.Fields("DocumentoNumero").value), "", IIf(IsNull(.Fields("DocumentoTipoNombre").value), .Fields("DocumentoNumero").value, .Fields("DocumentoTipoNombre").value & ": " & .Fields("DocumentoNumero").value))
                ListItem.SubItems(3) = Format(.Fields("Importe").value, "Currency")
                ListItem.SubItems(4) = Format(.Fields("ImportePagado").value, "Currency")
                If pParametro.ViajeDetalle_Paquete_Permite_Multiples_Pagos And mViaje.IDRuta = pParametro.Ruta_Paquete_ID Then
                    ListItem.SubItems(5) = Format(.Fields("Importe").value - .Fields("ImportePagado").value, "Currency")
                Else
                    ListItem.SubItems(5) = Format(.Fields("Debe").value, "Currency")
                End If
                If chkMostrarSaldo.value = vbChecked Then
                    ListItem.SubItems(6) = IIf(IsNull(.Fields("SaldoActual").value), " ", Format(.Fields("SaldoActual").value, "Currency"))
                Else
                    ListItem.SubItems(6) = " "
                End If
                ViajeDetalle.Estado = .Fields("Estado").value & ""
                ListItem.SubItems(7) = ViajeDetalle.Estado_ToString
                 ListItem.SubItems(8) = .Fields("AsientoIdentificacion").value & ""
                ListItem.SubItems(10) = .Fields("Origen").value
                ListItem.SubItems(11) = .Fields("Destino").value
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
    Set lvwData.SelectedItem = lvwData.ListItems(KeySave)
    lvwData.SelectedItem.EnsureVisible
    
    If frmMDI.ActiveForm.Name = Me.Name And GetForegroundWindow() = frmMDI.hwnd And frmMDI.WindowState <> vbMinimized Then
        lvwData.SetFocus
    End If
    
    FillListView = True
    Screen.MousePointer = MousePointerSave
    Exit Function
    
ErrorHandler:
    If Err.Number = ERROR_TYPE_MISMATCH Or Err.Number = ERROR_ELEMENT_NOT_FOUND Then
        mViaje.Asiento_Asignar
        Resume RESTART
    Else
        ShowErrorMessage "Forms.ViajeDetalle.FillListView", "Error al obtener el Detalle del Viaje."
    End If
End Function

Private Sub cboFilterRealizado_Click()
    FillListView mViaje.FechaHora, mViaje.IDRuta, 0
End Sub

Private Sub cbrMain_HeightChanged(ByVal NewHeight As Single)
    ResizeControls NewHeight
End Sub

Private Sub cboFilterEstado_Click()
    FillListView mViaje.FechaHora, mViaje.IDRuta, 0
End Sub

Private Sub chkMostrarSaldo_Click()
    FillListView mViaje.FechaHora, mViaje.IDRuta, 0
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
    tlbMain.Buttons("PRINT").Image = "PRINT"
    tlbMain.Buttons("ASISTENCIA").Image = "ASISTENCIA"
    tlbMain.Buttons("CHANGE_STATUS").Image = "CHANGE_STATUS"
    
    cbrMain.Bands("Toolbar").MinWidth = CSM_Control_Toolbar.GetTotalWidth(tlbMain)
    '//////////////////////////////////////////////////////////
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL LISTVIEW
    Set lvwData.ColumnHeaderIcons = frmMDI.ilsFormSortColumn
    Set lvwData.SmallIcons = ilsData
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
    cboFilterRealizado.ListIndex = 0
    
    CSM_Forms.ResizeAndPosition frmMDI, Me
    pParametro.GetCoolBarSettings "ViajeDetalle", cbrMain
    pParametro.GetListViewSettings "ViajeDetalle", lvwData
    lvwData.ColumnHeaders(lvwData.SortKey + 1).Icon = lvwData.SortOrder + 1
    tlbPin.Buttons("PIN").value = pParametro.Usuario_LeerNumero("ViajeDetalle_Pin", tlbPin.Buttons("PIN").value)
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
    pParametro.SaveCoolBarSettings "ViajeDetalle", cbrMain
    pParametro.SaveListViewSettings "ViajeDetalle", lvwData
    pParametro.Usuario_GuardarNumero "ViajeDetalle_Pin", tlbPin.Buttons("PIN").value
    Set mViaje = Nothing
    Set frmViajeDetalle = Nothing
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
    FillListView mViaje.FechaHora, mViaje.IDRuta, 0
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
    Dim ViajeDetalle As ViajeDetalle
    Dim Reporte As Reporte
    Dim Ruta As Ruta
    Dim RutaDetalleLimite As RutaDetalle
    Dim RutaDetalleOrigen As RutaDetalle
    
    Select Case Button.Key
        Case "NEW"
            If Button.Enabled Then
                If pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_ADD) Then
                    Screen.MousePointer = vbHourglass
                    
                    Set ViajeDetalle = New ViajeDetalle
                    ViajeDetalle.FechaHora = mViaje.FechaHora
                    ViajeDetalle.IDRuta = mViaje.IDRuta
                    frmViajeDetallePropiedad.LoadDataAndShow Me, ViajeDetalle
                    Set ViajeDetalle = Nothing
                    
                    Screen.MousePointer = vbDefault
                End If
            End If
        Case "PROPERTIES"
            If Button.Enabled Then
                If pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_MODIFY) Then
                    If lvwData.SelectedItem Is Nothing Then
                        MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                        lvwData.SetFocus
                        Exit Sub
                    End If
                    If Val(Mid(lvwData.SelectedItem.Key, 2)) < 0 Then
                        MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                        lvwData.SetFocus
                        Exit Sub
                    End If
                    
                    Screen.MousePointer = vbHourglass
                    
                    Set ViajeDetalle = New ViajeDetalle
                    ViajeDetalle.FechaHora = mViaje.FechaHora
                    ViajeDetalle.IDRuta = mViaje.IDRuta
                    ViajeDetalle.Indice = Val(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1))
                    If ViajeDetalle.Load() Then
                        frmViajeDetallePropiedad.LoadDataAndShow Me, ViajeDetalle
                    Else
                        lvwData.SetFocus
                    End If
                    Set ViajeDetalle = Nothing
                    Screen.MousePointer = vbDefault
                End If
            End If
        Case "DELETE"
            If Button.Enabled Then
                If pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_DELETE) Then
                    If lvwData.SelectedItem Is Nothing Then
                        MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                        lvwData.SetFocus
                        Exit Sub
                    End If
                    If Val(Mid(lvwData.SelectedItem.Key, 2)) < 0 Then
                        MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                        lvwData.SetFocus
                        Exit Sub
                    End If
                    If MsgBox("¿Desea eliminar el Detalle del Viaje seleccionado?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                        Set ViajeDetalle = New ViajeDetalle
                        ViajeDetalle.FechaHora = mViaje.FechaHora
                        ViajeDetalle.IDRuta = mViaje.IDRuta
                        ViajeDetalle.Indice = Val(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1))
                        If ViajeDetalle.Load() Then
                            Call ViajeDetalle.Delete
                        End If
                        Set ViajeDetalle = Nothing
                    End If
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
                If Val(Mid(lvwData.SelectedItem.Key, 2)) < 0 Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                Screen.MousePointer = vbHourglass
                Forms(FormIndex).ViajePasajeroSelected mViaje.FechaHora, mViaje.IDRuta, Val(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1))
                Forms(FormIndex).SetFocus
                If tlbPin.Buttons("PIN").value = tbrUnpressed Then
                    Unload Me
                End If
                Screen.MousePointer = vbDefault
            End If
            FormWaitingForSelect = ""
        Case "PRINT"
            If mViaje.Estado = VIAJE_ESTADO_ACTIVO Then
                If MsgBox("Este Viaje no está En Progreso." & vbCr & vbCr & "¿Desea imprimirlo de todos modos?", vbExclamation + vbYesNo, App.Title) = vbNo Then
                    Exit Sub
                End If
            End If
            
            If pCPermiso.GotPermission(PERMISO_REPORTE_REPORTE & "Viaje_Planilla_Observaciones") Then
                Set Reporte = New Reporte
                Reporte.IDReporte = "Viaje_Planilla_Observaciones"
                If Reporte.Load() Then
                    Reporte.Titulo = "Detalle del Viaje: " & mViaje.FechaHora_Formatted & " - " & mViaje.Ruta_DisplayName
                    Reporte.Parametros("FechaHora_FILTER").Valor = mViaje.FechaHora
                    Reporte.Parametros("IDRuta_FILTER").Valor = mViaje.IDRuta
                    
                    If Reporte.OpenReport() Then
                        Reporte.PrintReport pParametro.Report_Preview
                    End If
                    
                    Set Reporte = Nothing
                End If
            End If
        Case "ASISTENCIA"
            If pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_MODIFY) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                If Val(Mid(lvwData.SelectedItem.Key, 2)) < 0 Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                Screen.MousePointer = vbHourglass
                
                Set ViajeDetalle = New ViajeDetalle
                With ViajeDetalle
                    .FechaHora = mViaje.FechaHora
                    .IDRuta = mViaje.IDRuta
                    .Indice = Val(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1))
                    If Not .Load() Then
                        Set ViajeDetalle = Nothing
                        Exit Sub
                    End If
                    If .Estado <> VIAJE_DETALLE_ESTADO_CONFIRMADO Then
                        MsgBox "Sólo se puede cargar la Asistencia de los Pasajeros o Comisiones con Estado Confirmado.", vbInformation, App.Title
                        lvwData.SetFocus
                        Set ViajeDetalle = Nothing
                        Exit Sub
                    End If
                    If Not pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_FINALIZADO_MODIFY, False) Then
                        If ViajeDetalle.OcupanteTipo = OCUPANTE_TIPO_PASAJERO Then
                            If mViaje.Estado = VIAJE_ESTADO_FINALIZADO And (ViajeDetalle.Realizado = 2 Or (ViajeDetalle.Realizado = 1 And ViajeDetalle.ImporteContado = ViajeDetalle.Importe)) Then
                                MsgBox "Ya se le dió Asistencia a esta Reserva.", vbInformation, App.Title
                                lvwData.SetFocus
                                Set ViajeDetalle = Nothing
                                Exit Sub
                            End If
                        Else
                            If mViaje.Estado = VIAJE_ESTADO_FINALIZADO And ViajeDetalle.ImporteContado = ViajeDetalle.Importe And ViajeDetalle.Entregada Then
                                MsgBox "Ya se le dió Asistencia a esta Comisión.", vbInformation, App.Title
                                lvwData.SetFocus
                                Set ViajeDetalle = Nothing
                                Exit Sub
                            End If
                        End If
                    End If
                    frmViajeDetalleAsistencia.LoadDataAndShow Me, ViajeDetalle
                End With
                Set ViajeDetalle = Nothing
                Screen.MousePointer = vbDefault
            End If
        Case "CHANGE_STATUS"
            If Button.Enabled Then
                CambiarEstado
            End If
    End Select
End Sub

Private Sub tlbMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim Reporte As Reporte
    Dim ViajeDetalle As ViajeDetalle
    
    Select Case ButtonMenu.Key
        Case "PRINT_PLANILLA_OBSERVACIONES"
            If mViaje.Estado = VIAJE_ESTADO_ACTIVO Then
                If MsgBox("Este Viaje no está En Progreso." & vbCr & vbCr & "¿Desea imprimirlo de todos modos?", vbExclamation + vbYesNo, App.Title) = vbNo Then
                    Exit Sub
                End If
            End If
            
            If pCPermiso.GotPermission(PERMISO_REPORTE_REPORTE & "Viaje_Planilla_Observaciones") Then
                Set Reporte = New Reporte
                Reporte.IDReporte = "Viaje_Planilla_Observaciones"
                If Reporte.Load() Then
                    Reporte.Titulo = "Detalle del Viaje: " & mViaje.FechaHora_Formatted & " - " & mViaje.Ruta_DisplayName
                    Reporte.Parametros("FechaHora_FILTER").Valor = mViaje.FechaHora
                    Reporte.Parametros("IDRuta_FILTER").Valor = mViaje.IDRuta
                    
                    If Reporte.OpenReport() Then
                        Reporte.PrintReport pParametro.Report_Preview
                    End If
                    
                    Set Reporte = Nothing
                End If
            End If
        Case "PRINT_PLANILLA_PASAJERO_OBSERVACIONES"
            If mViaje.Estado = VIAJE_ESTADO_ACTIVO Then
                If MsgBox("Este Viaje no está En Progreso." & vbCr & vbCr & "¿Desea imprimirlo de todos modos?", vbExclamation + vbYesNo, App.Title) = vbNo Then
                    Exit Sub
                End If
            End If
            
            If pCPermiso.GotPermission(PERMISO_REPORTE_REPORTE & "Viaje_Planilla_Pasajero_Observaciones") Then
                Set Reporte = New Reporte
                Reporte.IDReporte = "Viaje_Planilla_Pasajero_Observaciones"
                If Reporte.Load() Then
                    Reporte.Titulo = "Detalle del Viaje: " & mViaje.FechaHora_Formatted & " - " & mViaje.Ruta_DisplayName
                    Reporte.Parametros("FechaHora_FILTER").Valor = mViaje.FechaHora
                    Reporte.Parametros("IDRuta_FILTER").Valor = mViaje.IDRuta
                    
                    If Reporte.OpenReport() Then
                        Reporte.PrintReport pParametro.Report_Preview
                    End If
                    
                    Set Reporte = Nothing
                End If
            End If
        Case "PRINT_PLANILLA_DOCUMENTO"
            If mViaje.Estado = VIAJE_ESTADO_ACTIVO Then
                If MsgBox("Este Viaje no está En Progreso." & vbCr & vbCr & "¿Desea imprimirlo de todos modos?", vbExclamation + vbYesNo, App.Title) = vbNo Then
                    Exit Sub
                End If
            End If
            
            If pCPermiso.GotPermission(PERMISO_REPORTE_REPORTE & "Viaje_Planilla_Documento") Then
                Set Reporte = New Reporte
                Reporte.IDReporte = "Viaje_Planilla_Documento"
                If Reporte.Load() Then
                    Reporte.Titulo = "Detalle del Viaje: " & mViaje.FechaHora_Formatted & " - " & mViaje.Ruta_DisplayName
                    Reporte.Parametros("FechaHora_FILTER").Valor = mViaje.FechaHora
                    Reporte.Parametros("IDRuta_FILTER").Valor = mViaje.IDRuta
                    
                    If Reporte.OpenReport() Then
                        Reporte.PrintReport pParametro.Report_Preview
                    End If
                    
                    Set Reporte = Nothing
                End If
            End If
        Case "PRINT_PLANILLA_PASAJERO_DOCUMENTO"
            If mViaje.Estado = VIAJE_ESTADO_ACTIVO Then
                If MsgBox("Este Viaje no está En Progreso." & vbCr & vbCr & "¿Desea imprimirlo de todos modos?", vbExclamation + vbYesNo, App.Title) = vbNo Then
                    Exit Sub
                End If
            End If
            
            If pCPermiso.GotPermission(PERMISO_REPORTE_REPORTE & "Viaje_Planilla_Pasajero_Documento") Then
                Set Reporte = New Reporte
                Reporte.IDReporte = "Viaje_Planilla_Pasajero_Documento"
                If Reporte.Load() Then
                    Reporte.Titulo = "Detalle del Viaje: " & mViaje.FechaHora_Formatted & " - " & mViaje.Ruta_DisplayName
                    Reporte.Parametros("FechaHora_FILTER").Valor = mViaje.FechaHora
                    Reporte.Parametros("IDRuta_FILTER").Valor = mViaje.IDRuta
                    
                    If Reporte.OpenReport() Then
                        Reporte.PrintReport pParametro.Report_Preview
                    End If
                    
                    Set Reporte = Nothing
                End If
            End If
        Case "PRINT_PLANILLA_PASAJERO_DOMICILIO"
            If mViaje.Estado = VIAJE_ESTADO_ACTIVO Then
                If MsgBox("Este Viaje no está En Progreso." & vbCr & vbCr & "¿Desea imprimirlo de todos modos?", vbExclamation + vbYesNo, App.Title) = vbNo Then
                    Exit Sub
                End If
            End If
            
            If pCPermiso.GotPermission(PERMISO_REPORTE_REPORTE & "Viaje_Planilla_Pasajero_Domicilio") Then
                Set Reporte = New Reporte
                Reporte.IDReporte = "Viaje_Planilla_Pasajero_Domicilio"
                If Reporte.Load() Then
                    Reporte.Titulo = "Pasajeros del Viaje: " & mViaje.FechaHora_Formatted & " - " & mViaje.Ruta_DisplayName
                    Reporte.Parametros("FechaHora_FILTER").Valor = mViaje.FechaHora
                    Reporte.Parametros("IDRuta_FILTER").Valor = mViaje.IDRuta
                    
                    If Reporte.OpenReport() Then
                        Reporte.PrintReport pParametro.Report_Preview
                    End If
                    
                    Set Reporte = Nothing
                End If
            End If
        Case "PRINT_PLANILLA_COMISION"
            If mViaje.Estado = VIAJE_ESTADO_ACTIVO Then
                If MsgBox("Este Viaje no está En Progreso." & vbCr & vbCr & "¿Desea imprimirlo de todos modos?", vbExclamation + vbYesNo, App.Title) = vbNo Then
                    Exit Sub
                End If
            End If
            
            If pCPermiso.GotPermission(PERMISO_REPORTE_REPORTE & "Viaje_Planilla_Comision") Then
                Set Reporte = New Reporte
                Reporte.IDReporte = "Viaje_Planilla_Comision"
                If Reporte.Load() Then
                    Reporte.Titulo = "Detalle del Viaje: " & mViaje.FechaHora_Formatted & " - " & mViaje.Ruta_DisplayName
                    Reporte.Parametros("FechaHora_FILTER").Valor = mViaje.FechaHora
                    Reporte.Parametros("IDRuta_FILTER").Valor = mViaje.IDRuta
                    
                    If Reporte.OpenReport() Then
                        Reporte.PrintReport pParametro.Report_Preview
                    End If
                    
                    Set Reporte = Nothing
                End If
            End If
        Case "PRINT_COMISION_REMITO"
            If pCPermiso.GotPermission(PERMISO_REPORTE_REPORTE & "Comision_Remito") Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                If Val(Mid(lvwData.SelectedItem.Key, 2)) < 0 Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                Set ViajeDetalle = New ViajeDetalle
                With ViajeDetalle
                    .FechaHora = mViaje.FechaHora
                    .IDRuta = mViaje.IDRuta
                    .Indice = Val(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1))
                    If Not .Load() Then
                        Set ViajeDetalle = Nothing
                        lvwData.SetFocus
                        Exit Sub
                    End If
                    
                    If .OcupanteTipo <> OCUPANTE_TIPO_COMISION Then
                        MsgBox "El Item seleccionado no es una Comisión.", vbInformation, App.Title
                        Set ViajeDetalle = Nothing
                        lvwData.SetFocus
                        Exit Sub
                    End If
                End With
                Set ViajeDetalle = Nothing
                
                Set Reporte = New Reporte
                Reporte.IDReporte = "Comision_Remito"
                If Reporte.Load() Then
                    Reporte.Parametros("FechaHora_FILTER").Valor = mViaje.FechaHora
                    Reporte.Parametros("IDRuta_FILTER").Valor = mViaje.IDRuta
                    Reporte.Parametros("Indice_FILTER").Valor = Val(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1))
                    
                    If Reporte.OpenReport() Then
                        Reporte.PrintReport pParametro.Report_Preview
                    End If
                    
                    Set Reporte = Nothing
                End If
            End If
        Case "PRINT_COMISION_LISTADO"
            If pCPermiso.GotPermission(PERMISO_REPORTE_REPORTE & "Comision_Listado") Then
                Set Reporte = New Reporte
                Reporte.IDReporte = "Viaje_Comision_Listado"
                If Reporte.Load() Then
                    Reporte.Parametros("FechaHoraDesde").Valor = mViaje.FechaHora
                    Reporte.Parametros("FechaHoraHasta").Valor = mViaje.FechaHora
                    Reporte.Parametros("IDRuta").Valor = mViaje.IDRuta
                    Reporte.Parametros("MostrarTodas").Valor = True
                    If Reporte.OpenReport() Then
                        Reporte.PrintReport pParametro.Report_Preview
                    End If
                    Set Reporte = Nothing
                End If
            End If
        Case "PRINT_FACTURA"
            'Call PrintFactura
        Case "ASISTENCIA_SIMPLE"
            Call tlbMain_ButtonClick(tlbMain.Buttons("ASISTENCIA"))
        Case "ASISTENCIA_MULTIPLE"
            If pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_MODIFY) Then
                If lvwData.ListItems.Count = 0 Then
                    MsgBox "No hay Items en el Viaje.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                frmViajeDetalleAsistenciaMultiple.LoadDataAndShow mViaje
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
    
    lvwData.Top = CoolBarHeight + CONTROL_SPACE
    lvwData.Left = CONTROL_SPACE
    lvwData.Width = ScaleWidth - (CONTROL_SPACE * 2)
    lvwData.Height = ScaleHeight - lvwData.Top - CONTROL_SPACE - stbMain.Height
    
    tlbPin.Top = ScaleHeight - 330
    tlbPin.Left = 15
End Sub

Private Sub CambiarEstado()
    Dim ViajeDetalle As ViajeDetalle
    Dim Ruta As Ruta
    Dim RutaDetalleLimite As RutaDetalle
    Dim RutaDetalleOrigen As RutaDetalle
    
    If lvwData.SelectedItem Is Nothing Then
        MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
        lvwData.SetFocus
        Exit Sub
    End If
    If Val(Mid(lvwData.SelectedItem.Key, 2)) < 0 Then
        MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
        lvwData.SetFocus
        Exit Sub
    End If
    
    ' Cargo el detalle del viaje
    Set ViajeDetalle = New ViajeDetalle
    ViajeDetalle.FechaHora = mViaje.FechaHora
    ViajeDetalle.IDRuta = mViaje.IDRuta
    ViajeDetalle.Indice = Val(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1))
    If Not ViajeDetalle.Load() Then
        lvwData.SetFocus
        Set ViajeDetalle = Nothing
        Exit Sub
    End If
    
    ' Verifico si tiene permiso
    If ViajeDetalle.IDUsuarioCreacion = pParametro.ReservaWebIdUsuario Then
        If Not pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_WEB_CHANGE_STATUS, False) Then
            MsgBox "No está autorizado a cambiar el estado de Reservas realizadas por la Web.", vbExclamation, App.Title
            Exit Sub
        End If
    Else
        If Not pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_CHANGE_STATUS) Then
            Exit Sub
        End If
    End If
            
    ' Verifico que no haya pasado el límite de tiempo
    Set Ruta = New Ruta
    Ruta.IDRuta = mViaje.IDRuta
    If Not Ruta.Load() Then
        lvwData.SetFocus
        Set ViajeDetalle = Nothing
        Set Ruta = Nothing
        Exit Sub
    End If
            
            
    If Ruta.LimiteCancelacionDuracion > 0 And Ruta.LimiteCancelacionIDLugar > 0 Then
        Set RutaDetalleLimite = New RutaDetalle
        RutaDetalleLimite.IDRuta = mViaje.IDRuta
        RutaDetalleLimite.IDLugar = Ruta.LimiteCancelacionIDLugar
        If RutaDetalleLimite.Load() Then
            Set RutaDetalleOrigen = New RutaDetalle
            RutaDetalleOrigen.IDRuta = mViaje.IDRuta
            RutaDetalleOrigen.IDLugar = ViajeDetalle.IDOrigen
            If RutaDetalleOrigen.Load() Then
                If RutaDetalleOrigen.Indice <= RutaDetalleLimite.Indice Then
                    If DateDiff("n", ViajeDetalle.FechaHora, Now) > Ruta.LimiteCancelacionDuracion Then
                        'Tiempo Vencido, habilito según Permiso
                        If pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_CHANGE_STATUS_AFTER_LIMIT, False) Then
                            'Permitido por Permiso
                            Select Case ViajeDetalle.Estado
                                Case VIAJE_DETALLE_ESTADO_CONFIRMADO, VIAJE_DETALLE_ESTADO_CONDICIONAL
                                    frmViajeDetalleCancelar.LoadDataAndShow Me, ViajeDetalle
                                Case VIAJE_DETALLE_ESTADO_CANCELADO
                                    If ViajeDetalle.OcupanteTipo = OCUPANTE_TIPO_PASAJERO Then
                                        If MsgBox("Esta Reserva está Cancelada." & vbCr & vbCr & "Orden: " & ViajeDetalle.Orden & vbCr & "Pasajero: " & lvwData.SelectedItem.SubItems(1) & ", " & lvwData.SelectedItem.SubItems(2) & vbCr & vbCr & "¿Desea Activar esta Reserva?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                                            ViajeDetalle.Estado = ""
                                            Call ViajeDetalle.CambiarEstado(pParametro.Viaje_Permite_RutaConexion)
                                        End If
                                    Else
                                        If MsgBox("Esta Comisión está Cancelada." & vbCr & vbCr & "Orden: " & ViajeDetalle.Orden & vbCr & "Remitente: " & lvwData.SelectedItem.SubItems(1) & ", " & lvwData.SelectedItem.SubItems(2) & vbCr & vbCr & "¿Desea Activar esta Comisión?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                                            ViajeDetalle.Estado = VIAJE_DETALLE_ESTADO_CONFIRMADO
                                            Call ViajeDetalle.CambiarEstado(pParametro.Viaje_Permite_RutaConexion)
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
                        Select Case ViajeDetalle.Estado
                            Case VIAJE_DETALLE_ESTADO_CONFIRMADO, VIAJE_DETALLE_ESTADO_CONDICIONAL
                                frmViajeDetalleCancelar.LoadDataAndShow Me, ViajeDetalle
                            Case VIAJE_DETALLE_ESTADO_CANCELADO
                                If ViajeDetalle.OcupanteTipo = OCUPANTE_TIPO_PASAJERO Then
                                    If MsgBox("Esta Reserva está Cancelada." & vbCr & vbCr & "Orden: " & ViajeDetalle.Orden & vbCr & "Pasajero: " & lvwData.SelectedItem.SubItems(1) & ", " & lvwData.SelectedItem.SubItems(2) & vbCr & vbCr & "¿Desea Activar esta Reserva?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                                        ViajeDetalle.Estado = ""
                                        Call ViajeDetalle.CambiarEstado(pParametro.Viaje_Permite_RutaConexion)
                                    End If
                                Else
                                    If MsgBox("Esta Comisión está Cancelada." & vbCr & vbCr & "Orden: " & ViajeDetalle.Orden & vbCr & "Remitente: " & lvwData.SelectedItem.SubItems(1) & ", " & lvwData.SelectedItem.SubItems(2) & vbCr & vbCr & "¿Desea Activar esta Comisión?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                                        ViajeDetalle.Estado = VIAJE_DETALLE_ESTADO_CONFIRMADO
                                        Call ViajeDetalle.CambiarEstado(pParametro.Viaje_Permite_RutaConexion)
                                    End If
                                End If
                            Case Else
                                MsgBox "Estado Incorrecto.", vbCritical, App.Title
                        End Select
                    End If
                Else
                    'Permitido porque el Origen está antes que el Límite
                    Select Case ViajeDetalle.Estado
                        Case VIAJE_DETALLE_ESTADO_CONFIRMADO, VIAJE_DETALLE_ESTADO_CONDICIONAL
                            frmViajeDetalleCancelar.LoadDataAndShow Me, ViajeDetalle
                        Case VIAJE_DETALLE_ESTADO_CANCELADO
                            If ViajeDetalle.OcupanteTipo = OCUPANTE_TIPO_PASAJERO Then
                                If MsgBox("Esta Reserva está Cancelada." & vbCr & vbCr & "Orden: " & ViajeDetalle.Orden & vbCr & "Pasajero: " & lvwData.SelectedItem.SubItems(1) & ", " & lvwData.SelectedItem.SubItems(2) & vbCr & vbCr & "¿Desea Activar esta Reserva?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                                    ViajeDetalle.Estado = ""
                                    Call ViajeDetalle.CambiarEstado(pParametro.Viaje_Permite_RutaConexion)
                                End If
                            Else
                                If MsgBox("Esta Comisión está Cancelada." & vbCr & vbCr & "Orden: " & ViajeDetalle.Orden & vbCr & "Remitente: " & lvwData.SelectedItem.SubItems(1) & ", " & lvwData.SelectedItem.SubItems(2) & vbCr & vbCr & "¿Desea Activar esta Comisión?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                                    ViajeDetalle.Estado = VIAJE_DETALLE_ESTADO_CONFIRMADO
                                    Call ViajeDetalle.CambiarEstado(pParametro.Viaje_Permite_RutaConexion)
                                End If
                            End If
                        Case Else
                            MsgBox "Estado Incorrecto.", vbCritical, App.Title
                    End Select
                End If
            Else
                lvwData.SetFocus
                Set ViajeDetalle = Nothing
                Set Ruta = Nothing
                Set RutaDetalleLimite = Nothing
                Set RutaDetalleOrigen = Nothing
                Exit Sub
            End If
            Set RutaDetalleOrigen = Nothing
        Else
            lvwData.SetFocus
            Set ViajeDetalle = Nothing
            Set Ruta = Nothing
            Set RutaDetalleLimite = Nothing
            Exit Sub
        End If
        Set RutaDetalleLimite = Nothing
    Else
        'Permitido porque la Ruta no tiene Límite
        Select Case ViajeDetalle.Estado
            Case VIAJE_DETALLE_ESTADO_CONFIRMADO, VIAJE_DETALLE_ESTADO_CONDICIONAL
                frmViajeDetalleCancelar.LoadDataAndShow Me, ViajeDetalle
            Case VIAJE_DETALLE_ESTADO_CANCELADO
                If ViajeDetalle.OcupanteTipo = OCUPANTE_TIPO_PASAJERO Then
                    If MsgBox("Esta Reserva está Cancelada." & vbCr & vbCr & "Orden: " & ViajeDetalle.Orden & vbCr & "Pasajero: " & lvwData.SelectedItem.SubItems(1) & ", " & lvwData.SelectedItem.SubItems(2) & vbCr & vbCr & "¿Desea Activar esta Reserva?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                        ViajeDetalle.Estado = ""
                        Call ViajeDetalle.CambiarEstado(pParametro.Viaje_Permite_RutaConexion)
                    End If
                Else
                    If MsgBox("Esta Comisión está Cancelada." & vbCr & vbCr & "Orden: " & ViajeDetalle.Orden & vbCr & "Remitente: " & lvwData.SelectedItem.SubItems(1) & ", " & lvwData.SelectedItem.SubItems(2) & vbCr & vbCr & "¿Desea Activar esta Comisión?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                        ViajeDetalle.Estado = VIAJE_DETALLE_ESTADO_CONFIRMADO
                        Call ViajeDetalle.CambiarEstado(pParametro.Viaje_Permite_RutaConexion)
                    End If
                End If
            Case Else
                MsgBox "Estado Incorrecto.", vbCritical, App.Title
        End Select
    End If
    
    Set Ruta = Nothing
            
    SetLastPersona ViajeDetalle.IDPersona
    Set ViajeDetalle = Nothing
End Sub
