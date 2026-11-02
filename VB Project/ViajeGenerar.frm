VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmViajeGenerar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar Viajes"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ViajeGenerar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5265
   ScaleWidth      =   7830
   Begin VB.CommandButton cmdDosSemanas 
      Caption         =   "Dos Semanas"
      Height          =   390
      Left            =   3120
      TabIndex        =   32
      Top             =   4740
      Width           =   1395
   End
   Begin VB.CommandButton cmdUnDia 
      Caption         =   "Un Día"
      Height          =   390
      Left            =   120
      TabIndex        =   30
      Top             =   4740
      Width           =   1395
   End
   Begin VB.CommandButton cmdUnaSemana 
      Caption         =   "Una Semana"
      Height          =   390
      Left            =   1620
      TabIndex        =   31
      Top             =   4740
      Width           =   1395
   End
   Begin VB.CheckBox chkDia 
      Caption         =   "Día 6"
      Height          =   210
      Index           =   13
      Left            =   4020
      TabIndex        =   28
      Top             =   4200
      Width           =   2235
   End
   Begin VB.ComboBox cboBasadoSemana 
      Height          =   330
      Index           =   13
      Left            =   6420
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   4140
      Width           =   1290
   End
   Begin VB.CheckBox chkDia 
      Caption         =   "Día 1"
      Height          =   210
      Index           =   7
      Left            =   4020
      TabIndex        =   16
      Top             =   1680
      Width           =   2235
   End
   Begin VB.ComboBox cboBasadoSemana 
      Height          =   330
      Index           =   7
      Left            =   6420
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1620
      Width           =   1290
   End
   Begin VB.CheckBox chkDia 
      Caption         =   "Día 2"
      Height          =   210
      Index           =   8
      Left            =   4020
      TabIndex        =   18
      Top             =   2100
      Width           =   2235
   End
   Begin VB.ComboBox cboBasadoSemana 
      Height          =   330
      Index           =   8
      Left            =   6420
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2040
      Width           =   1290
   End
   Begin VB.CheckBox chkDia 
      Caption         =   "Día 3"
      Height          =   210
      Index           =   9
      Left            =   4020
      TabIndex        =   20
      Top             =   2520
      Width           =   2235
   End
   Begin VB.ComboBox cboBasadoSemana 
      Height          =   330
      Index           =   9
      Left            =   6420
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   2460
      Width           =   1290
   End
   Begin VB.CheckBox chkDia 
      Caption         =   "Día 4"
      Height          =   210
      Index           =   10
      Left            =   4020
      TabIndex        =   22
      Top             =   2940
      Width           =   2235
   End
   Begin VB.ComboBox cboBasadoSemana 
      Height          =   330
      Index           =   10
      Left            =   6420
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   2880
      Width           =   1290
   End
   Begin VB.CheckBox chkDia 
      Caption         =   "Día 5"
      Height          =   210
      Index           =   11
      Left            =   4020
      TabIndex        =   24
      Top             =   3360
      Width           =   2235
   End
   Begin VB.ComboBox cboBasadoSemana 
      Height          =   330
      Index           =   11
      Left            =   6420
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   3300
      Width           =   1290
   End
   Begin VB.CheckBox chkDia 
      Caption         =   "Día 6"
      Height          =   210
      Index           =   12
      Left            =   4020
      TabIndex        =   26
      Top             =   3780
      Width           =   2235
   End
   Begin VB.ComboBox cboBasadoSemana 
      Height          =   330
      Index           =   12
      Left            =   6420
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   3720
      Width           =   1290
   End
   Begin VB.ComboBox cboBasadoSemana 
      Height          =   330
      Index           =   6
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   4140
      Width           =   1290
   End
   Begin VB.CheckBox chkDia 
      Caption         =   "Día 6"
      Height          =   210
      Index           =   6
      Left            =   120
      TabIndex        =   14
      Top             =   4200
      Width           =   2235
   End
   Begin VB.ComboBox cboBasadoSemana 
      Height          =   330
      Index           =   5
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3720
      Width           =   1290
   End
   Begin VB.CheckBox chkDia 
      Caption         =   "Día 5"
      Height          =   210
      Index           =   5
      Left            =   120
      TabIndex        =   12
      Top             =   3780
      Width           =   2235
   End
   Begin VB.ComboBox cboBasadoSemana 
      Height          =   330
      Index           =   4
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3300
      Width           =   1290
   End
   Begin VB.CheckBox chkDia 
      Caption         =   "Día 4"
      Height          =   210
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   2235
   End
   Begin VB.ComboBox cboBasadoSemana 
      Height          =   330
      Index           =   3
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2880
      Width           =   1290
   End
   Begin VB.CheckBox chkDia 
      Caption         =   "Día 3"
      Height          =   210
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   2940
      Width           =   2235
   End
   Begin VB.ComboBox cboBasadoSemana 
      Height          =   330
      Index           =   2
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2460
      Width           =   1290
   End
   Begin VB.CheckBox chkDia 
      Caption         =   "Día 2"
      Height          =   210
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   2235
   End
   Begin VB.ComboBox cboBasadoSemana 
      Height          =   330
      Index           =   1
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2040
      Width           =   1290
   End
   Begin VB.CheckBox chkDia 
      Caption         =   "Día 1"
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   2100
      Width           =   2235
   End
   Begin MSComctlLib.ProgressBar prbStatus 
      Height          =   315
      Left            =   120
      TabIndex        =   42
      Top             =   5520
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   33
      Top             =   4740
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6480
      TabIndex        =   34
      Top             =   4740
      Width           =   1215
   End
   Begin VB.ComboBox cboBasado 
      Height          =   330
      Left            =   6420
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1020
      Width           =   1290
   End
   Begin VB.Frame fraLine 
      Height          =   75
      Left            =   120
      TabIndex        =   39
      Top             =   780
      Width           =   7575
   End
   Begin VB.CommandButton cmdSiguiente 
      Height          =   315
      Left            =   4020
      Picture         =   "ViajeGenerar.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   37
      TabStop         =   0   'False
      ToolTipText     =   "Siguiente"
      Top             =   1020
      Width           =   300
   End
   Begin VB.CommandButton cmdAnterior 
      Height          =   315
      Left            =   1020
      Picture         =   "ViajeGenerar.frx":0596
      Style           =   1  'Graphical
      TabIndex        =   35
      TabStop         =   0   'False
      ToolTipText     =   "Anterior"
      Top             =   1020
      Width           =   300
   End
   Begin VB.CommandButton cmdHoy 
      Height          =   315
      Left            =   4320
      Picture         =   "ViajeGenerar.frx":0B20
      Style           =   1  'Graphical
      TabIndex        =   38
      TabStop         =   0   'False
      ToolTipText     =   "Hoy"
      Top             =   1020
      Width           =   315
   End
   Begin VB.TextBox txtDiaSemana 
      Height          =   315
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1020
      Width           =   1050
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Top             =   1020
      Width           =   1635
      _ExtentX        =   2884
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
   Begin VB.Line Line2 
      X1              =   3900
      X2              =   3900
      Y1              =   1500
      Y2              =   4440
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7680
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Label lblBasadoSemana 
      AutoSize        =   -1  'True
      Caption         =   "Basado en:"
      Height          =   210
      Left            =   2520
      TabIndex        =   44
      Top             =   1680
      Width           =   825
   End
   Begin VB.Label lblFechaSemana 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
      Height          =   210
      Left            =   120
      TabIndex        =   43
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Height          =   210
      Left            =   120
      TabIndex        =   41
      Top             =   5280
      Visible         =   0   'False
      Width           =   7605
   End
   Begin VB.Label lblBasado 
      AutoSize        =   -1  'True
      Caption         =   "Basado en:"
      Height          =   210
      Left            =   5220
      TabIndex        =   2
      Top             =   1080
      Width           =   825
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "ViajeGenerar.frx":0C6A
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Se Generarán los Viajes correspondientes"
      Height          =   210
      Left            =   780
      TabIndex        =   40
      Top             =   300
      Width           =   3090
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      Caption         =   "&Fecha:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   495
   End
End
Attribute VB_Name = "frmViajeGenerar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mRunning As Boolean
Private mCanceled As Boolean

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByVal Fecha As Date)
    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    dtpFecha.value = Fecha
    dtpFecha_Change
    
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Private Sub cmdDosSemanas_Click()
    Dim ControlIndex As Byte
    
    For ControlIndex = 1 To 13
        chkDia(ControlIndex).value = vbChecked
    Next ControlIndex
End Sub

Private Sub cmdOK_Click()
    Dim ControlIndex As Byte
    Dim MultipleLeyenda As String
    
    If DateDiff("d", Date, dtpFecha.value) < 0 Then
        If MsgBox("Está por Generar Viajes correspondientes a una Fecha anterior a Hoy." & vbCr & "Las Reservas Fijas no serán Generadas." & vbCr & vbCr & "¿Desea continuar de todos modos?", vbExclamation + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If
    End If
    
    For ControlIndex = 1 To 13
        If chkDia(ControlIndex).value = vbChecked Then
            MultipleLeyenda = MultipleLeyenda & vbCr & "Fecha: " & chkDia(ControlIndex).Caption & "    - Basado en: " & cboBasadoSemana(ControlIndex).Text
        End If
    Next ControlIndex
    
    If MsgBox("Se Generarán los Viajes correspondientes a:" & vbCr & vbCr & "Fecha: " & txtDiaSemana.Text & ", " & dtpFecha.value & " - Basado en: " & cboBasado.Text & MultipleLeyenda & vbCr & vbCr & "¿Desea continuar?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        EnableControls False
        lblStatus.Visible = True
        prbStatus.Visible = True
        Me.Height = 6360
        DoEvents
        
        mRunning = True
        mCanceled = False
        
        GenerarViajes dtpFecha.value, cboBasado.ListIndex + 1
        
        For ControlIndex = 1 To 13
            If chkDia(ControlIndex).value = vbChecked Then
                GenerarViajes DateAdd("d", ControlIndex, dtpFecha.value), cboBasadoSemana(ControlIndex).ListIndex + 1
            End If
        Next ControlIndex
        
        Me.Height = 5640
        lblStatus.Visible = False
        lblStatus.Caption = ""
        prbStatus.Visible = False
        prbStatus.value = 0
        EnableControls True
        
        mRunning = False
        
        If mCanceled Then
            MsgBox "La Generación de Viajes ha sido cancelada por el Usuario.", vbExclamation, App.Title
        Else
            MsgBox "Se han Generado los Viajes.", vbInformation, App.Title
            Unload Me
        End If
    End If
End Sub

Private Sub GenerarViajes(ByVal Fecha As Date, ByVal DiaSemana As Byte)
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    
    Dim Viaje As Viaje
    Dim Feriado As Feriado
    
    Dim errorMessage As String
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    'CHEQUEAR FERIADOS
    Set Feriado = New Feriado
    Feriado.NoMatchRaiseError = False
    Feriado.Fecha = Fecha
    If Feriado.Load() Then
        If Not Feriado.NoMatch Then
            If MsgBox("El Día " & Feriado.Fecha_Formatted & " es feriado" & IIf(Feriado.Nombre = "", ".", " (" & Feriado.Nombre & ").") & vbCr & vbCr & "¿Desea Generar los Viajes de este día?", vbExclamation + vbYesNo, App.Title) = vbNo Then
                Exit Sub
            End If
        End If
    End If
    Set Feriado = Nothing
    
    Screen.MousePointer = vbHourglass
        
    lblStatus.Caption = "Leyendo Horarios del " & WeekdayName(DiaSemana) & "..."
    lblStatus.Refresh
    errorMessage = "Error al leer los Horarios."
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_Horario_List_DiaSemana"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("DiaSemana_FILTER", adTinyInt, adParamInput, , DiaSemana)
    Set recData = New ADODB.Recordset
    recData.Open cmdData, , adOpenForwardOnly, adLockReadOnly
    Set cmdData = Nothing
    
    lblStatus.Caption = "Abriendo la tabla de Viajes..."
    lblStatus.Refresh
    
    errorMessage = "Error al abrir la tabla de Viajes."
    
    Set Viaje = New Viaje
    Viaje.NoMatchRaiseError = False
    Viaje.RefreshListSkip = True
    
    Do While (Not recData.EOF) And (Not mCanceled)
        'Busco a ver si existe el Viaje, y si no, lo genero
        lblStatus.Caption = "Buscando el Viaje: " & WeekdayName(Weekday(Fecha)) & ", " & Format(Fecha, "Short Date") & " " & Format(recData("Hora").value, "Short Time") & " - " & RTrim(recData("IDRuta").value) & "..."
        lblStatus.Refresh
        
        errorMessage = "Error al buscar el Viaje."
        
        Viaje.FechaHora = CDate(Format(Fecha, "Short Date") & " " & Format(recData("Hora").value, "Short Time"))
        Viaje.IDRuta = recData("IDRuta").value
        
        If Viaje.Load() Then
            If Viaje.NoMatch Then
                lblStatus.Caption = "Generando el Viaje: " & WeekdayName(Weekday(Fecha)) & ", " & Format(Fecha, "Short Date") & " " & Format(recData("Hora").value, "Short Time") & " - " & RTrim(recData("IDRuta").value) & "..."
                lblStatus.Refresh
                
                Select Case Viaje.IDRuta
                    Case pParametro.Ruta_ID_Otra
                        Viaje.RutaOtra = recData("RutaOtra").value & ""
                        Viaje.IDPersona = Val(recData("IDPersona").value & "")
                        Viaje.Importe = IIf(IsNull(recData("Importe").value), 0, recData("Importe").value)
                    Case pParametro.Ruta_Paquete_ID
                        Viaje.RutaOtra = recData("RutaOtra").value & ""
                        Viaje.IDPersona = 0
                        Viaje.Importe = IIf(IsNull(recData("Importe").value), 0, recData("Importe").value)
                    Case Else
                        Viaje.RutaOtra = ""
                        Viaje.IDPersona = 0
                        Viaje.Importe = 0
                End Select
                Viaje.Kilometro = Val(recData("Kilometro").value & "")
                Viaje.Duracion = Val(recData("Duracion").value & "")
                
                Viaje.IDVehiculo = Val(recData("IDVehiculo").value & "")
                Viaje.IDConductor = Val(recData("IDConductor").value & "")
                Viaje.IDConductor2 = Val(recData("IDConductor2").value & "")
                Viaje.DiaSemanaBase = DiaSemana
                Viaje.Personal = recData("Personal").value
                Viaje.Update
            End If
        End If
        
        prbStatus.value = ((recData.AbsolutePosition) / recData.RecordCount) * 100
        DoEvents
        recData.MoveNext
    Loop
    
    RefreshList_RefreshViaje Now, ""
    DoEvents
    
    recData.Close
    Set recData = Nothing
    
    Set Viaje = Nothing

    Screen.MousePointer = vbDefault
    Exit Sub

ErrorHandler:
    ShowErrorMessage "Forms.ViajeGenerar.GenerarViajes", errorMessage & vbCr & vbCr & "Fecha: " & Format(Fecha, "Short Date") & vbCr & "Basado en: " & WeekdayName(DiaSemana)
End Sub

Private Sub cmdCancel_Click()
    If mRunning Then
        mCanceled = True
    Else
        Unload Me
    End If
End Sub

Private Sub cmdUnaSemana_Click()
    Dim ControlIndex As Byte
    
    For ControlIndex = 1 To 6
        chkDia(ControlIndex).value = vbChecked
    Next ControlIndex
    For ControlIndex = 7 To 13
        chkDia(ControlIndex).value = vbUnchecked
    Next ControlIndex
End Sub

Private Sub cmdUnDia_Click()
    Dim ControlIndex As Byte
    
    For ControlIndex = 1 To 13
        chkDia(ControlIndex).value = vbUnchecked
    Next ControlIndex
End Sub

Private Sub dtpFecha_Change()
    Dim ControlIndex As Byte
    Dim Fecha As Date
    
    txtDiaSemana.Text = WeekdayName(Weekday(dtpFecha.value))
    cboBasado.ListIndex = Weekday(dtpFecha.value) - 1
    For ControlIndex = 1 To 13
        Fecha = DateAdd("d", ControlIndex, dtpFecha.value)
        chkDia(ControlIndex).Caption = WeekdayName(Weekday(Fecha)) & ", " & Format(Fecha, "Short Date")
        cboBasadoSemana(ControlIndex).ListIndex = Weekday(Fecha) - 1
    Next ControlIndex
End Sub

Private Sub cmdAnterior_Click()
    dtpFecha.value = DateAdd("d", -1, dtpFecha.value)
    dtpFecha.SetFocus
    dtpFecha_Change
End Sub

Private Sub cmdSiguiente_Click()
    dtpFecha.value = DateAdd("d", 1, dtpFecha.value)
    dtpFecha.SetFocus
    dtpFecha_Change
End Sub

Private Sub cmdHoy_Click()
    Dim OldValue As Date
    
    OldValue = dtpFecha.value
    dtpFecha.value = Date
    dtpFecha.SetFocus
    If OldValue <> dtpFecha.value Then
        dtpFecha_Change
    End If
End Sub

Private Sub Form_Load()
    Dim Weekday As Byte
    Dim ControlIndex As Byte
    
    For Weekday = 1 To 7
        cboBasado.AddItem WeekdayName(Weekday)
        For ControlIndex = 1 To 13
            cboBasadoSemana(ControlIndex).AddItem WeekdayName(Weekday)
        Next ControlIndex
    Next Weekday
    
    dtpFecha.value = Date
    dtpFecha_Change
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmViajeGenerar = Nothing
End Sub

Private Sub EnableControls(ByVal value As Boolean)
    Dim Dia As Integer
    
    cmdAnterior.Enabled = value
    txtDiaSemana.Enabled = value
    dtpFecha.Enabled = value
    cmdSiguiente.Enabled = value
    cmdHoy.Enabled = value
    cboBasado.Enabled = value
    
    For Dia = 1 To 13
        chkDia(Dia).Enabled = value
        cboBasadoSemana(Dia).Enabled = value
    Next Dia
    
    cmdUnDia.Enabled = value
    cmdUnaSemana.Enabled = value
    cmdDosSemanas.Enabled = value
    cmdOK.Enabled = value
End Sub
