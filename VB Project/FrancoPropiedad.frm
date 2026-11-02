VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFrancoPropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrancoPropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3300
   ScaleWidth      =   4680
   Begin VB.TextBox txtImporteConductor 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2400
      MaxLength       =   20
      TabIndex        =   5
      Tag             =   "CURRENCY|EMPTY|ZERO|POSITIVE"
      Top             =   1980
      Width           =   1155
   End
   Begin VB.ComboBox cboConductor 
      Height          =   330
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1500
      Width           =   3450
   End
   Begin VB.CommandButton cmdAuditoria 
      Caption         =   "Auditoría"
      Height          =   615
      Left            =   3600
      Picture         =   "FrancoPropiedad.frx":054A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   60
      Width           =   975
   End
   Begin VB.Frame fraLine 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   75
      Left            =   120
      TabIndex        =   10
      Top             =   780
      Width           =   4455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3300
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1980
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   315
      Left            =   1080
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
      MaxDate         =   73050
      MinDate         =   36526
   End
   Begin VB.Label lblImporteConductor 
      AutoSize        =   -1  'True
      Caption         =   "&Importe a pagar al Conductor:"
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   2130
   End
   Begin VB.Label lblConductor 
      AutoSize        =   -1  'True
      Caption         =   "Conductor:"
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   795
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
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Ingrese aquí los Datos del Franco"
      Height          =   210
      Left            =   780
      TabIndex        =   8
      Top             =   300
      Width           =   2415
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmFrancoPropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mFranco As Franco
Private mNew As Boolean
Private mKeyDecimal As Boolean

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef Franco As Franco)
    Set mFranco = Franco
    mNew = (mFranco.Fecha = DATE_TIME_FIELD_NULL_VALUE)

    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    With mFranco
        If mNew Then
            dtpFecha.value = Date
        Else
            dtpFecha.value = .Fecha
        End If
        dtpFecha_Change
        
        Call FillComboBoxConductor
        
        cboConductor.ListIndex = CSM_Control_ComboBox.GetListIndexByItemData(cboConductor, .IDPersona, cscpCurrentOrFirstIfUnique)
        
        txtImporteConductor.Text = .Importe_FormattedAsString
    End With
    
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call CSM_Control_TextBox.PrepareAll(Me)

    If cboConductor.ListIndex = -1 Then
        MsgBox "Debe seleccionar el Conductor.", vbInformation, App.Title
        cboConductor.SetFocus
        Exit Sub
    End If
    
    With mFranco
        .Fecha = dtpFecha.value
        .IDPersona = cboConductor.ItemData(cboConductor.ListIndex)
        .Importe_FormattedAsString = txtImporteConductor.Text
        If Not .Update Then
            Exit Sub
        End If
    End With
    
    Unload Me
End Sub

Private Sub cmdAuditoria_Click()
    frmAuditoriaGenerico.LoadDataAndShow mFranco
End Sub

Private Sub dtpFecha_Change()
    Caption = "Propiedades del Franco: " & dtpFecha.value
End Sub

Private Sub txtImporteConductor_GotFocus()
    Call CSM_Control_TextBox.SelAllText(txtImporteConductor)
End Sub

Private Sub txtImporteConductor_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CSM_Control_TextBox.CheckKeyDown(txtImporteConductor, KeyCode)
End Sub

Private Sub txtImporteConductor_KeyPress(KeyAscii As Integer)
    Call CSM_Control_TextBox.CheckKeyPress(txtImporteConductor, KeyAscii, mKeyDecimal)
End Sub

Private Sub txtImporteConductor_LostFocus()
    Call CSM_Control_TextBox.FormatValue_ByTag(txtImporteConductor)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mFranco = Nothing
    Set frmFrancoPropiedad = Nothing
End Sub

Public Sub FillComboBoxConductor()
    Dim KeySave As Long
    
    If cboConductor.ListCount > 0 Then
        KeySave = cboConductor.ItemData(cboConductor.ListIndex)
    End If
    Call CSM_Control_ComboBox.FillFromSQL(cboConductor, "SELECT IDPersona, dbo.udf_GetEntidadApellidoYNombre(Apellido, Nombre) AS ApellidoNombre FROM Persona WHERE Activo = 1 AND EntidadTipo = '" & ENTIDAD_TIPO_PERSONA_CONDUCTOR & "' ORDER BY Apellido, Nombre", "IDPersona", "ApellidoNombre", "Conductores", cscpItemOrFirst, KeySave)
End Sub
