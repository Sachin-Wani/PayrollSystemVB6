VERSION 5.00
Begin VB.Form frmSysLock 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1095
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6315
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSysLock.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLock 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin PayrollSystem.jcbutton cmdUnlock 
      Height          =   615
      Left            =   5400
      TabIndex        =   0
      Top             =   240
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   ""
      Picture         =   "frmSysLock.frx":000C
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "System Lock:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frmSysLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdUnlock_Click()
SQL = "SELECT * FROM Accounts WHERE Password = '" & txtLock.Text & "'"
connectDB
With RS
    If .RecordCount <> 0 And .Fields!Position = "Administrator" Then
        frmMainMenu.Enabled = True
        connClose
        Unload Me
        Exit Sub
    Else
        txtLock.Text = ""
        txtLock.BackColor = &HC0C0FF
        txtLock.SetFocus
    End If
End With
connClose
'With RS
'    While Not .EOF
'        If txtLock.Text = .Fields!Password Then
'            If .Fields!Position = "Administrator" Then
'                frmMainMenu.Enabled = True
'                connClose
'                Unload Me
'                Exit Sub
'            Else
'                txtLock.BackColor = &HC0C0FF
'                txtLock.SetFocus
'                connClose
'                Exit Sub
'            End If
'        Else
'            .MoveNext
'        End If
'    Wend
'    txtLock.BackColor = &HC0C0FF
'    txtLock.SetFocus
'End With
'connClose
End Sub

Private Sub Form_Load()
frmMainMenu.Enabled = False
End Sub

Private Sub txtLock_Change()
If txtLock.BackColor = &HC0C0FF Then txtLock.BackColor = &HFFFFFF
End Sub

Private Sub txtLock_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdUnlock_Click
End If
End Sub
