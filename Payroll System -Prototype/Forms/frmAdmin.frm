VERSION 5.00
Begin VB.Form frmAdmin 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3000
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   3720
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAdmin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3495
      Begin VB.TextBox txtUsername 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   720
         Width           =   1935
      End
      Begin PayrollSystem.jcbutton cmdCancel 
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14935011
         Caption         =   "&Cancel"
         Picture         =   "frmAdmin.frx":000C
         UseMaskCOlor    =   -1  'True
         MaskColor       =   16777215
      End
      Begin PayrollSystem.jcbutton cmdOk 
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14935011
         Caption         =   "&OK"
         Picture         =   "frmAdmin.frx":0288
         UseMaskCOlor    =   -1  'True
         MaskColor       =   16777215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Image cmdExit 
      Height          =   360
      Left            =   3360
      MousePointer    =   99  'Custom
      Picture         =   "frmAdmin.frx":04E2
      Top             =   0
      Width           =   360
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Security Check"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "Administrator"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmAdmin.frx":0BCC
      Top             =   0
      Width           =   11025
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Check_Fields()
If txtUsername.Text = "" Or txtPassword.Text = "" Then
    If txtUsername.Text = "" Then
        txtUsername.BackColor = &HC0C0FF
        txtUsername.SetFocus
        Exit Sub
    End If
    If txtPassword.Text = "" Then
        txtPassword.BackColor = &HC0C0FF
        txtPassword.SetFocus
        Exit Sub
    End If
End If
End Sub

Private Sub cmdCancel_Click()
frmLogIn.Show
Unload Me
End Sub

Private Sub cmdOK_Click()
SQL = "SELECT * FROM Accounts"
connOpen
With RS
    .MoveFirst
    While Not .EOF
        If txtUsername.Text = .Fields!UserName And txtPassword.Text = .Fields!Password Then
            If .Fields!Position = "Administrator" Then
                connClose
                frmMainMenu.Show
                Unload Me
            Else
                connClose
                dlgDTR.Show
                Unload Me
            End If
            Exit Sub
        Else
            .MoveNext
        End If
    Wend
    MsgBox "Invalid Username and Password!"
End With
connClose
End Sub

Private Sub cmdExit_Click()
Call cmdCancel_Click
End Sub


Private Sub txtPassword_Change()
If txtPassword.BackColor = &HC0C0FF Then txtPassword.BackColor = &H80000005
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdOK_Click
End If
End Sub

Private Sub txtUsername_Change()
If txtUsername.BackColor = &HC0C0FF Then txtUsername.BackColor = &H80000005
End Sub
