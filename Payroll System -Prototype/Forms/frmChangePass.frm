VERSION 5.00
Begin VB.Form frmChangePass 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4890
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   4470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmChangePass.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4215
      Begin VB.CheckBox chkUnMask 
         Caption         =   "Unmask Password"
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
         Left            =   1200
         TabIndex        =   14
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Frame Frame2 
         Height          =   1215
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   3975
         Begin VB.TextBox txtNewPass2 
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
            Left            =   2040
            PasswordChar    =   "*"
            TabIndex        =   8
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox txtNewPass 
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
            Left            =   2040
            PasswordChar    =   "*"
            TabIndex        =   7
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label5 
            Caption         =   "Confirm Password:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label Label4 
            Caption         =   "New Password:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   9
            Top             =   240
            Width           =   1575
         End
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
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   720
         Width           =   1815
      End
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
         Left            =   2160
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Old Password:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   495
      Left            =   2400
      Picture         =   "frmChangePass.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Height          =   495
      Left            =   600
      Picture         =   "frmChangePass.frx":224E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Save"
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "Change Password"
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
      TabIndex        =   13
      Top             =   0
      Width           =   3015
   End
   Begin VB.Image cmdExit 
      Height          =   360
      Left            =   4080
      MousePointer    =   99  'Custom
      Picture         =   "frmChangePass.frx":4490
      Top             =   0
      Width           =   360
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Old and New Password"
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
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   -360
      Picture         =   "frmChangePass.frx":4B7A
      Top             =   0
      Width           =   11025
   End
End
Attribute VB_Name = "frmChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Clear_Fields()
txtUsername.Text = ""
txtPassword.Text = ""
txtNewPass.Text = ""
txtNewPass2.Text = ""
chkUnMask.Value = 0
End Sub


Sub Check_Fields()
If txtUsername.Text = "" Or txtPassword.Text = "" Or txtNewPass.Text = "" Or txtNewPass2.Text = "" Then
    If txtUsername.Text = "" Then
        txtUsername.BackColor = &HC0C0FF
        txtUsername.SetFocus
    End If
    If txtPassword.Text = "" Then
        txtPassword.BackColor = &HC0C0FF
        txtPassword.SetFocus
    End If
    If txtNewPass.Text = "" Then
        txtNewPass.BackColor = &HC0C0FF
        txtNewPass.SetFocus
    End If
    If txtNewPass2.Text = "" Then
        txtNewPass2.BackColor = &HC0C0FF
        txtNewPass2.SetFocus
    End If
    Exit Sub
    Exit Sub
End If
End Sub

Sub Check_Password()
If txtNewPass.Text <> txtNewPass2.Text Then
    txtNewPass.BackColor = &HC0C0FF
    txtNewPass2.BackColor = &HC0C0FF
    txtNewPass.SetFocus
    Exit Sub
    Exit Sub
End If
End Sub

Private Sub chkUnMask_Click()
If chkUnMask.Value = 1 Then
    txtPassword.PasswordChar = ""
    txtNewPass.PasswordChar = ""
    txtNewPass2.PasswordChar = ""
Else
    txtPassword.PasswordChar = "*"
    txtNewPass.PasswordChar = "*"
    txtNewPass2.PasswordChar = "*"
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdExit_Click()
Call cmdCancel_Click
End Sub


Private Sub cmdSave_Click()
On Error Resume Next
' Checkpoint
Check_Fields
Check_Password

' If passed, Save New Password
SQL = "SELECT * FROM Accounts"
connectDB
With RS
    .MoveFirst
    While Not .EOF
        If txtUsername.Text = .Fields!UserName And txtPassword.Text = .Fields!Password Then
            .Fields!Password = txtNewPass.Text
            .Update
            connClose
            Clear_Fields
            MsgBox "Password changed successfully!"
            Exit Sub
        Else
            .MoveNext
        End If
    Wend
    txtUsername.BackColor = &HC0C0FF
    txtPassword.BackColor = &HC0C0FF
    txtUsername.SetFocus
End With
connClose
End Sub

Private Sub txtNewPass_Change()
If txtNewPass.BackColor = &HC0C0FF Then txtNewPass.BackColor = &H80000005
End Sub

Private Sub txtNewPass2_Change()
If txtNewPass2.BackColor = &HC0C0FF Then txtNewPass2.BackColor = &H80000005
End Sub

Private Sub txtPassword_Change()
If txtPassword.BackColor = &HC0C0FF Then txtPassword.BackColor = &H80000005
End Sub

Private Sub txtUsername_Change()
If txtUsername.BackColor = &HC0C0FF Then txtUsername.BackColor = &H80000005
End Sub
