VERSION 5.00
Begin VB.Form frmEmpNewAccount 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4815
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   4095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmEmpNewAccount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3855
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   3615
         Begin VB.TextBox txtEmpNo 
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
            Left            =   1560
            TabIndex        =   11
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Employee No:"
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
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2655
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   3615
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
            Left            =   1560
            TabIndex        =   4
            Top             =   240
            Width           =   1815
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
            Left            =   1560
            PasswordChar    =   "*"
            TabIndex        =   3
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox txtConfirm 
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
            Left            =   1560
            PasswordChar    =   "*"
            TabIndex        =   2
            Top             =   1320
            Width           =   1815
         End
         Begin PayrollSystem.jcbutton cmdCancel 
            Height          =   495
            Left            =   1920
            TabIndex        =   5
            Top             =   1920
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
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
            Caption         =   "Cancel"
            UseMaskCOlor    =   -1  'True
         End
         Begin PayrollSystem.jcbutton cmdRegister 
            Height          =   495
            Left            =   480
            TabIndex        =   6
            Top             =   1920
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
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
            Caption         =   "Register"
            UseMaskCOlor    =   -1  'True
         End
         Begin VB.Label Label3 
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
            Height          =   375
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
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
            Height          =   495
            Left            =   240
            TabIndex        =   8
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Password:"
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
            Left            =   360
            TabIndex        =   7
            Top             =   720
            Width           =   975
         End
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "Registration"
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
      Left            =   240
      TabIndex        =   14
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Create New Employee Account"
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
      TabIndex        =   13
      Top             =   600
      Width           =   3135
   End
   Begin VB.Image cmdExit 
      Height          =   360
      Left            =   3720
      MousePointer    =   99  'Custom
      Picture         =   "frmEmpNewAccount.frx":000C
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmEmpNewAccount.frx":06F6
      Top             =   0
      Width           =   11025
   End
End
Attribute VB_Name = "frmEmpNewAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim stat As Integer
Dim EmpPosition As String


Sub Check_Fields()
If txtEmpNo.Text = "" Or txtUsername.Text = "" Or txtPassword.Text = "" Or txtConfirm.Text = "" Then
    If txtEmpNo.Text = "" Then txtEmpNo.BackColor = &HC0C0FF
    If txtUsername.Text = "" Then txtUsername.BackColor = &HC0C0FF
    If txtPassword.Text = "" Then txtPassword.BackColor = &HC0C0FF
    If txtConfirm.Text = "" Then txtConfirm.BackColor = &HC0C0FF
    Exit Sub
    Exit Sub
End If
End Sub

Sub Match_Password()
If txtPassword.Text <> txtConfirm.Text Then
    txtPassword.BackColor = &HC0C0FF
    txtConfirm.BackColor = &HC0C0FF
    txtPassword.SetFocus
    Exit Sub
    Exit Sub
End If
End Sub

Sub Verify_EmpNo()
SQL = "SELECT * FROM Employees"
connectDB
With RS
    Do While Not .EOF
        If txtEmpNo.Text = .Fields!EmployeeNo Then
            EmpPosition = .Fields!Position
            stat = 1
            Exit Do
        Else
            .MoveNext
        End If
    Loop
    If stat = 0 Then
        txtEmpNo.BackColor = &HC0C0FF
        txtEmpNo.SetFocus
    End If
End With
connClose
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdRegister_Click()
On Error Resume Next
Check_Fields ' All Fields are Required
Match_Password ' Security Check

'==========================================

'If stat = 0 Then Exit Sub
' Passed? then Add New Account
If txtUsername.Text <> "" Or txtPassword.Text <> "" Then
    Verify_EmpNo ' Check if Employee Number is Valid
    
    
    ' ======== Add New Account ================================

    SQL = "SELECT * FROM Accounts"
    connectDB
    With RS
        .MoveLast
        .AddNew
            .Fields!EmployeeNo = txtEmpNo.Text
            .Fields!UserName = txtUsername.Text
            .Fields!Password = txtPassword.Text
            .Fields!Position = EmpPosition
        .Update
        MsgBox "Account added successfully!"
    End With
    connClose
End If
End Sub

Private Sub txtConfirm_Change()
txtConfirm.BackColor = &H80000005
End Sub

Private Sub txtEmpNo_Change()
txtEmpNo.BackColor = &H80000005
End Sub

Private Sub txtPassword_Change()
txtPassword.BackColor = &H80000005
End Sub

Private Sub txtUsername_Change()
txtUsername.BackColor = &H80000005
End Sub

