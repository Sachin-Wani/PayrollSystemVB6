VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEmpAccounts 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3750
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   6375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEmpAccounts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3240
      Picture         =   "frmEmpAccounts.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Save"
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   495
      Left            =   4800
      Picture         =   "frmEmpAccounts.frx":224E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CheckBox chkUnMask 
      Caption         =   "chkUnMask"
      Height          =   225
      Left            =   3600
      TabIndex        =   9
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox txtPassword2 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4560
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4560
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtUsername 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
   End
   Begin MSComctlLib.ListView Accounts 
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Employee No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Username"
         Object.Width           =   2540
      EndProperty
   End
   Begin PayrollSystem.jcbutton cmdRefresh 
      Height          =   375
      Left            =   5880
      TabIndex        =   14
      Top             =   2640
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
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
      Picture         =   "frmEmpAccounts.frx":4490
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
   End
   Begin VB.Label lblEmpNo 
      Height          =   375
      Left            =   600
      TabIndex        =   13
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Show Password"
      Height          =   255
      Left            =   3960
      TabIndex        =   10
      Top             =   2760
      Width           =   1455
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
      Left            =   3480
      TabIndex        =   7
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label3 
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
      Left            =   3480
      TabIndex        =   4
      Top             =   1560
      Width           =   975
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
      Left            =   3480
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Accounts"
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
      TabIndex        =   1
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "View, Edit, Remove Employee Accounts"
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
   Begin VB.Image cmdExit 
      Height          =   360
      Left            =   6000
      MousePointer    =   99  'Custom
      Picture         =   "frmEmpAccounts.frx":46F0
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmEmpAccounts.frx":4DDA
      Top             =   0
      Width           =   11025
   End
End
Attribute VB_Name = "frmEmpAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Accounts_DblClick()
If Accounts.ListItems.Count <> 0 Then
    cmdSave.Enabled = True
    lblEmpNo.Caption = Accounts.SelectedItem
    SQL = "SELECT * FROM Accounts WHERE EmployeeNo = '" & Trim(Accounts.SelectedItem) & "'"
    connectDB
    With RS
        txtUsername.Text = .Fields!UserName
        txtPassword.Text = .Fields!Password
        txtPassword2.Text = .Fields!Password
    End With
    connClose
End If
End Sub

Private Sub chkUnMask_Click()
If chkUnMask.Value = 0 Then
    txtPassword.PasswordChar = "*"
    txtPassword2.PasswordChar = "*"
Else
    txtPassword.PasswordChar = ""
    txtPassword2.PasswordChar = ""
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdRefresh_Click()
Dim EmpAccounts

Accounts.ListItems.Clear

SQL = "SELECT * FROM Accounts"
connectDB
With RS
    While Not .EOF
        Set EmpAccounts = Accounts.ListItems.Add(, , .Fields!EmployeeNo)
            EmpAccounts.SubItems(1) = .Fields!UserName
        .MoveNext
    Wend
End With
connClose
End Sub

Private Sub cmdSave_Click()

If txtUsername.Text = "" Then
    txtUsername.BackColor = &HC0C0FF
    Exit Sub
End If
If txtPassword.Text = "" Then
    txtPassword.BackColor = &HC0C0FF
    Exit Sub
End If
If txtPassword2.Text = "" Then
    txtPassword2.BackColor = &HC0C0FF
    Exit Sub
End If


If txtPassword2.Text = txtPassword.Text Then
     SQL = "SELECT * FROM Accounts WHERE EmployeeNo = '" & lblEmpNo.Caption & "'"
     connectDB
     With RS
         .Fields!UserName = txtUsername.Text
         .Fields!Password = txtPassword.Text
         .Update
         txtUsername.Text = ""
         txtPassword.Text = ""
         txtPassword2.Text = ""
         lblEmpNo.Caption = ""
     End With
     connClose
Else
    txtPassword2.BackColor = &HC0C0FF
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim EmpAccounts

SQL = "SELECT * FROM Accounts"
connectDB
With RS
    While Not .EOF
        Set EmpAccounts = Accounts.ListItems.Add(, , .Fields!EmployeeNo)
            EmpAccounts.SubItems(1) = .Fields!UserName
        .MoveNext
    Wend
End With
connClose
End Sub

Private Sub txtPassword_Change()
If txtPassword.BackColor = &HC0C0FF Then txtPassword.BackColor = &H80000005
End Sub

Private Sub txtPassword2_Change()
If txtPassword2.BackColor = &HC0C0FF Then txtPassword2.BackColor = &H80000005
End Sub

Private Sub txtUsername_Change()
If txtUsername.BackColor = &HC0C0FF Then txtUsername.BackColor = &H80000005
End Sub
