VERSION 5.00
Begin VB.Form frmDeactivateRecord 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8745
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   5055
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmDeleteRecord.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4815
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Width           =   4575
         Begin VB.TextBox txtEmpNo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   35
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdSearch 
            Height          =   375
            Left            =   4080
            Picture         =   "frmDeleteRecord.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblEmpNo 
            Caption         =   "NEPC -"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2040
            TabIndex        =   36
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Employee No:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   28
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Height          =   6615
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   4575
         Begin VB.TextBox txtGender 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtLastName 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox txtFirstName 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   720
            Width           =   2655
         End
         Begin VB.TextBox txtAge 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   1680
            Width           =   615
         End
         Begin VB.TextBox txtMName 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   1200
            Width           =   2655
         End
         Begin VB.TextBox txtDBirth 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   2160
            Width           =   2655
         End
         Begin VB.TextBox txtPBirth 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   2640
            Width           =   2655
         End
         Begin VB.TextBox txtAddress 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   3120
            Width           =   2655
         End
         Begin VB.TextBox txtContact 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   3600
            Width           =   2655
         End
         Begin VB.Frame Frame4 
            Height          =   1815
            Left            =   120
            TabIndex        =   4
            Top             =   4080
            Width           =   4335
            Begin VB.TextBox txtStatus 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
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
               Locked          =   -1  'True
               TabIndex        =   34
               Top             =   1200
               Width           =   2655
            End
            Begin VB.TextBox txtPosition 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   33
               Top             =   240
               Width           =   2655
            End
            Begin VB.ComboBox cboStatus 
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
               Height          =   390
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   6
               Top             =   1200
               Width           =   2655
            End
            Begin VB.TextBox txtHired 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
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
               Locked          =   -1  'True
               TabIndex        =   5
               Top             =   720
               Width           =   2655
            End
            Begin VB.Label Label12 
               Caption         =   "Position:"
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
               TabIndex        =   9
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label13 
               Caption         =   "Date Hired:"
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
               TabIndex        =   8
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label15 
               Caption         =   "Status:"
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
               Left            =   720
               TabIndex        =   7
               Top             =   1200
               Width           =   735
            End
         End
         Begin VB.CommandButton cmdDeactivate 
            Enabled         =   0   'False
            Height          =   495
            Left            =   600
            Picture         =   "frmDeleteRecord.frx":060E
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   6000
            Width           =   1815
         End
         Begin VB.CommandButton cmdCancel 
            Height          =   495
            Left            =   2520
            Picture         =   "frmDeleteRecord.frx":2978
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   6000
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Last Name:"
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
            TabIndex        =   26
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "First Name:"
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
            TabIndex        =   25
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Middle Name:"
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
            Left            =   240
            TabIndex        =   24
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Gender:"
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
            Left            =   720
            TabIndex        =   23
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Age:"
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
            Left            =   3240
            TabIndex        =   22
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label Label8 
            Caption         =   "Date of Birth:"
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
            TabIndex        =   21
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "Place of Birth:"
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
            Left            =   240
            TabIndex        =   20
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Address:"
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
            Left            =   720
            TabIndex        =   19
            Top             =   3120
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "Contact No:"
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
            TabIndex        =   18
            Top             =   3600
            Width           =   1095
         End
      End
   End
   Begin VB.Image cmdExit 
      Height          =   360
      Left            =   4680
      MousePointer    =   99  'Custom
      Picture         =   "frmDeleteRecord.frx":4BBA
      Top             =   0
      Width           =   360
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Deactivate Employee"
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
      TabIndex        =   30
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "Deactivate Record"
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
      TabIndex        =   29
      Top             =   0
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmDeleteRecord.frx":52A4
      Top             =   0
      Width           =   11025
   End
End
Attribute VB_Name = "frmDeactivateRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Disable_TextFields()
txtLastName.Enabled = False
txtFirstName.Enabled = False
txtMName.Enabled = False
txtGender.Enabled = False
txtAge.Enabled = False
txtDBirth.Enabled = False
txtPBirth.Enabled = False
txtAddress.Enabled = False
txtContact.Enabled = False
txtPosition.Enabled = False
txtHired.Enabled = False
txtStatus.Enabled = False
End Sub

Sub Clear_TextFields()
txtLastName.Text = ""
txtFirstName.Text = ""
txtMName.Text = ""
txtGender.Text = ""
txtAge.Text = ""
txtDBirth.Text = ""
txtPBirth.Text = ""
txtAddress.Text = ""
txtContact.Text = ""
txtPosition.Text = ""
txtHired.Text = ""
txtStatus.Text = ""
End Sub

Private Sub cmdCancel_Click()
frmMainMenu.Enabled = True
Unload Me
End Sub

Private Sub cmdDeactivate_Click()
On Error Resume Next
If txtEmpNo.Text <> "" Then
    SQL = "SELECT * FROM Employees WHERE EmployeeNo = 'NEPC-" & Trim(txtEmpNo.Text) & "'"
    connectDB
    With RS
        If .RecordCount <> 0 Then
            If .Fields!Position = "Administrator" Then
                txtEmpNo.BackColor = &HC0C0FF
                txtEmpNo.ToolTipText = "You cannot Deactivate the Administrator!"
                Clear_TextFields
                Disable_TextFields
                cmdDelete.Enabled = False
                connClose
                Exit Sub
            Else
                .Fields!CurrentStatus = "Inactive"
                .Update
                Clear_TextFields
                Disable_TextFields
                cmdDelete.Enabled = False
                txtEmpNo.BackColor = &HC0FFC0
                txtEmpNo.ToolTipText = "Employee Deactivated successfully!"
            End If
        End If
    End With
    connClose

'    SQL = "DELETE * FROM Employees WHERE EmployeeNo = '" & ("NEPC-" & Trim(txtEmpNo.Text)) & "'"
'    connectDB
'    Clear_TextFields
'    Disable_TextFields
'    cmdDelete.Enabled = False
'    connClose
    
'    SQL = "DELETE * FROM Accounts WHERE Employeeno = '" & ("NEPC-" & Trim(txtEmpNo.Text)) & "'"
'    connectDB
'    connClose
    
    SQL = "SELECT * FROM Employees"
    connOpen
    frmMainMenu.EmpRecs
    connClose
    
    frmMainMenu.PayrollDetails

Else
    txtEmpNo.BackColor = &HC0C0FF
    txtEmpNo.SetFocus
End If
End Sub

Private Sub cmdExit_Click()
frmMainMenu.Enabled = True
Unload Me
End Sub

Private Sub cmdSearch_Click()
SQL = "SELECT * FROM Employees WHERE CurrentStatus = 'Active'"
connectDB
With RS
    .MoveFirst
    While Not .EOF
        If ("NEPC-" & Trim(txtEmpNo.Text)) = .Fields!EmployeeNo Then
            txtLastName.Text = .Fields!LastName
            txtFirstName.Text = .Fields!FirstName
            txtMName.Text = .Fields!MiddleName
            txtGender.Text = .Fields!Gender
            txtAge.Text = .Fields!Age
            txtDBirth.Text = .Fields!Birthday
            txtPBirth.Text = .Fields!BirthPlace
            txtAddress.Text = .Fields!Address
            txtContact.Text = .Fields!ContactNo
            txtPosition.Text = .Fields!Position
            txtHired.Text = .Fields!DateHired
            txtStatus.Text = .Fields!Status
            cmdDeactivate.Enabled = True
            connClose
            Exit Sub
        Else
            .MoveNext
        End If
    Wend
    txtEmpNo.BackColor = &HC0C0FF
    txtEmpNo.ToolTipText = "Employee not Found!"
End With
connClose
End Sub

Private Sub Form_Load()
frmMainMenu.Enabled = False
End Sub

Private Sub txtEmpNo_Change()
If txtEmpNo.BackColor = &HC0C0FF Or txtEmpNo.BackColor = &HC0FFC0 Then
    txtEmpNo.BackColor = &H80000005
    txtEmpNo.ToolTipText = ""
End If
Disable_TextFields
Clear_TextFields
End Sub

Private Sub txtEmpNo_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = vbKeyBack) Then
    Exit Sub
Else
    KeyAscii = 0
End If
End Sub
