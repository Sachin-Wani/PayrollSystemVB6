VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSearch 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7410
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   7860
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   7695
      Begin VB.Frame Frame3 
         Height          =   5535
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   5535
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
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   23
            Top             =   1680
            Width           =   1215
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
            Left            =   1800
            TabIndex        =   22
            Top             =   2640
            Width           =   2775
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
            Left            =   1800
            TabIndex        =   21
            Top             =   3120
            Width           =   2775
         End
         Begin VB.TextBox txtBirthday 
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
            Left            =   1800
            TabIndex        =   20
            Top             =   2160
            Width           =   2775
         End
         Begin VB.Frame Frame4 
            Height          =   1815
            Left            =   120
            TabIndex        =   12
            Top             =   3600
            Width           =   5295
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
               Left            =   1680
               TabIndex        =   15
               Top             =   1200
               Width           =   1455
            End
            Begin VB.TextBox txtDesignation 
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
               TabIndex        =   14
               Top             =   240
               Width           =   2775
            End
            Begin VB.TextBox txtDateHired 
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
               Left            =   1680
               TabIndex        =   13
               Top             =   720
               Width           =   2775
            End
            Begin MSComCtl2.DTPicker DTPDateHired 
               Height          =   375
               Left            =   4200
               TabIndex        =   16
               Top             =   720
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   661
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   110821377
               CurrentDate     =   41181
            End
            Begin VB.Label Label9 
               Caption         =   "Position:"
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
               Height          =   255
               Left            =   720
               TabIndex        =   19
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label12 
               Caption         =   "Date Hired:"
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
               Height          =   255
               Left            =   600
               TabIndex        =   18
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label8 
               Caption         =   "Status:"
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
               Height          =   255
               Left            =   960
               TabIndex        =   17
               Top             =   1200
               Width           =   735
            End
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
            Left            =   1800
            TabIndex        =   11
            Top             =   240
            Width           =   2775
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
            Left            =   1800
            TabIndex        =   10
            Top             =   720
            Width           =   2775
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
            Left            =   3840
            MaxLength       =   2
            TabIndex        =   9
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox txtMI 
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
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   8
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label15 
            Caption         =   "Last Name:"
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
            Height          =   255
            Left            =   720
            TabIndex        =   31
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label14 
            Caption         =   "First Name:"
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
            Left            =   720
            TabIndex        =   30
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "MI:"
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
            Left            =   1440
            TabIndex        =   29
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label Label6 
            Caption         =   "Age:"
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
            Left            =   3360
            TabIndex        =   28
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label Label7 
            Caption         =   "Gender:"
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
            Left            =   960
            TabIndex        =   27
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "Address:"
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
            Height          =   255
            Left            =   960
            TabIndex        =   26
            Top             =   2640
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "Contact No.:"
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
            Height          =   255
            Left            =   600
            TabIndex        =   25
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label Label13 
            Caption         =   "Date of Birth:"
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
            Height          =   255
            Left            =   480
            TabIndex        =   24
            Top             =   2160
            Width           =   1335
         End
      End
      Begin VB.TextBox txtLName 
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
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   3360
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton cmdPrint 
         Enabled         =   0   'False
         Height          =   495
         Left            =   6000
         Picture         =   "frmSearch.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Print"
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtEmpNo 
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
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   4320
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel 
         Height          =   495
         Left            =   6000
         Picture         =   "frmSearch.frx":224E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5760
         Width           =   1455
      End
      Begin VB.CommandButton cmdSearch 
         Height          =   495
         Left            =   6000
         Picture         =   "frmSearch.frx":4490
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Search Record"
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cmbCategories 
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
         Height          =   345
         ItemData        =   "frmSearch.frx":66D2
         Left            =   1440
         List            =   "frmSearch.frx":66DC
         TabIndex        =   1
         Text            =   "Employee No."
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label NEPC 
         Caption         =   "NEPC"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   36
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Dash 
         Caption         =   "-"
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
         Left            =   4200
         TabIndex        =   33
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "Categories:"
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
         Left            =   360
         TabIndex        =   32
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "Search Record"
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
      TabIndex        =   35
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Employee Record"
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
      TabIndex        =   34
      Top             =   600
      Width           =   3135
   End
   Begin VB.Image cmdExit 
      Height          =   360
      Left            =   7440
      MousePointer    =   99  'Custom
      Picture         =   "frmSearch.frx":6701
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmSearch.frx":6DEB
      Top             =   0
      Width           =   11025
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Choice As String

Sub Categ(a, b)
txtEmpNo.Visible = a
NEPC.Visible = a
Dash.Visible = a
txtLName.Visible = b
End Sub

Private Sub cmbCategories_Click()
Select Case cmbCategories.Text
Case "Employee No.":
    Choice = "EmployeeNo"
    txtEmpNo.Visible = True
    NEPC.Visible = True
    Dash.Visible = True
    txtLName.Visible = False
Case "Employee LastName":
    Choice = "EmployeeName"
    txtEmpNo.Visible = False
    NEPC.Visible = False
    Dash.Visible = False
    txtLName.Visible = True
End Select
End Sub

Private Sub cmdCancel_Click()
frmMainMenu.Enabled = True
Unload Me
End Sub

Private Sub cmdExit_Click()
frmMainMenu.Enabled = True
Unload Me
End Sub

Private Sub cmdPrint_Click()
If Choice = "EmployeeNo" Then
    SQL = "SELECT * FROM Employees WHERE EmployeeNo = 'NEPC-" & Trim(txtEmpNo.Text) & "'"
ElseIf Choice = "EmployeeName" Then
    SQL = "SELECT * FROM Employees WHERE LastName = '" & txtLName.Text & "'"
End If

connectDB
With RS
    If .RecordCount <> 0 Then
        Set rptEmpInfo.DataSource = RS
        rptEmpInfo.Show
    End If
End With
connClose
End Sub

Private Sub cmdSearch_Click()
SQL = "SELECT * FROM Employees"

If Choice = "EmployeeNo" Then
    connectDB
    With RS
        While Not .EOF
            If Trim("NEPC-") + Trim(txtEmpNo.Text) = .Fields!EmployeeNo Then
                txtLastName.Text = .Fields!LastName
                txtFirstName.Text = .Fields!FirstName
                txtMI.Text = .Fields!MiddleName
                txtGender.Text = .Fields!Gender
                txtAge.Text = .Fields!Age
                txtBirthday.Text = .Fields!Birthday
                txtAddress.Text = .Fields!Address
                txtContact.Text = .Fields!ContactNo
                txtDesignation.Text = .Fields!Position
                txtDateHired.Text = .Fields!DateHired
                txtStatus.Text = .Fields!Status
                connClose
                cmdPrint.Enabled = True
                Exit Sub
            Else
                .MoveNext
            End If
        Wend
        txtEmpNo.BackColor = &HC0C0FF
        txtEmpNo.SetFocus
    End With
    connClose
ElseIf Choice = "EmployeeName" Then
    connectDB
    With RS
        While Not .EOF
            If txtLName.Text = .Fields!LastName Then
                txtLastName.Text = .Fields!LastName
                txtFirstName.Text = .Fields!FirstName
                txtMI.Text = .Fields!MiddleName
                txtGender.Text = .Fields!Gender
                txtAge.Text = .Fields!Age
                txtBirthday.Text = .Fields!Birthday
                txtAddress.Text = .Fields!Address
                txtContact.Text = .Fields!ContactNo
                txtDesignation.Text = .Fields!Position
                txtDateHired.Text = .Fields!DateHired
                txtStatus.Text = .Fields!Status
                connClose
                Exit Sub
            Else
                .MoveNext
            End If
        Wend
        txtLName.BackColor = &HC0C0FF
        txtLName.SetFocus
    End With
    connClose
End If


End Sub

Private Sub Form_Load()
frmMainMenu.Enabled = False
Choice = "EmployeeNo"
End Sub

Private Sub txtEmpNo_Change()
cmdPrint.Enabled = False
txtLastName.Text = ""
txtFirstName.Text = ""
txtMI.Text = ""
txtGender.Text = ""
txtAge.Text = ""
txtBirthday.Text = ""
txtAddress.Text = ""
txtContact.Text = ""
txtDesignation.Text = ""
txtDateHired.Text = ""
txtStatus.Text = ""
End Sub

Private Sub txtLName_Change()
cmdPrint.Enabled = False
txtLastName.Text = ""
txtFirstName.Text = ""
txtMI.Text = ""
txtGender.Text = ""
txtAge.Text = ""
txtBirthday.Text = ""
txtAddress.Text = ""
txtContact.Text = ""
txtDesignation.Text = ""
txtDateHired.Text = ""
txtStatus.Text = ""
End Sub

Private Sub txtLName_KeyPress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then
    KeyAscii = 0
End If
End Sub
