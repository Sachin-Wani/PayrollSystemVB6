VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmEditRecord 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8760
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   5055
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmEditRecord.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4815
      Begin VB.Frame Frame3 
         Height          =   6615
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   4575
         Begin MSACAL.Calendar calHired 
            Height          =   2175
            Left            =   720
            TabIndex        =   37
            Top             =   4080
            Visible         =   0   'False
            Width           =   3615
            _Version        =   524288
            _ExtentX        =   6376
            _ExtentY        =   3836
            _StockProps     =   1
            BackColor       =   4210688
            Year            =   2013
            Month           =   1
            Day             =   1
            DayLength       =   1
            MonthLength     =   1
            DayFontColor    =   16777152
            FirstDay        =   7
            GridCellEffect  =   0
            GridFontColor   =   12632064
            GridLinesColor  =   8421376
            ShowDateSelectors=   -1  'True
            ShowDays        =   -1  'True
            ShowHorizontalGrid=   -1  'True
            ShowTitle       =   -1  'True
            ShowVerticalGrid=   -1  'True
            TitleFontColor  =   16776960
            ValueIsNull     =   0   'False
            BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Calligraphy"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSACAL.Calendar calBirth 
            Height          =   2175
            Left            =   720
            TabIndex        =   36
            Top             =   2160
            Visible         =   0   'False
            Width           =   3615
            _Version        =   524288
            _ExtentX        =   6376
            _ExtentY        =   3836
            _StockProps     =   1
            BackColor       =   4210688
            Year            =   1980
            Month           =   1
            Day             =   1
            DayLength       =   1
            MonthLength     =   1
            DayFontColor    =   16777152
            FirstDay        =   7
            GridCellEffect  =   0
            GridFontColor   =   12632064
            GridLinesColor  =   8421376
            ShowDateSelectors=   -1  'True
            ShowDays        =   -1  'True
            ShowHorizontalGrid=   -1  'True
            ShowTitle       =   -1  'True
            ShowVerticalGrid=   -1  'True
            TitleFontColor  =   16776960
            ValueIsNull     =   0   'False
            BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Calligraphy"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.CommandButton cmdSave 
            Enabled         =   0   'False
            Height          =   495
            Left            =   840
            Picture         =   "frmEditRecord.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   6000
            Width           =   1455
         End
         Begin VB.CommandButton cmdCancel 
            Height          =   495
            Left            =   2520
            Picture         =   "frmEditRecord.frx":224E
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   6000
            Width           =   1455
         End
         Begin VB.Frame Frame4 
            Height          =   1815
            Left            =   120
            TabIndex        =   14
            Top             =   4080
            Width           =   4335
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
               TabIndex        =   17
               Top             =   720
               Width           =   2655
            End
            Begin VB.ComboBox cboStatus 
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
               Height          =   390
               Left            =   1560
               TabIndex        =   16
               Top             =   1200
               Width           =   2655
            End
            Begin VB.ComboBox cboPosition 
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
               Height          =   390
               Left            =   1560
               TabIndex        =   15
               Top             =   240
               Width           =   2655
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
               TabIndex        =   20
               Top             =   1200
               Width           =   735
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
               TabIndex        =   19
               Top             =   720
               Width           =   1095
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
               TabIndex        =   18
               Top             =   240
               Width           =   855
            End
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
            MaxLength       =   11
            TabIndex        =   13
            Top             =   3600
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
            TabIndex        =   12
            Top             =   3120
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
            TabIndex        =   11
            Top             =   2640
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
            TabIndex        =   10
            Top             =   2160
            Width           =   2655
         End
         Begin VB.ComboBox cboGender 
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
            Height          =   390
            ItemData        =   "frmEditRecord.frx":4490
            Left            =   1680
            List            =   "frmEditRecord.frx":449A
            TabIndex        =   9
            Top             =   1680
            Width           =   1335
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
            TabIndex        =   8
            Top             =   1200
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
            MaxLength       =   2
            TabIndex        =   7
            Top             =   1680
            Width           =   615
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
            TabIndex        =   6
            Top             =   720
            Width           =   2655
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
            TabIndex        =   5
            Top             =   240
            Width           =   2655
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
            TabIndex        =   30
            Top             =   3600
            Width           =   1095
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
            TabIndex        =   29
            Top             =   3120
            Width           =   855
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
            TabIndex        =   28
            Top             =   2640
            Width           =   1335
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
            TabIndex        =   27
            Top             =   2160
            Width           =   1335
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
            TabIndex        =   26
            Top             =   1680
            Width           =   495
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
            TabIndex        =   25
            Top             =   1680
            Width           =   855
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
            TabIndex        =   23
            Top             =   720
            Width           =   1095
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
            TabIndex        =   22
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   4575
         Begin VB.CommandButton cmdSearch 
            Height          =   375
            Left            =   3960
            Picture         =   "frmEditRecord.frx":44AC
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   240
            Width           =   375
         End
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
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   2
            Top             =   240
            Width           =   1095
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
            Left            =   1920
            TabIndex        =   35
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
            TabIndex        =   3
            Top             =   240
            Width           =   1335
         End
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Record"
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
      TabIndex        =   32
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Update Employee Record"
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
      TabIndex        =   31
      Top             =   600
      Width           =   3135
   End
   Begin VB.Image cmdExit 
      Height          =   360
      Left            =   4680
      MousePointer    =   99  'Custom
      Picture         =   "frmEditRecord.frx":4AAE
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   -120
      Picture         =   "frmEditRecord.frx":5198
      Top             =   0
      Width           =   11025
   End
End
Attribute VB_Name = "frmEditRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Enable_TextFields()
txtLastName.Enabled = True
txtFirstName.Enabled = True
txtMName.Enabled = True
cboGender.Enabled = True
txtAge.Enabled = True
txtDBirth.Enabled = True
txtPBirth.Enabled = True
txtAddress.Enabled = True
txtContact.Enabled = True
cboPosition.Enabled = True
txtHired.Enabled = True
cboStatus.Enabled = True
cmdSave.Enabled = True
End Sub

Sub Disable_TextFields()
txtLastName.Enabled = False
txtFirstName.Enabled = False
txtMName.Enabled = False
cboGender.Enabled = False
txtAge.Enabled = False
txtDBirth.Enabled = False
txtPBirth.Enabled = False
txtAddress.Enabled = False
txtContact.Enabled = False
cboPosition.Enabled = False
txtHired.Enabled = False
cboStatus.Enabled = False
cmdSave.Enabled = False
End Sub

Sub Clear_TextFields()
txtLastName.Text = ""
txtFirstName.Text = ""
txtMName.Text = ""
cboGender.Text = ""
txtAge.Text = ""
txtDBirth.Text = ""
txtPBirth.Text = ""
txtAddress.Text = ""
txtContact.Text = ""
cboPosition.Text = ""
txtHired.Text = ""
cboStatus.Text = ""
End Sub

Sub Display_Positions()
cboPosition.Clear

SQL = "SELECT * FROM Positions"
connectDB
With RS
    .MoveFirst
    While Not .EOF
        cboPosition.AddItem .Fields!Designation
        .MoveNext
    Wend
End With
connClose
End Sub

Sub Stats()
With cboStatus
    .Clear
    .AddItem "Contractual"
    .AddItem "Regular"
End With
End Sub



Private Sub calBirth_Click()
calBirth.Visible = False
txtDBirth.Text = calBirth.Value
txtAge.Text = Year(Now) - Year(calBirth.Value)
End Sub

Private Sub calHired_Click()
calHired.Visible = False
txtHired.Text = calHired.Value
End Sub

Private Sub cboGender_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cboPosition_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
frmMainMenu.Enabled = True
Unload Me
End Sub

Private Sub cmdExit_Click()
Call cmdCancel_Click
End Sub



Private Sub cmdSave_Click()
On Error Resume Next
If txtEmpNo.Text <> "" Then
    SQL = "SELECT * FROM Employees"
    connectDB
    With RS
        .MoveFirst
        While Not .EOF
            If "NEPC-" & Trim(txtEmpNo.Text) = .Fields!EmployeeNo Then
                .Fields!LastName = txtLastName.Text
                .Fields!FirstName = txtFirstName.Text
                .Fields!MiddleName = txtMName.Text
                .Fields!Gender = cboGender.Text
                .Fields!Age = txtAge.Text
                .Fields!Birthday = txtDBirth.Text
                .Fields!BirthPlace = txtPBirth.Text
                .Fields!Address = txtAddress.Text
                .Fields!ContactNo = txtContact.Text
                .Fields!Position = cboPosition.Text
                .Fields!DateHired = txtHired.Text
                .Fields!Status = cboStatus.Text
                .Update
                connClose
                txtEmpNo.BackColor = &HC0FFC0
                txtEmpNo.ToolTipText = "Employee Record updated successfully!"
                'MsgBox "Employee's Record updated!", vbInformation + vbOKOnly, "Update Record"
                Clear_TextFields
                
                ' Refresh all Records
                connOpen
                frmMainMenu.EmpRecs
                connClose
                Exit Sub
            Else
                .MoveNext
            End If
        Wend
    End With
    connClose
End If
End Sub

Private Sub cmdSearch_Click()
SQL = "SELECT * FROM Employees"
connectDB
With RS
    .MoveFirst
    While Not .EOF
        If ("NEPC-" & Trim(txtEmpNo.Text)) = .Fields!EmployeeNo Then
            txtLastName.Text = .Fields!LastName
            txtFirstName.Text = .Fields!FirstName
            txtMName.Text = .Fields!MiddleName
            cboGender.Text = .Fields!Gender
            txtAge.Text = .Fields!Age
            txtDBirth.Text = .Fields!Birthday
            txtPBirth.Text = .Fields!BirthPlace
            txtAddress.Text = .Fields!Address
            txtContact.Text = .Fields!ContactNo
            cboPosition.Text = .Fields!Position
            txtHired.Text = .Fields!DateHired
            cboStatus.Text = .Fields!Status
            
            
            Enable_TextFields
            connClose
            Exit Sub
        Else
            .MoveNext
        End If
    Wend
    txtEmpNo.BackColor = &HC0C0FF
    txtEmpNo.ToolTipText = "Employee not found!"
    'MsgBox "Employee's Record not Found!", vbExclamation + vbOKOnly, "Record Error"
End With
connClose
End Sub

Private Sub Form_Load()
On Error Resume Next
frmMainMenu.Enabled = False
calHired.Value = Date
Display_Positions
Stats
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = vbKeyBack) Then
    Exit Sub
Else
    KeyAscii = 0
End If
End Sub

Private Sub txtDBirth_Click()
calBirth.Visible = True
End Sub

Private Sub txtEmpNo_Change()
If txtEmpNo.BackColor = &HC0C0FF Then txtEmpNo.BackColor = &H80000005
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

Private Sub txtHired_Click()
calHired.Visible = True
End Sub
