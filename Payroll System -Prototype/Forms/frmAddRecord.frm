VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmAddRecord 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8745
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   5055
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAddRecord.frx":0000
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
         TabIndex        =   29
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
            ForeColor       =   &H00400000&
            Height          =   375
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   240
            Width           =   2295
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
            TabIndex        =   31
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
         Begin MSACAL.Calendar calBirth 
            Height          =   2175
            Left            =   720
            TabIndex        =   35
            Top             =   2040
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
         Begin MSACAL.Calendar calHired 
            Height          =   2175
            Left            =   720
            TabIndex        =   34
            Top             =   4200
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
         Begin VB.TextBox txtLastName 
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
            Left            =   1680
            TabIndex        =   19
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox txtFirstName 
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
            Left            =   1680
            TabIndex        =   18
            Top             =   720
            Width           =   2655
         End
         Begin VB.TextBox txtAge 
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
            Left            =   3720
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   17
            Top             =   1680
            Width           =   615
         End
         Begin VB.TextBox txtMName 
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
            Left            =   1680
            TabIndex        =   16
            Top             =   1200
            Width           =   2655
         End
         Begin VB.ComboBox cboGender 
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
            ItemData        =   "frmAddRecord.frx":000C
            Left            =   1680
            List            =   "frmAddRecord.frx":0016
            TabIndex        =   15
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtDBirth 
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
            Left            =   1680
            TabIndex        =   14
            Top             =   2160
            Width           =   2655
         End
         Begin VB.TextBox txtPBirth 
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
            Left            =   1680
            TabIndex        =   13
            Top             =   2640
            Width           =   2655
         End
         Begin VB.TextBox txtAddress 
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
            Left            =   1680
            TabIndex        =   12
            Top             =   3120
            Width           =   2655
         End
         Begin VB.TextBox txtContact 
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
            Left            =   1680
            MaxLength       =   11
            TabIndex        =   11
            Top             =   3600
            Width           =   2655
         End
         Begin VB.Frame Frame4 
            Height          =   1815
            Left            =   120
            TabIndex        =   4
            Top             =   4080
            Width           =   4335
            Begin VB.ComboBox cboPosition 
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
               TabIndex        =   7
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
               TabIndex        =   6
               Top             =   1200
               Width           =   2655
            End
            Begin VB.TextBox txtHired 
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
               TabIndex        =   10
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
               TabIndex        =   9
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
               TabIndex        =   8
               Top             =   1200
               Width           =   735
            End
         End
         Begin VB.CommandButton cmdAdd 
            Height          =   495
            Left            =   720
            Picture         =   "frmAddRecord.frx":0028
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   6000
            Width           =   1455
         End
         Begin VB.CommandButton cmdCancel 
            Height          =   495
            Left            =   2520
            Picture         =   "frmAddRecord.frx":226A
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
            TabIndex        =   28
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
            TabIndex        =   27
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
            TabIndex        =   26
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
            TabIndex        =   25
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
            TabIndex        =   24
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
            TabIndex        =   23
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
            TabIndex        =   22
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
            TabIndex        =   21
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
            TabIndex        =   20
            Top             =   3600
            Width           =   1095
         End
      End
   End
   Begin VB.Image cmdExit 
      Height          =   360
      Left            =   4680
      MousePointer    =   99  'Custom
      Picture         =   "frmAddRecord.frx":44AC
      Top             =   0
      Width           =   360
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Add New Employee Record"
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
      TabIndex        =   33
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   " New Record"
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
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   -240
      Picture         =   "frmAddRecord.frx":4B96
      Top             =   0
      Width           =   11025
   End
End
Attribute VB_Name = "frmAddRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Generate_EmpNo()
' Generate Employee Number
SQL = "SELECT * FROM Employees"
connectDB
With RS
    txtEmpNo.Text = Trim("NEPC-" & Trim(Str(Minute(Now) + .RecordCount + (Val(Format(Now, "ssdd")) * Second(Now)))))
End With
connClose
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

Sub Stats()
With cboStatus
    .Clear
    .AddItem "Contractual"
    .AddItem "Regular"
End With
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

Private Sub cmdAdd_Click()
On Error Resume Next
If txtLastName.Text = "" Or txtFirstName.Text = "" Or txtMName.Text = "" Or cboGender.Text = "" Or txtAge.Text = "" Or txtDBirth.Text = "" Or txtPBirth.Text = "" Or txtAddress.Text = "" Or txtContact.Text = "" Or cboPosition.Text = "" Or txtHired.Text = "" Or cboStatus.Text = "" Then
    MsgBox "All fields are required!", vbExclamation + vbOKOnly, "Error"
    Exit Sub
End If

SQL = "SELECT * FROM Employees WHERE EmployeeNo = '" & Trim("NEPC-") & Trim(txtEmpNo.Text) & "'"
connectDB
With RS
    If .RecordCount >= 1 Then
        MsgBox "Employee Already exist!"
        Exit Sub
    Else
        frmNewAccount.txtEmpNo.Text = txtEmpNo.Text
        frmNewAccount.Show
    End If
End With
connClose

'SQL = "SELECT * FROM Employees"
'connectDB
'With RS
'    .MoveLast
'    .AddNew
'    .Fields!EmployeeNo = txtEmpNo.Text
'    .Fields!LastName = txtLastName.Text
'    .Fields!FirstName = txtFirstName.Text
'    .Fields!MiddleName = txtMName.Text
'    .Fields!Gender = cboGender.Text
'    .Fields!Age = txtAge.Text
'    .Fields!Birthday = txtDBirth.Text
'    .Fields!BirthPlace = txtPBirth.Text
'    .Fields!Address = txtAddress.Text
'    .Fields!ContactNo = txtContact.Text
'    .Fields!Position = cboPosition.Text
'    .Fields!DateHired = txtHired.Text
'    .Fields!Status = cboStatus.Text
'    .Update

'    MsgBox "Record added successfully!"
'    ' Close connection
'    connClose
'
'    frmNewAccount.txtEmpNo.Text = txtEmpNo.Text
'    frmNewAccount.Show
'
'    ' Clear all Fields then Generate new Employee Number
'    Clear_TextFields
'    Generate_EmpNo
'
'    ' Refresh Records
'    connectDB
'    .MoveFirst
'    frmMainMenu.EmpRecs
'    connClose
'End With
End Sub

Private Sub cmdCancel_Click()
frmMainMenu.Enabled = True
Unload Me
End Sub

Private Sub cmdExit_Click()
frmMainMenu.Enabled = True
Unload Me
End Sub


Private Sub Form_Load()
frmMainMenu.Enabled = False
calHired.Value = Date
txtHired.Text = Format(Now, "mm/dd/yyyy")
Generate_EmpNo
Display_Positions
Stats
End Sub


Private Sub txtAddress_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "'" Then
    KeyAscii = 0
End If
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = vbKeyBack) Then
    Exit Sub
Else
    KeyAscii = 0
End If
End Sub

Private Sub txtContact_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = vbKeyBack) Then
    Exit Sub
Else
    KeyAscii = 0
End If
End Sub

Private Sub txtDBirth_Click()
calBirth.Visible = True
End Sub

Private Sub txtDBirth_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "'" Then
    KeyAscii = 0
End If
End Sub

Private Sub txtFirstName_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "'" Then
    KeyAscii = 0
End If
End Sub

Private Sub txtHired_Click()
calHired.Visible = True
End Sub

Private Sub txtLastName_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "'" Then
    KeyAscii = 0
End If
End Sub

Private Sub txtMName_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "'" Then
    KeyAscii = 0
End If
End Sub

Private Sub txtPBirth_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "'" Then
    KeyAscii = 0
End If
End Sub
