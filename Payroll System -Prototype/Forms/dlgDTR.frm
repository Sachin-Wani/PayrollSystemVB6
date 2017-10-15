VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form dlgDTR 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5655
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   5520
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "dlgDTR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTimeDate 
      Interval        =   50
      Left            =   4320
      Top             =   10
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3000
      Picture         =   "dlgDTR.frx":000C
      ScaleHeight     =   495
      ScaleWidth      =   2295
      TabIndex        =   13
      Top             =   1080
      Width           =   2295
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Current Date:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      Picture         =   "dlgDTR.frx":1F490
      ScaleHeight     =   495
      ScaleWidth      =   2295
      TabIndex        =   11
      Top             =   1080
      Width           =   2295
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Current Time:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   120
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView DTR 
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Employee No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Time-in"
         Object.Width           =   2118
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Time-out"
         Object.Width           =   2118
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Status"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   5295
      Begin PayrollSystem.jcbutton cmdRefresh 
         Height          =   375
         Left            =   4200
         TabIndex        =   21
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "Refresh"
         Picture         =   "dlgDTR.frx":3E914
         UseMaskCOlor    =   -1  'True
         MaskColor       =   16777215
      End
      Begin PayrollSystem.jcbutton cmdCancel 
         Height          =   375
         Left            =   4200
         TabIndex        =   10
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14935011
         Caption         =   "&Cancel"
         Picture         =   "dlgDTR.frx":3EB74
         PictureAlign    =   0
         UseMaskCOlor    =   -1  'True
         MaskColor       =   16777215
         CaptionAlign    =   2
      End
      Begin PayrollSystem.jcbutton cmdOut 
         Height          =   375
         Left            =   2760
         TabIndex        =   9
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14935011
         Caption         =   "Time &out"
         Picture         =   "dlgDTR.frx":3EDF0
         UseMaskCOlor    =   -1  'True
         MaskColor       =   16777215
      End
      Begin PayrollSystem.jcbutton cmdIn 
         Height          =   375
         Left            =   2760
         TabIndex        =   8
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14935011
         Caption         =   "Time &in"
         Picture         =   "dlgDTR.frx":3F217
         UseMaskCOlor    =   -1  'True
         MaskColor       =   16777215
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
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   1575
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
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   720
         Width           =   1575
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
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   975
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
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label lblUserLoc 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   6840
      TabIndex        =   22
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "You can leave now NAME"
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
      TabIndex        =   20
      Top             =   5640
      Width           =   5175
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "00/00/0000"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   18
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackColor       =   &H00808080&
      Height          =   1215
      Left            =   2880
      TabIndex        =   16
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808080&
      Height          =   1215
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   2535
   End
   Begin VB.Image cmdExit 
      Height          =   360
      Left            =   5160
      MousePointer    =   99  'Custom
      Picture         =   "dlgDTR.frx":3F63E
      Top             =   0
      Width           =   360
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Time-in and Time-out"
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
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "Daily Time Record"
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
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "dlgDTR.frx":3FD28
      Top             =   0
      Width           =   11025
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      Height          =   1455
      Left            =   0
      TabIndex        =   19
      Top             =   840
      Width           =   5655
   End
End
Attribute VB_Name = "dlgDTR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim usrloc As String

Dim EmpNo As String
Dim stat As String

' Time in
Dim intHH_in As Integer
Dim intMM_in As Integer
Dim intSS_in As Integer

'Time out
Dim intHH_out As Integer
Dim intMM_out As Integer
Dim intSS_out As Integer

' Date
Dim intMonth As Integer
Dim intDay As Integer
Dim intYear As Integer

' Computation
Dim intCompute_HH As Integer
Dim intCompute_MM As Integer
Dim intCompute_SS As Integer

Private Function TimeDiff(Time1 As String, Time2 As String) As String
Dim Time_Hours As String
Dim Time_Minutes As String
Dim Time_HR As String

Time_HR = DateDiff("n", Time1, Time2)
Time_HR = IIf(Time_HR < 0, Time_HR + 1440, Time_HR)
Time_Hours = Format(Int(Time_HR / 60), "00")
Time_Minutes = Format(Int(Time_HR Mod 60), "00")
TimeDiff = Format(Time_Hours & ":" & Time_Minutes, "hh:mm")
End Function

Sub Check_Fields()
If txtUsername.Text = "" Or txtPassword.Text = "" Then
    If txtUsername.Text = "" Then
        txtUsername.BackColor = &HC0C0FF
        txtUsername.SetFocus
    End If

    If txtPassword.Text = "" Then
        txtPassword.BackColor = &HC0C0FF
        txtPassword.SetFocus
    End If
    Exit Sub
    Exit Sub
End If
End Sub

Sub Verify_Account()
stat = 0
SQL = "SELECT * FROM Accounts"
connOpen
With RS
    .MoveFirst
    While Not .EOF
        If txtUsername.Text = .Fields!UserName And txtPassword.Text = .Fields!Password Then
            EmpNo = .Fields!EmployeeNo
            connClose
            stat = 1
            Exit Sub
        Else
            .MoveNext
        End If
    Wend
    MsgBox "Invalid Username and Password!"
End With
connClose
End Sub


Sub ConnectDTR()
Dim recs

SQL = "SELECT * FROM DTR"
connectDB

DTR.ListItems.Clear

With RS
    While Not .EOF
        If .Fields!WorkDate = Format(Now, "mm/dd/yyyy") Then
            Set recs = DTR.ListItems.Add(, , .Fields!EmployeeNo)
                recs.SubItems(1) = .Fields!TimeIn
                recs.SubItems(2) = .Fields!TimeOut
                recs.SubItems(3) = .Fields!Status
                .MoveNext
        Else
            .MoveNext
        End If
    Wend
End With
connClose
End Sub

Sub EmpDTR_in()
SQL = "SELECT * FROM DTR"
connectDB
With RS
    While Not .EOF
        If EmpNo = .Fields!EmployeeNo And Format(Now, "mm/dd/yyyy") = .Fields!WorkDate Then
            If .Fields!Status = "In" Then
                connClose
                MsgBox "Already logged in!"
                Exit Sub
            Else
                connClose
                MsgBox "Can't Log-in again!"
                Exit Sub
            End If
        Else
            .MoveNext
        End If
    Wend
    
    txtUsername.Text = ""
    txtPassword.Text = ""
    
    ' Time in na :D
    .AddNew
    .Fields!EmployeeNo = EmpNo
    .Fields!WorkDate = Format(Now, "mm/dd/yyyy")
    .Fields!TimeIn = Format(Now, "hh:mm AM/PM")
    .Fields!TimeOut = "---"
    .Fields!WorkingHours = "---"
    .Fields!OverTime = "---"
    .Fields!Late = "---"
    .Fields!Status = "In"
    .Update
    dlgDTR.Height = 6150
    lblStatus.Caption = "Welcome " + EmpNo + "!"
    txtUsername.Text = ""
    txtPassword.Text = ""
End With
connClose
ConnectDTR
End Sub

Sub EmpDTR_out()
Dim HH As Integer
Dim MM As Integer
Dim SS As Integer
Dim EmpOT

SQL = "SELECT * FROM DTR"
connectDB
With RS
    While Not .EOF
        If EmpNo = .Fields!EmployeeNo And .Fields!WorkDate = Format(Now, "mm/dd/yyyy") And .Fields!Status = "In" Then
            txtUsername.Text = ""
            txtPassword.Text = ""
            .Fields!TimeOut = Format(Now, "hh:mm AM/PM")
            'HH = Hour(Now) - Hour(.Fields!TimeIn)
            'MM = Minute(Now) - Minute(.Fields!TimeIn)
            '.Fields!WorkingHours = Trim(Str(HH)) + ":" + Trim(Str(MM))
            .Fields!WorkingHours = TimeDiff(.Fields!TimeIn, Format(Now, "hh:mm AM/PM"))
            EmpOT = Format(TimeDiff(.Fields!TimeIn, Format(Now, "hh:mm AM/PM")), "hh")
            .Fields!OverTime = IIf(TimeDiff(.Fields!TimeIn, Format(Now, "hh:mm AM/PM")) >= "08:00", Format(Str(Val(EmpOT) - 8), "00") & ":" & Trim(Minute(.Fields!WorkingHours)), "00:00")
            If .Fields!TimeIn >= "08:00 AM" Then
                .Fields!Late = TimeDiff("08:00 AM", .Fields!TimeIn)
            Else
                .Fields!Late = "00:00"
            End If
            .Fields!Status = "Out"
            .Update
           connClose
           dlgDTR.Height = 6150
           
           lblStatus.Caption = "You can leave now " + EmpNo + "."
           ConnectDTR
           
           Exit Sub
        ElseIf EmpNo = .Fields!EmployeeNo And .Fields!WorkDate = Format(Now, "mm/dd/yyyy") And .Fields!Status = "Out" Then
            MsgBox "You already Time-out!"
            connClose
            Exit Sub
        Else
            .MoveNext
        End If
    Wend
End With
connClose
ConnectDTR
End Sub

Private Sub cmdCancel_Click()
Unload Me
frmLogIn.Show
End Sub


Private Sub cmdExit_Click()
frmMainMenu.Enabled = True
Unload Me
End Sub

Private Sub cmdIn_Click()
Check_Fields
Verify_Account
' If stat is 1 then Go to DTR Time-in
If stat = 1 Then EmpDTR_in
End Sub

Private Sub cmdOut_Click()
Check_Fields
Verify_Account
If stat = 1 Then EmpDTR_out
End Sub

Private Sub cmdRefresh_Click()
Unload Me
dlgDTR.Show 1
End Sub

Private Sub Form_Load()
ConnectDTR
End Sub

Private Sub tmrTimeDate_Timer()
lblTime.Caption = Format(Now, "hh:mm:ss")
lblDate.Caption = Format(Now, "mm/dd/yyyy")
End Sub

Private Sub txtPassword_Change()
If dlgDTR.Height = 6150 Then dlgDTR.Height = 5685
If txtPassword.BackColor = &HC0C0FF Then txtPassword.BackColor = &H80000005
End Sub

Private Sub txtUsername_Change()
If dlgDTR.Height = 6150 Then dlgDTR.Height = 5685
If txtUsername.BackColor = &HC0C0FF Then txtUsername.BackColor = &H80000005
End Sub
