VERSION 5.00
Begin VB.Form frmActions 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4005
   ClientLeft      =   5835
   ClientTop       =   -675
   ClientWidth     =   4020
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmActions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin PayrollSystem.jcbutton cmdDeductions 
      Height          =   615
      Left            =   2040
      TabIndex        =   12
      Top             =   3240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "Deductions"
      Picture         =   "frmActions.frx":000C
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
   End
   Begin PayrollSystem.jcbutton cmdActivate 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "Activate"
      Picture         =   "frmActions.frx":05D3
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
   End
   Begin PayrollSystem.jcbutton cmdPaySlip 
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "PaySlip"
      Picture         =   "frmActions.frx":10ED
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
   End
   Begin PayrollSystem.jcbutton cmdDeactivate 
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "Deactivate"
      Picture         =   "frmActions.frx":1665
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
   End
   Begin PayrollSystem.jcbutton cmdPrint 
      Height          =   615
      Left            =   2040
      TabIndex        =   3
      Top             =   2520
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "Print"
      Picture         =   "frmActions.frx":22B7
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
      CaptionAlign    =   0
   End
   Begin VB.Label Label5 
      Caption         =   "Employee No:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblEmpNo 
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
      Left            =   1800
      TabIndex        =   10
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblStatus 
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
      Left            =   840
      TabIndex        =   8
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblName 
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
      Left            =   840
      TabIndex        =   7
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Action"
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
      TabIndex        =   1
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Take an action"
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
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
   Begin VB.Image cmdExit 
      Height          =   360
      Left            =   3600
      MousePointer    =   99  'Custom
      Picture         =   "frmActions.frx":2F09
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmActions.frx":35F3
      Top             =   0
      Width           =   11025
   End
End
Attribute VB_Name = "frmActions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdActivate_Click()
SQL = "SELECT * FROM Employees WHERE EmployeeNo='" & EmpNumb & "'"
connectDB
With RS
    If .RecordCount <> 0 Then
        .Fields!CurrentStatus = "Active"
        .Update
        cmdPaySlip.Enabled = True
        cmdDeactivate.Visible = True
        cmdActivate.Visible = False
        MsgBox "The Employee was Activated successfully!"
    End If
End With
connClose

SQL = "SELECT * FROM Employees"
connOpen
frmMainMenu.EmpRecs
connClose

frmMainMenu.PayrollDetails
End Sub

Private Sub cmdDeactivate_Click()
SQL = "SELECT * FROM Employees WHERE EmployeeNo='" & EmpNumb & "'"
connectDB
With RS
    If .RecordCount <> 0 Then
        .Fields!CurrentStatus = "Inactive"
        .Update
        cmdPaySlip.Enabled = False
        cmdDeactivate.Visible = False
        cmdActivate.Visible = True
        MsgBox "The Employee was Deactivated successfully!"
    End If
End With
connClose

SQL = "SELECT * FROM Employees"
connOpen
frmMainMenu.EmpRecs
connClose

frmMainMenu.PayrollDetails
End Sub

Private Sub cmdDeductions_Click()
SQL = "SELECT * FROM Employees WHERE EmployeeNo = '" & lblEmpNo.Caption & "'"
connectDB
With RS
    frmEmpDeductions.txtSSS.Text = IIf(IsNull(.Fields!SSS), "0", .Fields!SSS)
    frmEmpDeductions.txtPagibig.Text = IIf(IsNull(.Fields!Pagibig), "0", .Fields!Pagibig)
    frmEmpDeductions.txtPHealth.Text = IIf(IsNull(.Fields!PhilHealth), "0", .Fields!PhilHealth)
End With
connClose
frmEmpDeductions.lblEmpNo.Caption = lblEmpNo.Caption
frmEmpDeductions.Show
End Sub

Private Sub cmdExit_Click()
frmMainMenu.Enabled = True
Unload Me
End Sub

Private Sub cmdPaySlip_Click()
frmEmpPay.Show
End Sub

Private Sub cmdPrint_Click()
SQL = "SELECT * FROM Employees WHERE EmployeeNo = '" & lblEmpNo.Caption & "'"
connectDB
With RS
    If .RecordCount <> 0 Then
        Set rptEmpInfo.DataSource = RS
        rptEmpInfo.Show
    End If
End With
connClose
End Sub

Private Sub Form_Load()
frmMainMenu.Enabled = False
End Sub
