VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPaySlip 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5565
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   8865
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmPaySlip.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   8865
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.MonthView calStart 
      Height          =   2610
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   4604
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   97976321
      TitleBackColor  =   4210816
      TitleForeColor  =   16777215
      CurrentDate     =   41362
      MinDate         =   36526
   End
   Begin PayrollSystem.jcbutton cmdCancel 
      Height          =   375
      Left            =   7080
      TabIndex        =   9
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
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
      Picture         =   "frmPaySlip.frx":000C
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
   End
   Begin PayrollSystem.jcbutton cmdSelected 
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "Print &Selected"
      Picture         =   "frmPaySlip.frx":0288
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
   End
   Begin PayrollSystem.jcbutton cmdAll 
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "Print &All"
      Picture         =   "frmPaySlip.frx":0405
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
   End
   Begin VB.TextBox txtStart 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdSearch 
      Height          =   375
      Left            =   3000
      Picture         =   "frmPaySlip.frx":067A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Width           =   375
   End
   Begin MSComctlLib.ListView PaySlip 
      Height          =   3495
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   6165
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
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Employee No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Employee Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Working Hours"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "OverTime"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Absences"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Late"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Date Transaction"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Date Processed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Net Pay"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblNetPay 
      Caption         =   "00.00"
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
      Left            =   7440
      TabIndex        =   14
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Php"
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
      Left            =   7080
      TabIndex        =   13
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "Total Net Pay:"
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
      Left            =   5760
      TabIndex        =   12
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label lblTRecords 
      Alignment       =   2  'Center
      Caption         =   "0"
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
      Left            =   1440
      TabIndex        =   11
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Total Records:"
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
      TabIndex        =   10
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Starting Date:"
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
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Image cmdExit 
      Height          =   360
      Left            =   8520
      MousePointer    =   99  'Custom
      Picture         =   "frmPaySlip.frx":0C7C
      Top             =   0
      Width           =   360
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Employees PaySlip"
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
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "PaySlip"
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
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmPaySlip.frx":1366
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20700
   End
End
Attribute VB_Name = "frmPaySlip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub DisplayPay()
Dim NetPayTotal
Dim pay
Dim inf As Integer

NetPayTotal = 0
inf = 0

connectDB
With RS
    While Not .EOF
        If txtStart.Text = .Fields!DateProcessed Then
            Set pay = PaySlip.ListItems.Add(, , .Fields!EmployeeNo)
                pay.SubItems(1) = .Fields!LastName + ", " + .Fields!FirstName + " " + .Fields!MiddleName
                pay.SubItems(2) = .Fields!WorkingHours
                pay.SubItems(3) = .Fields!OverTime
                pay.SubItems(4) = .Fields!Absences
                pay.SubItems(5) = .Fields!Late
                pay.SubItems(6) = .Fields!DateTransaction
                pay.SubItems(7) = .Fields!DateProcessed
                pay.SubItems(8) = .Fields!NetPay
                NetPayTotal = Val(NetPayTotal) + Val(.Fields!NetPay)
                .MoveNext
                inf = 1
        Else
            .MoveNext
        End If
    Wend
    If inf = 0 Then
        txtStart.BackColor = &HC0C0FF
    End If
End With
connClose
lblTRecords.Caption = PaySlip.ListItems.Count
lblNetPay.Caption = Format(NetPayTotal, "#,##0")
End Sub

Private Sub calEnd_DateClick(ByVal DateClicked As Date)
txtEnd.Text = calEnd.Value
calEnd.Visible = False
End Sub

Private Sub calStart_DateClick(ByVal DateClicked As Date)
txtStart.Text = Format(calStart.Value, "mm/dd/yyyy")
calStart.Visible = False
End Sub

Private Sub cmdAll_Click()
If PaySlip.ListItems.Count <> 0 Then
    SQL = "SELECT * FROM PaySlip WHERE DateProcessed = '" & txtStart.Text & "'"
    connectDB
    With RS
        If .RecordCount <> 0 Then
            Set rptPaySlip.DataSource = RS
            rptPaySlip.Show
        End If
    End With
    connClose
End If
End Sub

Private Sub cmdCancel_Click()
frmMainMenu.Enabled = True
Unload Me
End Sub

Private Sub cmdExit_Click()
frmMainMenu.Enabled = True
Unload Me
End Sub

Private Sub cmdSearch_Click()
PaySlip.ListItems.Clear
SQL = "SELECT * FROM PaySlip"
DisplayPay
End Sub

Private Sub cmdSelected_Click()
If PaySlip.ListItems.Count <> 0 Then
    SQL = "SELECT * FROM PaySlip WHERE EmployeeNo = '" & PaySlip.SelectedItem & "'"
    connectDB
    With RS
        If .RecordCount <> 0 Then
            .Fields!MiddleName = Left(.Fields!MiddleName, 1) + "."
            Set rptPaySlip.DataSource = RS
            rptPaySlip.Show
        End If
    End With
    connClose
End If
End Sub

Private Sub Form_Load()
calStart.Value = Date
txtStart.Text = Format(Date, "mm/dd/yyyy")
Dim pay
Dim inf As Integer

inf = 0

PaySlip.ListItems.Clear
SQL = "SELECT * FROM PaySlip WHERE DateProcessed = '" & txtStart.Text & "'"
DisplayPay
frmMainMenu.Enabled = False
End Sub

Private Sub txtEnd_Click()
txtEnd.BackColor = &H80000005
calEnd.Visible = True
End Sub

Private Sub txtStart_Click()
PaySlip.ListItems.Clear
txtStart.BackColor = &H80000005
calStart.Visible = True
End Sub
