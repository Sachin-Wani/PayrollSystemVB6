VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDTRHistory 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4935
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   9015
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDTRHistory.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSearch 
      Height          =   375
      Left            =   4200
      Picture         =   "frmDTRHistory.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Width           =   375
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   1080
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CalendarForeColor=   4210752
      CalendarTitleBackColor=   4210688
      CalendarTitleForeColor=   16777215
      Format          =   112787456
      CurrentDate     =   41387
   End
   Begin MSComctlLib.ListView DTR 
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5530
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Employee No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Time-in"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Time-out"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Working Hours"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Over Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Late"
         Object.Width           =   2540
      EndProperty
   End
   Begin PayrollSystem.jcbutton cmdRefresh 
      Height          =   375
      Left            =   8520
      TabIndex        =   8
      Top             =   1080
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
      Picture         =   "frmDTRHistory.frx":060E
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
   End
   Begin VB.Label lblRecCount 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   7
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Records:"
      Height          =   375
      Left            =   5760
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Date:"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "Daily Time Record History"
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
      Width           =   4455
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Review or Modify Daily Time Record"
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
   Begin VB.Image cmdExit 
      Height          =   360
      Left            =   8640
      MousePointer    =   99  'Custom
      Picture         =   "frmDTRHistory.frx":086E
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmDTRHistory.frx":0F58
      Top             =   0
      Width           =   11025
   End
End
Attribute VB_Name = "frmDTRHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
frmMainMenu.Enabled = True
Unload Me
End Sub

Private Sub cmdRefresh_Click()
Call cmdSearch_Click
End Sub

Private Sub cmdSearch_Click()
Dim recs

DTR.ListItems.Clear

SQL = "SELECT * FROM DTR WHERE WorkDate = '" & Format(dtpDate.Value, "mm/dd/yyyy") & "'"
connectDB
With RS
    If .RecordCount <> 0 Then
        lblRecCount.Caption = .RecordCount
        While Not .EOF
            Set recs = DTR.ListItems.Add(, , .Fields!EmployeeNo)
                recs.SubItems(1) = .Fields!TimeIn
                recs.SubItems(2) = .Fields!TimeOut
                recs.SubItems(3) = .Fields!WorkingHours
                recs.SubItems(4) = .Fields!OverTime
                recs.SubItems(5) = .Fields!Late
            .MoveNext
        Wend
    End If
End With
connClose
End Sub

Private Sub dtpDate_Change()
DTR.ListItems.Clear
lblRecCount.Caption = "0"
End Sub

Private Sub dtpDate_Click()
DTR.ListItems.Clear
lblRecCount.Caption = "0"
End Sub

Private Sub DTR_DblClick()
'On Error GoTo NoRecord:
If DTR.ListItems.Count <> 0 Then
    frmModifyDTR.lblEmpNo.Caption = DTR.SelectedItem
    frmModifyDTR.lblDate.Caption = Format(dtpDate.Value, "mm/dd/yyyy")
    frmModifyDTR.dtpIn.Value = DTR.SelectedItem.SubItems(1)
    frmModifyDTR.dtpOut.Value = IIf(DTR.SelectedItem.SubItems(2) = "---", "00:00 AM", DTR.SelectedItem.SubItems(2))
    frmModifyDTR.txtWHours.Text = DTR.SelectedItem.SubItems(3)
    frmModifyDTR.txtOT.Text = DTR.SelectedItem.SubItems(4)
    frmModifyDTR.txtLate.Text = DTR.SelectedItem.SubItems(5)
    frmModifyDTR.Show
End If
'NoRecord:
'    MsgBox "No Selected Record!"
End Sub

Private Sub Form_Load()
dtpDate.Value = Date
frmMainMenu.Enabled = False
End Sub
