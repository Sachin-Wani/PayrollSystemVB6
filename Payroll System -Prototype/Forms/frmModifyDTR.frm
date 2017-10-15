VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmModifyDTR 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5910
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   3840
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
   Icon            =   "frmModifyDTR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Height          =   495
      Left            =   2040
      Picture         =   "frmModifyDTR.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Height          =   495
      Left            =   360
      Picture         =   "frmModifyDTR.frx":224E
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Height          =   3375
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   3615
      Begin MSComCtl2.DTPicker dtpOut 
         Height          =   375
         Left            =   1800
         TabIndex        =   15
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   162988034
         CurrentDate     =   41387
      End
      Begin MSComCtl2.DTPicker dtpIn 
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "hh:mm"
         Format          =   162988034
         CurrentDate     =   41387
      End
      Begin VB.TextBox txtLate 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1800
         TabIndex        =   13
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox txtOT 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1800
         TabIndex        =   11
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox txtWHours 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   9
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         Caption         =   "00/00/0000"
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
         TabIndex        =   19
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Date Worked:"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Late:"
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Over Time:"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Working Hours:"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Time-out:"
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Time-in:"
         Height          =   375
         Left            =   720
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3615
      Begin VB.Label lblEmpNo 
         Caption         =   "NEPC-00000000000"
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Employee No:"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "Modify DTR"
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
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Update Daily Time Record"
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
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
   Begin VB.Image cmdExit 
      Height          =   360
      Left            =   3480
      MousePointer    =   99  'Custom
      Picture         =   "frmModifyDTR.frx":4490
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmModifyDTR.frx":4B7A
      Top             =   0
      Width           =   11025
   End
End
Attribute VB_Name = "frmModifyDTR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub


Private Sub cmdSave_Click()
SQL = "SELECT * FROM DTR WHERE EmployeeNo = '" & lblEmpNo.Caption & "' AND WorkDate = '" & lblDate.Caption & "'"
connectDB
With RS
    If .RecordCount <> 0 Then
        .Fields!TimeIn = Format(dtpIn.Value, "hh:mm AM/PM")
        .Fields!TimeOut = Format(dtpOut.Value, "hh:mm AM/PM")
        .Fields!WorkingHours = txtWHours.Text
        .Fields!OverTime = txtOT.Text
        .Fields!Late = txtLate.Text
        .Fields!Status = "Out"
        .Update
        Unload Me
    End If
End With
connClose
End Sub

Private Sub dtpIn_Change()
'IIf(Format(dtpOut.Value, "hh:mm AM/PM") < Format(dtpIn.Value, "hh:mm AM/PM"), "00:00", TimeDiff(Format(dtpIn.Value, "hh:mm AM/PM"), Format(dtpOut.Value, "hh:mm AM/PM")))
txtWHours.Text = TimeDiff(Format(dtpIn.Value, "hh:mm AM/PM"), Format(dtpOut.Value, "hh:mm AM/PM"))
txtLate.Text = IIf(Format(Format(dtpIn.Value, "hh:mm AM/PM"), "hh:mm AM/PM") >= "08:00 AM", TimeDiff("08:00 AM", Format(Format(dtpIn.Value, "hh:mm AM/PM"), "hh:mm AM/PM")), "00:00")
End Sub

Private Sub dtpOut_Change()
'txtWHours.Text = IIf(Format(dtpOut.Value, "hh:mm AM/PM") < Format(dtpIn.Value, "hh:mm AM/PM"), "00:00", TimeDiff(Format(dtpIn.Value, "hh:mm AM/PM"), Format(dtpOut.Value, "hh:mm AM/PM")))
txtWHours.Text = TimeDiff(Format(dtpIn.Value, "hh:mm AM/PM"), Format(dtpOut.Value, "hh:mm AM/PM"))
txtOT.Text = IIf(Format(dtpOut.Value, "hh:mm AM/PM") > "04:00 PM", TimeDiff("04:00 PM", Format(dtpOut.Value, "hh:mm AM/PM")), "00:00")
End Sub
