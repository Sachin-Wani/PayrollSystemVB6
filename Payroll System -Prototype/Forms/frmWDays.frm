VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmWDays 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3675
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   6825
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmWDays.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Height          =   495
      Left            =   4920
      Picture         =   "frmWDays.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Height          =   495
      Left            =   3360
      Picture         =   "frmWDays.frx":224E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Save"
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox txtDays 
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
      Left            =   5880
      TabIndex        =   2
      Top             =   1560
      Width           =   855
   End
   Begin MSComCtl2.MonthView calStart 
      Height          =   2610
      Left            =   120
      TabIndex        =   6
      Top             =   960
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
      StartOfWeek     =   113049601
      TitleBackColor  =   4210816
      TitleForeColor  =   16777215
      CurrentDate     =   41362
      MinDate         =   36526
   End
   Begin VB.Label Label3 
      Caption         =   "No. of Working Days (15-30):"
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
      Left            =   3120
      TabIndex        =   5
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "Working Days"
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
      Width           =   2535
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Default: (15 days)"
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
      Width           =   2055
   End
   Begin VB.Image cmdExit 
      Height          =   360
      Left            =   6480
      MousePointer    =   99  'Custom
      Picture         =   "frmWDays.frx":4490
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmWDays.frx":4B7A
      Top             =   0
      Width           =   11025
   End
End
Attribute VB_Name = "frmWDays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
If txtDays.Text = "" Or txtDays.Text <= "14" Then
    txtDays.BackColor = &HC0C0FF
    txtDays.SetFocus
Else
    SQL = "SELECT WDays FROM WorkingDays"
    connectDB
    With RS
        .Fields!WDays = txtDays.Text
        .Update
        connClose
        Unload Me
    End With
    connClose
End If
End Sub

Private Sub Form_Load()
calStart.Value = Date

SQL = "SELECT * FROM WorkingDays"
connectDB
With RS
    txtDays.Text = .Fields!WDays
End With
connClose
End Sub

Private Sub txtDays_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = vbKeyBack) Then
    Exit Sub
Else
    KeyAscii = 0
End If
End Sub
