VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPosition 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5280
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   4890
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmPosition.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView Position 
      Height          =   3615
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   6376
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
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Designation"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Contractual Rate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Regular Rate"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   495
      Left            =   3240
      Picture         =   "frmPosition.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   495
      Left            =   1680
      Picture         =   "frmPosition.frx":224E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Delete Record"
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   495
      Left            =   120
      Picture         =   "frmPosition.frx":4490
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Add Record"
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
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
      TabIndex        =   4
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Add, Edit and Delete Position"
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
      TabIndex        =   3
      Top             =   600
      Width           =   3135
   End
   Begin VB.Image cmdExit 
      Height          =   360
      Left            =   4440
      MousePointer    =   99  'Custom
      Picture         =   "frmPosition.frx":66D2
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmPosition.frx":6DBC
      Top             =   0
      Width           =   11025
   End
End
Attribute VB_Name = "frmPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub EmpPositions()
Dim EmpPos

SQL = "SELECT * FROM Positions"
connectDB

Position.ListItems.Clear

With RS
    .MoveFirst
    While Not .EOF
        Set EmpPos = Position.ListItems.Add(, , .Fields!Designation)
            EmpPos.SubItems(1) = .Fields!ConRate
            EmpPos.SubItems(2) = .Fields!RegRate
            .MoveNext
    Wend
End With
connClose
End Sub

Private Sub cmdAdd_Click()
frmAddPosition.Show
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
If Position.ListItems.Count <> 0 Then
    SQL = "DELETE * FROM Positions WHERE Designation = '" & Position.SelectedItem & "'"
    connectDB
    connClose
    EmpPositions
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
EmpPositions
End Sub

Private Sub Position_DblClick()
EmpPosition = Position.SelectedItem
frmEditPosition.Show
End Sub
