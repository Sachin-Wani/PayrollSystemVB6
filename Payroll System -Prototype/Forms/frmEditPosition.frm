VERSION 5.00
Begin VB.Form frmEditPosition 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3720
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   4095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmEditPosition.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3855
      Begin VB.CommandButton cmdSave 
         Height          =   495
         Left            =   360
         Picture         =   "frmEditPosition.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtPosition 
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
         Left            =   1800
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtConRate 
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
         Left            =   1800
         TabIndex        =   5
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtRegRate 
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
         Left            =   1800
         TabIndex        =   4
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancel 
         Height          =   495
         Left            =   1920
         Picture         =   "frmEditPosition.frx":224E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label3 
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
         Height          =   375
         Left            =   840
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Contractual Rate:"
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
         TabIndex        =   8
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Regular Rate:"
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
         Left            =   480
         TabIndex        =   7
         Top             =   1320
         Width           =   1335
      End
   End
   Begin VB.Image cmdExit 
      Height          =   360
      Left            =   3720
      MousePointer    =   99  'Custom
      Picture         =   "frmEditPosition.frx":4490
      Top             =   0
      Width           =   360
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Existing Position"
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
      Caption         =   "Edit Position"
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
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmEditPosition.frx":4B7A
      Top             =   0
      Width           =   11025
   End
End
Attribute VB_Name = "frmEditPosition"
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
SQL = "SELECT * FROM Positions WHERE Designation = '" & EmpPosition & "'"
connectDB
With RS
    .Fields!Designation = txtPosition.Text
    .Fields!ConRate = txtConRate.Text
    .Fields!RegRate = txtRegRate.Text
    .Update
End With
connClose
frmPosition.EmpPositions
Unload Me
End Sub

Private Sub Form_Load()
SQL = "SELECT * FROM Positions WHERE Designation = '" & EmpPosition & "'"
connectDB
With RS
    txtPosition.Text = .Fields!Designation
    txtConRate.Text = .Fields!ConRate
    txtRegRate.Text = .Fields!RegRate
End With
connClose
End Sub

Private Sub txtConRate_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = vbKeyBack) Then
    Exit Sub
Else
    KeyAscii = 0
End If
End Sub


Private Sub txtRegRate_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = vbKeyBack) Then
    Exit Sub
ElseIf KeyAscii = 13 Then
    Call cmdSave_Click
Else
    KeyAscii = 0
End If
End Sub
