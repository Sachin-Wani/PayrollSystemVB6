VERSION 5.00
Begin VB.Form frmDeductions 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3465
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   3465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmDeductions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3255
      Begin VB.CommandButton cmdCancel 
         Height          =   495
         Left            =   1680
         Picture         =   "frmDeductions.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmdSave 
         Height          =   495
         Left            =   120
         Picture         =   "frmDeductions.frx":224E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Save"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtPHealth 
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
         Left            =   1440
         TabIndex        =   5
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtPagibig 
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
         Left            =   1440
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtSSS 
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
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Phil-Health:"
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
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Pag-ibig:"
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
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "SSS:"
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
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Image cmdExit 
      Height          =   360
      Left            =   3120
      MousePointer    =   99  'Custom
      Picture         =   "frmDeductions.frx":4490
      Top             =   0
      Width           =   360
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Update Deductions"
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
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "Deductions"
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
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmDeductions.frx":4B7A
      Top             =   0
      Width           =   11025
   End
End
Attribute VB_Name = "frmDeductions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Check_Fields()
If txtSSS.Text = "" Or txtPagibig.Text = "" Or txtPHealth.Text = "" Then
    If txtSSS.Text = "" Then
        txtSSS.BackColor = &HC0C0FF
        txtSSS.SetFocus
    End If
    If txtPagibig.Text = "" Then
        txtPagibig.BackColor = &HC0C0FF
        txtSSS.SetFocus
    End If
    If txtPHealth.Text = "" Then
        txtPHealth.BackColor = &HC0C0FF
        txtPHealth.SetFocus
    End If
    Exit Sub
End If
End Sub

Sub Clear_Fields()
txtSSS.Text = ""
txtPagibig.Text = "'"
txtPHealth = ""
End Sub


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdExit_Click()
Call cmdCancel_Click
End Sub


Private Sub cmdSave_Click()
On Error Resume Next
Check_Fields

SQL = "SELECT * FROM Deductions"
connectDB
With RS
    .MoveFirst
    .Fields!SSS = txtSSS.Text
    .Fields!Pagibig = txtPagibig.Text
    .Fields!PhilHealth = txtPHealth.Text
    .Update
    Clear_Fields
    MsgBox "Deductions updated!"
End With
connClose
End Sub

Private Sub Form_Load()
SQL = "SELECT * FROM Deductions"
connectDB
With RS
    txtSSS.Text = .Fields!SSS
    txtPagibig.Text = .Fields!Pagibig
    txtPHealth.Text = .Fields!PhilHealth
End With
connClose
End Sub

Private Sub txtPagibig_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = vbKeyBack) Or (KeyAscii = 46) Then
    Exit Sub
Else
    KeyAscii = 0
End If
End Sub

Private Sub txtPHealth_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = vbKeyBack) Or (KeyAscii = 46) Then
    Exit Sub
ElseIf KeyAscii = 13 Then
    Call cmdSave_Click
Else
    KeyAscii = 0
End If
End Sub

Private Sub txtSSS_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = vbKeyBack) Or (KeyAscii = 46) Then
    Exit Sub
Else
    KeyAscii = 0
End If
End Sub
