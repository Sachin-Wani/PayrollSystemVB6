VERSION 5.00
Begin VB.Form frmEmpDeductions 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3885
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   3525
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmEmpDeductions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   3525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   3255
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
         TabIndex        =   5
         Top             =   240
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
         TabIndex        =   3
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         Height          =   495
         Left            =   120
         Picture         =   "frmEmpDeductions.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Save"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Height          =   495
         Left            =   1680
         Picture         =   "frmEmpDeductions.frx":224E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1800
         Width           =   1455
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
         TabIndex        =   7
         Top             =   720
         Width           =   855
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
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
      End
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
      Left            =   1680
      TabIndex        =   12
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Employee No:"
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
      TabIndex        =   11
      Top             =   960
      Width           =   1335
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
      TabIndex        =   10
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Update Employee Deductions"
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
      TabIndex        =   9
      Top             =   600
      Width           =   3135
   End
   Begin VB.Image cmdExit 
      Height          =   360
      Left            =   3120
      MousePointer    =   99  'Custom
      Picture         =   "frmEmpDeductions.frx":4490
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmEmpDeductions.frx":4B7A
      Top             =   0
      Width           =   11025
   End
End
Attribute VB_Name = "frmEmpDeductions"
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
If txtSSS.Text = "" Then
    txtSSS.BackColor = &HC0C0FF
    Exit Sub
End If
If txtPagibig.Text = "" Then
    txtPagibig.BackColor = &HC0C0FF
    Exit Sub
End If
If txtPHealth.Text = "" Then
    txtPHealth.BackColor = &HC0C0FF
    Exit Sub
End If

SQL = "SELECT * FROM Employees WHERE EmployeeNo = '" & lblEmpNo.Caption & "'"
connectDB
With RS
    If .RecordCount <> 0 Then
        .Fields!SSS = txtSSS.Text
        .Fields!Pagibig = txtPagibig.Text
        .Fields!PhilHealth = txtPHealth.Text
        .Update
        connClose
        Unload Me
    End If
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
