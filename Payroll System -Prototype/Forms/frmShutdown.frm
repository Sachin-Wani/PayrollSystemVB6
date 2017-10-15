VERSION 5.00
Begin VB.Form frmShutdown 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1680
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   3495
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmShutdown.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin PayrollSystem.jcbutton cmdOff 
      Height          =   615
      Left            =   1800
      TabIndex        =   3
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "Log &Off"
      Picture         =   "frmShutdown.frx":000C
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
   End
   Begin PayrollSystem.jcbutton cmdShutdown 
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "&Shutdown"
      Picture         =   "frmShutdown.frx":0C5E
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Close Payroll System"
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
      Width           =   1935
   End
   Begin VB.Image cmdExit 
      Height          =   360
      Left            =   3120
      MousePointer    =   99  'Custom
      Picture         =   "frmShutdown.frx":18B0
      Top             =   0
      Width           =   360
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "Shutdown"
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
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmShutdown.frx":1F9A
      Top             =   0
      Width           =   11025
   End
End
Attribute VB_Name = "frmShutdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
frmMainMenu.Enabled = True
Unload Me
End Sub

Private Sub cmdOff_Click()
frmAdmin.Show
Unload Me
Unload frmMainMenu
End Sub

Private Sub cmdShutdown_Click()
End
End Sub

Private Sub Form_Load()
frmMainMenu.Enabled = False
End Sub
