VERSION 5.00
Begin VB.Form frmPrint 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2745
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   6060
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   5775
      Begin PayrollSystem.jcbutton cmdCancel 
         Height          =   615
         Left            =   3000
         TabIndex        =   6
         Top             =   960
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1085
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
         BackColor       =   12640511
         Caption         =   "&Cancel"
         Picture         =   "frmPrint.frx":000C
         UseMaskCOlor    =   -1  'True
         MaskColor       =   16777215
      End
      Begin PayrollSystem.jcbutton cmdPaySlip 
         Height          =   615
         Left            =   3000
         TabIndex        =   3
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1085
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
         Caption         =   "Employees &PaySlip"
         Picture         =   "frmPrint.frx":063B
         PictureAlign    =   0
         UseMaskCOlor    =   -1  'True
         MaskColor       =   16777215
         CaptionAlign    =   0
      End
      Begin PayrollSystem.jcbutton cmdRecords 
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1085
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
         Caption         =   "Employees &Records"
         Picture         =   "frmPrint.frx":0CD7
         PictureAlign    =   0
         UseMaskCOlor    =   -1  'True
         MaskColor       =   16777215
         CaptionAlign    =   0
      End
      Begin PayrollSystem.jcbutton cmdInfo 
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1085
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
         Caption         =   "Employees &Information"
         Picture         =   "frmPrint.frx":125C
         PictureAlign    =   0
         UseMaskCOlor    =   -1  'True
         MaskColor       =   16777215
         CaptionAlign    =   0
      End
   End
   Begin VB.Image Image8 
      Height          =   720
      Left            =   120
      Picture         =   "frmPrint.frx":17F3
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "Select to Print"
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
      Left            =   960
      TabIndex        =   1
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Select one to Print Records"
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
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   4575
   End
   Begin VB.Image cmdExit 
      Height          =   360
      Left            =   5640
      MousePointer    =   99  'Custom
      Picture         =   "frmPrint.frx":26BD
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmPrint.frx":2DA7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20700
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
frmMainMenu.Enabled = True
Unload Me
End Sub

Private Sub cmdExit_Click()
frmMainMenu.Enabled = True
Unload Me
End Sub

Private Sub cmdInfo_Click()
SQL = "SELECT * FROM Employees"
connectDB
With RS
    If .RecordCount <> 0 Then
        Set rptEmpInfo.DataSource = RS
        rptEmpInfo.Show
    End If
End With
connClose
End Sub

Private Sub cmdPaySlip_Click()
frmPaySlip.Show
Unload Me
End Sub

Private Sub cmdRecords_Click()
SQL = "SELECT * FROM Employees"
connectDB
With RS
    If .RecordCount <> 0 Then
        Set rptRecords.DataSource = RS
        rptRecords.Show
    End If
End With
connClose
End Sub

Private Sub Form_Load()
frmMainMenu.Enabled = False
End Sub
