VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5790
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   3405
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSettings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3135
      Begin PayrollSystem.jcbutton cmdChangePass 
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
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
         Caption         =   "Change &Password"
         Picture         =   "frmSettings.frx":000C
         UseMaskCOlor    =   -1  'True
         MaskColor       =   16777215
         CaptionAlign    =   0
      End
      Begin PayrollSystem.jcbutton cmdEmpAccounts 
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
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
         Caption         =   "Employee &Accounts"
         Picture         =   "frmSettings.frx":0423
         UseMaskCOlor    =   -1  'True
         MaskColor       =   16777215
         CaptionAlign    =   0
      End
      Begin VB.Frame Frame2 
         Height          =   2655
         Left            =   120
         TabIndex        =   3
         Top             =   1920
         Width           =   2895
         Begin PayrollSystem.jcbutton cmdClose 
            Height          =   495
            Left            =   120
            TabIndex        =   10
            Top             =   2040
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   873
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
            Caption         =   "&Close"
            Picture         =   "frmSettings.frx":0842
            UseMaskCOlor    =   -1  'True
            MaskColor       =   16777215
            CaptionAlign    =   0
         End
         Begin PayrollSystem.jcbutton cmdDeductions 
            Height          =   495
            Left            =   120
            TabIndex        =   4
            Top             =   1440
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   873
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
            BackColor       =   12648447
            Caption         =   "D&eductions Settings"
            Picture         =   "frmSettings.frx":0ABE
            UseMaskCOlor    =   -1  'True
            MaskColor       =   16777215
            CaptionAlign    =   0
         End
         Begin PayrollSystem.jcbutton cmdPosition 
            Height          =   495
            Left            =   120
            TabIndex        =   5
            Top             =   840
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   873
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
            BackColor       =   12648447
            Caption         =   "Change P&osition"
            Picture         =   "frmSettings.frx":0ED5
            UseMaskCOlor    =   -1  'True
            MaskColor       =   16777215
            CaptionAlign    =   0
         End
         Begin PayrollSystem.jcbutton cmdWDays 
            Height          =   495
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   873
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
            BackColor       =   12648447
            Caption         =   "Working &Days Settings"
            Picture         =   "frmSettings.frx":12F5
            UseMaskCOlor    =   -1  'True
            MaskColor       =   16777215
            CaptionAlign    =   0
         End
      End
      Begin PayrollSystem.jcbutton cmdNewAccount 
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
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
         Caption         =   "Create &New Account"
         Picture         =   "frmSettings.frx":1715
         UseMaskCOlor    =   -1  'True
         MaskColor       =   16777215
         CaptionAlign    =   0
      End
   End
   Begin VB.Image cmdExit 
      Height          =   360
      Left            =   3000
      MousePointer    =   99  'Custom
      Picture         =   "frmSettings.frx":1B23
      Top             =   0
      Width           =   360
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Update Accounts, Days, Taxes, etc.."
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
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
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
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmSettings.frx":220D
      Top             =   0
      Width           =   11025
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChangePass_Click()
frmChangePass.Show
End Sub

Private Sub cmdClose_Click()
frmMainMenu.Enabled = True
Unload Me
End Sub

Private Sub cmdDeductions_Click()
frmDeductions.Show
End Sub

Private Sub cmdEmpAccounts_Click()
frmEmpAccounts.Show
End Sub

Private Sub cmdExit_Click()
frmMainMenu.Enabled = True
Unload Me
End Sub

Private Sub cmdNewAccount_Click()
frmEmpNewAccount.Show
End Sub

Private Sub cmdPosition_Click()
frmPosition.Show
End Sub

Private Sub cmdWDays_Click()
frmWDays.Show
End Sub

Private Sub Form_Load()
frmMainMenu.Enabled = False
End Sub
