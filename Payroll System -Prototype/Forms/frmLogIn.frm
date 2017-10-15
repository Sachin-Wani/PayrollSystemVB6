VERSION 5.00
Begin VB.Form frmLogIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select User"
   ClientHeight    =   3135
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1852.261
   ScaleMode       =   0  'User
   ScaleWidth      =   4070.331
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdmin 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   240
      Picture         =   "frmLogIn.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   3855
   End
   Begin VB.CommandButton cmdUser 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   240
      Picture         =   "frmLogIn.frx":8182
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   3855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "Select User"
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
      TabIndex        =   5
      Top             =   0
      Width           =   3495
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Level"
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
      TabIndex        =   4
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404040&
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmLogIn.frx":10304
      Top             =   0
      Width           =   11025
   End
End
Attribute VB_Name = "frmLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdmin_Click()
frmAdmin.Show
Unload Me
End Sub

Private Sub cmdUser_Click()
dlgDTR.Show
Unload Me
End Sub

