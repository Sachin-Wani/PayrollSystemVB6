VERSION 5.00
Begin VB.Form frmLock 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1500
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6195
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmLock.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin PayrollSystem.jcbutton cmdUnlock 
      Height          =   615
      Left            =   5280
      TabIndex        =   2
      Top             =   360
      Width           =   615
      _extentx        =   1085
      _extenty        =   1085
      buttonstyle     =   10
      font            =   "frmLock.frx":000C
      backcolor       =   8421504
      caption         =   ""
      picture         =   "frmLock.frx":0034
      maskcolor       =   16777215
      usemaskcolor    =   -1  'True
   End
   Begin VB.TextBox txtLock 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "System Lock:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmLOck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdUnlock_Click()
If txtLock.Text = "admin" Then
    frmMainMenu.Enabled = True
    Unload Me
End If
End Sub

Private Sub Form_Load()
frmMainMenu.Enabled = False
End Sub
