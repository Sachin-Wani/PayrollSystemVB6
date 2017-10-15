VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00808000&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   3915
   ClientLeft      =   330
   ClientTop       =   1485
   ClientWidth     =   7230
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrLoading 
      Interval        =   50
      Left            =   0
      Top             =   3360
   End
   Begin VB.Timer tmrInfo 
      Interval        =   100
      Left            =   480
      Top             =   600
   End
   Begin VB.Label lblSplash 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait ..."
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   6495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait.."
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   735
      Left            =   2400
      TabIndex        =   2
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Image Image4 
      Height          =   2280
      Left            =   5040
      Picture         =   "frmSplash.frx":000C
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   2280
   End
   Begin VB.Image Image3 
      Height          =   2280
      Left            =   4080
      Picture         =   "frmSplash.frx":1618
      Top             =   1800
      Width           =   2280
   End
   Begin VB.Image Disk1 
      Height          =   2280
      Left            =   360
      Picture         =   "frmSplash.frx":2C1B
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2280
   End
   Begin VB.Image Disk2 
      Height          =   2280
      Left            =   360
      Picture         =   "frmSplash.frx":4227
      Top             =   1800
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.Label Period 
      BackStyle       =   0  'Transparent
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
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Loading 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nippon Express Philippines Corporation"
      BeginProperty Font 
         Name            =   "Freehand521 BT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub tmrInfo_Timer()
Static counting As Integer
counting = counting + 1

If counting = 3 Then
    lblSplash.Caption = "Preparing ..."
ElseIf counting = 10 Then
    lblSplash.Caption = "Loading forms ..."
ElseIf counting = 20 Then
    lblSplash.Caption = "Checking Databases ..."
ElseIf counting = 30 Then
    lblSplash.Caption = "Loading Databases ..."
ElseIf counting = 40 Then
    tmrLoading.Enabled = False
    lblSplash.Caption = "Done ..."
ElseIf counting = 45 Then
    lblSplash.Caption = "Starting Payroll System ..."
ElseIf counting = 60 Then
    tmrInfo.Enabled = False
    Unload Me
    frmLogIn.Show
End If
End Sub

Private Sub tmrLoading_Timer()
If Disk1.Visible = True Then
    Disk1.Visible = False
    Disk2.Visible = True
Else
    Disk1.Visible = True
    Disk2.Visible = False
End If
End Sub
