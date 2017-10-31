VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6210
   LinkTopic       =   "Form2"
   ScaleHeight     =   4530
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "ADM LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   9
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   7
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EMP LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   2520
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   1680
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   360
      ScaleHeight     =   795
      ScaleWidth      =   1035
      TabIndex        =   1
      Top             =   360
      Width           =   1095
      Begin VB.PictureBox Picture2 
         Height          =   855
         Left            =   0
         ScaleHeight     =   795
         ScaleWidth      =   1035
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Label Label3 
      Caption         =   "PASSWORD :"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "USERNAME :"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "RKLS PRODUCTIONS PVT. LTD."
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   600
      Width           =   4095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show
Unload Me

End Sub

Private Sub Command3_Click()
If (Text1.Text = "admin") And (Text2.Text = "admin123") Then
Form1.Show
Unload Me
Else
MsgBox "Invalid Credentials"
End If

End Sub
