VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   9465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12765
   LinkTopic       =   "Form3"
   ScaleHeight     =   9465
   ScaleWidth      =   12765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   40
      Top             =   8640
      Width           =   1935
   End
   Begin VB.TextBox Text19 
      Height          =   285
      Left            =   8160
      TabIndex        =   38
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox Text18 
      Height          =   285
      Left            =   2280
      TabIndex        =   36
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   2280
      TabIndex        =   35
      Top             =   1080
      Width           =   3855
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   2280
      TabIndex        =   34
      Top             =   600
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   8160
      TabIndex        =   27
      Top             =   600
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   8160
      TabIndex        =   26
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Caption         =   "ALLOWANCES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   840
      TabIndex        =   14
      Top             =   2040
      Width           =   5295
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   2400
         TabIndex        =   41
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   2520
         TabIndex        =   19
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   2520
         TabIndex        =   18
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox Text6 
         Height          =   495
         Left            =   2520
         TabIndex        =   17
         Top             =   3000
         Width           =   2415
      End
      Begin VB.TextBox Text7 
         Height          =   495
         Left            =   2520
         TabIndex        =   16
         Top             =   3960
         Width           =   2415
      End
      Begin VB.TextBox Text13 
         Height          =   495
         Left            =   2520
         TabIndex        =   15
         Top             =   4680
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "BASIC :"
         Height          =   255
         Left            =   720
         TabIndex        =   25
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "H.R.A. :"
         Height          =   255
         Left            =   720
         TabIndex        =   24
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "D.A. :"
         Height          =   375
         Left            =   840
         TabIndex        =   23
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "MEDICAL :"
         Height          =   375
         Left            =   600
         TabIndex        =   22
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "SPECIAL ALLOWANCE :"
         Height          =   375
         Left            =   360
         TabIndex        =   21
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label Label16 
         Caption         =   "TOTAL :"
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   4800
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "DEDUCTIONS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   6600
      TabIndex        =   1
      Top             =   2040
      Width           =   5175
      Begin VB.TextBox Text8 
         Height          =   495
         Left            =   2280
         TabIndex        =   7
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox Text9 
         Height          =   525
         Left            =   2280
         TabIndex        =   6
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox Text10 
         Height          =   495
         Left            =   2280
         TabIndex        =   5
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox Text11 
         Height          =   495
         Left            =   2280
         TabIndex        =   4
         Top             =   2880
         Width           =   2535
      End
      Begin VB.TextBox Text12 
         Height          =   495
         Left            =   2280
         TabIndex        =   3
         Top             =   3840
         Width           =   2535
      End
      Begin VB.TextBox Text14 
         Height          =   495
         Left            =   2280
         TabIndex        =   2
         Top             =   4680
         Width           =   2535
      End
      Begin VB.Label Label11 
         Caption         =   "INSURANCE :"
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "LOAN  :"
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "TAX :"
         Height          =   375
         Left            =   840
         TabIndex        =   11
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "LOSS OF PAY :"
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label15 
         Caption         =   "SPECIAL DEDUCTION :"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label Label17 
         Caption         =   "TOTAL :"
         Height          =   375
         Left            =   720
         TabIndex        =   8
         Top             =   4800
         Width           =   975
      End
   End
   Begin VB.TextBox Text15 
      Height          =   495
      Left            =   4800
      TabIndex        =   0
      Top             =   7680
      Width           =   4695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "RKLS PRODUCTIONS  PVT. LTD."
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   39
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label20 
      Caption         =   "MAIL ID  :"
      Height          =   255
      Left            =   7080
      TabIndex        =   37
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label19 
      Caption         =   "EMPLOYEE NAME :"
      Height          =   255
      Left            =   600
      TabIndex        =   33
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "ISSUE  DATE :"
      Height          =   255
      Left            =   6720
      TabIndex        =   32
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "DEPARTMENT :"
      Height          =   255
      Left            =   720
      TabIndex        =   31
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "DESIGNATION :"
      Height          =   255
      Left            =   6720
      TabIndex        =   30
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "EMPLOYEE CODE :"
      Height          =   255
      Left            =   480
      TabIndex        =   29
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label18 
      Caption         =   "TOTAL SALARY   :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   28
      Top             =   7800
      Width           =   2175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Show
Unload Me

End Sub
