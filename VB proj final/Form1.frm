VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAYROLL - PAGE 1"
   ClientHeight    =   11115
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11115
   ScaleWidth      =   16125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Delete"
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
      Left            =   6120
      TabIndex        =   47
      Top             =   10320
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Edit"
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
      Left            =   2040
      TabIndex        =   46
      Top             =   10320
      Width           =   1575
   End
   Begin VB.TextBox Text20 
      Height          =   285
      Left            =   2280
      TabIndex        =   44
      Top             =   1320
      Width           =   3735
   End
   Begin VB.CommandButton Command3 
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
      Left            =   9840
      TabIndex        =   41
      Top             =   9600
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Search"
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
      Left            =   6120
      TabIndex        =   40
      Top             =   9600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
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
      Left            =   2040
      TabIndex        =   39
      Top             =   9600
      Width           =   1575
   End
   Begin VB.TextBox Text19 
      Height          =   285
      Left            =   8400
      TabIndex        =   38
      Top             =   2400
      Width           =   3495
   End
   Begin VB.TextBox Text18 
      Height          =   285
      Left            =   2280
      TabIndex        =   36
      Top             =   2280
      Width           =   3735
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   2280
      TabIndex        =   35
      Top             =   1800
      Width           =   3735
   End
   Begin VB.TextBox Text16 
      Height          =   405
      Left            =   5520
      TabIndex        =   34
      Top             =   720
      Width           =   3855
   End
   Begin VB.TextBox Text15 
      Height          =   495
      Left            =   5160
      TabIndex        =   32
      Top             =   8760
      Width           =   4695
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Top             =   3360
      Width           =   2415
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
      Left            =   6360
      TabIndex        =   5
      Top             =   3000
      Width           =   5175
      Begin VB.TextBox Text14 
         Height          =   495
         Left            =   2280
         TabIndex        =   30
         Top             =   4680
         Width           =   2535
      End
      Begin VB.TextBox Text12 
         Height          =   495
         Left            =   2280
         TabIndex        =   26
         Top             =   3840
         Width           =   2535
      End
      Begin VB.TextBox Text11 
         Height          =   495
         Left            =   2280
         TabIndex        =   25
         Top             =   2880
         Width           =   2535
      End
      Begin VB.TextBox Text10 
         Height          =   495
         Left            =   2280
         TabIndex        =   24
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox Text9 
         Height          =   525
         Left            =   2280
         TabIndex        =   23
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox Text8 
         Height          =   495
         Left            =   2280
         TabIndex        =   22
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label17 
         Caption         =   "TOTAL :"
         Height          =   375
         Left            =   720
         TabIndex        =   29
         Top             =   4800
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "SPECIAL DEDUCTION :"
         Height          =   375
         Left            =   360
         TabIndex        =   21
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label Label14 
         Caption         =   "LOSS OF PAY :"
         Height          =   375
         Left            =   480
         TabIndex        =   20
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label13 
         Caption         =   "TAX :"
         Height          =   375
         Left            =   840
         TabIndex        =   19
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "LOAN  :"
         Height          =   255
         Left            =   840
         TabIndex        =   18
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "INSURANCE :"
         Height          =   255
         Left            =   600
         TabIndex        =   17
         Top             =   720
         Width           =   1215
      End
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
      Left            =   480
      TabIndex        =   4
      Top             =   3000
      Width           =   5295
      Begin VB.TextBox Text13 
         Height          =   495
         Left            =   2520
         TabIndex        =   28
         Top             =   4680
         Width           =   2295
      End
      Begin VB.TextBox Text7 
         Height          =   495
         Left            =   2520
         TabIndex        =   16
         Top             =   3960
         Width           =   2415
      End
      Begin VB.TextBox Text6 
         Height          =   495
         Left            =   2520
         TabIndex        =   14
         Top             =   3000
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   2520
         TabIndex        =   12
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   2520
         TabIndex        =   10
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label16 
         Caption         =   "TOTAL :"
         Height          =   255
         Left            =   600
         TabIndex        =   27
         Top             =   4800
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "SPECIAL ALLOWANCE :"
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "MEDICAL :"
         Height          =   375
         Left            =   600
         TabIndex        =   13
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "D.A. :"
         Height          =   375
         Left            =   840
         TabIndex        =   11
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "H.R.A. :"
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "BASIC :"
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   8400
      TabIndex        =   3
      Top             =   1920
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   8400
      TabIndex        =   0
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label Label22 
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
      Left            =   4560
      TabIndex        =   45
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label5 
      Caption         =   "PASSWORD :"
      Height          =   255
      Left            =   840
      TabIndex        =   43
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "ISSUE  DATE :"
      Height          =   255
      Left            =   6960
      TabIndex        =   42
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label20 
      Caption         =   "EMAIL :"
      Height          =   255
      Left            =   7440
      TabIndex        =   37
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label19 
      Caption         =   "EMPLOYEE NAME :"
      Height          =   255
      Left            =   480
      TabIndex        =   33
      Top             =   1920
      Width           =   1575
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
      Left            =   2400
      TabIndex        =   31
      Top             =   8880
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "USERNAME :"
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "DESIGNATION :"
      Height          =   255
      Left            =   6960
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "DEPARTMENT :"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()

End Sub

Private Sub Combo2_Change()

End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub Command1_Click()
DataEnvironment1.Add Text16.Text, Text20.Text, Text17.Text, Text18.Text, Text1.Text, Text2.Text, Text19.Text, Text3.Text, Text4.Text, Text5.Text, Text6.Text, Text7.Text, Text13.Text, Text8.Text, Text9.Text, Text10.Text, Text11.Text, Text12.Text, Text14.Text, Text15.Text, Text16.Text
MsgBox "Employee Record Successfully Added"
End Sub

Private Sub Command2_Click()
DataEnvironment1.Search Text16.Text

With DataEnvironment1.rsSearch

If DataEnvironment1.rsSearch.EOF Then
MsgBox "Record Of Employee Not Found"
Else
MsgBox "Record Found"

Text16.Text = DataEnvironment1.rsSearch.Fields("UserID")
Text20.Text = DataEnvironment1.rsSearch.Fields("Password")
Text17.Text = DataEnvironment1.rsSearch.Fields("Emp Name")
Text18.Text = DataEnvironment1.rsSearch.Fields("Dept")
Text1.Text = DataEnvironment1.rsSearch.Fields("Issue Date")
Text2.Text = DataEnvironment1.rsSearch.Fields("Designation")
Text19.Text = DataEnvironment1.rsSearch.Fields("Email ID")
Text3.Text = DataEnvironment1.rsSearch.Fields("Basic Sal")
Text4.Text = DataEnvironment1.rsSearch.Fields("HRA")
Text5.Text = DataEnvironment1.rsSearch.Fields("DA")
Text6.Text = DataEnvironment1.rsSearch.Fields("Medical")
Text7.Text = DataEnvironment1.rsSearch.Fields("Special Allowance")
Text13.Text = DataEnvironment1.rsSearch.Fields("Total Allowance")
Text8.Text = DataEnvironment1.rsSearch.Fields("Insurance")
Text9.Text = DataEnvironment1.rsSearch.Fields("Loan")
Text10.Text = DataEnvironment1.rsSearch.Fields("Tax")
Text11.Text = DataEnvironment1.rsSearch.Fields("Leave")
Text12.Text = DataEnvironment1.rsSearch.Fields("Special Deduction")
Text14.Text = DataEnvironment1.rsSearch.Fields("Total Deduction")
Text15.Text = DataEnvironment1.rsSearch.Fields("Total Salary")
End If
.Close
End With



End Sub

Private Sub Command3_Click()
Form2.Show
Unload Me

End Sub

Private Sub Command4_Click()
DataEnvironment1.Edit Text16.Text, Text20.Text, Text17.Text, Text18.Text, Text1.Text, Text2.Text, Text19.Text, Text3.Text, Text4.Text, Text5.Text, Text6.Text, Text7.Text, Text13.Text, Text8.Text, Text9.Text, Text10.Text, Text11.Text, Text12.Text, Text14.Text, Text15.Text, Text16.Text
MsgBox "Employee Record Successfully Edited"
End Sub

Private Sub Command5_Click()
DataEnvironment1.Delete Text16.Text
MsgBox "Record Deleted"
End Sub
