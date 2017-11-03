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
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
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
Public Sub Command1_Click()
If ((Text1.Text = "160501") And (Text2.Text = "emp1")) Or ((Text1.Text = "160502") And (Text2.Text = "emp2")) Then
DataEnvironment1.Search Text1.Text

With DataEnvironment1.rsSearch
Form3.Show
Unload Me


If DataEnvironment1.rsSearch.EOF Then
MsgBox "Record Of Employee Not Found"


Text17.Text = ""
Text18.Text = ""
Text1.Text = ""
Text2.Text = ""
Text19.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text13.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text14.Text = ""
Text15.Text = ""
Else
MsgBox "Record Found"

Text16.Text = Val(DataEnvironment1.rsSearch.Fields("UserID"))

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
Else
MsgBox "Invalid Credentials"
End If

End Sub

Private Sub Command3_Click()
If (Text1.Text = "admin") And (Text2.Text = "admin123") Then
Form1.Show
Unload Me
Else
MsgBox "Invalid Credentials"
End If

End Sub
