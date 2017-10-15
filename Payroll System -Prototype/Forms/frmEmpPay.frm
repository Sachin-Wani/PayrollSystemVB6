VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEmpPay 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   9315
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmEmpPay.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4455
      Begin MSComCtl2.MonthView calStart 
         Height          =   2610
         Left            =   120
         TabIndex        =   36
         Top             =   1440
         Visible         =   0   'False
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   4604
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   16777215
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         StartOfWeek     =   114884609
         TitleBackColor  =   4210816
         TitleForeColor  =   16777215
         CurrentDate     =   41362
         MinDate         =   36526
      End
      Begin MSComCtl2.MonthView calEnd 
         Height          =   2610
         Left            =   1440
         TabIndex        =   37
         Top             =   1440
         Visible         =   0   'False
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   4604
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   16777215
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         StartOfWeek     =   114884609
         TitleBackColor  =   4210816
         TitleForeColor  =   16777215
         CurrentDate     =   41362
         MinDate         =   36526
      End
      Begin VB.TextBox txtDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   42
         ToolTipText     =   "Days"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtStart 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   28
         ToolTipText     =   "Starting Date"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtEnd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   29
         ToolTipText     =   "Ending Date"
         Top             =   960
         Width           =   1095
      End
      Begin VB.Frame Frame8 
         Height          =   1455
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Width           =   4215
         Begin VB.TextBox txtAbsent 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   40
            Text            =   "0"
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtLate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   39
            Text            =   "0"
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtRegHours 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   3
            Text            =   "0"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtOTHours 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   2
            Text            =   "0"
            Top             =   960
            Width           =   735
         End
         Begin PayrollSystem.jcbutton cmdTimeSheet 
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   0
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   661
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
            Caption         =   "Time Sheet"
            UseMaskCOlor    =   -1  'True
         End
         Begin VB.Label Label16 
            Caption         =   "Late:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   41
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label15 
            Caption         =   "Absences:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2400
            TabIndex        =   38
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "Hour Worked:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Overtime:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2400
            TabIndex        =   4
            Top             =   960
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5295
         Left            =   120
         TabIndex        =   6
         Top             =   2880
         Width           =   4215
         Begin VB.TextBox txtHourlyRate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   44
            Text            =   "0"
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtOTRate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   43
            Text            =   "0"
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtGross 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3240
            TabIndex        =   35
            Text            =   "0"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtBasic 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   34
            Text            =   "0"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtNetPay 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   9
            Text            =   "0"
            Top             =   4200
            Width           =   1815
         End
         Begin VB.CommandButton cmdCancel 
            Height          =   495
            Left            =   2160
            Picture         =   "frmEmpPay.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   4680
            Width           =   1455
         End
         Begin VB.CommandButton cmdProcess 
            Enabled         =   0   'False
            Height          =   495
            Left            =   600
            Picture         =   "frmEmpPay.frx":224E
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   4680
            Width           =   1455
         End
         Begin VB.Frame Frame6 
            Caption         =   "Deductions"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3015
            Left            =   120
            TabIndex        =   16
            Top             =   1080
            Width           =   3975
            Begin VB.TextBox txtTotal 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2040
               Locked          =   -1  'True
               TabIndex        =   21
               Text            =   "0"
               Top             =   2280
               Width           =   1335
            End
            Begin VB.TextBox txtOthers 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2040
               TabIndex        =   20
               Text            =   "0"
               Top             =   1800
               Width           =   1335
            End
            Begin VB.TextBox txtPH 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2040
               Locked          =   -1  'True
               TabIndex        =   19
               Text            =   "0"
               Top             =   1320
               Width           =   1095
            End
            Begin VB.TextBox txtPI 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2040
               Locked          =   -1  'True
               TabIndex        =   18
               Text            =   "0"
               Top             =   840
               Width           =   1095
            End
            Begin VB.TextBox txtSSS 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2040
               Locked          =   -1  'True
               TabIndex        =   17
               Text            =   "0"
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label10 
               Caption         =   "Total Deductions:"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   480
               TabIndex        =   26
               Top             =   2400
               Width           =   1455
            End
            Begin VB.Label Label9 
               Caption         =   "Other Deduction:"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   375
               Left            =   480
               TabIndex        =   25
               Top             =   1800
               Width           =   1455
            End
            Begin VB.Label Label8 
               Caption         =   "Phil Health:"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   960
               TabIndex        =   24
               Top             =   1320
               Width           =   1095
            End
            Begin VB.Label Label7 
               Caption         =   "Pag-ibig:"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1200
               TabIndex        =   23
               Top             =   840
               Width           =   855
            End
            Begin VB.Label Label6 
               Caption         =   "SSS:"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1560
               TabIndex        =   22
               Top             =   360
               Width           =   375
            End
         End
         Begin VB.Label Label19 
            Caption         =   "Hourly Rate:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label18 
            Caption         =   "OT Rate:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2280
            TabIndex        =   45
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label12 
            Caption         =   "Basic Pay:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   33
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "Gross Pay:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            TabIndex        =   32
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            Caption         =   "Net Pay:"
            BeginProperty Font 
               Name            =   "Candara"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   375
            Left            =   480
            TabIndex        =   10
            Top             =   4200
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   4215
         Begin VB.TextBox txtEmpNo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label3 
            Caption         =   "Employee No:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   13
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Label Label5 
         Caption         =   "End:"
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
         Left            =   2040
         TabIndex        =   31
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Start:"
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
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   615
      End
   End
   Begin VB.Image cmdExit 
      Height          =   360
      Left            =   4320
      MousePointer    =   99  'Custom
      Picture         =   "frmEmpPay.frx":4490
      Top             =   0
      Width           =   360
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Process Employee's Payroll"
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
      TabIndex        =   15
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "Payslip"
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
      Left            =   240
      TabIndex        =   14
      Top             =   0
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmEmpPay.frx":4B7A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15465
   End
End
Attribute VB_Name = "frmEmpPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SSS
Dim PI
Dim PH
Dim EmpPos
Dim EmpLName As String
Dim EmpFName As String
Dim EmpMName As String
Dim EmpTHours
Dim EmpOTHours
Dim EmpLHours
Dim EmpPresent As Integer

Sub ClearAll()
txtLate.Text = ""
txtAbsent.Text = ""
txtRegHours.Text = ""
txtOTHours.Text = ""
txtGross.Text = ""
txtNetPay.Text = ""
End Sub



' === Time Sheet =================================
Sub EmpTimeSheet()
Dim Sunday As Integer
Dim tsheet

Sunday = 2

SQL = "SELECT * FROM DTR WHERE WorkDate BETWEEN '" & Format(txtStart.Text, "mm/dd/yyyy") & "' AND '" & Format(txtEnd.Text, "mm/dd/yyyy") & "'"
connectDB
With RS
    While Not .EOF
        If txtEmpNo.Text = .Fields!EmployeeNo Then
            Set tsheet = frmTimeSheet.Attendance.ListItems.Add(, , UCase(Format(.Fields!WorkDate, "dddd")))
                tsheet.SubItems(1) = .Fields!WorkDate
                tsheet.SubItems(2) = .Fields!TimeIn
                tsheet.SubItems(3) = .Fields!TimeOut
                tsheet.SubItems(4) = .Fields!WorkingHours
                tsheet.SubItems(5) = .Fields!OverTime
            .MoveNext
        Else
            .MoveNext
        End If
    Wend
End With
connClose
End Sub



Sub EmpName()
SQL = "SELECT * FROM Employees"
connectDB
With RS
    .MoveFirst
    While Not .EOF
        If txtEmpNo.Text = .Fields!EmployeeNo Then
            frmTimeSheet.lblName.Caption = .Fields!LastName & ", " & "" & .Fields!FirstName & " " & .Fields!MiddleName
            connClose
            Exit Sub
        Else
            .MoveNext
        End If
    Wend
End With
connClose
End Sub


' === Deductions =================================

Sub EmpDeductions()
'SQL = "SELECT * FROM Deductions"
'connectDB
'With RS
'    If txtEmpNo.ToolTipText = "Regular" Then
'        txtSSS.Text = .Fields!SSS
'        txtPI.Text = .Fields!Pagibig
'        txtPH.Text = .Fields!PhilHealth
'    Else
'        txtRegHours.Text = 0
'        txtSSS.Text = 0
'        txtPI.Text = 0
'        txtPH.Text = 0
'    End If
'End With
'connClose
End Sub

' === Basic Pay =================================

Sub EmpBPay()
Dim EmpStatus
Dim BasicPay

SQL = "SELECT * FROM Employees WHERE EmployeeNo = '" & txtEmpNo.Text & "'"
connectDB
With RS
    EmpLName = .Fields!LastName
    EmpFName = .Fields!FirstName
    EmpMName = .Fields!MiddleName
    EmpPos = .Fields!Position
    EmpStatus = .Fields!Status
    txtSSS.Text = IIf(IsNull(.Fields!SSS), "0", .Fields!SSS)
    txtPI.Text = IIf(IsNull(.Fields!Pagibig), "0", .Fields!Pagibig)
    txtPH.Text = IIf(IsNull(.Fields!PhilHealth), "0", .Fields!PhilHealth)
End With
connClose

SQL = "SELECT * FROM Positions"
connectDB
With RS
    If EmpStatus = "Contractual" Then
        txtEmpNo.ToolTipText = "Contractual"
        txtBasic.Text = .Fields!ConRate
        txtGross.Enabled = False
        txtOTHours.Enabled = False
        txtLate.Enabled = False
        txtAbsent.Enabled = False
        txtRegHours.Enabled = False
        txtSSS.Enabled = False
        txtPI.Enabled = False
        txtPH.Enabled = False
    ElseIf EmpStatus = "Regular" Then
        txtEmpNo.ToolTipText = "Regular"
        txtBasic.Text = .Fields!RegRate
    End If
    txtHourlyRate.Text = Val(txtBasic.Text) / 8
End With
connClose
End Sub

Sub Computations()
If txtEmpNo.ToolTipText = "Regular" Then
    ' Gross Pay
    txtGross.Text = (Val(txtBasic.Text) * Val(txtRegHours.Text) + Val(txtOTHours.Text))
    ' Deductions
    txtTotal.Text = (Val(txtSSS.Text) + Val(txtPI.Text) + Val(txtPH.Text) + Val(txtOthers.Text))
    ' Net Pay
    txtNetPay.Text = (Val(txtGross.Text) - Val(txtTotal.Text))
    txtOTRate.Text = ((Val(txtBasic.Text) / 8) * 1.1) * Val(txtRegHours.Text)
Else
    txtTotal.Text = txtOthers.Text
    txtNetPay.Text = (Val(txtBasic) * Val(txtRegHours.Text) - Val(txtTotal.Text))
End If
End Sub

Private Sub calEnd_DateClick(ByVal DateClicked As Date)
txtEnd.Text = calEnd.Value
calEnd.Visible = False

If calStart.Value <= calEnd.Value Then
    txtDays.Text = calEnd.Value - calStart.Value
Else
    txtDays.Text = 0
End If

' Present

Dim Present As Integer

SQL = "SELECT * FROM DTR WHERE WorkDate BETWEEN '" & Format(txtStart.Text, "mm/dd/yyyy") & "' AND '" & Format(txtEnd.Text, "mm/dd/yyyy") & "'"
connectDB
With RS
    While Not .EOF
        If txtEmpNo.Text = .Fields!EmployeeNo Then
            EmpTHours = EmpTHours + Val(.Fields!WorkingHours)
            EmpOTHours = EmpOTHours + Val(.Fields!OverTime)
            EmpLHours = EmpLHours + Val(.Fields!Late)
            Present = Present + 1
            .MoveNext
        Else
            .MoveNext
        End If
    Wend
    'If Present > txtDays.Text And txtDays.Text <> 0 Then
    '    txtAbsent.Text = Present - Val(txtDays.Text)
    If txtDays.Text > Present Then
        txtAbsent.Text = Val(txtDays.Text) - Present
        If EmpStatus = "Regular" Then
            txtRegHours.Text = Present * Val(txtBasic.Text)
            txtLate.Text = EmpLHours
            EmpLHours = 0
            EmpTHours = 0
            EmpOTHours = 0
        Else
            txtRegHours.Text = EmpTHours
            txtNetPay.Text = (Val(txtBasic.Text) * Val(txtRegHours.Text))
            EmpLHours = 0
            EmpTHours = 0
            EmpOTHours = 0
        End If
    Else
        txtAbsent.Text = 0
    End If
End With
connClose
End Sub

Private Sub calStart_DateClick(ByVal DateClicked As Date)
Dim EmpWDays As Integer

SQL = "SELECT * FROM WorkingDays"
connectDB
    EmpWDays = Val(RS.Fields!WDays)
connClose

cmdProcess.Enabled = True
If txtStart.Text <> "" Then
    If calStart.Value = txtStart.Text Then Exit Sub
End If

txtStart.Text = calStart.Value
txtEnd.Text = calStart.Value + (EmpWDays - 1)
calEnd.Value = txtEnd.Text
calStart.Visible = False
txtEnd.Enabled = True

If calStart.Value <= calEnd.Value Then
    txtDays.Text = calEnd.Value - calStart.Value
Else
    txtDays.Text = 0
End If

Dim Present As Integer

SQL = "SELECT * FROM DTR WHERE WorkDate BETWEEN '" & Format(txtStart.Text, "mm/dd/yyyy") & "' AND '" & Format(txtEnd.Text, "mm/dd/yyyy") & "'"
connectDB
With RS
    While Not .EOF
        If .Fields!EmployeeNo = txtEmpNo.Text Then
            EmpTHours = EmpTHours + Val(.Fields!WorkingHours)
            EmpOTHours = EmpOTHours + Val(.Fields!OverTime)
            EmpLHours = EmpLHours + Val(.Fields!Late)
            Present = Present + 1
            .MoveNext
        Else
            .MoveNext
        End If
    Wend
    If txtDays.Text > Present Then
        txtAbsent.Text = Val(txtDays.Text) - Present
        If txtEmpNo.ToolTipText = "Regular" Then
            txtRegHours.Text = Val(EmpTHours) 'Present * Val(txtBasic.Text)
            txtOTHours.Text = Val(EmpOTHours)
            txtLate.Text = Val(EmpLHours)
            EmpLHours = 0
            EmpTHours = 0
            EmpOTHours = 0
            Computations
        Else
            txtRegHours.Text = EmpTHours
            txtNetPay.Text = (Val(txtBasic.Text) * Val(txtRegHours.Text))
            EmpLHours = 0
            EmpTHours = 0
            EmpOTHours = 0
        End If
    Else
        txtAbsent.Text = 0
    End If
    
    'If EmpStatus = "Regular" Then
    '    txtAbsent.Text = Val(txtDays.Text) - Present
    '    txtRegHours.Text = Present * Val(txtBasic.Text)
    '    txtOTHours.Text = EmpOTHours
    'Else
    '    txtRegHours.Text = EmpTHours
    '    txtNetPay.Text = (Val(txtBasic.Text) * Val(txtRegHours.Text))
    '    EmpTHours = 0
    '    EmpOTHours = 0
    'End If
End With
connClose
End Sub

Private Sub cmdCancel_Click()
frmMainMenu.Enabled = True
Unload Me
End Sub

Private Sub cmdExit_Click()
frmMainMenu.Enabled = True
'frmMainMenu.mnuEmpRec.Enabled = False
'frmMainMenu.mnuHide.Enabled = True
'frmMainMenu.lblEmpRec.Visible = True
'frmMainMenu.lblEmpCount.Visible = True
'frmMainMenu.lblCount.Visible = True
'frmMainMenu.imgBar.Visible = True
'frmMainMenu.Records.Visible = True
Unload Me
End Sub

Private Sub cmdProcess_Click()
On Error Resume Next
SQL = "SELECT * FROM PaySlip"
connectDB
With RS
    While Not .EOF
        If txtEmpNo.Text = .Fields!EmployeeNo And .Fields!DateProcessed = Format(Now, "mm/dd/yyyy") Then
            If MsgBox("This Payroll has been processed before. Are you sure to overwrite the existing record?", vbYesNo + vbExclamation, "Overwrite Payslip") = vbYes Then
                .Fields!EmployeeNo = txtEmpNo.Text
                .Fields!LastName = EmpLName
                .Fields!FirstName = EmpFName
                .Fields!MiddleName = EmpMName
                .Fields!Position = EmpPos
                .Fields!Status = txtEmpNo.ToolTipText
                .Fields!WorkingHours = txtRegHours.Text
                .Fields!OverTime = txtOTHours.Text
                .Fields!Late = txtLate.Text
                .Fields!Absences = txtAbsent.Text
                .Fields!BasicPay = txtBasic.Text
                .Fields!Grosspay = txtGross.Text
                .Fields!SSS = txtSSS.Text
                .Fields!Pagibig = txtPI.Text
                .Fields!PhilHealth = txtPH.Text
                .Fields!Others = txtOthers.Text
                .Fields!Deductions = txtTotal.Text
                .Fields!DateTransaction = txtStart.Text & " to " & txtEnd.Text
                .Fields!DateProcessed = Format(Now, "mm/dd/yyyy")
                .Fields!NetPay = txtNetPay.Text
                .Update
                MsgBox "PaySlip processed successfully!"
            End If
            connClose
            Exit Sub
        Else
            .MoveNext
        End If
    Wend
    .AddNew
    .Fields!EmployeeNo = txtEmpNo.Text
    .Fields!LastName = EmpLName
    .Fields!FirstName = EmpFName
    .Fields!MiddleName = EmpMName
    .Fields!Position = EmpPos
    .Fields!Status = txtEmpNo.ToolTipText
    .Fields!WorkingHours = txtRegHours.Text
    .Fields!OverTime = txtOTHours.Text
    .Fields!Late = txtLate.Text
    .Fields!Absences = txtAbsent.Text
    .Fields!BasicPay = Val(txtBasic.Text)
    .Fields!Grosspay = Val(txtGross.Text)
    .Fields!SSS = txtSSS.Text
    .Fields!Pagibig = txtPI.Text
    .Fields!PhilHealth = txtPH.Text
    .Fields!Others = Val(txtOthers.Text)
    .Fields!Deductions = Val(txtTotal.Text)
    .Fields!DateTransaction = txtStart.Text & " to " & txtEnd.Text
    .Fields!DateProcessed = Format(Now, "mm/dd/yyyy")
    .Fields!NetPay = Val(txtNetPay.Text)
    MsgBox "PaySlip processed successfully!"
    .Update
    'MsgBox "An error has occurred! Try to process again.", vbOKOnly + vbInformation
End With
connClose
End Sub

Private Sub cmdTimeSheet_Click()
EmpName
EmpTimeSheet
frmTimeSheet.txtEmpNo.Text = txtEmpNo.Text
frmTimeSheet.Show
End Sub

Private Sub Form_Load()
txtEmpNo.Text = EmpNumb
frmMainMenu.Enabled = False

EmpBPay
'EmpDeductions
End Sub


Private Sub txtBasic_Change()
Computations
End Sub

Private Sub txtEnd_Click()
txtEnd.BackColor = &H80000005
calEnd.Visible = True
End Sub

Private Sub txtGross_Change()
Computations
End Sub

Private Sub txtNetPay_Change()
On Error Resume Next
txtTotal.Text = Format(txtTotal.Text, "#,##0.#0")
End Sub

Private Sub txtOthers_Change()
Computations
End Sub

Private Sub txtOthers_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = vbKeyBack) Or (KeyAscii = 46) Then
    Exit Sub
Else
    KeyAscii = 0
End If
End Sub

Private Sub txtPH_Change()
Computations
End Sub

Private Sub txtPI_Change()
Computations
End Sub

Private Sub txtSSS_Change()
Computations
End Sub

Private Sub txtStart_Click()
txtStart.BackColor = &H80000005
calStart.Visible = True
End Sub

Private Sub txtTotal_Change()
Computations
End Sub
