VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTimeSheet 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5190
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   9285
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTimeSheet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   62
      Top             =   960
      Width           =   9015
      Begin VB.TextBox txtEmpNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Employee No:"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblName 
         Caption         =   "Employee Name here.."
         Height          =   375
         Left            =   3360
         TabIndex        =   64
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Employees Attendance"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   60
      Top             =   1920
      Width           =   9015
      Begin MSComctlLib.ListView Attendance 
         Height          =   2655
         Left            =   120
         TabIndex        =   61
         Top             =   360
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Day"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Time In"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Time Out"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Working Hours"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Over Time"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   11880
      TabIndex        =   0
      Top             =   1200
      Width           =   8775
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   8535
         Begin VB.Frame Frame4 
            Caption         =   "First Week"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1815
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   8295
            Begin VB.TextBox txtSun_IN 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   840
               TabIndex        =   55
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox txtSun_OUT 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   840
               TabIndex        =   54
               Top             =   1200
               Width           =   855
            End
            Begin VB.TextBox txtSat_OUT 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   7320
               TabIndex        =   45
               Top             =   1200
               Width           =   855
            End
            Begin VB.TextBox txtFri_OUT 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   6240
               TabIndex        =   44
               Top             =   1200
               Width           =   855
            End
            Begin VB.TextBox txtThu_OUT 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   5160
               TabIndex        =   43
               Top             =   1200
               Width           =   855
            End
            Begin VB.TextBox txtWed_OUT 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   4080
               TabIndex        =   42
               Top             =   1200
               Width           =   855
            End
            Begin VB.TextBox txtTue_OUT 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   3000
               TabIndex        =   41
               Top             =   1200
               Width           =   855
            End
            Begin VB.TextBox txtMon_OUT 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   1920
               TabIndex        =   40
               Top             =   1200
               Width           =   855
            End
            Begin VB.TextBox txtSat_IN 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   7320
               TabIndex        =   39
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox txtFri_IN 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   6240
               TabIndex        =   38
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox txtThu_IN 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   5160
               TabIndex        =   37
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox txtWed_IN 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   4080
               TabIndex        =   36
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox txtTue_IN 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   3000
               TabIndex        =   35
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox txtMon_IN 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   1920
               TabIndex        =   34
               Top             =   720
               Width           =   855
            End
            Begin VB.Label Label35 
               Alignment       =   2  'Center
               Caption         =   "Sunday"
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
               TabIndex        =   56
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label5 
               Caption         =   "OUT:"
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
               Left            =   240
               TabIndex        =   53
               Top             =   1200
               Width           =   495
            End
            Begin VB.Label Label4 
               Caption         =   "IN:"
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
               Left            =   360
               TabIndex        =   52
               Top             =   720
               Width           =   375
            End
            Begin VB.Label Label29 
               Alignment       =   2  'Center
               Caption         =   "Monday"
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
               Left            =   1920
               TabIndex        =   51
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label30 
               Alignment       =   2  'Center
               Caption         =   "Tuesday"
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
               Left            =   3000
               TabIndex        =   50
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label31 
               Alignment       =   2  'Center
               Caption         =   "Wednesday"
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
               Left            =   3960
               TabIndex        =   49
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label32 
               Alignment       =   2  'Center
               Caption         =   "Thursday"
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
               Left            =   5160
               TabIndex        =   48
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label33 
               Alignment       =   2  'Center
               Caption         =   "Friday"
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
               Left            =   6240
               TabIndex        =   47
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label34 
               Alignment       =   2  'Center
               Caption         =   "Saturday"
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
               Left            =   7320
               TabIndex        =   46
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Second Week"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1815
            Left            =   120
            TabIndex        =   12
            Top             =   2040
            Width           =   8295
            Begin VB.TextBox txtSun_OUT2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   840
               TabIndex        =   58
               Top             =   1200
               Width           =   855
            End
            Begin VB.TextBox txtSun_IN2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   840
               TabIndex        =   57
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox txtMon_IN2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   1920
               TabIndex        =   24
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox txtTue_IN2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   3000
               TabIndex        =   23
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox txtWed_IN2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   4080
               TabIndex        =   22
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox txtThu_IN2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   5160
               TabIndex        =   21
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox txtFri_IN2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   6240
               TabIndex        =   20
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox txtSat_IN2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   7320
               TabIndex        =   19
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox txtMon_OUT2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   1920
               TabIndex        =   18
               Top             =   1200
               Width           =   855
            End
            Begin VB.TextBox txtTue_OUT2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   3000
               TabIndex        =   17
               Top             =   1200
               Width           =   855
            End
            Begin VB.TextBox txtWed_OUT2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   4080
               TabIndex        =   16
               Top             =   1200
               Width           =   855
            End
            Begin VB.TextBox txtThu_OUT2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   5160
               TabIndex        =   15
               Top             =   1200
               Width           =   855
            End
            Begin VB.TextBox txtFri_OUT2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   6240
               TabIndex        =   14
               Top             =   1200
               Width           =   855
            End
            Begin VB.TextBox txtSat_OUT2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   7320
               TabIndex        =   13
               Top             =   1200
               Width           =   855
            End
            Begin VB.Label Label28 
               Alignment       =   2  'Center
               Caption         =   "Sunday"
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
               TabIndex        =   59
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label27 
               Alignment       =   2  'Center
               Caption         =   "Saturday"
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
               Left            =   7320
               TabIndex        =   32
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label26 
               Alignment       =   2  'Center
               Caption         =   "Friday"
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
               Left            =   6240
               TabIndex        =   31
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label25 
               Alignment       =   2  'Center
               Caption         =   "Thursday"
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
               Left            =   5160
               TabIndex        =   30
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label24 
               Alignment       =   2  'Center
               Caption         =   "Wednesday"
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
               Left            =   3960
               TabIndex        =   29
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label23 
               Alignment       =   2  'Center
               Caption         =   "Tuesday"
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
               Left            =   3000
               TabIndex        =   28
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label22 
               Alignment       =   2  'Center
               Caption         =   "Monday"
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
               Left            =   1920
               TabIndex        =   27
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label21 
               Caption         =   "OUT:"
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
               Left            =   240
               TabIndex        =   26
               Top             =   1200
               Width           =   495
            End
            Begin VB.Label Label20 
               Caption         =   "IN:"
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
               Left            =   360
               TabIndex        =   25
               Top             =   720
               Width           =   375
            End
         End
         Begin VB.TextBox txtTotal_Sun 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7440
            Locked          =   -1  'True
            TabIndex        =   10
            Text            =   "0.00"
            Top             =   3960
            Width           =   855
         End
         Begin VB.TextBox txtTotal_Sat 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   9
            Text            =   "0.00"
            Top             =   3960
            Width           =   855
         End
         Begin VB.TextBox txtTotal_Fri 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   8
            Text            =   "0.00"
            Top             =   3960
            Width           =   855
         End
         Begin VB.TextBox txtTotal_Thu 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   7
            Text            =   "0.00"
            Top             =   3960
            Width           =   855
         End
         Begin VB.TextBox txtTotal_Wed 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   6
            Text            =   "0.00"
            Top             =   3960
            Width           =   855
         End
         Begin VB.TextBox txtTotal_Tue 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   5
            Text            =   "0.00"
            Top             =   3960
            Width           =   855
         End
         Begin VB.TextBox txtTotal_Mon 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   4
            Text            =   "0.00"
            Top             =   3960
            Width           =   855
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Caption         =   "Working Hours:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   11
            Top             =   3960
            Width           =   735
         End
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "Time Sheet"
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
      TabIndex        =   2
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee's Time Sheet"
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Image cmdExit 
      Height          =   360
      Left            =   8880
      MousePointer    =   99  'Custom
      Picture         =   "frmTimeSheet.frx":000C
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmTimeSheet.frx":06F6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20700
   End
End
Attribute VB_Name = "frmTimeSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub

