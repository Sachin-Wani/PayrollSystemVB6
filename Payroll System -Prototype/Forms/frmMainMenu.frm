VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmMainMenu 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nippon Express Philippines Corporation"
   ClientHeight    =   10710
   ClientLeft      =   2820
   ClientTop       =   3180
   ClientWidth     =   15240
   Icon            =   "frmMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10710
   ScaleMode       =   0  'User
   ScaleWidth      =   15033.65
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList EmpStatus 
      Left            =   3120
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMenu.frx":4888A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMenu.frx":48964
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   20640
      TabIndex        =   20
      Top             =   10320
      Width           =   20640
      Begin PayrollSystem.jcbutton cmdCP 
         Height          =   495
         Left            =   -120
         TabIndex        =   21
         Top             =   0
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16776960
         Caption         =   "Admin Panel"
         Picture         =   "frmMainMenu.frx":48A3D
         PictureAlign    =   0
         UseMaskCOlor    =   -1  'True
         MaskColor       =   16777215
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   0
         Picture         =   "frmMainMenu.frx":4A22F
         Stretch         =   -1  'True
         Top             =   0
         Width           =   20685
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   0
      ScaleHeight     =   1785
      ScaleWidth      =   2745
      TabIndex        =   3
      Top             =   1080
      Width           =   2775
      Begin VB.Timer tmrTime 
         Interval        =   50
         Left            =   2280
         Top             =   0
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   0
         Picture         =   "frmMainMenu.frx":5E193
         ScaleHeight     =   495
         ScaleWidth      =   8895
         TabIndex        =   4
         Top             =   0
         Width           =   8895
         Begin VB.Image Image3 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   120
            Picture         =   "frmMainMenu.frx":7D617
            Stretch         =   -1  'True
            Top             =   120
            Width           =   285
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Welcome Admin!"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   255
            Left            =   0
            TabIndex        =   5
            Top             =   120
            Width           =   2895
         End
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "TIME:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "DATE:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblTime 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "Current Time"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   840
         TabIndex        =   7
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Current Date"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   840
         TabIndex        =   6
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404000&
         Height          =   1095
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   2535
      End
   End
   Begin MSComctlLib.ListView Records 
      Height          =   8295
      Left            =   2880
      TabIndex        =   1
      ToolTipText     =   "Select an Action (Double click)"
      Top             =   1920
      Visible         =   0   'False
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   14631
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "EmpStatus"
      ForeColor       =   0
      BackColor       =   -2147483648
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Employee No"
         Object.Width           =   3040
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Employee Name"
         Object.Width           =   5364
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Gender"
         Object.Width           =   2575
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Age"
         Object.Width           =   2575
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Birthday"
         Object.Width           =   2575
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Address"
         Object.Width           =   5364
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Contact No"
         Object.Width           =   2575
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Position"
         Object.Width           =   2575
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Date Hired"
         Object.Width           =   2575
      EndProperty
      Picture         =   "frmMainMenu.frx":7E031
   End
   Begin VB.PictureBox CtrlP 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   -960
      ScaleHeight     =   5775
      ScaleWidth      =   3855
      TabIndex        =   2
      Top             =   5160
      Visible         =   0   'False
      Width           =   3855
      Begin PayrollSystem.jcbutton cmdEdit 
         Height          =   975
         Left            =   2280
         TabIndex        =   10
         Top             =   1920
         Width           =   1095
         _ExtentX        =   2143
         _ExtentY        =   1720
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
         Caption         =   "Edit"
         Picture         =   "frmMainMenu.frx":748BB5
         PictureAlign    =   5
         UseMaskCOlor    =   -1  'True
         MaskColor       =   16777215
      End
      Begin PayrollSystem.jcbutton cmdAdd 
         Height          =   975
         Left            =   1080
         TabIndex        =   11
         Top             =   840
         Width           =   1095
         _ExtentX        =   2143
         _ExtentY        =   1720
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14935011
         Caption         =   "Add"
         Picture         =   "frmMainMenu.frx":7494E3
         PictureAlign    =   5
         UseMaskCOlor    =   -1  'True
         MaskColor       =   16777215
      End
      Begin PayrollSystem.jcbutton cmdSearch 
         Height          =   975
         Left            =   2280
         TabIndex        =   12
         Top             =   840
         Width           =   1095
         _ExtentX        =   2143
         _ExtentY        =   1720
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
         Caption         =   "Search"
         Picture         =   "frmMainMenu.frx":749FFD
         PictureAlign    =   5
         UseMaskCOlor    =   -1  'True
         MaskColor       =   16777215
      End
      Begin PayrollSystem.jcbutton cmdDelete 
         Height          =   975
         Left            =   1080
         TabIndex        =   13
         Top             =   1920
         Width           =   1095
         _ExtentX        =   2143
         _ExtentY        =   1720
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
         Caption         =   "Deactivate"
         Picture         =   "frmMainMenu.frx":74AC4F
         PictureAlign    =   5
         UseMaskCOlor    =   -1  'True
         MaskColor       =   16777215
      End
      Begin PayrollSystem.jcbutton cmdPrint 
         Height          =   975
         Left            =   1080
         TabIndex        =   14
         Top             =   3000
         Width           =   1095
         _ExtentX        =   2143
         _ExtentY        =   1720
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
         Caption         =   "Print"
         Picture         =   "frmMainMenu.frx":74AE97
         PictureAlign    =   5
         UseMaskCOlor    =   -1  'True
         MaskColor       =   16777215
      End
      Begin PayrollSystem.jcbutton cmdSettings 
         Height          =   975
         Left            =   2280
         TabIndex        =   15
         Top             =   3000
         Width           =   1095
         _ExtentX        =   2143
         _ExtentY        =   1720
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
         Caption         =   "Settings"
         Picture         =   "frmMainMenu.frx":74BAE9
         PictureAlign    =   5
         UseMaskCOlor    =   -1  'True
         MaskColor       =   16777215
      End
      Begin PayrollSystem.jcbutton cmdAbout 
         Height          =   975
         Left            =   1080
         TabIndex        =   16
         Top             =   4080
         Width           =   1095
         _ExtentX        =   2143
         _ExtentY        =   1720
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
         Caption         =   "About"
         Picture         =   "frmMainMenu.frx":74C73B
         PictureAlign    =   5
         UseMaskCOlor    =   -1  'True
         MaskColor       =   16777215
      End
      Begin PayrollSystem.jcbutton cmdClose 
         Height          =   975
         Left            =   2280
         TabIndex        =   17
         Top             =   4080
         Width           =   1095
         _ExtentX        =   2143
         _ExtentY        =   1720
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
         BackColor       =   -2147483633
         Caption         =   "Shut Down"
         Picture         =   "frmMainMenu.frx":74D255
         PictureAlign    =   5
         UseMaskCOlor    =   -1  'True
         MaskColor       =   16777215
      End
      Begin PayrollSystem.jcbutton cmdRecords 
         Height          =   495
         Left            =   1080
         TabIndex        =   19
         Top             =   120
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14935011
         Caption         =   "View Records"
         Picture         =   "frmMainMenu.frx":74DD6F
         PictureAlign    =   0
         UseMaskCOlor    =   -1  'True
         MaskColor       =   16777215
      End
      Begin VB.Shape l2 
         BorderColor     =   &H00C0C000&
         FillColor       =   &H00404000&
         FillStyle       =   0  'Solid
         Height          =   5295
         Left            =   720
         Top             =   720
         Width           =   2775
      End
      Begin VB.Shape l1 
         BorderColor     =   &H00C0C000&
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   6735
         Left            =   0
         Top             =   0
         Width           =   3735
      End
   End
   Begin VB.PictureBox EmpPayrollInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2145
      ScaleWidth      =   2745
      TabIndex        =   33
      Top             =   5040
      Width           =   2775
      Begin VB.PictureBox Picture6 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   0
         Picture         =   "frmMainMenu.frx":74E3A1
         ScaleHeight     =   495
         ScaleWidth      =   8895
         TabIndex        =   34
         Top             =   0
         Width           =   8895
         Begin VB.Image Image4 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   120
            Picture         =   "frmMainMenu.frx":76D825
            Stretch         =   -1  'True
            Top             =   120
            Width           =   285
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Payroll Details"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   255
            Left            =   600
            TabIndex        =   35
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.Label lblInactive 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   2040
         TabIndex        =   42
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Inactive:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblActive 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   2040
         TabIndex        =   40
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Active:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   495
         Left            =   240
         TabIndex        =   39
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   2040
         TabIndex        =   38
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblEmpCount 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Employees:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label12 
         BackColor       =   &H00404000&
         Height          =   1455
         Left            =   120
         TabIndex        =   37
         Top             =   600
         Width           =   2535
      End
   End
   Begin VB.PictureBox QSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1665
      ScaleWidth      =   2745
      TabIndex        =   23
      Top             =   3120
      Width           =   2775
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
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
         Left            =   360
         TabIndex        =   31
         Top             =   2160
         Width           =   2055
      End
      Begin VB.ComboBox cboFilter 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmMainMenu.frx":76DC42
         Left            =   360
         List            =   "frmMainMenu.frx":76DC44
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1080
         Width           =   2055
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   0
         Picture         =   "frmMainMenu.frx":76DC46
         ScaleHeight     =   495
         ScaleWidth      =   8895
         TabIndex        =   24
         Top             =   0
         Width           =   8895
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Quick Search"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   255
            Left            =   600
            TabIndex        =   25
            Top             =   120
            Width           =   1455
         End
         Begin VB.Image Image1 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   120
            Picture         =   "frmMainMenu.frx":78D0CA
            Stretch         =   -1  'True
            Top             =   120
            Width           =   285
         End
      End
      Begin VB.ComboBox cboPos 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Filter by:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label10 
         BackColor       =   &H00404000&
         Height          =   975
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404000&
         Height          =   1095
         Left            =   120
         TabIndex        =   29
         Top             =   1680
         Width           =   2535
      End
   End
   Begin VB.Label lblEmpRec 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Records"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3120
      TabIndex        =   22
      Top             =   1200
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Image imgLogo 
      Height          =   840
      Left            =   0
      Picture         =   "frmMainMenu.frx":78DD0C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7830
   End
   Begin VB.Image imgBar 
      Height          =   735
      Left            =   2880
      Picture         =   "frmMainMenu.frx":7A654E
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   12345
   End
   Begin VB.Label White 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20655
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuDTR 
         Caption         =   "Daily Time Record"
         Begin VB.Menu mnuVDTR 
            Caption         =   "View DTR"
         End
         Begin VB.Menu mnuModify 
            Caption         =   "Modify DTR"
         End
      End
      Begin VB.Menu mnuEmpRec 
         Caption         =   "View Employee Records"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "Hide Employee Records"
      End
      Begin VB.Menu mnuPaySlip 
         Caption         =   "PaySlip"
      End
      Begin VB.Menu mnuSet 
         Caption         =   "Settings"
      End
   End
   Begin VB.Menu mnuSys 
      Caption         =   "System"
      Begin VB.Menu mnuLock 
         Caption         =   "Lock"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuNotepad 
         Caption         =   "Notepad"
      End
      Begin VB.Menu mnuCalculator 
         Caption         =   "Calculator"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
      Begin VB.Menu mnuProgram 
         Caption         =   "Program"
      End
   End
   Begin VB.Menu mnuOff 
      Caption         =   "Log off"
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ustat

Sub Filter()
cboFilter.Clear
cboFilter.AddItem "Name"
cboFilter.AddItem "Gender"
cboFilter.AddItem "Age"
cboFilter.AddItem "Status"
End Sub

Sub EmpRecs()
Dim X

Records.ListItems.Clear

With RS
    While .EOF = False
        Set X = Records.ListItems.Add(, , .Fields!EmployeeNo)
            X.SubItems(1) = .Fields!LastName + ", " + .Fields!FirstName + " " + Left(.Fields!MiddleName, 1)
            X.SubItems(2) = .Fields!Gender
            X.SubItems(3) = .Fields!Age
            X.SubItems(4) = .Fields!Birthday
            X.SubItems(5) = .Fields!Address
            X.SubItems(6) = .Fields!ContactNo
            X.SubItems(7) = .Fields!Position
            X.SubItems(8) = .Fields!DateHired
        
            If .Fields!CurrentStatus = "Active" Then
                X.SmallIcon = 1
            Else
                X.SmallIcon = 2
            End If
            
            .MoveNext
            
    Wend
End With
End Sub


Sub PayrollDetails()
' Record Count
SQL = "SELECT * FROM Employees"
connectDB
lblCount.Caption = RS.RecordCount
connClose

' Active/Inactive Employees
SQL = "SELECT * FROM Employees WHERE CurrentStatus = 'Active'"
connectDB
lblActive.Caption = RS.RecordCount
lblInactive.Caption = (Val(lblCount.Caption) - Val(lblActive.Caption))
connClose
End Sub

Private Sub cboFilter_Click()
QSearch.Height = 2895
EmpPayrollInfo.Top = 6240
If CtrlP.Visible = True Then CtrlP.Visible = False

Select Case cboFilter.Text:
Case "Name":
    txtSearch.Visible = True
    cboPos.Visible = False
    lblName.Caption = "Employee Name:"
    ustat = "EmployeeName"
Case "Gender":
    txtSearch.Visible = False
    lblName.Caption = "Gender:"
    With cboPos
        .Visible = True
        .Clear
        .AddItem "All"
        .AddItem "Male"
        .AddItem "Female"
    End With
Case "Age":
    txtSearch.Visible = True
    cboPos.Visible = False
    lblName.Caption = "Age:"
    ustat = "EmployeeName"
Case "Status":
    txtSearch.Visible = False
    lblName.Caption = "Status:"
    With cboPos
        .Visible = True
        .Clear
        .AddItem "All"
        .AddItem "Active"
        .AddItem "Inactive"
    End With
End Select
End Sub

Private Sub cboFilter_GotFocus()
If CtrlP.Visible = True Then CtrlP.Visible = False
If Records.Visible = False Then
    cmdRecords.Caption = "Hide Records"
    lblEmpRec.Visible = True
    lblEmpCount.Visible = True
    lblCount.Visible = True
    imgBar.Visible = True
    Records.Visible = True
End If
End Sub


Private Sub cboPos_Click()
If cboPos.Text = "" Then
    SQL = "SELECT * FROM Employees"
    connOpen
    Me.EmpRecs
    connClose
    Exit Sub
ElseIf cboPos.Text = "All" Then
    SQL = "SELECT * FROM Employees"
    connOpen
    Me.EmpRecs
    connClose
    Exit Sub
ElseIf cboFilter.Text = "Gender" Then
    SQL = "SELECT * FROM Employees WHERE Gender LIKE '" & Replace(cboPos.Text, "'", "") & "%'"
ElseIf cboFilter.Text = "Status" Then
    SQL = "SELECT * FROM Employees WHERE CurrentStatus = '" & cboPos.Text & "'"
End If

connOpen

Records.ListItems.Clear

With RS
    Do While .EOF = False
        Set rec = Records.ListItems.Add(, , .Fields!EmployeeNo)
            rec.SubItems(1) = .Fields!LastName + ", " + .Fields!FirstName + " " + Left(.Fields!MiddleName, 1)
            rec.SubItems(2) = .Fields!Gender
            rec.SubItems(3) = .Fields!Age
            rec.SubItems(4) = .Fields!Birthday
            rec.SubItems(5) = .Fields!Address
            rec.SubItems(6) = .Fields!ContactNo
            rec.SubItems(7) = .Fields!Position
            rec.SubItems(8) = .Fields!DateHired
        
        If .Fields!CurrentStatus = "Active" Then
            rec.SmallIcon = 1
        Else
            rec.SmallIcon = 2
        End If
            
        .MoveNext
        
    Loop
End With
connClose
End Sub

Private Sub cmdAbout_Click()
frmAbout.Show
End Sub

Private Sub cmdAdd_Click()
frmAddRecord.Show
End Sub

Private Sub cmdClose_Click()
frmShutdown.Show
'If MsgBox("Are you sure to close the Program?", vbQuestion + vbYesNo, "Close Program") = vbYes Then
'    End
'End If
End Sub

Private Sub cmdCP_Click()
If CtrlP.Visible = False Then
    QSearch.Height = 1695
    EmpPayrollInfo.Top = 5040
    'cboFilter.Text = ""
    CtrlP.Visible = True
    'l1.Visible = True
    'l2.Visible = True
    'cmdRecords.Visible = True
    'CP.Visible = True
    'cmdAdd.Visible = True
    'cmdEdit.Visible = True
    'cmdDelete.Visible = True
    'cmdSearch.Visible = True
    'cmdPrint.Visible = True
    'cmdSettings.Visible = True
    'cmdAbout.Visible = True
    'cmdClose.Visible = True
Else
    CtrlP.Visible = False
    'l1.Visible = False
    'l2.Visible = False
    'cmdRecords.Visible = False
    'CP.Visible = False
    'cmdAdd.Visible = False
    'cmdEdit.Visible = False
    'cmdDelete.Visible = False
    'cmdSearch.Visible = False
    'cmdPrint.Visible = False
    'cmdSettings.Visible = False
    'cmdAbout.Visible = False
    'cmdClose.Visible = False
End If
    
End Sub

Private Sub cmdDelete_Click()
frmDeactivateRecord.Show
End Sub

Private Sub cmdEdit_Click()
frmEditRecord.Show
End Sub

Private Sub cmdPrint_Click()
frmPrint.Show
End Sub

Private Sub cmdRecords_Click()
If lblEmpRec.Visible = False Then
    mnuEmpRec.Enabled = False
    mnuHide.Enabled = True
    lblEmpRec.Visible = True
    imgBar.Visible = True
    Records.Visible = True
    cmdRecords.Caption = "Hide Records"
Else
    mnuEmpRec.Enabled = True
    mnuHide.Enabled = False
    lblEmpRec.Visible = False
    imgBar.Visible = False
    Records.Visible = False
    cmdRecords.Caption = "View Records"
End If
End Sub

Private Sub cmdSearch_Click()
frmSearch.Show
End Sub

Private Sub cmdSettings_Click()
frmSettings.Show
End Sub





Private Sub Form_Load()
imgBar.Width = (Me.ScaleWidth - 2900)
Records.Width = (Me.ScaleWidth - 2900)
Filter
mnuHide.Enabled = False

PayrollDetails

' Display all Employees Records
SQL = "SELECT * FROM Employees"
connOpen
Me.EmpRecs
connClose
End Sub


Private Sub Image5_Click()
Unload Me
End Sub

Private Sub mnuCalculator_Click()
Shell "calc.exe", vbNormalFocus
End Sub


Private Sub mnuEmpRec_Click()
cmdRecords.Caption = "Hide Records"
mnuHide.Enabled = True
mnuEmpRec.Enabled = False
lblEmpRec.Visible = True
'lblEmpCount.Visible = True
'lblCount.Visible = True
imgBar.Visible = True
Records.Visible = True
End Sub


Private Sub mnuHide_Click()
cmdRecords.Caption = "View Records"
mnuEmpRec.Enabled = True
lblEmpRec.Visible = False
'lblEmpCount.Visible = False
'lblCount.Visible = False
imgBar.Visible = False
Records.Visible = False
mnuHide.Enabled = False
End Sub

Private Sub mnuLock_Click()
frmSysLock.Show
End Sub

Private Sub mnuModify_Click()
frmDTRHistory.Show
End Sub

Private Sub mnuNotepad_Click()
Shell "notepad.exe", vbNormalFocus
End Sub

Private Sub mnuOff_Click()
frmLogIn.Show
Unload Me
End Sub

Private Sub mnuPaySlip_Click()
frmPaySlip.Show
End Sub

Private Sub mnuProgram_Click()
frmAbout.Show
End Sub

Private Sub mnuSet_Click()
frmSettings.Show
End Sub


Private Sub mnuVDTR_Click()
dlgDTR.lblUserLoc.Caption = "Admin Panel"
dlgDTR.Show
End Sub

Private Sub Records_DblClick()
'mnuEmpRec.Enabled = True
'mnuHide.Enabled = False
'lblEmpRec.Visible = False
'lblEmpCount.Visible = False
'lblCount.Visible = False
'imgBar.Visible = False
'Records.Visible = False
EmpNumb = Records.SelectedItem
frmActions.lblName.Caption = Records.SelectedItem.SubItems(1)
frmActions.lblStatus.Caption = IIf(Records.SelectedItem.SmallIcon = 1, "Active", "Inactive")
frmActions.lblEmpNo.Caption = EmpNumb
If Records.SelectedItem.SmallIcon = 1 Then
    frmActions.lblStatus.ForeColor = &HC000&
    frmActions.cmdPaySlip.Enabled = True
    frmActions.cmdActivate.Visible = False
    frmActions.cmdDeactivate.Visible = True
Else
    frmActions.lblStatus.ForeColor = &HFF&
    frmActions.cmdPaySlip.Enabled = False
    frmActions.cmdDeactivate.Visible = False
    frmActions.cmdActivate.Visible = True
End If
frmActions.Show
End Sub

Private Sub Records_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'mnuEmpRec.Enabled = True
    'mnuHide.Enabled = False
    'lblEmpRec.Visible = False
    'lblEmpCount.Visible = False
    'lblCount.Visible = False
    'imgBar.Visible = False
    'Records.Visible = False
    'EmpNumb = Records.SelectedItem
    'frmEmpPay.Show
    Call Records_DblClick
Else
    KeyAscii = 0
End If
End Sub

Private Sub tmrTime_Timer()
lblTime.Caption = Format(Now, "hh:mm:ss AM/PM")
lblDate.Caption = Format(Now, "mmmm dd, yyyy")
End Sub

Private Sub txtSearch_Change()
If txtSearch.Text = "" Then
    SQL = "SELECT * FROM Employees"
    connOpen
    Me.EmpRecs
    connClose
    Exit Sub
ElseIf cboFilter.Text = "Name" Then
    SQL = "SELECT * FROM Employees WHERE LastName LIKE '" & Replace(txtSearch.Text, "'", "") & "%'"
ElseIf cboFilter.Text = "Age" Then
    SQL = "SELECT * FROM Employees WHERE Age LIKE '" & Replace(txtSearch.Text, "'", "") & "%'"
End If

connOpen

Records.ListItems.Clear

With RS
    Do While .EOF = False
        Set rec = Records.ListItems.Add(, , .Fields!EmployeeNo)
            rec.SubItems(1) = .Fields!LastName + ", " + .Fields!FirstName + " " + Left(.Fields!MiddleName, 1)
            rec.SubItems(2) = .Fields!Gender
            rec.SubItems(3) = .Fields!Age
            rec.SubItems(4) = .Fields!Birthday
            rec.SubItems(5) = .Fields!Address
            rec.SubItems(6) = .Fields!ContactNo
            rec.SubItems(7) = .Fields!Position
            rec.SubItems(8) = .Fields!DateHired
            
        If .Fields!CurrentStatus = "Active" Then
            rec.SmallIcon = 1
        Else
            rec.SmallIcon = 2
        End If
        
        .MoveNext
    Loop
End With
connClose
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If cboFilter.Text = "Age" Then
    If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = vbKeyBack) Then
        Exit Sub
    Else
        KeyAscii = 0
    End If
End If
End Sub

