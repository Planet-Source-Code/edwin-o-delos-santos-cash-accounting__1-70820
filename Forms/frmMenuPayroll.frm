VERSION 5.00
Begin VB.Form frmMenuPayroll 
   BackColor       =   &H000080FF&
   BorderStyle     =   0  'None
   Caption         =   "Payroll Menu"
   ClientHeight    =   6645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmMenuPayroll.frx":0000
   ScaleHeight     =   6645
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicClose 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   8400
      MouseIcon       =   "frmMenuPayroll.frx":A86F
      MousePointer    =   99  'Custom
      Picture         =   "frmMenuPayroll.frx":B139
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   120
      Width           =   270
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   960
      Picture         =   "frmMenuPayroll.frx":B6C3
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdVacation 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Vacation Leave"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CommandButton CmdSSS 
      BackColor       =   &H00C0E0FF&
      Caption         =   "SSS Contribution"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton cmdDeduction 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Set  Employee Deduction "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton CmdDTR 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Daily Time Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton cmdPayroll 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Employee Payroll"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Payroll"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1920
      TabIndex        =   7
      Top             =   960
      Width           =   750
   End
End
Attribute VB_Name = "frmMenuPayroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDeduction_Click()
 Load frmDeduction
 frmDeduction.show
End Sub

Private Sub CmdDTR_Click()
 Load FrmDTR
 FrmDTR.show
End Sub

Private Sub cmdPayroll_Click()
  Load FrmPayroll
  FrmPayroll.show
End Sub

Private Sub CmdSSS_Click()
 Load FrmSS
 FrmSS.show
End Sub

Private Sub cmdVacation_Click()
 Load frmVacation
 frmVacation.show
End Sub

Private Sub Form_Load()
   FormRndCorner Me, 600, 400
End Sub

Private Sub Form_Resize()
 Me.Top = 0
 Me.Left = 0
End Sub

Private Sub PicClose_Click()
 Me.Hide
End Sub
