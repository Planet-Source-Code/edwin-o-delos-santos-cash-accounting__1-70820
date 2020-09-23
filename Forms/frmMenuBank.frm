VERSION 5.00
Begin VB.Form frmMenuBank 
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmMenuBank.frx":0000
   ScaleHeight     =   7065
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCheckTrans 
      BackColor       =   &H00DEBA83&
      Caption         =   "Check Tansaction"
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton CmdCheckBook 
      BackColor       =   &H00DEBA83&
      Caption         =   "Check Book N/A"
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
      TabIndex        =   3
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton CmdCheckJournal 
      BackColor       =   &H00DEBA83&
      Caption         =   "Check Disbursment Journal  N/A"
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
      Top             =   2040
      Width           =   5055
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
      Picture         =   "frmMenuBank.frx":A86F
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin VB.PictureBox PicClose 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   8400
      MouseIcon       =   "frmMenuBank.frx":C695
      MousePointer    =   99  'Custom
      Picture         =   "frmMenuBank.frx":CF5F
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   120
      Width           =   270
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank"
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
      TabIndex        =   5
      Top             =   960
      Width           =   540
   End
End
Attribute VB_Name = "frmMenuBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCheckTrans_Click()
Load frmCashDisburse
frmCashDisburse.show
End Sub

Private Sub Form_Activate()
 Me.Top = 0
 Me.Left = 0

End Sub

Private Sub Form_Load()
  FormRndCorner Me, 600, 400
End Sub

Private Sub PicClose_Click()
Me.Hide
End Sub
