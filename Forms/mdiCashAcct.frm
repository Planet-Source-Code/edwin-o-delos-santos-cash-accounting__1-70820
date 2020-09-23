VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm mdiCashAcct 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   10380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13170
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture6 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7710
      Left            =   0
      ScaleHeight     =   7710
      ScaleWidth      =   3660
      TabIndex        =   7
      Top             =   2295
      Width           =   3660
      Begin VB.PictureBox PicMenuPointer 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   240
         Picture         =   "mdiCashAcct.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   20
         Top             =   3045
         Width           =   240
      End
      Begin VB.PictureBox PicMenuPointer 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   240
         Picture         =   "mdiCashAcct.frx":014A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   19
         Top             =   2805
         Width           =   240
      End
      Begin VB.PictureBox PicMenuPointer 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   240
         Picture         =   "mdiCashAcct.frx":0294
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   18
         Top             =   2565
         Width           =   240
      End
      Begin VB.PictureBox PicMenuPointer 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "mdiCashAcct.frx":03DE
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   17
         Top             =   2325
         Width           =   240
      End
      Begin VB.TextBox txtPass 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         MaxLength       =   15
         MousePointer    =   99  'Custom
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtname 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   960
         MousePointer    =   99  'Custom
         TabIndex        =   12
         ToolTipText     =   "Down arrow key to view user list!"
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton CmdLogIn 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2640
         Picture         =   "mdiCashAcct.frx":0528
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1200
         Width           =   975
      End
      Begin VB.CheckBox chkAdmin 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Admin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin InstantReport.Hline ctrlLiner3 
         Height          =   30
         Left            =   0
         TabIndex        =   9
         Top             =   600
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   53
      End
      Begin InstantReport.Hline ctrlLiner2 
         Height          =   30
         Left            =   0
         TabIndex        =   10
         Top             =   1680
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   53
      End
      Begin VB.Label lblMenuOver 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Edwin Delos Santos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   240
         MouseIcon       =   "mdiCashAcct.frx":1132
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   1965
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.Label lblMenu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User File Maintenance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   3
         Left            =   600
         TabIndex        =   24
         Top             =   3030
         Width           =   2010
      End
      Begin VB.Label lblMenu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "References"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   600
         TabIndex        =   23
         Top             =   2760
         Width           =   1050
      End
      Begin VB.Label lblMenu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Calculator"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   600
         TabIndex        =   22
         Top             =   2520
         Width           =   900
      End
      Begin VB.Label lblMenu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Calendar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   600
         TabIndex        =   21
         Top             =   2280
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Log-In-User"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   15
         Top             =   1200
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User ID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   720
      End
      Begin VB.Image imgHelp 
         Height          =   360
         Left            =   3720
         MouseIcon       =   "mdiCashAcct.frx":19FC
         MousePointer    =   99  'Custom
         Picture         =   "mdiCashAcct.frx":22C6
         Top             =   600
         Width           =   360
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   10005
      Width           =   13170
      _ExtentX        =   23230
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   13170
      _ExtentX        =   23230
      _ExtentY        =   635
      ButtonWidth     =   1799
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "File"
            Key             =   "menuFile"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reports "
            Key             =   "menuReport"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Key             =   "MenuHelp"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture5 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   0
      Picture         =   "mdiCashAcct.frx":2A30
      ScaleHeight     =   1935
      ScaleWidth      =   13170
      TabIndex        =   0
      Top             =   360
      Width           =   13170
      Begin VB.PictureBox PicPettyCash 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   12000
         MouseIcon       =   "mdiCashAcct.frx":16A74
         MousePointer    =   99  'Custom
         Picture         =   "mdiCashAcct.frx":1733E
         ScaleHeight     =   1215
         ScaleWidth      =   1470
         TabIndex        =   4
         Top             =   600
         Width           =   1470
      End
      Begin VB.PictureBox PicPurchaseOrder 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   10440
         MouseIcon       =   "mdiCashAcct.frx":1D640
         MousePointer    =   99  'Custom
         Picture         =   "mdiCashAcct.frx":1DF0A
         ScaleHeight     =   1215
         ScaleWidth      =   1470
         TabIndex        =   3
         Top             =   600
         Width           =   1470
      End
      Begin VB.PictureBox PicPayroll 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   13560
         MouseIcon       =   "mdiCashAcct.frx":2420C
         MousePointer    =   99  'Custom
         Picture         =   "mdiCashAcct.frx":24AD6
         ScaleHeight     =   1215
         ScaleWidth      =   1470
         TabIndex        =   2
         Top             =   600
         Width           =   1470
      End
      Begin VB.PictureBox PicBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   8880
         MouseIcon       =   "mdiCashAcct.frx":2ADD8
         MousePointer    =   99  'Custom
         Picture         =   "mdiCashAcct.frx":2B6A2
         ScaleHeight     =   1215
         ScaleWidth      =   1470
         TabIndex        =   1
         Top             =   600
         Width           =   1470
      End
   End
End
Attribute VB_Name = "mdiCashAcct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
pIndex = Index
 Call HK(lblMenu(pIndex), lblMenuOver)
End Sub


Private Sub lblMenuOver_Click()
Select Case pIndex
 Case Is = 0
'      frmcalendar.show
     'Call allowACCESS(lblMenu(pIndex), FrmSQL)
 Case Is = 1
      FrmCalcu.show
     'Call allowACCESS(lblMenu(pIndex), FrmProdList)
Case Is = 2
     frmReference.show
    'Call allowACCESS(lblMenu(pIndex), FrmStockReceive)
 Case Is = 3
    'MDIPayroll.show
    'Call allowACCESS(lblMenu(pIndex), MDIPayroll, True)
End Select
  lockMenu True
End Sub
Private Sub lockMenu(ByVal sconfirm As Boolean)
Dim i As Integer
For i = 0 To lblMenu.UBound
    If sconfirm = True Then
      lblMenu(i).Enabled = False
      lblMenuOver.Enabled = False
    Else
      lblMenu(i).Enabled = True
      lblMenuOver.Enabled = True
    End If
    Next i
End Sub

Private Sub MDIForm_Load()

End Sub

Private Sub PicBank_Click()
  frmMenuBank.show
  frmMenuPurchaseOrder.Hide
  frmMenuPayroll.Hide
  frmMenuPettyCash.Hide
End Sub

Private Sub PicMenuPointer_Click(Index As Integer)
lockMenu False
End Sub

Private Sub PicPayroll_Click()
  frmMenuPayroll.show
  frmMenuPurchaseOrder.Hide
  frmMenuBank.Hide
  frmMenuPettyCash.Hide
End Sub

Private Sub PicPettyCash_Click()
 frmMenuPettyCash.show
 frmMenuPurchaseOrder.Hide
 frmMenuBank.Hide
 frmMenuPayroll.Hide
End Sub

Private Sub PicPurchaseOrder_Click()
 frmMenuPurchaseOrder.show
 frmMenuBank.Hide
 frmMenuPayroll.Hide
 frmMenuPettyCash.Hide
End Sub
