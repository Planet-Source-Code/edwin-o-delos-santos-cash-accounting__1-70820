VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCheckBook 
   Caption         =   "Bank Account"
   ClientHeight    =   9210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11385
   Icon            =   "frmCheckBook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmCheckBook.frx":1782
   ScaleHeight     =   9210
   ScaleWidth      =   11385
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PicEntry 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   240
      Picture         =   "frmCheckBook.frx":2429
      ScaleHeight     =   2745
      ScaleWidth      =   11145
      TabIndex        =   3
      Top             =   720
      Width           =   11175
      Begin VB.TextBox txtEntry 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Index           =   8
         Left            =   7320
         TabIndex        =   20
         Top             =   1440
         Width           =   3555
      End
      Begin VB.TextBox txtEntry 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Index           =   1
         Left            =   1920
         TabIndex        =   11
         Top             =   720
         Width           =   3555
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00E8FBFB&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Index           =   0
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   360
         Width           =   3555
      End
      Begin VB.TextBox txtEntry 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   7320
         TabIndex        =   9
         Top             =   1080
         Width           =   3555
      End
      Begin VB.TextBox txtEntry 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   7320
         TabIndex        =   8
         Top             =   720
         Width           =   3555
      End
      Begin VB.TextBox txtEntry 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   7320
         TabIndex        =   7
         Top             =   360
         Width           =   3555
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Index           =   3
         Left            =   1920
         TabIndex        =   6
         Top             =   1440
         Width           =   3555
      End
      Begin VB.TextBox txtEntry 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Index           =   2
         Left            =   1920
         TabIndex        =   5
         Top             =   1080
         Width           =   3555
      End
      Begin VB.TextBox txtEntry 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   1920
         TabIndex        =   4
         Top             =   1800
         Width           =   3555
      End
      Begin VB.Label lblFLDi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   8
         Left            =   5760
         TabIndex        =   21
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblFLDi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   7
         Left            =   5760
         TabIndex        =   19
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblFLDi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   6
         Left            =   5760
         TabIndex        =   18
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblFLDi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   5
         Left            =   5760
         TabIndex        =   17
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblFLDi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   4
         Left            =   360
         TabIndex        =   16
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lblFLDi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   3
         Left            =   360
         TabIndex        =   15
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblFLDi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   2
         Left            =   360
         TabIndex        =   14
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblFLDi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   360
         TabIndex        =   13
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblFLDi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   8865
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   14420
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "7/22/2008"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "9:08 PM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgMenu 
      Left            =   0
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   49
      ImageHeight     =   49
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":FE669
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":10049F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":1022D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":10410B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":105F41
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicTopBar 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   915
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComctlLib.ImageList ImgToolbar 
      Left            =   480
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":107D77
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":1084F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":108C6B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":1093E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":109B5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":10A2D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":10AA53
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   2760
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":10B1CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":10B767
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":10BB01
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":10BE9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":10C235
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":10C5CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":10C969
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":10CD03
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":10D715
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":10D769
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":10DB03
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":10DE9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":10E237
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":10E5D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":10EFE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":10F9F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":110407
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":110E19
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":11182B
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":11223D
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":112C4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":1131EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":113787
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   3435
      Left            =   0
      TabIndex        =   22
      Top             =   4800
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   6059
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      Icons           =   "i16x16"
      SmallIcons      =   "i16x16"
      ForeColor       =   -2147483640
      BackColor       =   16579829
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
      Picture         =   "frmCheckBook.frx":113D21
   End
   Begin MSComctlLib.Toolbar TbMenu 
      Height          =   450
      Left            =   360
      TabIndex        =   23
      Top             =   3600
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   794
      ButtonWidth     =   820
      ButtonHeight    =   794
      Style           =   1
      ImageList       =   "ImgToolbar"
      DisabledImageList=   "ImgToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            Object.ToolTipText     =   "Add"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "edit"
            Object.ToolTipText     =   "Edit"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "update"
            Object.ToolTipText     =   "Update"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cancel"
            Object.ToolTipText     =   "Cancel"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "delete"
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check Book"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "frmCheckBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsChkBook As Recordset
Dim rsTemp As Recordset

Private Sub Form_Load()
TbButtonShow "1010011"

Set rsChkBook = New ADODB.Recordset
rsChkBook.Open "SELECT * From CheckBook order by SN", cnBank, adOpenStatic, adLockOptimistic
Call ShowFldsLabel(Me, rsChkBook)
Load_DATA

End Sub

Private Sub Load_DATA()
On Error GoTo ERRORHANDLE
'// set columnheaders
Call InsertColumn(lvList, rsChkBook)
'//set details
Call FillListView(lvList, rsChkBook, 2)
'//get total
 Call Listview_Total(lvList, rsChkBook)
ERRORHANDLE:
    errorMsg Err, Me.Name, "Load_Data"
End Sub
Private Sub Form_Resize()
On Error Resume Next
  If WindowState <> vbMinimized Then
       If Me.Width < 13695 Then Me.Width = 13695
       If Me.Height < 9525 Then Me.Height = 9525
  
          CoolBar1.Width = ScaleWidth
          lvList.Width = Me.ScaleWidth
          lvList.Top = PicTopBar.Top
          lvList.Height = (Me.ScaleHeight - 4575)

  End If
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub
Private Sub TbButtonShow(ByRef buttonString As String)
'< syntax:  TbButtonShow ("0001111")
''--------------------------------------------------
''-- This routine handles setting the enabled --
''-- to true / false on the buttons.                --
''-------------------------------------------------
''-- A string of 0101 passed. If 0, disabled   --
''-------------------------------------------------
Dim indx As Integer
buttonString = Trim$(buttonString)
For indx = 1 To Len(buttonString)
  If (Mid$(buttonString, indx, 1) = "1") Then
     Me.TbMenu.Buttons(indx).Visible = True    '(index-1) use only if index start from 0
  Else
     Me.TbMenu.Buttons(indx).Visible = False
  End If
  Next indx
End Sub

Private Sub LvList_Click()
 Call BindDatasource(Me, rsChkBook, lvList, True)
End Sub

Private Sub LvList_KeyUp(KeyCode As Integer, Shift As Integer)
 LvList_Click
End Sub

Private Sub Picture2_Click()

End Sub



Private Sub PicEntry_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then
        Call Drag_It(PicEntry.hWnd)
  End If
End Sub

Private Sub nextItemNumber()
      Dim nextNo As Long
      Set rsTemp = New ADODB.Recordset
       Set rsTemp = rsItem
        If rsTemp.State = adStateOpen Then
           rsTemp.Close
        End If
      rsTemp.Open "SELECT * From CheckBook order by SN", cnBank
      nextNo = Last_Recc(rsTemp)
      If nextNo > 0 Then
       TextEntry(0).text = nextNo
       TextEntry(1).SetFocus
      Else
       nextNo = nextNo = 1
       TextEntry(0).text = nextNo
       TextEntry(1).SetFocus
      End If
      Set rsTemp = Nothing
End Sub
Private Sub TbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case "new"
     addRec = True
     TbButtonShow "0100100"
     nextNumber
  Case "save"
     TbButtonShow "1010011"
     Call WriteItems(rsItem, True)
     lvList.SetFocus
     addRec = False
  Case "edit"
     editRec = True
     TbButtonShow "0001100"
  Case "update"
     TbButtonShow "1010011"
     Call WriteItems(rsItem, False)
     lvList.SetFocus
     editRec = False
  Case "cancel"
     addRec = False
     editRec = False
     TbButtonShow "1010011"
     lvList.SetFocus
  Case "delete"
     Call Delete_Record(rsItem, lvList)
  Case "refresh"
       If rsItem.State = adStateOpen Then
          rsItem.Close
       End If
       rsItem.Open "SELECT * From CheckBook order by SN", cnBank, adOpenStatic, adLockOptimistic
       Load_ITEMS
       lvList.SetFocus
End Select
End Sub

