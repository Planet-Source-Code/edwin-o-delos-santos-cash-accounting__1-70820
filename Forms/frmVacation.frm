VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmVacation 
   BackColor       =   &H00E4DBC2&
   Caption         =   "Vacation Schedule"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   9870
   Icon            =   "frmVacation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmVacation.frx":109A
   ScaleHeight     =   6705
   ScaleWidth      =   9870
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Edit"
      Height          =   315
      Index           =   2
      Left            =   8280
      TabIndex        =   34
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Add"
      Height          =   315
      Index           =   0
      Left            =   8280
      TabIndex        =   33
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Update"
      Height          =   315
      Index           =   3
      Left            =   8280
      TabIndex        =   32
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Save"
      Height          =   315
      Index           =   1
      Left            =   8280
      TabIndex        =   31
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Cancel"
      Height          =   315
      Index           =   4
      Left            =   8280
      TabIndex        =   30
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Refresh"
      Height          =   315
      Index           =   6
      Left            =   8280
      TabIndex        =   29
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Delete"
      Height          =   315
      Index           =   5
      Left            =   8280
      TabIndex        =   28
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ListBox List1Type 
      Appearance      =   0  'Flat
      BackColor       =   &H00E6F1FD&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2430
      ItemData        =   "frmVacation.frx":FD2DA
      Left            =   1800
      List            =   "frmVacation.frx":FD2F3
      Sorted          =   -1  'True
      TabIndex        =   20
      Top             =   3600
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox PicNameList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3180
      Left            =   5880
      Picture         =   "frmVacation.frx":FD346
      ScaleHeight     =   3150
      ScaleWidth      =   3885
      TabIndex        =   16
      Top             =   3600
      Visible         =   0   'False
      Width           =   3915
      Begin VB.PictureBox PicNameClose 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3600
         Picture         =   "frmVacation.frx":142D2A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   17
         Top             =   0
         Width           =   270
      End
      Begin MSComctlLib.ListView lvName 
         Height          =   2595
         Left            =   0
         TabIndex        =   18
         Top             =   480
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   4577
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "i16x16"
         SmallIcons      =   "i16x16"
         ForeColor       =   12582912
         BackColor       =   15268859
         Appearance      =   0
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
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
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
         TabIndex        =   19
         Top             =   120
         Width           =   1365
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   6330
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9948
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            Picture         =   "frmVacation.frx":1432B4
            Text            =   "Print"
            TextSave        =   "Print"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "2:21 AM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "7/16/2008"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00FFFFFF&
      Height          =   1965
      Index           =   6
      Left            =   5400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   600
      Width           =   2595
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   5
      Top             =   1800
      Width           =   2835
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   4
      Top             =   1080
      Width           =   2835
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   3
      Top             =   1440
      Width           =   2835
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   2
      Top             =   2160
      Width           =   2835
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00008080&
      Height          =   285
      Index           =   0
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   1155
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   1155
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   2595
      Left            =   240
      TabIndex        =   7
      Top             =   3600
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   4577
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      Icons           =   "i16x16"
      SmallIcons      =   "i16x16"
      ForeColor       =   -2147483640
      BackColor       =   15003831
      Appearance      =   0
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
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":14364E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":143BE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":144182
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":14471C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":144AB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":144E50
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":1451EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":145584
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":14591E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":145CB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":1466CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":14671E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":146AB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":146E52
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":1471EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":147586
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":147F98
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":1489AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":1493BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":149DCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":14A7E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":14B1F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":14BC04
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":14C1A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":14C73C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      ScaleHeight     =   375
      ScaleWidth      =   9015
      TabIndex        =   22
      Top             =   3000
      Width           =   9015
      Begin VB.ComboBox CboSched 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   5280
         Sorted          =   -1  'True
         TabIndex        =   27
         Top             =   0
         Width           =   1455
      End
      Begin VB.OptionButton OptBySched 
         BackColor       =   &H00808080&
         Caption         =   "By Date Schedule"
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
         Height          =   255
         Left            =   3120
         TabIndex        =   26
         Top             =   50
         Width           =   1935
      End
      Begin VB.OptionButton OptByID 
         BackColor       =   &H00808080&
         Caption         =   "By ID"
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
         Height          =   255
         Left            =   1080
         TabIndex        =   25
         Top             =   50
         Width           =   975
      End
      Begin VB.TextBox TxtID 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   325
         Left            =   2160
         TabIndex        =   24
         Top             =   10
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter >>"
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
         TabIndex        =   23
         Top             =   75
         Width           =   705
      End
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000C0&
      Height          =   615
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   9375
   End
   Begin VB.Label F2Key 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[F2]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   4920
      TabIndex        =   21
      Top             =   1080
      Width           =   345
   End
   Begin VB.Label lblFLDi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   5
      Left            =   480
      TabIndex        =   14
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label lblFLDi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   4
      Left            =   480
      TabIndex        =   13
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lblFLDi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   480
      TabIndex        =   12
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblFLDi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   2
      Left            =   480
      TabIndex        =   11
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblFLDi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   3
      Left            =   480
      TabIndex        =   10
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblFLDi 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   6
      Left            =   5400
      TabIndex        =   9
      Top             =   360
      Width           =   2565
   End
   Begin VB.Label lblFLDi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   8
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmVacation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsEmp As ADODB.Recordset
Dim rs As ADODB.Recordset


Private Sub cmdButton_Click(Index As Integer)
'//                  A S E U C D R
'On Error GoTo ERRORHANDLE
Select Case Index
   Case BtnAdd                       '<------ add new record ------->'
     addRec = True
     cmdButtonShow ("0100100"), Me
'     If isFilter = True Then
'        MsgBox "Data is Filtered", vbCritical, "Refresh Record First!"
'        Exit Sub
'     End If
'     Dim NextNo As Long
'     '//initialize//
'     txtEntry(29).text = Format(Now(), "Short Date")
'     txtEntry(30).text = CurrUser.user_id
'     '//assign next number//
      nextNo = Last_Recc(rs)
      If nextNo > 0 Then
       txtEntry(0).text = nextNo
       txtEntry(2).SetFocus
      Else
       txtEntry(0).Locked = False
       txtEntry(0).SetFocus
      End If
   Case BtnSave                       '<------ save new record ------>'
        cmdButtonShow ("1010011"), Me
        Call WriteData(Me, rs, True)
        Call lvwPopulateData(lvList, rs, 2)
        addRec = False
   Case BtnEdit                       '<------ edit record ---------->'
        editRec = True
        cmdButtonShow ("0001100"), Me
        txtEntry(2).SetFocus
   Case BtnUpdate                     '<------ update record -------->'
        cmdButtonShow ("1010001"), Me
        Call WriteData(Me, rs, False)
        LvwReplaceData Me, rs, lvList
        editRec = False
   Case BtnCancel                     '<------ cancel update -------->'
        cmdButtonShow ("1010001"), Me
        addRec = False
        editRec = False
   Case BtnDelete                     '<------ delete record -------->'
        '// no delete here please !
        Call Delete_Record(rs, lvList)
   Case BtnRefresh                    '<------ Refresh record ------->'
        addRec = False
        editRec = False
       If rs.State = adStateOpen Then
          rs.Close
        End If
        rs.Open "SELECT * From VACATION order by SN", CnPay, adOpenStatic, adLockOptimistic
        Load_DATA
        isFilter = False
        lvList.SetFocus
End Select
'ERRORHANDLE:
' errorMsg Err, Me.Name, "Command Button"

End Sub




Private Sub Form_Load()
'// initialized
cmdButtonShow ("1010011"), Me
AlignObj txtEntry(2), PicNameList, 1, False
AlignObj txtEntry(3), List1Type, 1, False
'// set focus
show
lvList.SetFocus
      
Set rs = New ADODB.Recordset
rs.Open "SELECT * From VACATION order by SN", CnPay, adOpenStatic, adLockOptimistic
Load_DATA
Call ShowFldsLabel(Me, rs)
Call Add_Item(rs, "date_schedule", CboSched, True)
'//
Set rsEmp = New ADODB.Recordset
Dim SQL As String
SQL = "SELECT Employee_Name,ID_Code "
SQL = SQL & "From PAYROLL order by Employee_name"
rsEmp.Open SQL, CnPay, adOpenStatic, adLockOptimistic
Load_Employee


End Sub
Private Sub Load_DATA()
'On Error GoTo ERRORHANDLE
'// set columnheaders
'Insert_ExtraCol lvList, rsDed

Call InsertColumn(lvList, rs)
'//set details
 Call FillListView(lvList, rs, 2)
'ERRORHANDLE:
'    errorMsg Err, Me.Name
End Sub

Private Sub Load_Employee()
On Error GoTo ERRORHANDLE
'// set columnheaders
'Insert_ExtraCol lvList, rsDed
Call InsertColumn(lvName, rsEmp)
'//set details
Call FillListView(lvName, rsEmp, 8)
autoAlignCol lvName
ERRORHANDLE:
    errorMsg Err, Me.Name, "Load_Employee proc"
End Sub

Private Sub Form_Resize()
With Me
  If .WindowState = 0 Then
   .Height = 7185
   .Width = 9990
  End If
End With
 SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub


Private Sub List1Type_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
   txtEntry(3).text = List1Type.text
   txtEntry(3).SetFocus
   List1Type.Visible = False
 End If
End Sub

Private Sub lvList_Click()
On Error GoTo ERRORHANDLE
If addRec = True Or editRec = True Then Exit Sub
Call BindDatasource(Me, rs, lvList, True)
ERRORHANDLE:
 errorMsg Err, Me.Name
End Sub

Private Sub lvList_KeyUp(KeyCode As Integer, Shift As Integer)
 lvList_Click
End Sub

Private Sub lvName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   txtEntry(2).text = lvName.SelectedItem.text
   txtEntry(1).text = lvName.SelectedItem.ListSubItems(1).text  'ID
  txtEntry(2).SetFocus
  PicNameList.Visible = False
ElseIf KeyCode = 27 Then
  txtEntry(2).SetFocus
  PicNameList.Visible = False
End If
End Sub

Private Sub OptByID_Click()
  CboSched.Enabled = False
  TxtID.Enabled = (OptByID.Value = True)
End Sub

Private Sub OptBySched_Click()
 TxtID.Enabled = False
 CboSched.Enabled = (OptBySched.Value = True)
End Sub

Private Sub PicNameClose_Click()
 PicNameList.Visible = False
 txtEntry(2).SetFocus
End Sub

Private Sub txtId_Change()
Dim SQL As String
Dim lngID As Long
lngID = Val(TxtID.text)
If lngID = 0 Then Exit Sub
       If rs.State = adStateOpen Then
          rs.Close
        End If

SQL = "Select * From VACATION WHERE [ID_number]=" & lngID & " order by name"
rs.Open SQL, CnPay, adOpenStatic, adLockOptimistic
If rsEmp.RecordCount > 0 Then
   Load_DATA
Else
  MsgBox "No Record Found!", vbInformation, "Vacation File"
  Exit Sub
End If
'autoAlignCol LvDTR

End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
On Error GoTo errorMsg
nxTab = Index
txtEntry(nxTab).SelStart = 0
txtEntry(nxTab).SelLength = Len(txtEntry(nxTab).text)
Select Case nxTab
  Case Is = 2, 3
     If addRec = True Or editRec = True Then
      F2Key.Top = txtEntry(nxTab).Top
     End If
'     If addRec = True Or editRec = True Then
'       AlignObj txtEntry(2), PicNameList, 1, False
'     End If
End Select
errorMsg:
 errorMsg Err, Me.Name, "txtEntry_GotFocus"

End Sub

Private Sub txtEntry_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim lastTab As Integer
On Error GoTo ERRORHANDLE
lastTab = 5
If KeyCode = 13 Then
    If nxTab = lastTab Then Exit Sub
    If nxTab = 6 Then Exit Sub
    nxTab = nxTab + 1
ElseIf KeyCode = 38 Then  'up arrow key
     If nxTab = 0 Or nxTab = 1 Then Exit Sub
     If nxTab = 6 Then Exit Sub
     nxTab = nxTab - 1
End If
txtEntry(nxTab).SetFocus
ERRORHANDLE:
 errorMsg Err, Me.Name
End Sub

Private Sub txtEntry_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
Case Is = 2
  If addRec = True Or editRec = True Then
    If KeyCode = 113 Then 'F2
       PicNameList.Visible = True
       lvName.SetFocus
    End If
  End If
Case Is = 3
  If addRec = True Or editRec = True Then
    If KeyCode = 113 Then 'F2
       List1Type.Visible = True
       List1Type.SetFocus
    End If
  End If
Case Is = 27
  lvList.SetFocus
End Select
End Sub

Private Sub CboSched_Click()
      Dim sqlStatement As String
      Dim m_Table As String
      Dim m_field1 As String
      Dim m_Value1 As String
      m_Table = "VACATION"
      m_field1 = "DATE_SCHEDULE"
      m_Value1 = "#" & CDate(CboSched.text) & "#"
      sqlStatement = "SELECT * FROM [" & m_Table & "] WHERE [" & m_field1 & "]=" & m_Value1
      If rs.State = adStateOpen Then
          rs.Close
       End If
         rs.Open sqlStatement, CnPay
       If rs.RecordCount > 0 Then
         Load_DATA
       Else
        MsgBox "No record found", vbInformation, "Filter by Date Schedule"
       End If
End Sub
