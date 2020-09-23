VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmReference 
   Caption         =   "Reference Entry Form"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6270
   Icon            =   "frmReference.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CboTable 
      Height          =   315
      ItemData        =   "frmReference.frx":1CCA
      Left            =   120
      List            =   "frmReference.frx":1CD4
      TabIndex        =   11
      Text            =   "Select Reference Table..."
      Top             =   120
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CheckBox ChkSelect 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Select Reference Table"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   2895
   End
   Begin InstantReport.Hline Hline1 
      Height          =   30
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   53
   End
   Begin VB.PictureBox Picture1 
      Height          =   525
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   6195
      TabIndex        =   1
      Top             =   840
      Width           =   6255
      Begin VB.PictureBox PicNav 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         ScaleHeight     =   255
         ScaleWidth      =   1455
         TabIndex        =   24
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
         Begin VB.CommandButton CmdFirst 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   0
            MaskColor       =   &H00404040&
            MousePointer    =   99  'Custom
            Picture         =   "frmReference.frx":1CE7
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "First"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton CmdPrev 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   275
            Left            =   360
            MousePointer    =   99  'Custom
            Picture         =   "frmReference.frx":1F9C
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Previous"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton CmdNext 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   275
            Left            =   720
            MousePointer    =   99  'Custom
            Picture         =   "frmReference.frx":2251
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Next"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton CmdLast 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   275
            Left            =   1080
            MousePointer    =   99  'Custom
            Picture         =   "frmReference.frx":2506
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Last"
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   6
         Left            =   5760
         Picture         =   "frmReference.frx":27BB
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Refresh"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   5
         Left            =   5280
         Picture         =   "frmReference.frx":2F25
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   3
         Left            =   1560
         Picture         =   "frmReference.frx":368F
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Update "
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   2
         Left            =   1080
         Picture         =   "frmReference.frx":3DF9
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   1
         Left            =   600
         Picture         =   "frmReference.frx":4563
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Save New"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   4
         Left            =   4800
         Picture         =   "frmReference.frx":4CCD
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   0
         Left            =   120
         Picture         =   "frmReference.frx":5437
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Add"
         Top             =   0
         Width           =   495
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   6800
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Listing"
      TabPicture(0)   =   "frmReference.frx":5BA1
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lvList"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Details"
      TabPicture(1)   =   "frmReference.frx":5BBD
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblFLDi(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblFLDi(3)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblFLDi(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblFLDi(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblFLDi(4)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtEntry(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtEntry(0)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtEntry(3)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtEntry(2)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtEntry(4)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      Begin VB.TextBox txtEntry 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   325
         Index           =   4
         Left            =   1920
         TabIndex        =   22
         Top             =   2040
         Width           =   3915
      End
      Begin VB.TextBox txtEntry 
         Appearance      =   0  'Flat
         Height          =   325
         Index           =   2
         Left            =   1920
         TabIndex        =   17
         Top             =   1320
         Width           =   3915
      End
      Begin VB.TextBox txtEntry 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   325
         Index           =   3
         Left            =   1920
         TabIndex        =   16
         Top             =   1680
         Width           =   3915
      End
      Begin VB.TextBox txtEntry 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8FBF9&
         ForeColor       =   &H00008080&
         Height          =   325
         Index           =   0
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   600
         Width           =   3915
      End
      Begin VB.TextBox txtEntry 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   325
         Index           =   1
         Left            =   1920
         TabIndex        =   14
         Top             =   960
         Width           =   3915
      End
      Begin MSComctlLib.ListView lvList 
         Height          =   3495
         Left            =   -75000
         TabIndex        =   2
         Top             =   360
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   6165
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "i16x16"
         SmallIcons      =   "i16x16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
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
         Height          =   325
         Index           =   4
         Left            =   360
         TabIndex        =   23
         Top             =   2040
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
         Height          =   325
         Index           =   1
         Left            =   360
         TabIndex        =   21
         Top             =   960
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
         Height          =   325
         Index           =   2
         Left            =   360
         TabIndex        =   20
         Top             =   1320
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
         Height          =   325
         Index           =   3
         Left            =   360
         TabIndex        =   19
         Top             =   1680
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
         Height          =   325
         Index           =   0
         Left            =   360
         TabIndex        =   18
         Top             =   600
         Width           =   1575
      End
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   3360
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":5BD9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":65EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":6745
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":689F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":69F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":6CF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":708D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":7427
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":7E39
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":7E8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":8227
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":85C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":895B
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":8CF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":9707
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":A119
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":AB2B
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":B53D
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":BF4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":C961
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":D373
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":D90F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTABLE 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reference"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   5160
      TabIndex        =   13
      Top             =   240
      Width           =   900
   End
End
Attribute VB_Name = "frmReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rcdSet As Recordset
Dim rsTemp As Recordset


Private Sub CboTable_Click()
 lvList.ListItems.Clear
 lblTABLE.Caption = CboTable.text
 ChkSelect.Value = 0
 '//
  On Error GoTo ERRORHANDLE
    Dim sqlStr As String
    Dim lblStr As String

    Set rcdSet = New ADODB.Recordset
    rcdSet.CursorLocation = adUseClient
    sqlStr = "SELECT * FROM [" & lblTABLE.Caption & "]"
    rcdSet.Open sqlStr, cnRef, adOpenStatic, adLockOptimistic
    Load_DATA
    Call ShowFldsLabel(Me, rcdSet)

 If lvList.ListItems.Count > 0 Then
    If SSTab1.Tab = 0 Then
         cmdButtonShow ("0000011"), Me
    Else
         cmdButtonShow ("1010001"), Me
    End If
  Else
    cmdButtonShow ("0000000"), Me
 End If

ERRORHANDLE:
    errorMsg Err, Me.Name, "CboTables_Click()"
End Sub

Private Sub ChkSelect_Click()
 CboTable.Visible = (ChkSelect.Value = 1)
End Sub

Private Sub nextNumber()
      Dim nextNo As Long
      Set rsTemp = New ADODB.Recordset
      Set rsTemp = rcdSet
      If rsTemp.State = adStateOpen Then
        rsTemp.Close
      End If
      rsTemp.Open "SELECT * From [" & lblTABLE.Caption & "]"
      nextNo = Last_Recc(rsTemp)
      If nextNo > 0 Then
       TxtEntry(0).text = nextNo
       TxtEntry(1).SetFocus
      Else
       nextNo = nextNo = 1
       TxtEntry(0).text = nextNo
       TxtEntry(1).SetFocus
      End If
      Set rsTemp = Nothing
End Sub


Private Sub cmdButton_Click(Index As Integer)
'//                  A S E U C D R
On Error GoTo ERRORHANDLE
Dim sqlStr As String
'Dim nextNo As Long
Select Case Index
   Case BtnAdd                       '<------ add new record ------->'
      addRec = True
      cmdButtonShow ("0100100"), Me
      nextNumber
   Case BtnSave                       '<------ save new record ------>'
        cmdButtonShow ("1010001"), Me
        If rcdSet Is Nothing Then Exit Sub
        Call WriteData(Me, rcdSet, True)
        Call lvwPopulateData(lvList, rcdSet, 2, TxtEntry(0).text)
        addRec = False
   Case BtnEdit                       '<------ edit record ---------->'
        editRec = True
        cmdButtonShow ("0001100"), Me
        TxtEntry(1).SetFocus
   Case BtnUpdate                     '<------ update record -------->'
        cmdButtonShow ("1010001"), Me
        If rcdSet Is Nothing Then Exit Sub
        Call WriteData(Me, rcdSet, False)
        LvwReplaceData Me, rcdSet, lvList
        editRec = False
   Case BtnCancel                     '<------ cancel update -------->'
        cmdButtonShow ("1010001"), Me
        addRec = False
        editRec = False
   Case BtnDelete                     '<------ delete record -------->'
        Call Delete_Record(rcdSet, lvList)
   Case BtnRefresh                    '<------ Refresh record ------->'
        addRec = False
        editRec = False
       If rcdSet Is Nothing Then Exit Sub
       If rcdSet.State = adStateOpen Then
          rcdSet.Close
        End If
         sqlStr = "SELECT * FROM [" & lblTABLE.Caption & "]"
         rcdSet.Open sqlStr, cnRef, adOpenStatic, adLockOptimistic
        Load_DATA
        lvList.SetFocus
End Select
ERRORHANDLE:
 errorMsg Err, Me.Name, "Command Button"

End Sub

Private Sub CmdFirst_Click()
If rcdSet Is Nothing Then Exit Sub
   rcdSet.MoveFirst
 Call BindDatasource(Me, rcdSet, lvList, False)
End Sub

Private Sub CmdLast_Click()
If rcdSet Is Nothing Then Exit Sub
 rcdSet.MoveLast
 Call BindDatasource(Me, rcdSet, lvList, False)
End Sub

Private Sub CmdNext_Click()
 If rcdSet Is Nothing Then Exit Sub
 If rcdSet.EOF = True Then rcdSet.MoveLast
 rcdSet.MoveNext
 Call BindDatasource(Me, rcdSet, lvList, False)

End Sub

Private Sub CmdPrev_Click()
If rcdSet Is Nothing Then Exit Sub
If rcdSet.BOF = True Then rcdSet.MoveFirst
rcdSet.MovePrevious
Call BindDatasource(Me, rcdSet, lvList, False)

End Sub




Private Sub Form_Load()
  CboTable.Clear
  cmdButtonShow ("0000000"), Me
  SSTab1.Tab = 0
  LoadTables
  Call OpenDB("REFERENCE.MDB", cnRef)
End Sub

Private Sub Form_Resize()
With Me
  If .WindowState = 0 Then
   .Height = 5820
   .Width = 6390
  End If
End With
 SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub

Private Sub LoadTables()

    Dim db As Database
    Dim qdef As QueryDef
    Dim td As TableDef
    Dim dbname As String

    ' Open the database. replace "c:\DBfile.mdb" with your
    ' database file name
    
    Set db = OpenDatabase(App.Path & "\DB\REFERENCE.mdb")
    ' List the table names.
    For Each td In db.TableDefs
    ' if you want to display also the system tables, replace the line
    ' below with:  List1.AddItem td.Name
       If td.Attributes = 0 Then CboTable.AddItem td.Name
    Next td
    db.Close
End Sub


Private Sub Load_DATA()
On Error GoTo ERRORHANDLE
'// set columnheaders
 Call InsertColumn(lvList, rcdSet)
'//set details
 Call FillListView(lvList, rcdSet, 3)
ERRORHANDLE:
      errorMsg Err, Me.Name, "load data"

End Sub

Private Sub LvList_Click()
On Error GoTo ERRORHANDLE
If addRec = True Or editRec = True Then Exit Sub
Call BindDatasource(Me, rcdSet, lvList, True)
ERRORHANDLE:
    If Err.Number = 91 Then
       Exit Sub
    Else
      errorMsg Err, Me.Name
    End If
End Sub

Private Sub LvList_KeyUp(KeyCode As Integer, Shift As Integer)
 LvList_Click
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
 With SSTab1
  If .Tab = 1 Then
    cmdButtonShow ("1010001"), Me
    PicNav.Visible = True
  Else '=0
    PicNav.Visible = False
     If lvList.ListItems.Count > 0 Then
       cmdButtonShow ("0000011"), Me
     Else
       cmdButtonShow ("0000000"), Me
     End If
  End If
 End With
End Sub

