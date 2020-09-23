VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCashDisburse 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Bank Account"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13665
   Icon            =   "frmCashDisburse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmCashDisburse.frx":1E26
   ScaleHeight     =   9015
   ScaleWidth      =   13665
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImgMenu 
      Left            =   1320
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   49
      ImageHeight     =   49
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashDisburse.frx":2ACD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashDisburse.frx":4903
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashDisburse.frx":6739
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashDisburse.frx":856F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashDisburse.frx":A3A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashDisburse.frx":C1DB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LvMenu 
      Height          =   7815
      Left            =   0
      TabIndex        =   15
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   13785
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      PictureAlignment=   5
      _Version        =   393217
      Icons           =   "ImgMenu"
      SmallIcons      =   "ImgMenu"
      ForeColor       =   12582912
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   8640
      Width           =   13665
      _ExtentX        =   24104
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   18441
            Text            =   "<:- System designed by:  Edwin delos Santos   (c)-2008  All rights reserved -:>"
            TextSave        =   "<:- System designed by:  Edwin delos Santos   (c)-2008  All rights reserved -:>"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "7/22/2008"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "10:21 PM"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   2040
      TabIndex        =   0
      Top             =   480
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   6
      TabHeight       =   794
      BackColor       =   16777215
      ForeColor       =   12582912
      TabCaption(0)   =   "Expenses"
      TabPicture(0)   =   "frmCashDisburse.frx":E011
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblFLDi(9)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "TxtEntry(9)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lvList"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ChkPrint"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdButton(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdButton(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdButton(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdButton(3)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdButton(4)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdButton(5)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdButton(6)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Picture1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "i16x16"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Items"
      TabPicture(1)   =   "frmCashDisburse.frx":E02D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ImgToolbar"
      Tab(1).Control(1)=   "CboNumber"
      Tab(1).Control(2)=   "FraItems"
      Tab(1).Control(3)=   "lvItems"
      Tab(1).Control(4)=   "TbMenu"
      Tab(1).Control(5)=   "lblSelected"
      Tab(1).Control(6)=   "Label9"
      Tab(1).Control(7)=   "ImgItemHelp"
      Tab(1).ControlCount=   8
      Begin MSComctlLib.ImageList ImgToolbar 
         Left            =   -74880
         Top             =   3480
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
               Picture         =   "frmCashDisburse.frx":E049
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":E7C3
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":EF3D
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":F6B7
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":FE31
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":105AB
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":10D25
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList i16x16 
         Left            =   2280
         Top             =   5040
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
               Picture         =   "frmCashDisburse.frx":1149F
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":11A39
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":11DD3
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":1216D
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":12507
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":128A1
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":12C3B
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":12FD5
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":139E7
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":13A3B
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":13DD5
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":1416F
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":14509
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":148A3
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":152B5
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":15CC7
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":166D9
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":170EB
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":17AFD
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":1850F
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":18F21
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":194BD
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":19A59
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox CboNumber 
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
         Left            =   -66000
         TabIndex        =   76
         Text            =   "Select ...."
         Top             =   4080
         Width           =   2295
      End
      Begin VB.PictureBox FraItems 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEBA83&
         ForeColor       =   &H80000008&
         Height          =   3315
         Left            =   -74640
         Picture         =   "frmCashDisburse.frx":19FF3
         ScaleHeight     =   3285
         ScaleWidth      =   10905
         TabIndex        =   54
         Top             =   600
         Width           =   10935
         Begin VB.TextBox TextEntry 
            Appearance      =   0  'Flat
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
            Height          =   1800
            Index           =   6
            Left            =   5880
            MultiLine       =   -1  'True
            TabIndex        =   74
            Top             =   1200
            Width           =   4095
         End
         Begin VB.TextBox TextEntry 
            Appearance      =   0  'Flat
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
            Index           =   5
            Left            =   1920
            TabIndex        =   60
            Top             =   2640
            Width           =   3375
         End
         Begin VB.TextBox TextEntry 
            Appearance      =   0  'Flat
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
            TabIndex        =   59
            Top             =   2280
            Width           =   3375
         End
         Begin VB.TextBox TextEntry 
            Appearance      =   0  'Flat
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
            Index           =   3
            Left            =   1920
            TabIndex        =   58
            Top             =   1920
            Width           =   3375
         End
         Begin VB.TextBox TextEntry 
            Appearance      =   0  'Flat
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
            Height          =   360
            Index           =   2
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   57
            Top             =   1560
            Width           =   3375
         End
         Begin VB.TextBox TextEntry 
            Appearance      =   0  'Flat
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
            Index           =   1
            Left            =   1920
            TabIndex        =   56
            Top             =   1200
            Width           =   3375
         End
         Begin VB.TextBox TextEntry 
            Appearance      =   0  'Flat
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
            Height          =   360
            Index           =   0
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   840
            Width           =   3375
         End
         Begin VB.Label LblFLD 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Memo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   360
            Index           =   6
            Left            =   5880
            TabIndex        =   75
            Top             =   840
            Width           =   4095
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Entry Form (Items)"
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
            TabIndex        =   67
            Top             =   0
            Width           =   1545
         End
         Begin VB.Label LblFLD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "SN:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   360
            Index           =   0
            Left            =   480
            TabIndex        =   66
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label LblFLD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Check Number:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   360
            Index           =   2
            Left            =   480
            TabIndex        =   65
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label LblFLD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vendor:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   360
            Index           =   5
            Left            =   480
            TabIndex        =   64
            Top             =   2640
            Width           =   1455
         End
         Begin VB.Label LblFLD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Date:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   360
            Index           =   1
            Left            =   480
            TabIndex        =   63
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label LblFLD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Description:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   360
            Index           =   3
            Left            =   480
            TabIndex        =   62
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label LblFLD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Amount:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   360
            Index           =   4
            Left            =   480
            TabIndex        =   61
            Top             =   2280
            Width           =   1455
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   240
         Picture         =   "frmCashDisburse.frx":1AC9A
         ScaleHeight     =   3735
         ScaleWidth      =   11055
         TabIndex        =   16
         Top             =   600
         Width           =   11055
         Begin VB.PictureBox PicAccount 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1995
            Left            =   6600
            ScaleHeight     =   1965
            ScaleWidth      =   4065
            TabIndex        =   17
            Top             =   480
            Visible         =   0   'False
            Width           =   4095
            Begin VB.CommandButton CmdRef 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Refresh"
               Height          =   325
               Left            =   2400
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   120
               Width           =   855
            End
            Begin VB.CommandButton CmdExit 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Exit"
               Height          =   325
               Left            =   3360
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   120
               Width           =   615
            End
            Begin MSComctlLib.ListView ListView2 
               Height          =   1575
               Left            =   0
               TabIndex        =   20
               Top             =   480
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   2778
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               Icons           =   "i16x16"
               SmallIcons      =   "i16x16"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   0
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vendor List"
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
               TabIndex        =   22
               Top             =   0
               Width           =   975
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "<:- Enter to Select !  -:>"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   21
               Top             =   120
               Width           =   1635
            End
         End
         Begin VB.PictureBox PicVendor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1995
            Left            =   2040
            ScaleHeight     =   1965
            ScaleWidth      =   4425
            TabIndex        =   68
            Top             =   1680
            Visible         =   0   'False
            Width           =   4455
            Begin VB.CommandButton CmdRefresh 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Refresh"
               Height          =   325
               Left            =   2760
               Style           =   1  'Graphical
               TabIndex        =   70
               Top             =   1600
               Width           =   855
            End
            Begin VB.CommandButton BtnOK 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Exit"
               Height          =   325
               Left            =   3720
               Style           =   1  'Graphical
               TabIndex        =   69
               Top             =   1600
               Width           =   615
            End
            Begin MSComctlLib.ListView ListView1 
               Height          =   1575
               Left            =   0
               TabIndex        =   71
               Top             =   0
               Width           =   4455
               _ExtentX        =   7858
               _ExtentY        =   2778
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               Icons           =   "i16x16"
               SmallIcons      =   "i16x16"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   0
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vendor List"
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
               TabIndex        =   73
               Top             =   0
               Width           =   975
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "<:- Enter to Select !  -:>"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   72
               Top             =   1680
               Width           =   1635
            End
         End
         Begin VB.TextBox TxtEntry 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   8520
            MaxLength       =   20
            TabIndex        =   35
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox TxtEntry 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   8520
            TabIndex        =   34
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox TxtEntry 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   8520
            TabIndex        =   33
            Top             =   1320
            Width           =   2175
         End
         Begin VB.TextBox TxtEntry 
            Appearance      =   0  'Flat
            BackColor       =   &H00E4F0B7&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   1320
            Width           =   4455
         End
         Begin VB.TextBox TxtEntry 
            Appearance      =   0  'Flat
            BackColor       =   &H00E4F0B7&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Index           =   7
            Left            =   1440
            MultiLine       =   -1  'True
            TabIndex        =   31
            Top             =   2160
            Width           =   2655
         End
         Begin VB.CommandButton BtnVendor 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   275
            Left            =   6600
            Picture         =   "frmCashDisburse.frx":116EDA
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   1320
            Width           =   375
         End
         Begin VB.CheckBox ChkShow 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00F3F4DF&
            Caption         =   "Details Visible When Checked !"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7920
            TabIndex        =   29
            Top             =   3360
            Width           =   2775
         End
         Begin VB.CommandButton BtnExpenses 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   275
            Left            =   10320
            Picture         =   "frmCashDisburse.frx":117264
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   2520
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox TxtEntry 
            Appearance      =   0  'Flat
            BackColor       =   &H00E4F0B7&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   7080
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   2520
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.TextBox TxtEntry 
            Appearance      =   0  'Flat
            BackColor       =   &H00E4F0B7&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   8880
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   2160
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox TxtEntry 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   435
            Index           =   8
            Left            =   7080
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   25
            Top             =   2880
            Visible         =   0   'False
            Width           =   3615
         End
         Begin VB.TextBox TxtEntry 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E4F0B7&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   7080
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   2160
            Visible         =   0   'False
            Width           =   615
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   285
            Left            =   10440
            TabIndex        =   23
            Top             =   720
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   39650
         End
         Begin VB.Line Line1 
            X1              =   8520
            X2              =   10680
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line Line2 
            X1              =   8520
            X2              =   10680
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Line Line3 
            X1              =   8520
            X2              =   10680
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Line Line4 
            X1              =   2040
            X2              =   6480
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Label lblAmtInWord 
            AutoSize        =   -1  'True
            BackColor       =   &H000000FF&
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
            Left            =   360
            TabIndex        =   52
            Top             =   1800
            Width           =   2445
         End
         Begin VB.Line Line5 
            X1              =   240
            X2              =   9960
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pesos"
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
            Left            =   10080
            TabIndex        =   51
            Top             =   1800
            Width           =   585
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Php"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   7920
            TabIndex        =   50
            Top             =   1320
            Width           =   360
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblFLDi"
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
            Index           =   1
            Left            =   7680
            TabIndex        =   49
            Top             =   360
            Width           =   630
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblFLDi"
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
            Index           =   2
            Left            =   7680
            TabIndex        =   48
            Top             =   720
            Width           =   630
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblFLDi"
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
            Index           =   4
            Left            =   7200
            TabIndex        =   47
            Top             =   1080
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label lblFLDi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblFLDi"
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
            Index           =   6
            Left            =   240
            TabIndex        =   46
            Top             =   1560
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label lblFLDi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblFLDi"
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
            Index           =   7
            Left            =   240
            TabIndex        =   45
            Top             =   2520
            Width           =   630
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0920-6747-545"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   300
            Left            =   1440
            TabIndex        =   44
            Top             =   3360
            Width           =   1920
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "edwinSoftware"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3240
            TabIndex        =   43
            Top             =   120
            Width           =   1710
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Makati City, Philippines"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   3120
            TabIndex        =   42
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Pay to the &Order of"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   240
            TabIndex        =   41
            Top             =   1320
            Width           =   1650
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblFLDi"
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
            Height          =   240
            Index           =   3
            Left            =   6240
            TabIndex        =   40
            Top             =   2520
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblFLDi"
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
            Height          =   240
            Index           =   5
            Left            =   8160
            TabIndex        =   39
            Top             =   2160
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblFLDi"
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
            Height          =   240
            Index           =   8
            Left            =   6240
            TabIndex        =   38
            Top             =   2880
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblFLDi"
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
            Height          =   240
            Index           =   0
            Left            =   6240
            TabIndex        =   37
            Top             =   2160
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "cyber_edu2005@yahoo.com"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   2895
            TabIndex        =   36
            Top             =   600
            Width           =   2505
         End
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00DCCD78&
         Caption         =   "&Refresh"
         Height          =   325
         Index           =   6
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4440
         Width           =   900
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00DCCD78&
         Caption         =   "&Delete"
         Height          =   325
         Index           =   5
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   4440
         Width           =   900
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00DCCD78&
         Caption         =   "&Cancel"
         Height          =   325
         Index           =   4
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4440
         Width           =   900
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00DCCD78&
         Caption         =   "&Update"
         Height          =   325
         Index           =   3
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4440
         Width           =   900
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00DCCD78&
         Caption         =   "&Edit"
         Height          =   325
         Index           =   2
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4440
         Width           =   900
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00DCCD78&
         Caption         =   "&Save"
         Height          =   325
         Index           =   1
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4440
         Width           =   900
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00DCCD78&
         Caption         =   "&Add"
         Height          =   325
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4440
         Width           =   900
      End
      Begin VB.CheckBox ChkPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9720
         TabIndex        =   1
         Top             =   4560
         Width           =   255
      End
      Begin MSComctlLib.ListView lvItems 
         Height          =   3195
         Left            =   -74880
         TabIndex        =   6
         Top             =   4560
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   5636
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
         BackColor       =   15135229
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
         Picture         =   "frmCashDisburse.frx":1175EE
      End
      Begin MSComctlLib.ListView lvList 
         Height          =   2835
         Left            =   240
         TabIndex        =   7
         Top             =   4920
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   5001
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
         Picture         =   "frmCashDisburse.frx":11D946
      End
      Begin VB.TextBox TxtEntry 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   9600
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSComctlLib.Toolbar TbMenu 
         Height          =   450
         Left            =   -74640
         TabIndex        =   53
         Top             =   4080
         Width           =   3495
         _ExtentX        =   6165
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
      Begin VB.Label lblSelected 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Check Number:"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   -70920
         TabIndex        =   78
         Top             =   120
         Width           =   1785
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Check Number -:>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   -67920
         TabIndex        =   77
         Top             =   4080
         Width           =   1590
      End
      Begin VB.Image ImgItemHelp 
         Height          =   360
         Left            =   -64200
         MouseIcon       =   "frmCashDisburse.frx":1214EA
         MousePointer    =   99  'Custom
         Picture         =   "frmCashDisburse.frx":121DB4
         Top             =   120
         Width           =   360
      End
      Begin VB.Label lblFLDi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblFLDi"
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
         Index           =   9
         Left            =   9960
         TabIndex        =   3
         Top             =   4560
         Width           =   630
      End
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Check Transaction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   240
      TabIndex        =   4
      Top             =   0
      Width           =   1680
   End
End
Attribute VB_Name = "frmCashDisburse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsCash As Recordset
Private rsVend As Recordset
Private rsItem As Recordset
Private rsAcct As Recordset
Dim rsTemp As Recordset
Dim convert As numTOword
Dim MyMenu As String  'lvmenu temporary storage


Private Sub BtnExpenses_Click()
 PicAccount.Visible = True
 ListView2.SetFocus
End Sub

Private Sub BtnOK_Click()
 PicVendor.Visible = False
End Sub






Private Sub CboNumber_Click()
      Dirty
      If rsItem.State = adStateOpen Then
          rsItem.Close
       End If
       rsItem.Open "SELECT * From [CheckItems] WHERE [Check_Number]like'" & CboNumber.text & "'"
       If rsItem.RecordCount > 0 Then
         Load_ITEMS
         lvItems.SetFocus
       Else
         lvItems.ListItems.Clear
       End If
exitSub:
End Sub

Private Sub cmdExit_Click()
 PicAccount.Visible = False
End Sub

Private Sub CmdRef_Click()
Dim strSQL As String
If rsAcct.State = adStateOpen Then
    rsAcct.Close
End If
strSQL = "SELECT Account,Description "
strSQL = strSQL & "From Expenses order by Description"
rsAcct.Open strSQL, cnRef, adOpenStatic, adLockOptimistic
Load_Expenses

End Sub

Private Sub cmdRefresh_Click()
Dim SQL As String
If rsVend.State = adStateOpen Then
    rsVend.Close
End If
SQL = "SELECT Vendor_Name,Address,Vendor_ID "
SQL = SQL & "From Vendor order by Vendor_Name"
rsVend.Open SQL, cnRef, adOpenStatic, adLockOptimistic
Load_Vendor
End Sub

Private Sub BtnVendor_Click()
 PicVendor.Visible = True
 ListView1.SetFocus
End Sub


Private Sub ChkPrint_Click()
 TxtEntry(9).text = CStr(ChkPrint.Value)
End Sub

Private Sub ChkShow_Click()
 TxtEntry(0).Visible = (ChkShow.Value = 1)
 TxtEntry(3).Visible = (ChkShow.Value = 1)
 TxtEntry(5).Visible = (ChkShow.Value = 1)
 TxtEntry(8).Visible = (ChkShow.Value = 1)
 lblFLDi(0).Visible = (ChkShow.Value = 1)
 lblFLDi(3).Visible = (ChkShow.Value = 1)
 lblFLDi(5).Visible = (ChkShow.Value = 1)
 lblFLDi(8).Visible = (ChkShow.Value = 1)
 BtnExpenses.Visible = (ChkShow.Value = 1)

End Sub

Private Sub nextNumber()
      Dim nextNo As Long
      Set rsTemp = New ADODB.Recordset
      Set rsTemp = rsCash
       If rsTemp.State = adStateOpen Then
          rsTemp.Close
       End If
      rsTemp.Open "SELECT * From CheckTrans order by SN", cnBank
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

Private Sub nextItemNumber()
      Dim nextNo As Long
      Set rsTemp = New ADODB.Recordset
       Set rsTemp = rsItem
        If rsTemp.State = adStateOpen Then
           rsTemp.Close
        End If
      rsTemp.Open "SELECT * From CheckItems order by SN", cnBank
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

Private Sub cmdButton_Click(Index As Integer)
'//                  A S E U C D R
'On Error GoTo ERRORHANDLE
Dim nextNo As Long
Select Case Index
   Case BtnAdd                       '<------ add new record ------->'
     addRec = True
     cmdButtonShow ("0100100"), Me
     ChkShow.Value = 1
     nextNumber
   Case BtnSave                       '<------ save new record ------>'
        cmdButtonShow ("1010011"), Me
        Call WriteData(Me, rsCash, True)
        Dim lngSN As Long
        lngSN = Val(TxtEntry(0).text)
        Call lvwPopulateData(lvList, rsCash, 2, lngSN)
        addRec = False
   Case BtnEdit                       '<------ edit record ---------->'
        editRec = True
        cmdButtonShow ("0001100"), Me
        ChkShow.Value = 1
        TxtEntry(1).SetFocus
   Case BtnUpdate                     '<------ update record -------->'
        cmdButtonShow ("1010011"), Me
        Call WriteData(Me, rsCash, False)
        LvwReplaceData Me, rsCash, lvList
        editRec = False
   Case BtnCancel                     '<------ cancel update -------->'
        cmdButtonShow ("1010011"), Me
        addRec = False
        editRec = False
   Case BtnDelete                     '<------ delete record -------->'
        '// no delete here please !
        Call Delete_Record(rsCash, lvList)
   Case BtnRefresh                    '<------ Refresh record ------->'
        addRec = False
        editRec = False
       If rsCash.State = adStateOpen Then
          rsCash.Close
        End If
        rsCash.Open "SELECT * From CheckTrans order by SN", cnBank, adOpenStatic, adLockOptimistic
        Load_DATA
        isFilter = False
        lvList.SetFocus
End Select
'ERRORHANDLE:
' errorMsg Err, Me.Name, "Command Button"

End Sub






Private Sub DTPicker1_CloseUp()
   TxtEntry(2).text = Format(DTPicker1.Value, "mmm-dd-yyyy")
   TxtEntry(2).SetFocus
End Sub

Private Sub Form_Load()
'// initialized
lblAmtInWord.BackStyle = 0
Set convert = New numTOword
isFilter = False
cmdButtonShow ("1010011"), Me
TbButtonShow "1010011"
SSTab1.Tab = 0
'//
    With LvMenu
        
'        Set .SmallIcons = ImageList2
'        Set .Icons = ImageList2
        'For Sales
        .ListItems.Add , "frmPrint", "Print Summary", 1, 1
        .ListItems.Add , "frmSearch", "Search / Filter", 2, 2
        .ListItems.Add , "frmReference", "Reference", 3, 3
        .ListItems.Add , "frmCalcu", "Calculator", 4, 4
        .ListItems.Add , "frmCheckBook", "Check Book", 5, 5
        .ListItems.Add , "help", "Help", 6, 6
        
     End With
'//
Set rsCash = New ADODB.Recordset
rsCash.Open "SELECT * From CheckTrans order by date", cnBank, adOpenStatic, adLockOptimistic
Load_DATA
Call ShowFldsLabel(Me, rsCash)
Call Add_Item(rsCash, "Check_Number", CboNumber)

Set rsItem = New ADODB.Recordset
rsItem.Open "SELECT * From CheckItems order by date", cnBank, adOpenStatic, adLockOptimistic
Load_ITEMS

Set rsVend = New ADODB.Recordset
Dim SQL As String
SQL = "SELECT Vendor_Name,Address,Vendor_ID "
SQL = SQL & "From Vendor order by Vendor_Name"
rsVend.Open SQL, cnRef, adOpenStatic, adLockOptimistic
Load_Vendor

Set rsAcct = New ADODB.Recordset
Dim strSQL As String
strSQL = "SELECT Account,Description "
strSQL = strSQL & "From Expenses order by Description"
rsAcct.Open strSQL, cnRef, adOpenStatic, adLockOptimistic
Load_Expenses

errMsg:
  errorMsg Err, Me.Name, "Form Load"
End Sub

Private Sub Load_DATA()
On Error GoTo ERRORHANDLE
'// set columnheaders
Call InsertColumn(lvList, rsCash)
'//set details
Call FillListView(lvList, rsCash, 2)
'//get total
 Call Listview_Total(lvList, rsCash)
ERRORHANDLE:
    errorMsg Err, Me.Name, "Load_Data"
End Sub

Private Sub Load_ITEMS()
On Error GoTo ERRORHANDLE
'// set columnheaders
Call InsertColumn(lvItems, rsItem)
'//set details
Call FillListView(lvItems, rsItem, 2)
'//get total
 Call Listview_Total(lvItems, rsItem)
ERRORHANDLE:
    errorMsg Err, Me.Name, "Load_Data"
End Sub

Private Sub Load_Vendor()
'// set columnheaders
'Insert_ExtraCol lvList, rsDed
If rsVend.RecordCount = 0 Then Exit Sub
Call InsertColumn(ListView1, rsVend)
'//set details
Call FillListView(ListView1, rsVend, 6)
End Sub

Private Sub Load_Expenses()
'// set columnheaders
'Insert_ExtraCol lvList, rsDed
If rsAcct.RecordCount = 0 Then Exit Sub
Call InsertColumn(ListView2, rsAcct)
'//set details
Call FillListView(ListView2, rsAcct, 4)
End Sub

Private Sub Form_Resize()
With Me
  If .WindowState = 0 Then
   .Height = 9525
   .Width = 13785
  End If
End With
' SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub ImgItemHelp_Click()
  myMsg "Items entry must be used only " _
  & "for itemized payment!" & vbCrLf & vbCrLf _
  & "", "Items Help", 2, True
End Sub

Private Sub Listview1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then
    PicVendor.Visible = False
    TxtEntry(6).SetFocus
  ElseIf KeyCode = 13 Then
   TxtEntry(6).text = ListView1.SelectedItem.text
   TxtEntry(7).text = ListView1.SelectedItem.ListSubItems(1).text
   TxtEntry(5).text = ListView1.SelectedItem.ListSubItems(2).text
   TxtEntry(7).SetFocus
   PicVendor.Visible = False
  End If
End Sub




Private Sub ListView2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then
    PicVendor.Visible = False
    TxtEntry(3).SetFocus
  ElseIf KeyCode = 13 Then
   TxtEntry(3).text = ListView2.SelectedItem.text & ":" & ListView2.SelectedItem.ListSubItems(1).text
   TxtEntry(3).SetFocus
   PicAccount.Visible = False
  End If
End Sub

Private Sub lvItems_Click()
  If addRec = True Or editRec = True Then Exit Sub
  Call BindDataItems(rsItem, lvItems, True)
End Sub

Private Sub lvItems_KeyUp(KeyCode As Integer, Shift As Integer)
  lvItems_Click
End Sub


Private Sub LvList_KeyUp(KeyCode As Integer, Shift As Integer)
  Call BindDatasource(Me, rsCash, lvList, True)
  If Val(TxtEntry(4).text) = 0 Then
     Exit Sub
  Else
    lblAmtInWord = "***" & convert.TOword(TxtEntry(4)) & "***"
  End If
End Sub
Private Sub LvList_Click()
  If addRec = True Or editRec = True Then Exit Sub
  Call BindDatasource(Me, rsCash, lvList, True)
  If Val(TxtEntry(4).text) = 0 Then
     Exit Sub
  Else
    lblAmtInWord = "***" & convert.TOword(TxtEntry(4)) & "***"
  End If
End Sub


Private Sub lvMenu_DblClick()
MyMenu = LvMenu.SelectedItem.text
    Select Case LvMenu.SelectedItem.Key
        Case "frmSearch" '//: loadForm frmSearch
           SSTab1.Tab = 0
            With frmSearch
               Set .pFindForm = Me
               Set .pFindRecset = rsCash
               Set .pFindCon = cnBank
                   .pFindTABLE = "CheckTrans"
                   .Caption = .Caption & " <:- Check Transaction -:>"
                   .show
            End With
        Case "frmPrint"  '//: loadForm frmPrint
          With frmPrint
               Set .pPrintForm = Me
               Set .pPrintRecset = rsCash
               Set .pPrintCon = cnBank
                   .pPrintTABLE = "CheckTrans"
                   .Caption = .Caption & " <:- Check Transaction -:>"
                  .show
            End With
        Case "frmReference": loadForm frmReference
        Case "frmCalcu": loadForm FrmCalcu
        Case "frmCheckBook": loadForm frmCheckBook
   End Select
End Sub


Private Sub PicAccount_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then
        Call Drag_It(PicAccount.hWnd)
  End If
End Sub





Private Sub PicVendor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then
        Call Drag_It(PicVendor.hWnd)
  End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
  Dirty
  If SSTab1.Tab = 1 Then
  lblSelected.Caption = ""
  lblSelected.Caption = "Selected Check Number: " & "( " & TxtEntry(1).text & " )"
  End If
End Sub


Private Sub TbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case "new"
     addRec = True
     TbButtonShow "0100100"
     TextEntry(2).text = TxtEntry(1).text
     TextEntry(5).text = TxtEntry(6).text
     nextItemNumber
  Case "save"
     TbButtonShow "1010011"
     Call WriteItems(rsItem, True)
     lvItems.SetFocus
     addRec = False
  Case "edit"
     editRec = True
     TbButtonShow "0001100"
  Case "update"
     TbButtonShow "1010011"
     Call WriteItems(rsItem, False)
     lvItems.SetFocus
     editRec = False
  Case "cancel"
     addRec = False
     editRec = False
     TbButtonShow "1010011"
     lvItems.SetFocus
  Case "delete"
     Call Delete_Record(rsItem, lvItems)
  Case "refresh"
       If rsItem.State = adStateOpen Then
          rsItem.Close
       End If
       rsItem.Open "SELECT * From CheckItems order by SN", cnBank, adOpenStatic, adLockOptimistic
       Load_ITEMS
       lvItems.SetFocus
End Select
End Sub



Private Sub TxtEntry_Change(Index As Integer)
Select Case Index
Case Is = 9
  ChkPrint.Value = Val(TxtEntry(9).text)
End Select
End Sub

Private Sub TbButtonShow(ByRef buttonString As String)
'< syntax:  TbButtonShow ("0001111")
''--------------------------------------------------
''-- This routine handles setting the visible --
''-- to true / false on the buttons.                --
''-------------------------------------------------
''-- A string of 0101 passed. If 0, visible   --
''-------------------------------------------------
Dim indx As Integer
buttonString = Trim$(buttonString)
For indx = 1 To Len(buttonString)
  If (Mid$(buttonString, indx, 1) = "1") Then
     Me.TbMenu.Buttons(indx).Visible = True    '(index-1) use only if index start from 0
  Else
     Me.TbMenu.Buttons(indx).Visible = False
  End If
Next

End Sub


Private Sub WriteItems(ByRef srcRS As Recordset, _
                      ByVal newRec As Boolean, _
                      Optional ByVal srcNumFlds As Integer = 0)
'//addnew = true for new record else false > forced
'//srcnumflds = number of fields loaded in textbox  > optional
                'if not all fields are loaded, srcnumflds is equal to text upperbound indeces
                'based on the numbers of textbox showed in the form (see enabled textbox procedures)
If srcRS Is Nothing Then Exit Sub
If srcRS.RecordCount > 0 Then
If srcRS.EOF = True Or srcRS.BOF = True Then
   'MsgBox "Either EOF or BOF reached.", vbInformation, "Write Data!"
   'Exit Sub
   srcRS.MoveLast
End If
End If
Dim i As Integer
Dim NOF As Integer 'Number Of Feilds
If srcNumFlds > 0 Then
   NOF = srcNumFlds
Else
   NOF = (srcRS.Fields.Count - 1)  'remember that indeces are zero based
End If
ReDim entries(NOF) As TextBox
For i = 0 To NOF
    Set entries(i) = TextEntry(i)  'm tired of using frm, set number of elements allowed
    Next i
i = 0
With srcRS
  If newRec = True Then
      .AddNew
  End If
      For i = 0 To NOF
      Select Case srcRS.Fields.Item(i).Type
       Case Is = 3   'integer
           If IsNumeric(entries(i).text) Then
              srcRS.Fields(i) = toNumber(entries(i).text)
           End If
      Case Is = 5, 6  'currency or double
           If IsNumeric(entries(i).text) Then
             srcRS.Fields(i) = toMoney(entries(i).text)
           End If
       Case Is = 7   'date
           If IsDate(entries(i).text) Then
               srcRS.Fields(i) = CDate(entries(i).text)
           Else '//save empty entry
               srcRS.Fields(i) = Null
           End If
       Case Is = 202, 203    'text, memo
             srcRS.Fields(i) = CStr(entries(i).text)
      End Select
      Next i
      .Update
End With
End Sub
Private Sub BindDataItems(ByRef srcRS As Recordset, _
                          ByRef lv As ListView, _
                          Optional ByVal findFirst As Boolean = True, _
                          Optional ByVal numOfFlds As Integer = 0)
'//findFIRST - optional/false when use for next,previous,last,first
   If srcRS Is Nothing Then Exit Sub
With srcRS
  If .RecordCount = 0 Then
      Exit Sub
   End If
End With
If findFirst = False Then
 If srcRS.EOF = True Then
    MsgBox "EOF reached.", vbInformation, "Bind Data!"
    Exit Sub
 ElseIf srcRS.BOF = True Then
   MsgBox "BOF reached.", vbInformation, "Bind Data!"
   Exit Sub
 End If
End If
Dim abPos As Boolean   'absolutePosition
Dim i As Integer
Dim strFind As String
Dim strMatch As String
Dim NOF As Integer 'Number Of Feilds
'//
If srcRS.RecordCount = 0 Then Exit Sub
'// initialized
If numOfFlds > 0 Then
   NOF = numOfFlds
Else
   NOF = (srcRS.Fields.Count - 1)  'remember that indeces are zero based
End If
For i = 0 To NOF
   TextEntry(i) = Empty
   Next i
If IsNumeric(TrimSpaces(CStr(lv.SelectedItem.text))) Then
    strFind = TrimSpaces(CStr(lv.SelectedItem.text))
    abPos = False
Else
    strFind = lv.SelectedItem.Index
    abPos = True
End If
If findFirst = True Then
 With srcRS
 .MoveFirst
   Do Until srcRS.EOF
   If abPos = False Then
        lv.MousePointer = vbHourglass
       strMatch = TrimSpaces(CStr(toNumber(srcRS.Fields(0))))
     Else
       lv.MousePointer = vbHourglass
       'slower//i use only on alpha type// so you can show the value one
       'row even if there is duplicate reference for viewing record
       'remember that reference must be a unique key
       strMatch = srcRS.Bookmark '// .AbsolutePosition
   End If
   If strMatch = strFind Then
         lv.MousePointer = vbDefault

         GoTo iFound
   Else
     .MoveNext
   End If
   Loop
 End With
 lv.MousePointer = vbDefault
End If 'findFirst
iFound:
With srcRS
         If srcRS.EOF = True Or srcRS.BOF = True Then Exit Sub
         For i = 0 To NOF
          If Not IsNull(srcRS.Fields(i)) Then
             TextEntry(i) = FormatRS(srcRS.Fields(i))
              If srcRS.Fields(i).Type = 6 Or srcRS.Fields(i).Type = 5 Then
                TextEntry(i).Alignment = 1
                 If Val(TextEntry(i)) = 0 Then
                      TxtEntry(i).ForeColor = &HD38545
                 ElseIf Val(TextEntry(i)) < 0 Then
                      TextEntry(i).ForeColor = vbRed      ' if the value is negative
                 Else
                      TextEntry(i).ForeColor = vbBlack
                End If
             Else                                          'string value and non-zero value
                   TextEntry(i).ForeColor = vbBlack
            End If
          Else
              TextEntry(i) = Empty
          End If
         Next i
    '//end of Search
End With

End Sub

Private Sub Dirty()
  If addRec = True Then
    MsgBox "You have pending records to save !", vbCritical, "Add"
    Exit Sub
  ElseIf editRec = True Then
   MsgBox "You have pending records to update !", vbCritical, "Edit"
    Exit Sub
  End If
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
nxTab = Index
Select Case nxTab
Case Is = 7, 8
 TxtEntry(nxTab).SelStart = Len(TxtEntry(nxTab).text)
Case Else
 TxtEntry(nxTab).SelStart = 0
 TxtEntry(nxTab).SelLength = Len(TxtEntry(nxTab).text)
End Select
End Sub

Private Sub txtEntry_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim lastTab As Integer
On Error GoTo ERRORHANDLE
lastTab = 8  ' rsCash.Fields.Count - 1 'or txtEntry Upper Bound if kung limitado lang ang textbox
If KeyCode = 13 Then
ChkShow.Value = 1
     If nxTab = lastTab Then Exit Sub
'     If nxTab = 3 Then
'        nxTab = Index                              'stay foot ka lang
'        Exit Sub
'     End If
     nxTab = nxTab + 1
     If nxTab = 3 Then nxTab = 4                   '//current tab is 2 *passed 3 punta ka ng 4
     If nxTab = 5 Then nxTab = 6                   '//current tab is 4 *Passed 5 punta ka ng 6
     If nxTab = 7 Then nxTab = 8                   '//current tab is 6 *Passed 7 punta ka ng 8
ElseIf KeyCode = 38 Then  'up arrow key
ChkShow.Value = 1
     If nxTab = 0 Or nxTab = 1 Then Exit Sub
     nxTab = nxTab - 1
     If nxTab = 2 Then nxTab = 6                   '//current tab is 3 *passed 2 balik ka 6
     If nxTab = 5 Then nxTab = 4                   '//current tab is 6 *passed 5 balik ka sa 4
     If nxTab = 3 Then nxTab = 2

End If
TxtEntry(nxTab).SetFocus
ERRORHANDLE:
 errorMsg Err, Me.Name
End Sub
