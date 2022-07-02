VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAccountMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Master"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13110
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   13110
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   7470
      Left            =   15
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   0
      Width           =   13095
      _Version        =   65536
      _ExtentX        =   23098
      _ExtentY        =   13176
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TintColor       =   16711935
      Alignment       =   0
      AutoSize        =   0   'False
      BevelSize       =   0
      BevelStyle      =   0
      BorderColor     =   -2147483642
      BorderStyle     =   1
      FillColor       =   -2147483633
      FontStyle       =   0
      FontTransparent =   0   'False
      LightColor      =   -2147483643
      ShadowColor     =   -2147483632
      TextColor       =   -2147483640
      WallPaper       =   0
      NoPrefix        =   0   'False
      FormatString    =   ""
      Caption         =   ""
      Picture         =   "AccountMaster.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   7110
         Left            =   120
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   120
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   12541
         _Version        =   393216
         Style           =   1
         Tabs            =   10
         TabsPerRow      =   8
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "&List"
         TabPicture(0)   =   "AccountMaster.frx":001C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Mh3dLabel1(3)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Mh3dLabel1(2)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "DataGrid1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Text1"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "AccountMaster.frx":0038
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "btnNotes"
         Tab(1).Control(1)=   "Mh3dFrame2(0)"
         Tab(1).Control(2)=   "txtNotes"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "&Details"
         TabPicture(2)   =   "AccountMaster.frx":0054
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         TabCaption(3)   =   "&Details"
         TabPicture(3)   =   "AccountMaster.frx":0070
         Tab(3).ControlEnabled=   0   'False
         Tab(3).ControlCount=   0
         TabCaption(4)   =   "&Details"
         TabPicture(4)   =   "AccountMaster.frx":008C
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Mh3dFrame2(3)"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "&Details"
         TabPicture(5)   =   "AccountMaster.frx":00A8
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Mh3dFrame2(4)"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "&Details"
         TabPicture(6)   =   "AccountMaster.frx":00C4
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "Mh3dFrame2(5)"
         Tab(6).ControlCount=   1
         TabCaption(7)   =   "&Details"
         TabPicture(7)   =   "AccountMaster.frx":00E0
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "Mh3dFrame2(6)"
         Tab(7).ControlCount=   1
         TabCaption(8)   =   "&Details"
         TabPicture(8)   =   "AccountMaster.frx":00FC
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "Mh3dFrame2(7)"
         Tab(8).ControlCount=   1
         TabCaption(9)   =   "&Op.Bal."
         TabPicture(9)   =   "AccountMaster.frx":0118
         Tab(9).ControlEnabled=   0   'False
         Tab(9).Control(0)=   "Mh3dFrame2(8)"
         Tab(9).ControlCount=   1
         Begin VB.CommandButton btnNotes 
            Caption         =   " Notes"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   -74880
            TabIndex        =   14
            Top             =   6600
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   840
            TabIndex        =   80
            Top             =   6630
            Width           =   7575
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   6225
            Left            =   120
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   360
            Width           =   12615
            _ExtentX        =   22251
            _ExtentY        =   10980
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BackColor       =   9164542
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   18
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   "Name"
               Caption         =   "Name"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "Alias"
               Caption         =   "Alias"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               ScrollBars      =   3
               AllowRowSizing  =   0   'False
               AllowSizing     =   0   'False
               Locked          =   -1  'True
               BeginProperty Column00 
                  Locked          =   -1  'True
                  ColumnWidth     =   9884.977
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   2145.26
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   3060
            Index           =   0
            Left            =   -74880
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   720
            Width           =   12615
            _Version        =   65536
            _ExtentX        =   22251
            _ExtentY        =   5397
            _StockProps     =   77
            Enabled         =   0   'False
            TintColor       =   16711935
            Alignment       =   0
            AutoSize        =   0   'False
            BevelSize       =   0
            BevelStyle      =   0
            BorderColor     =   -2147483642
            BorderStyle     =   1
            FillColor       =   -2147483633
            FontStyle       =   0
            FontTransparent =   0   'False
            LightColor      =   -2147483643
            ShadowColor     =   -2147483632
            TextColor       =   -2147483640
            WallPaper       =   0
            NoPrefix        =   0   'False
            FormatString    =   ""
            Caption         =   ""
            Picture         =   "AccountMaster.frx":0134
            Begin VB.TextBox Text14 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   0
               Left            =   9000
               MaxLength       =   40
               TabIndex        =   9
               Top             =   1995
               Width           =   3495
            End
            Begin VB.TextBox Text13 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   0
               Left            =   10560
               MaxLength       =   40
               TabIndex        =   2
               Top             =   420
               Width           =   1935
            End
            Begin VB.TextBox Text10 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   0
               Left            =   1680
               MaxLength       =   40
               TabIndex        =   3
               Top             =   740
               Width           =   7335
            End
            Begin VB.TextBox Text12 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   0
               Left            =   10560
               MaxLength       =   40
               TabIndex        =   11
               Top             =   2310
               Width           =   1935
            End
            Begin VB.TextBox Text11 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   0
               Left            =   1680
               MaxLength       =   80
               TabIndex        =   12
               Top             =   2625
               Width           =   7335
            End
            Begin VB.TextBox Text9 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   0
               Left            =   10560
               MaxLength       =   40
               TabIndex        =   13
               Top             =   2625
               Width           =   1935
            End
            Begin VB.TextBox Text8 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   0
               Left            =   1680
               MaxLength       =   40
               TabIndex        =   10
               Top             =   2310
               Width           =   7335
            End
            Begin VB.TextBox Text7 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   0
               Left            =   1680
               MaxLength       =   40
               TabIndex        =   8
               Top             =   1995
               Width           =   5775
            End
            Begin VB.TextBox Text6 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   0
               Left            =   1680
               MaxLength       =   40
               TabIndex        =   7
               Top             =   1680
               Width           =   10815
            End
            Begin VB.TextBox Text5 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   0
               Left            =   1680
               MaxLength       =   40
               TabIndex        =   6
               Top             =   1365
               Width           =   10815
            End
            Begin VB.TextBox Text4 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   0
               Left            =   1680
               MaxLength       =   40
               TabIndex        =   5
               Top             =   1050
               Width           =   10815
            End
            Begin VB.TextBox Text2 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   0
               Left            =   1680
               MaxLength       =   40
               TabIndex        =   0
               ToolTipText     =   "Hidden Notes"
               Top             =   100
               Width           =   10815
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   75
               Top             =   420
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Print Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0150
               Picture         =   "AccountMaster.frx":016C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   74
               Top             =   105
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0188
               Picture         =   "AccountMaster.frx":01A4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   1275
               Index           =   0
               Left            =   120
               TabIndex        =   83
               Top             =   1050
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   2249
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Address"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":01C0
               Picture         =   "AccountMaster.frx":01DC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   84
               Top             =   2310
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Phone"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":01F8
               Picture         =   "AccountMaster.frx":0214
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   85
               Top             =   2625
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " E-mail"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0230
               Picture         =   "AccountMaster.frx":024C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Index           =   0
               Left            =   9000
               TabIndex        =   86
               Top             =   2625
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " GST No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0268
               Picture         =   "AccountMaster.frx":0284
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   0
               Left            =   9000
               TabIndex        =   87
               Top             =   2310
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Mobile"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":02A0
               Picture         =   "AccountMaster.frx":02BC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   13
               Left            =   9000
               TabIndex        =   131
               Top             =   420
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Alias"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":02D8
               Picture         =   "AccountMaster.frx":02F4
            End
            Begin VB.TextBox Text3 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   0
               Left            =   1680
               MaxLength       =   40
               TabIndex        =   1
               Top             =   420
               Width           =   7335
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Index           =   2
               Left            =   120
               TabIndex        =   140
               Top             =   740
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Group"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0310
               Picture         =   "AccountMaster.frx":032C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   2
               Left            =   9000
               TabIndex        =   155
               Top             =   740
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Opening Balance"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0348
               Picture         =   "AccountMaster.frx":0364
            End
            Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
               Height          =   330
               Left            =   10560
               TabIndex        =   4
               Top             =   740
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
               _ExtentY        =   582
               Calculator      =   "AccountMaster.frx":0380
               Caption         =   "AccountMaster.frx":03A0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "AccountMaster.frx":040C
               Keys            =   "AccountMaster.frx":042A
               Spin            =   "AccountMaster.frx":0474
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "#########0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "#########0.00;-#########0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999999.99
               MinValue        =   -9999999999.99
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   8
               Left            =   7440
               TabIndex        =   156
               Top             =   1995
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " State"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":049C
               Picture         =   "AccountMaster.frx":04B8
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   6210
            Index           =   4
            Left            =   -74880
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   720
            Width           =   12615
            _Version        =   65536
            _ExtentX        =   22251
            _ExtentY        =   10954
            _StockProps     =   77
            Enabled         =   0   'False
            TintColor       =   16711935
            Alignment       =   0
            AutoSize        =   0   'False
            BevelSize       =   0
            BevelStyle      =   0
            BorderColor     =   -2147483642
            BorderStyle     =   1
            FillColor       =   -2147483633
            FontStyle       =   0
            FontTransparent =   0   'False
            LightColor      =   -2147483643
            ShadowColor     =   -2147483632
            TextColor       =   -2147483640
            WallPaper       =   0
            NoPrefix        =   0   'False
            FormatString    =   ""
            Caption         =   ""
            Picture         =   "AccountMaster.frx":04D4
            Begin VB.TextBox Text8 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   4
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   145
               TabStop         =   0   'False
               Top             =   2025
               Width           =   6975
            End
            Begin VB.TextBox Text12 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   4
               Left            =   10560
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   144
               TabStop         =   0   'False
               Top             =   2025
               Width           =   1935
            End
            Begin VB.TextBox Text9 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   4
               Left            =   10560
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   143
               TabStop         =   0   'False
               Top             =   2340
               Width           =   1935
            End
            Begin VB.TextBox Text2 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   4
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   100
               Width           =   10815
            End
            Begin VB.TextBox Text3 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   4
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   410
               Width           =   6975
            End
            Begin VB.TextBox Text4 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   4
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   725
               Width           =   10815
            End
            Begin VB.TextBox Text5 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   4
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   1040
               Width           =   10815
            End
            Begin VB.TextBox Text6 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   4
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   1355
               Width           =   10815
            End
            Begin VB.TextBox Text7 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   380
               Index           =   4
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   1670
               Width           =   10815
            End
            Begin VB.TextBox Text11 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   4
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   80
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   2340
               Width           =   6975
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   320
               Index           =   4
               Left            =   120
               TabIndex        =   89
               Top             =   420
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   564
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Print Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":04F0
               Picture         =   "AccountMaster.frx":050C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   4
               Left            =   120
               TabIndex        =   90
               Top             =   105
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0528
               Picture         =   "AccountMaster.frx":0544
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   1320
               Index           =   4
               Left            =   120
               TabIndex        =   91
               Top             =   720
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   2328
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   "  Address"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0560
               Picture         =   "AccountMaster.frx":057C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Index           =   4
               Left            =   120
               TabIndex        =   92
               Top             =   2340
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " E-mail"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0598
               Picture         =   "AccountMaster.frx":05B4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Index           =   4
               Left            =   8640
               TabIndex        =   93
               Top             =   2340
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " GST No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":05D0
               Picture         =   "AccountMaster.frx":05EC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   4
               Left            =   8640
               TabIndex        =   94
               Top             =   2025
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Mobile"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0608
               Picture         =   "AccountMaster.frx":0624
            End
            Begin MSDataGridLib.DataGrid DataGrid2 
               Height          =   2925
               Index           =   4
               Left            =   120
               TabIndex        =   41
               Top             =   3180
               Width           =   12375
               _ExtentX        =   21828
               _ExtentY        =   5159
               _Version        =   393216
               AllowUpdate     =   0   'False
               AllowArrows     =   -1  'True
               Appearance      =   0
               BackColor       =   9164542
               HeadLines       =   1
               RowHeight       =   18
               TabAction       =   2
               FormatLocked    =   -1  'True
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnCount     =   5
               BeginProperty Column00 
                  DataField       =   "SizeName"
                  Caption         =   "         Item Size"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   "Range1"
                  Caption         =   "Range (1)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2057
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column02 
                  DataField       =   "Range2"
                  Caption         =   "Range (2)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2057
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column03 
                  DataField       =   "Range4"
                  Caption         =   "Range (4)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2057
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column04 
                  DataField       =   "Range6"
                  Caption         =   "Range (SPL. Col)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2057
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  MarqueeStyle    =   3
                  ScrollBars      =   3
                  AllowRowSizing  =   0   'False
                  AllowSizing     =   0   'False
                  Locked          =   -1  'True
                  BeginProperty Column00 
                     Locked          =   -1  'True
                     ColumnWidth     =   5025.26
                  EndProperty
                  BeginProperty Column01 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   1695.118
                  EndProperty
                  BeginProperty Column02 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   1695.118
                  EndProperty
                  BeginProperty Column03 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   1695.118
                  EndProperty
                  BeginProperty Column04 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   1695.118
                  EndProperty
               EndProperty
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   11
               Left            =   8640
               TabIndex        =   129
               Top             =   405
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Alias"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0640
               Picture         =   "AccountMaster.frx":065C
            End
            Begin VB.TextBox Text13 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   4
               Left            =   10560
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   410
               Width           =   1935
            End
            Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame4 
               Height          =   330
               Left            =   1680
               TabIndex        =   141
               TabStop         =   0   'False
               Top             =   2655
               Width           =   10815
               _Version        =   65536
               _ExtentX        =   19076
               _ExtentY        =   582
               _StockProps     =   77
               TintColor       =   16711935
               Alignment       =   0
               AutoSize        =   0   'False
               BevelSize       =   0
               BevelStyle      =   0
               BorderColor     =   -2147483642
               BorderStyle     =   1
               FillColor       =   16777215
               FontStyle       =   0
               FontTransparent =   0   'False
               LightColor      =   -2147483643
               ShadowColor     =   -2147483632
               TextColor       =   -2147483640
               WallPaper       =   0
               NoPrefix        =   0   'False
               FormatString    =   ""
               Caption         =   ""
               Picture         =   "AccountMaster.frx":0678
               Begin VB.CheckBox chkRound 
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   290
                  Left            =   120
                  TabIndex        =   40
                  Top             =   30
                  Width           =   255
               End
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Index           =   2
               Left            =   120
               TabIndex        =   142
               Top             =   2655
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Round Off Qty"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0694
               Picture         =   "AccountMaster.frx":06B0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Index           =   4
               Left            =   120
               TabIndex        =   146
               Top             =   2025
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Phone"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":06CC
               Picture         =   "AccountMaster.frx":06E8
            End
            Begin VB.Line Line4 
               X1              =   0
               X2              =   12600
               Y1              =   3075
               Y2              =   3075
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   6210
            Index           =   5
            Left            =   -74880
            TabIndex        =   95
            Top             =   720
            Width           =   12615
            _Version        =   65536
            _ExtentX        =   22251
            _ExtentY        =   10954
            _StockProps     =   77
            Enabled         =   0   'False
            TintColor       =   16711935
            Alignment       =   0
            AutoSize        =   0   'False
            BevelSize       =   0
            BevelStyle      =   0
            BorderColor     =   -2147483642
            BorderStyle     =   1
            FillColor       =   -2147483633
            FontStyle       =   0
            FontTransparent =   0   'False
            LightColor      =   -2147483643
            ShadowColor     =   -2147483632
            TextColor       =   -2147483640
            WallPaper       =   0
            NoPrefix        =   0   'False
            FormatString    =   ""
            Caption         =   ""
            Picture         =   "AccountMaster.frx":0704
            Begin VB.TextBox Text13 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   5
               Left            =   10560
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   151
               TabStop         =   0   'False
               Top             =   420
               Width           =   1935
            End
            Begin VB.TextBox Text2 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   5
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   147
               TabStop         =   0   'False
               Top             =   100
               Width           =   10815
            End
            Begin VB.TextBox Text3 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   5
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   420
               Width           =   7335
            End
            Begin VB.TextBox Text4 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   5
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   725
               Width           =   10815
            End
            Begin VB.TextBox Text5 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   5
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   1040
               Width           =   10815
            End
            Begin VB.TextBox Text6 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   5
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   1355
               Width           =   10815
            End
            Begin VB.TextBox Text7 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   5
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   1670
               Width           =   10815
            End
            Begin VB.TextBox Text8 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   5
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   1980
               Width           =   7335
            End
            Begin VB.TextBox Text9 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   5
               Left            =   10560
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   50
               TabStop         =   0   'False
               Top             =   2300
               Width           =   1935
            End
            Begin VB.TextBox Text11 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   5
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   80
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   2300
               Width           =   7335
            End
            Begin VB.TextBox Text12 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   5
               Left            =   10560
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   1980
               Width           =   1935
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   320
               Index           =   5
               Left            =   120
               TabIndex        =   96
               Top             =   420
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   564
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Print Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0720
               Picture         =   "AccountMaster.frx":073C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   1280
               Index           =   5
               Left            =   120
               TabIndex        =   97
               Top             =   720
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   2258
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Address"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0758
               Picture         =   "AccountMaster.frx":0774
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Index           =   5
               Left            =   120
               TabIndex        =   98
               Top             =   1980
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Phone"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0790
               Picture         =   "AccountMaster.frx":07AC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Index           =   5
               Left            =   120
               TabIndex        =   99
               Top             =   2300
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " E-mail"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":07C8
               Picture         =   "AccountMaster.frx":07E4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Index           =   5
               Left            =   9000
               TabIndex        =   100
               Top             =   2300
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " GST No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0800
               Picture         =   "AccountMaster.frx":081C
            End
            Begin MSDataGridLib.DataGrid DataGrid2 
               Height          =   3165
               Index           =   5
               Left            =   120
               TabIndex        =   118
               Top             =   2940
               Width           =   12375
               _ExtentX        =   21828
               _ExtentY        =   5583
               _Version        =   393216
               AllowUpdate     =   0   'False
               AllowArrows     =   -1  'True
               Appearance      =   0
               BackColor       =   9164542
               HeadLines       =   1
               RowHeight       =   18
               TabAction       =   2
               FormatLocked    =   -1  'True
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnCount     =   5
               BeginProperty Column00 
                  DataField       =   "SizeName"
                  Caption         =   "     Item Size"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   "Range1"
                  Caption         =   "  Range (1)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2057
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column02 
                  DataField       =   "Range2"
                  Caption         =   "  Range (2)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2057
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column03 
                  DataField       =   "Range4"
                  Caption         =   "  Range (4)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2057
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column04 
                  DataField       =   "Range6"
                  Caption         =   "  Range (Spl. Col)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2057
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  MarqueeStyle    =   3
                  ScrollBars      =   3
                  AllowRowSizing  =   0   'False
                  AllowSizing     =   0   'False
                  Locked          =   -1  'True
                  BeginProperty Column00 
                     Locked          =   -1  'True
                     ColumnWidth     =   4844.977
                  EndProperty
                  BeginProperty Column01 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   1755.213
                  EndProperty
                  BeginProperty Column02 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   1785.26
                  EndProperty
                  BeginProperty Column03 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   1860.095
                  EndProperty
                  BeginProperty Column04 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   1544.882
                  EndProperty
               EndProperty
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   5
               Left            =   120
               TabIndex        =   148
               Top             =   100
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0838
               Picture         =   "AccountMaster.frx":0854
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   5
               Left            =   9000
               TabIndex        =   149
               Top             =   1980
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Mobile"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0870
               Picture         =   "AccountMaster.frx":088C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   10
               Left            =   9000
               TabIndex        =   150
               Top             =   420
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Alias"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":08A8
               Picture         =   "AccountMaster.frx":08C4
            End
            Begin VB.Line Line5 
               X1              =   0
               X2              =   12600
               Y1              =   2780
               Y2              =   2780
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   6210
            Index           =   6
            Left            =   -74880
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   720
            Width           =   12615
            _Version        =   65536
            _ExtentX        =   22251
            _ExtentY        =   10954
            _StockProps     =   77
            Enabled         =   0   'False
            TintColor       =   16711935
            Alignment       =   0
            AutoSize        =   0   'False
            BevelSize       =   0
            BevelStyle      =   0
            BorderColor     =   -2147483642
            BorderStyle     =   1
            FillColor       =   -2147483633
            FontStyle       =   0
            FontTransparent =   0   'False
            LightColor      =   -2147483643
            ShadowColor     =   -2147483632
            TextColor       =   -2147483640
            WallPaper       =   0
            NoPrefix        =   0   'False
            FormatString    =   ""
            Caption         =   ""
            Picture         =   "AccountMaster.frx":08E0
            Begin VB.TextBox Text2 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   6
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   100
               Width           =   10815
            End
            Begin VB.TextBox Text3 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   6
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   52
               TabStop         =   0   'False
               Top             =   410
               Width           =   7335
            End
            Begin VB.TextBox Text4 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   6
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   725
               Width           =   10815
            End
            Begin VB.TextBox Text5 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   6
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   1040
               Width           =   10815
            End
            Begin VB.TextBox Text6 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   6
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   56
               TabStop         =   0   'False
               Top             =   1355
               Width           =   10815
            End
            Begin VB.TextBox Text7 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   6
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   1670
               Width           =   10815
            End
            Begin VB.TextBox Text8 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   6
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   1980
               Width           =   7335
            End
            Begin VB.TextBox Text9 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   6
               Left            =   10560
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   61
               TabStop         =   0   'False
               Top             =   2300
               Width           =   1935
            End
            Begin VB.TextBox Text11 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   6
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   80
               TabIndex        =   60
               TabStop         =   0   'False
               Top             =   2300
               Width           =   7335
            End
            Begin VB.TextBox Text12 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   6
               Left            =   10560
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   59
               TabStop         =   0   'False
               Top             =   1980
               Width           =   1935
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   320
               Index           =   6
               Left            =   120
               TabIndex        =   102
               Top             =   420
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   564
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Print Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":08FC
               Picture         =   "AccountMaster.frx":0918
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   6
               Left            =   120
               TabIndex        =   103
               Top             =   105
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0934
               Picture         =   "AccountMaster.frx":0950
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   1270
               Index           =   6
               Left            =   120
               TabIndex        =   104
               Top             =   720
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   2240
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Address"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":096C
               Picture         =   "AccountMaster.frx":0988
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Index           =   6
               Left            =   120
               TabIndex        =   105
               Top             =   1980
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Phone"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":09A4
               Picture         =   "AccountMaster.frx":09C0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Index           =   6
               Left            =   120
               TabIndex        =   106
               Top             =   2300
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " E-mail"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":09DC
               Picture         =   "AccountMaster.frx":09F8
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Index           =   6
               Left            =   9000
               TabIndex        =   107
               Top             =   2295
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " GST No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0A14
               Picture         =   "AccountMaster.frx":0A30
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   6
               Left            =   9000
               TabIndex        =   108
               Top             =   1980
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Mobile"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0A4C
               Picture         =   "AccountMaster.frx":0A68
            End
            Begin MSDataGridLib.DataGrid DataGrid2 
               Height          =   3285
               Index           =   6
               Left            =   120
               TabIndex        =   117
               Top             =   2820
               Width           =   12375
               _ExtentX        =   21828
               _ExtentY        =   5794
               _Version        =   393216
               AllowUpdate     =   0   'False
               AllowArrows     =   -1  'True
               Appearance      =   0
               BackColor       =   9164542
               Enabled         =   -1  'True
               HeadLines       =   1
               RowHeight       =   18
               TabAction       =   2
               FormatLocked    =   -1  'True
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnCount     =   5
               BeginProperty Column00 
                  DataField       =   "OperationName"
                  Caption         =   "    Operation"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   "CalcModeName"
                  Caption         =   "    Calc Mode"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column02 
                  DataField       =   "SizeName"
                  Caption         =   "   Size"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column03 
                  DataField       =   "Range"
                  Caption         =   "             Range"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column04 
                  DataField       =   "Rate"
                  Caption         =   "           Rate"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.000"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  MarqueeStyle    =   3
                  ScrollBars      =   3
                  AllowRowSizing  =   0   'False
                  AllowSizing     =   0   'False
                  Locked          =   -1  'True
                  BeginProperty Column00 
                     Locked          =   -1  'True
                     ColumnWidth     =   3674.835
                  EndProperty
                  BeginProperty Column01 
                     Locked          =   -1  'True
                     ColumnWidth     =   2280.189
                  EndProperty
                  BeginProperty Column02 
                     Locked          =   -1  'True
                     ColumnWidth     =   2025.071
                  EndProperty
                  BeginProperty Column03 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   1980.284
                  EndProperty
                  BeginProperty Column04 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   1830.047
                  EndProperty
               EndProperty
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   9
               Left            =   9000
               TabIndex        =   128
               Top             =   420
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Alias"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0A84
               Picture         =   "AccountMaster.frx":0AA0
            End
            Begin VB.TextBox Text13 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   6
               Left            =   10560
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   53
               TabStop         =   0   'False
               Top             =   420
               Width           =   1935
            End
            Begin VB.Line Line6 
               X1              =   0
               X2              =   12600
               Y1              =   2720
               Y2              =   2720
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   6225
            Index           =   7
            Left            =   -74880
            TabIndex        =   109
            TabStop         =   0   'False
            Top             =   720
            Width           =   12615
            _Version        =   65536
            _ExtentX        =   22251
            _ExtentY        =   10980
            _StockProps     =   77
            Enabled         =   0   'False
            TintColor       =   16711935
            Alignment       =   0
            AutoSize        =   0   'False
            BevelSize       =   0
            BevelStyle      =   0
            BorderColor     =   -2147483642
            BorderStyle     =   1
            FillColor       =   -2147483633
            FontStyle       =   0
            FontTransparent =   0   'False
            LightColor      =   -2147483643
            ShadowColor     =   -2147483632
            TextColor       =   -2147483640
            WallPaper       =   0
            NoPrefix        =   0   'False
            FormatString    =   ""
            Caption         =   ""
            Picture         =   "AccountMaster.frx":0ABC
            Begin VB.TextBox Text2 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   7
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   62
               TabStop         =   0   'False
               Top             =   100
               Width           =   10815
            End
            Begin VB.TextBox Text3 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   7
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   63
               TabStop         =   0   'False
               Top             =   420
               Width           =   7335
            End
            Begin VB.TextBox Text4 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   7
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   65
               TabStop         =   0   'False
               Top             =   725
               Width           =   10815
            End
            Begin VB.TextBox Text5 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   7
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   66
               TabStop         =   0   'False
               Top             =   1040
               Width           =   10815
            End
            Begin VB.TextBox Text6 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   7
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   67
               TabStop         =   0   'False
               Top             =   1355
               Width           =   10815
            End
            Begin VB.TextBox Text7 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   7
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   68
               TabStop         =   0   'False
               Top             =   1670
               Width           =   10815
            End
            Begin VB.TextBox Text8 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   7
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   69
               TabStop         =   0   'False
               Top             =   1980
               Width           =   7335
            End
            Begin VB.TextBox Text9 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   7
               Left            =   10560
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   72
               TabStop         =   0   'False
               Top             =   2300
               Width           =   1935
            End
            Begin VB.TextBox Text11 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   7
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   80
               TabIndex        =   71
               TabStop         =   0   'False
               Top             =   2300
               Width           =   7335
            End
            Begin VB.TextBox Text12 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   7
               Left            =   10560
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   70
               TabStop         =   0   'False
               Top             =   1980
               Width           =   1935
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   315
               Index           =   7
               Left            =   120
               TabIndex        =   110
               Top             =   420
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   556
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Print Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0AD8
               Picture         =   "AccountMaster.frx":0AF4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   7
               Left            =   120
               TabIndex        =   111
               Top             =   105
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0B10
               Picture         =   "AccountMaster.frx":0B2C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   1270
               Index           =   7
               Left            =   120
               TabIndex        =   112
               Top             =   720
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   2240
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Address"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0B48
               Picture         =   "AccountMaster.frx":0B64
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Index           =   7
               Left            =   120
               TabIndex        =   113
               Top             =   1980
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Phone"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0B80
               Picture         =   "AccountMaster.frx":0B9C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Index           =   7
               Left            =   120
               TabIndex        =   114
               Top             =   2300
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " E-mail"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0BB8
               Picture         =   "AccountMaster.frx":0BD4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Index           =   7
               Left            =   9000
               TabIndex        =   115
               Top             =   2295
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " GST No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0BF0
               Picture         =   "AccountMaster.frx":0C0C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   7
               Left            =   9000
               TabIndex        =   116
               Top             =   1980
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Mobile"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0C28
               Picture         =   "AccountMaster.frx":0C44
            End
            Begin MSDataGridLib.DataGrid DataGrid2 
               Height          =   3300
               Index           =   7
               Left            =   120
               TabIndex        =   73
               Top             =   2820
               Width           =   12375
               _ExtentX        =   21828
               _ExtentY        =   5821
               _Version        =   393216
               AllowUpdate     =   0   'False
               AllowArrows     =   -1  'True
               Appearance      =   0
               BackColor       =   9164542
               HeadLines       =   1
               RowHeight       =   18
               TabAction       =   2
               FormatLocked    =   -1  'True
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnCount     =   9
               BeginProperty Column00 
                  DataField       =   "SizeName"
                  Caption         =   "    Item Size"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   "BindingTypeName"
                  Caption         =   "Binding Type"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column02 
                  DataField       =   "Range04"
                  Caption         =   "Range (04)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column03 
                  DataField       =   "Range08"
                  Caption         =   "Range (08)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column04 
                  DataField       =   "Range12"
                  Caption         =   "Range (12)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column05 
                  DataField       =   "Range16"
                  Caption         =   "Range (16)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column06 
                  DataField       =   "Range24"
                  Caption         =   "Range (24)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column07 
                  DataField       =   "Range32"
                  Caption         =   "Range (32)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column08 
                  DataField       =   "Range64"
                  Caption         =   "Range (64)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  MarqueeStyle    =   3
                  ScrollBars      =   3
                  AllowRowSizing  =   0   'False
                  AllowSizing     =   0   'False
                  Locked          =   -1  'True
                  BeginProperty Column00 
                     Locked          =   -1  'True
                     ColumnWidth     =   2684.977
                  EndProperty
                  BeginProperty Column01 
                     Locked          =   -1  'True
                     ColumnWidth     =   2654.929
                  EndProperty
                  BeginProperty Column02 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   915.024
                  EndProperty
                  BeginProperty Column03 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   915.024
                  EndProperty
                  BeginProperty Column04 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   915.024
                  EndProperty
                  BeginProperty Column05 
                     ColumnWidth     =   929.764
                  EndProperty
                  BeginProperty Column06 
                     ColumnWidth     =   929.764
                  EndProperty
                  BeginProperty Column07 
                     ColumnWidth     =   929.764
                  EndProperty
                  BeginProperty Column08 
                     ColumnWidth     =   929.764
                  EndProperty
               EndProperty
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   3
               Left            =   9000
               TabIndex        =   127
               Top             =   420
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Alias1"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0C60
               Picture         =   "AccountMaster.frx":0C7C
            End
            Begin VB.TextBox Text13 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   7
               Left            =   10560
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   64
               TabStop         =   0   'False
               Top             =   420
               Width           =   1935
            End
            Begin VB.Line Line7 
               X1              =   0
               X2              =   12600
               Y1              =   2715
               Y2              =   2715
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   6210
            Index           =   3
            Left            =   -74880
            TabIndex        =   119
            TabStop         =   0   'False
            Top             =   720
            Width           =   12615
            _Version        =   65536
            _ExtentX        =   22251
            _ExtentY        =   10954
            _StockProps     =   77
            Enabled         =   0   'False
            TintColor       =   16711935
            Alignment       =   0
            AutoSize        =   0   'False
            BevelSize       =   0
            BevelStyle      =   0
            BorderColor     =   -2147483642
            BorderStyle     =   1
            FillColor       =   -2147483633
            FontStyle       =   0
            FontTransparent =   0   'False
            LightColor      =   -2147483643
            ShadowColor     =   -2147483632
            TextColor       =   -2147483640
            WallPaper       =   0
            NoPrefix        =   0   'False
            FormatString    =   ""
            Caption         =   ""
            Picture         =   "AccountMaster.frx":0C98
            Begin VB.TextBox Text2 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   3
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   100
               Width           =   10815
            End
            Begin VB.TextBox Text3 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   3
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   410
               Width           =   7335
            End
            Begin VB.TextBox Text4 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   3
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   725
               Width           =   10815
            End
            Begin VB.TextBox Text5 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   3
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   1040
               Width           =   10815
            End
            Begin VB.TextBox Text6 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   3
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   1355
               Width           =   10815
            End
            Begin VB.TextBox Text7 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   3
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   1670
               Width           =   10815
            End
            Begin VB.TextBox Text8 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   3
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   1980
               Width           =   7335
            End
            Begin VB.TextBox Text9 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   3
               Left            =   10560
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   2300
               Width           =   1935
            End
            Begin VB.TextBox Text11 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   3
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   80
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   2300
               Width           =   7335
            End
            Begin VB.TextBox Text12 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   3
               Left            =   10560
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   1980
               Width           =   1935
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   320
               Index           =   1
               Left            =   120
               TabIndex        =   120
               Top             =   420
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   564
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Print Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0CB4
               Picture         =   "AccountMaster.frx":0CD0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   1
               Left            =   120
               TabIndex        =   121
               Top             =   105
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0CEC
               Picture         =   "AccountMaster.frx":0D08
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   1270
               Index           =   1
               Left            =   120
               TabIndex        =   122
               Top             =   720
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   2240
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Address"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0D24
               Picture         =   "AccountMaster.frx":0D40
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Index           =   1
               Left            =   120
               TabIndex        =   123
               Top             =   1980
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Phone"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0D5C
               Picture         =   "AccountMaster.frx":0D78
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Index           =   1
               Left            =   120
               TabIndex        =   124
               Top             =   2300
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " E-mail"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0D94
               Picture         =   "AccountMaster.frx":0DB0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Index           =   1
               Left            =   9000
               TabIndex        =   125
               Top             =   2295
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " GST No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0DCC
               Picture         =   "AccountMaster.frx":0DE8
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   1
               Left            =   9000
               TabIndex        =   126
               Top             =   1980
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Mobile"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0E04
               Picture         =   "AccountMaster.frx":0E20
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   12
               Left            =   9000
               TabIndex        =   130
               Top             =   420
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Alias"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0E3C
               Picture         =   "AccountMaster.frx":0E58
            End
            Begin VB.TextBox Text13 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
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
               Index           =   3
               Left            =   10560
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   410
               Width           =   1935
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   14
               Left            =   120
               TabIndex        =   132
               Top             =   3135
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Negative"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0E74
               Picture         =   "AccountMaster.frx":0E90
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Index           =   2
               Left            =   120
               TabIndex        =   133
               Top             =   3450
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Positive"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0EAC
               Picture         =   "AccountMaster.frx":0EC8
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
               Height          =   330
               Left            =   120
               TabIndex        =   134
               Top             =   2820
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " "
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0EE4
               Picture         =   "AccountMaster.frx":0F00
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
               Height          =   330
               Left            =   1680
               TabIndex        =   26
               Top             =   3135
               Width           =   3620
               _Version        =   65536
               _ExtentX        =   6385
               _ExtentY        =   582
               Calculator      =   "AccountMaster.frx":0F1C
               Caption         =   "AccountMaster.frx":0F3C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "AccountMaster.frx":0FA8
               Keys            =   "AccountMaster.frx":0FC6
               Spin            =   "AccountMaster.frx":1010
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "##0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "##0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999.99
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   330366981
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput2 
               Height          =   330
               Left            =   5280
               TabIndex        =   27
               Top             =   3135
               Width           =   3620
               _Version        =   65536
               _ExtentX        =   6385
               _ExtentY        =   582
               Calculator      =   "AccountMaster.frx":1038
               Caption         =   "AccountMaster.frx":1058
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "AccountMaster.frx":10C4
               Keys            =   "AccountMaster.frx":10E2
               Spin            =   "AccountMaster.frx":112C
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "##0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "##0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999.99
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1638405
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
               Height          =   330
               Left            =   8875
               TabIndex        =   28
               Top             =   3135
               Width           =   3620
               _Version        =   65536
               _ExtentX        =   6385
               _ExtentY        =   582
               Calculator      =   "AccountMaster.frx":1154
               Caption         =   "AccountMaster.frx":1174
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "AccountMaster.frx":11E0
               Keys            =   "AccountMaster.frx":11FE
               Spin            =   "AccountMaster.frx":1248
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "##0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "##0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999.99
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1638405
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
               Height          =   330
               Left            =   1680
               TabIndex        =   135
               Top             =   2820
               Width           =   3620
               _Version        =   65536
               _ExtentX        =   6385
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   "Cut Piece Rate"
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":1270
               Picture         =   "AccountMaster.frx":128C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
               Height          =   330
               Left            =   5280
               TabIndex        =   136
               Top             =   2820
               Width           =   3620
               _Version        =   65536
               _ExtentX        =   6385
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   "One Piece Rate"
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":12A8
               Picture         =   "AccountMaster.frx":12C4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   8875
               TabIndex        =   137
               Top             =   2820
               Width           =   3620
               _Version        =   65536
               _ExtentX        =   6385
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   "Pasting Rate"
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":12E0
               Picture         =   "AccountMaster.frx":12FC
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput6 
               Height          =   330
               Left            =   8875
               TabIndex        =   31
               Top             =   3450
               Width           =   3620
               _Version        =   65536
               _ExtentX        =   6385
               _ExtentY        =   582
               Calculator      =   "AccountMaster.frx":1318
               Caption         =   "AccountMaster.frx":1338
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "AccountMaster.frx":13A4
               Keys            =   "AccountMaster.frx":13C2
               Spin            =   "AccountMaster.frx":140C
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "##0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "##0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999.99
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1638405
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
               Height          =   330
               Left            =   1680
               TabIndex        =   29
               Top             =   3450
               Width           =   3620
               _Version        =   65536
               _ExtentX        =   6385
               _ExtentY        =   582
               Calculator      =   "AccountMaster.frx":1434
               Caption         =   "AccountMaster.frx":1454
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "AccountMaster.frx":14C0
               Keys            =   "AccountMaster.frx":14DE
               Spin            =   "AccountMaster.frx":1528
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "##0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "##0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999.99
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1638405
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
               Height          =   330
               Left            =   5280
               TabIndex        =   30
               Top             =   3450
               Width           =   3620
               _Version        =   65536
               _ExtentX        =   6385
               _ExtentY        =   582
               Calculator      =   "AccountMaster.frx":1550
               Caption         =   "AccountMaster.frx":1570
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "AccountMaster.frx":15DC
               Keys            =   "AccountMaster.frx":15FA
               Spin            =   "AccountMaster.frx":1644
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "##0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "##0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999.99
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1638405
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin VB.Line Line2 
               X1              =   0
               X2              =   12600
               Y1              =   3870
               Y2              =   3870
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   12600
               Y1              =   2720
               Y2              =   2720
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   6255
            Index           =   8
            Left            =   -74880
            TabIndex        =   138
            TabStop         =   0   'False
            Top             =   720
            Width           =   12495
            _Version        =   65536
            _ExtentX        =   22040
            _ExtentY        =   11033
            _StockProps     =   77
            Enabled         =   0   'False
            TintColor       =   16711935
            Alignment       =   0
            AutoSize        =   0   'False
            BevelSize       =   0
            BevelStyle      =   0
            BorderColor     =   -2147483642
            BorderStyle     =   1
            FillColor       =   -2147483633
            FontStyle       =   0
            FontTransparent =   0   'False
            LightColor      =   -2147483643
            ShadowColor     =   -2147483632
            TextColor       =   -2147483640
            WallPaper       =   0
            NoPrefix        =   0   'False
            FormatString    =   ""
            Caption         =   ""
            Picture         =   "AccountMaster.frx":166C
            Begin FPSpreadADO.fpSpread fpSpread1 
               Height          =   6045
               Left            =   120
               TabIndex        =   139
               Top             =   105
               Width           =   12255
               _Version        =   524288
               _ExtentX        =   21616
               _ExtentY        =   10663
               _StockProps     =   64
               EditEnterAction =   5
               EditModeReplace =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               GridColor       =   4227327
               MaxCols         =   5
               MaxRows         =   100
               OperationMode   =   2
               SpreadDesigner  =   "AccountMaster.frx":1688
            End
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   2
            Left            =   8400
            TabIndex        =   152
            Top             =   6630
            Width           =   4335
            _Version        =   65536
            _ExtentX        =   7646
            _ExtentY        =   582
            _StockProps     =   77
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TintColor       =   16711935
            Caption         =   " Ctrl+A->Add  Ctrl+E->Edit  F8->Delete  Ctrl+S->Save"
            Alignment       =   0
            FillColor       =   8421504
            TextColor       =   16777215
            Picture         =   "AccountMaster.frx":1D40
            Picture         =   "AccountMaster.frx":1D5C
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   3
            Left            =   8760
            TabIndex        =   153
            Top             =   0
            Width           =   3975
            _Version        =   65536
            _ExtentX        =   7011
            _ExtentY        =   582
            _StockProps     =   77
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TintColor       =   16711935
            Caption         =   " F5-> Refresh-> F12-> Create Duplicate Account"
            Alignment       =   0
            FillColor       =   8421504
            TextColor       =   16777215
            Picture         =   "AccountMaster.frx":1D78
            Picture         =   "AccountMaster.frx":1D94
         End
         Begin VB.TextBox txtNotes 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataSource      =   "Adodc1"
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
            Left            =   -73440
            MultiLine       =   -1  'True
            TabIndex        =   154
            ToolTipText     =   "Open Notes"
            Top             =   6600
            Visible         =   0   'False
            Width           =   5175
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H008BD6FE&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Find"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   330
            Left            =   120
            TabIndex        =   82
            Top             =   6630
            Width           =   735
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   77
      Top             =   0
      Width           =   13110
      _ExtentX        =   23125
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Add"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Edit"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Save"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancel"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Filter"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Print"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Print Preview"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Mail"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "First"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Previous"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Next"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Last"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   4
      Left            =   2760
      Top             =   2280
   End
End
Attribute VB_Name = "FrmAccountMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SL As Boolean 'Selection List
Public MasterCode As String  'Master to Modify
Public AccountType As String, AccountGroup As String
Public StateCode As String
Dim cnAccountMaster As New ADODB.Connection
Dim rstAccountList As New ADODB.Recordset
Dim rstAccountMaster As New ADODB.Recordset
Dim rstSizeList As New ADODB.Recordset
Dim rstSizeGroupList As New ADODB.Recordset
Dim rstBindingTypeList As New ADODB.Recordset
Dim rstOutsourceItemList As New ADODB.Recordset
Dim rstPaperList As New ADODB.Recordset
Dim rstFreshBookList As New ADODB.Recordset
Dim rstRepairBookList As New ADODB.Recordset
Dim rstStateList As New ADODB.Recordset
Dim rstAccountChild As New ADODB.Recordset
Dim rstCheckRef As New ADODB.Recordset
Dim rstAccountGroupList As New ADODB.Recordset
Dim PrevStr As String
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Dim SizeCode As String
Dim SizeGroupCode As String
Dim BindingTypeCode As String
Dim LaminationTypeCode As String
Dim OutsourceItem As String
Dim Paper As String
Dim FreshBook As String
Dim RepairBook As String
Dim Title As String
Dim AccountGroupCode As String
Dim SortOrder As String
Dim EditMode As Boolean
Private Sub btnNotes_Click()
frmNotes.NotesFlag = 1
frmNotes.Label1.Caption = "Notes : " & Text2(Val(AccountType) - 1).Text
frmNotes.Show (vbModal)
End Sub
Private Sub Form_Load()
    If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
    If Not SL Then MasterCode = ""
    Dim Cnt As Integer
    On Error GoTo ErrorHandler
    If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
    CenterForm Me
    BusySystemIndicator True
    For Cnt = 1 To 8
        If Cnt <> Val(AccountType) Then SSTab1.TabVisible(Cnt) = False
    Next
    AccountGroup = IIf(CheckEmpty(AccountGroup, False), "[Group]<>'*99999'", "[Group]='" & AccountGroup & "'")
    If AccountType <> "08" Then SSTab1.TabVisible(9) = False
    cnAccountMaster.CursorLocation = adUseClient
    cnAccountMaster.Open cnDatabase.ConnectionString
    rstAccountList.Open "SELECT Name,Alias,Code,State FROM AccountMaster WHERE " & AccountGroup & " ORDER BY Name", cnAccountMaster, adOpenKeyset, adLockOptimistic
    rstSizeList.Open "Select Name As Col0, Code From GeneralMaster Where Type = '1' Order By Name", cnAccountMaster, adOpenKeyset, adLockReadOnly
    rstSizeGroupList.Open "Select Name As Col0, Code From GeneralMaster Where Type = '10' Order By Name", cnAccountMaster, adOpenKeyset, adLockReadOnly
    rstBindingTypeList.Open "Select Name As Col0, Code From GeneralMaster Where Type = '6' Order By Name", cnAccountMaster, adOpenKeyset, adLockReadOnly
    If AccountType = "08" Then
        rstOutsourceItemList.Open "Select Name,'1'+Code As NCode From OutsourceItemMaster Order By Name", cnAccountMaster, adOpenKeyset, adLockOptimistic
        rstPaperList.Open "Select LTRIM(M.Name)+' (UOM : '+LTRIM(C.Name)+')' As Name,'2'+M.Code As NCode From PaperMaster M INNER JOIN GeneralMaster C ON M.UOM=C.Code ORDER BY M.Name", cnAccountMaster, adOpenKeyset, adLockOptimistic
        rstFreshBookList.Open "Select Name,Board,'3'+Code As NCode From BookMaster Where Type='F' Order By Name", cnAccountMaster, adOpenKeyset, adLockOptimistic
        rstRepairBookList.Open "Select Name,'4'+Code As NCode From BookMaster Where Type='R' Order By Name", cnAccountMaster, adOpenKeyset, adLockOptimistic
    End If
    rstAccountGroupList.Open "Select Name As Col0, Code From GeneralMaster Where Type = '12' Or Type = '26' Order By Name", cnAccountMaster, adOpenKeyset, adLockReadOnly
    rstStateList.Open "Select Name As Col0, Code From GeneralMaster Where Type = '56'  Order By Name", cnAccountMaster, adOpenKeyset, adLockReadOnly
    rstAccountMaster.CursorLocation = adUseClient
    rstAccountList.Filter = adFilterNone
    If rstAccountList.RecordCount Then
        If CheckEmpty(MasterCode, False) Then
            rstAccountList.MoveFirst
        Else
            rstAccountList.MoveFirst
            rstAccountList.Find "[Code]='" & MasterCode & "'"
        End If
    End If
    Set DataGrid1.DataSource = rstAccountList
    BusySystemIndicator False
    SSTab1.Tab = 0
    SortOrder = "Name"
    If Not (rstAccountList.EOF Or rstAccountList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstAccountList.ActiveConnection = Nothing
    rstSizeList.ActiveConnection = Nothing
    rstSizeGroupList.ActiveConnection = Nothing
    rstOutsourceItemList.ActiveConnection = Nothing
    rstPaperList.ActiveConnection = Nothing
    rstFreshBookList.ActiveConnection = Nothing
    rstRepairBookList.ActiveConnection = Nothing
    rstBindingTypeList.ActiveConnection = Nothing
    rstAccountGroupList.ActiveConnection = Nothing
    rstStateList.ActiveConnection = Nothing
    If AccountType = "08" Then
        Call RefreshDropDownList("A")
        fpSpread1.Col = 4
        fpSpread1.ColHidden = True
        fpSpread1.Col = 5
        fpSpread1.ColHidden = True
    End If
    SetButtonsForNoRecord
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Unload Me
End Sub
Private Sub Form_Activate()
    SetMenuOptions False
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        If SSTab1.Tab = 0 Then
            Unload Me
        Else
            If Toolbar1.Buttons.Item(1).Enabled Then
                SSTab1.Tab = 0
            Else
                If Not EditMode Then
                    If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then
                        Me.ActiveControl.SetFocus
                    Else
                        Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
                    End If
                End If
            End If
        End If
        If Not EditMode Then KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA And Toolbar1.Buttons.Item(1).Enabled Then
        If AccountType = "01" Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(1)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE And Toolbar1.Buttons.Item(2).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF8 And Toolbar1.Buttons.Item(3).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(3)
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF12 And Toolbar1.Buttons.Item(1).Enabled Then
        If MsgBox("Are you sure to make a duplicate copy of the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then DuplicateRecord
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS And Toolbar1.Buttons.Item(4).Enabled Then
        If Not EditMode Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF5 And Toolbar1.Buttons.Item(6).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(13)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyP And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(14)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyN And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(15)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyL And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(16)
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyReturn Then
        If Toolbar1.Buttons.Item(1).Enabled Then
            If SL Then
                If SSTab1.Tab = 0 Then Me.Tag = "S": slCode = rstAccountList.Fields("Code").Value: slName = rstAccountList.Fields("Name").Value: slStateCode = rstAccountList.Fields("State").Value: KeyCode = 0: Unload Me: Exit Sub
            Else
                SSTab1.Tab = Val(AccountType): SSTab1.SetFocus
            End If
        Else
            If Me.ActiveControl.Name <> "fpSpread1" Then Sendkeys "{TAB}"
        End If
        If Me.ActiveControl.Name <> "fpSpread1" Then KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then
        Call Form_KeyDown(vbKeyEscape, 0): Cancel = 1
    Else
        If Me.Tag <> "S" Then slCode = "": slName = ""
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstAccountMaster)
    Call CloseRecordset(rstSizeList)
    Call CloseRecordset(rstSizeGroupList)
    Call CloseRecordset(rstBindingTypeList)
    Call CloseRecordset(rstAccountChild)
    Call CloseConnection(cnAccountMaster)
    Call CloseRecordset(rstCheckRef)
    Call CloseRecordset(rstOutsourceItemList)
    Call CloseRecordset(rstPaperList)
    Call CloseRecordset(rstFreshBookList)
    Call CloseRecordset(rstRepairBookList)
    Call CloseRecordset(rstAccountGroupList)
    Call CloseRecordset(rstStateList)
    ShowProgressInStatusBar False
    SetMenuOptions True
    AccountGroup = ""
End Sub
Private Sub Text1_Change()
    If rstAccountList.RecordCount = 0 Then Exit Sub
    rstAccountList.MoveFirst
    If Len(Text1.Text) > 0 Then
        rstAccountList.Filter = "[" & SortOrder & "] Like '%" & FixQuote(Text1.Text) & "%'"
        If rstAccountList.EOF Then  'if Spelling mistake
            rstAccountList.Filter = adFilterNone
            rstAccountList.MoveFirst
            Beep
            DisplayError ("Spelling Error")
            Text1.Text = PrevStr
            Sendkeys "{End}"
        Else    'if Spelling alright
            PrevStr = Text1.Text
        End If
    Else
        rstAccountList.Filter = adFilterNone
        rstAccountList.MoveFirst
        Set DataGrid1.DataSource = rstAccountList
        PrevStr = ""
    End If
    If Not (rstAccountList.EOF Or rstAccountList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    
    If rstAccountList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstAccountList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstAccountList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstAccountList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstAccountList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstAccountList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstAccountList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstAccountList
            .MoveLast
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    End If
    If KeyProcessed Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
        KeyProcessed = False
        KeyCode = 0
    End If
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
    On Error Resume Next
    If Toolbar1.Buttons.Item(1).Enabled Then
        If SSTab1.Tab = Val(AccountType) Then
            ViewRecord
        Else
            If Not (rstAccountList.EOF Or rstAccountList.BOF) Then
                With DataGrid1.SelBookmarks
                    If .Count <> 0 Then .Remove 0
                    .Add DataGrid1.Bookmark
                End With
            End If
            Text1.SetFocus
        End If
        SSTab1.TabEnabled(0) = True
    Else
        SSTab1.TabEnabled(0) = False
        If AccountType <> "08" Then
            Text2(Val(AccountType) - 1).SetFocus
        Else
            If SSTab1.Tab = 8 Then  'Binding Rate
                Mh3dFrame2(Val(AccountType) - 1).Enabled = True
                Mh3dFrame2(Val(AccountType)).Enabled = False
                Text2(Val(AccountType) - 1).SetFocus
            Else
                Mh3dFrame2(Val(AccountType) - 1).Enabled = False
                Mh3dFrame2(Val(AccountType)).Enabled = True
                fpSpread1.SetFocus
            End If
        End If
    End If
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim HiLiteRecord As Boolean
    Dim UpdateFlag As Integer, i As Integer
    Dim CellVal As Variant, Imported As Variant
    If Button.Index = 1 Then
        If rstAccountMaster.State = adStateOpen Then rstAccountMaster.Close
        rstAccountMaster.Open "Select * From AccountMaster Where Code = ''", cnAccountMaster, adOpenKeyset, adLockOptimistic
        ClearFields
        Call LoadRateList("")
        If InStr(1, "01", AccountType) = 0 Then
            If rstAccountChild.State = adStateClosed Then SSTab1.Tab = 0: Exit Sub
        End If
        If AddRecord(rstAccountMaster) Then
            Call SetButtons(False)
            SSTab1.Tab = Val(AccountType)
            Text2(Val(AccountType) - 1).SetFocus
            blnRecordExist = False
            cnAccountMaster.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstAccountList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = Val(AccountType)
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstAccountList.RecordCount = 0 Then Exit Sub
        If AllowMastersDeletion = 0 Then
            Call DisplayError("You don't have the rights to Delete this Master")
            Exit Sub
        End If
        SSTab1.Tab = Val(AccountType)
        If CheckRef Then
            DisplayError ("Failed to delete the record")
        ElseIf MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            UpdateFlag = 0
            cnAccountMaster.BeginTrans
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            blnRecordExist = True
            If UpdateRateList("D") Then
                cnAccountMaster.Execute "DELETE FROM AccountMaster WHERE Code='" & rstAccountList.Fields("Code").Value & "'"
                If Err.Number = 0 Then
                    cnAccountMaster.CommitTrans
                    rstAccountList.Delete
                    rstAccountList.MoveNext
                    If rstAccountList.RecordCount > 0 And rstAccountList.EOF Then rstAccountList.MoveLast
'                    Call UpdateUserAction(Choose(Val(AccountType), "Account", "Artist", "Composer", "Processor", "Book Printer", "Title Printer", "Laminator", "Binder") & " Master", "D", Trim(Text2(Val(AccountType) - 1).Text), cnAccountMaster)
                    ShowProgressInStatusBar True
                    Timer1.Enabled = True
                    UpdateFlag = 1
                End If
            End If
            If UpdateFlag = 0 Then
                DisplayError ("Failed to delete the record")
                cnAccountMaster.RollbackTrans
            End If
            MdiMainMenu.MousePointer = vbNormal
            On Error GoTo 0
        End If
        SetButtons (True)
        SetButtonsForNoRecord
        SSTab1.Tab = 0
        HiLiteRecord = True
    ElseIf Button.Index = 4 Then
        If CheckMandatoryFields Then Exit Sub
        If blnRecordExist And AllowMastersModification = 0 Then
            Call DisplayError("You don't have the rights to Edit this Master")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
            Exit Sub
        End If
        SaveFields
        UpdateFlag = 0
        If UpdateRecord(rstAccountMaster) Then
            UpdateFlag = 1
            If UpdateRateList("D") Then
                If InStr(1, "01_09", AccountType) = 0 Then
                    If rstAccountChild.RecordCount <> 0 Then
                        rstAccountChild.MoveFirst
                        Do While Not rstAccountChild.EOF
                            If Not UpdateRateList("I") Then
                                UpdateFlag = 0
                                Exit Do
                            End If
                            rstAccountChild.MoveNext
                        Loop
                    End If
                End If
            End If
            If AccountType = "08" Then
                If UpdateFlag Then
                    If UpdateMaterialList("D") Then
                        For i = 1 To fpSpread1.DataRowCnt
                            fpSpread1.SetActiveCell 3, i
                            fpSpread1.GetText 3, i, CellVal
                            fpSpread1.GetText 5, i, Imported
                            If Val(CellVal) <> 0 And (Imported = "N" Or Imported = "") Then
                                If Not UpdateMaterialList("I") Then
                                    UpdateFlag = 0
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                End If
            End If
        End If
        If UpdateFlag Then
'            Call UpdateUserAction(Choose(Val(AccountType), "Account", "Artist", "Composer", "Processor", "Book Printer", "Title Printer", "Laminator", "Binder") & " Master", IIf(blnRecordExist, "M", "A"), Trim(Text2(Val(AccountType) - 1).Text), cnAccountMaster)
            AddToList
            cnAccountMaster.CommitTrans
            If rstAccountMaster.State = adStateOpen Then rstAccountMaster.Close
            rstAccountMaster.CursorLocation = adUseClient
            Call SetButtons(True)
            SSTab1.Tab = 0
            ShowProgressInStatusBar True
            Timer1.Enabled = True
            Call MsgBox("Record updated !!!", vbInformation, App.Title)
        Else
            DisplayError ("Failed to save the record")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
        End If
    ElseIf Button.Index = 5 Then
        If CancelRecordUpdate(rstAccountMaster) Then
            cnAccountMaster.RollbackTrans
            If rstAccountMaster.State = adStateOpen Then rstAccountMaster.Close
            rstAccountMaster.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        RefreshData rstAccountList
        Set DataGrid1.DataSource = rstAccountList
        RefreshData rstSizeList
        RefreshData rstSizeGroupList
        RefreshData rstBindingTypeList
        RefreshData rstAccountGroupList
        RefreshData rstStateList
        HiLiteRecord = True
    ElseIf Button.Index = 7 Then
        SSTab1.Tab = 0
        With FrmFilter
            .Combo1.AddItem "Name", 0
            .Combo1.ListIndex = 0
            Set .srcForm = Me
            .Show vbModal
        End With
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then
        If rstAccountList.RecordCount > 0 Then rstAccountList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstAccountList.RecordCount > 0 Then
           rstAccountList.MovePrevious
           If rstAccountList.BOF Then
              rstAccountList.MoveNext
           End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstAccountList.RecordCount > 0 Then
           rstAccountList.MoveNext
           If rstAccountList.EOF Then
              rstAccountList.MovePrevious
           End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstAccountList.RecordCount > 0 Then rstAccountList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstAccountList.EOF Or rstAccountList.BOF) Then
            With DataGrid1.SelBookmarks
                If .Count <> 0 Then .Remove 0
                .Add DataGrid1.Bookmark
            End With
        End If
        Text1.SetFocus
    End If
End Sub
Private Sub DataGrid1_DblClick()
    If Toolbar1.Buttons.Item(2).Enabled Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    Static AD As String
    SortOrder = DataGrid1.Columns(ColIndex).DataField
    If AD = "Asc" Then
        rstAccountList.Sort = "[" + SortOrder & "] Desc"
        AD = "Desc"
    Else
        rstAccountList.Sort = "[" + SortOrder & "] Asc"
        AD = "Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstAccountList.EOF Or rstAccountList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    Text1.Text = ""
    Text1.SetFocus
End Sub
Private Sub SetButtons(bVal As Boolean)
    Toolbar1.Buttons.Item(1).Enabled = bVal
    Toolbar1.Buttons.Item(2).Enabled = bVal
    Toolbar1.Buttons.Item(3).Enabled = bVal
    Toolbar1.Buttons.Item(4).Enabled = Not bVal
    Toolbar1.Buttons.Item(5).Enabled = Not bVal
    Toolbar1.Buttons.Item(6).Enabled = bVal
    Toolbar1.Buttons.Item(7).Enabled = bVal
    Toolbar1.Buttons.Item(13).Enabled = bVal
    Toolbar1.Buttons.Item(14).Enabled = bVal
    Toolbar1.Buttons.Item(15).Enabled = bVal
    Toolbar1.Buttons.Item(16).Enabled = bVal
    Toolbar1.Buttons.Item(18).Enabled = bVal
    Mh3dFrame2(Val(AccountType) - 1).Enabled = Not bVal
    Mh3dFrame2(8).Enabled = False
End Sub
Private Sub SetButtonsForNoRecord()
    If rstAccountList.RecordCount = 0 Then
        Toolbar1.Buttons.Item(2).Enabled = False
        Toolbar1.Buttons.Item(3).Enabled = False
        Toolbar1.Buttons.Item(13).Enabled = False
        Toolbar1.Buttons.Item(14).Enabled = False
        Toolbar1.Buttons.Item(15).Enabled = False
        Toolbar1.Buttons.Item(16).Enabled = False
    End If
End Sub
Private Sub Text2_Validate(Index As Integer, Cancel As Boolean)
    If rstAccountMaster.EOF Or rstAccountMaster.BOF Then Exit Sub
    If CheckEmpty(Text2(Val(AccountType) - 1), True) Then
        Cancel = True
    ElseIf CheckDuplicate(cnAccountMaster, "AccountMaster", "Code", "Name", Trim(Text2(Val(AccountType) - 1).Text), rstAccountMaster.Fields("Code").Value, False) Then
        Cancel = True
    ElseIf CheckEmpty(Text3(Val(AccountType) - 1), False) Then
        Text3(Val(AccountType) - 1).Text = Text2(Val(AccountType) - 1).Text
    End If
End Sub
Private Sub Text14_Change(Index As Integer)
    If Text14(Val(AccountType) - 1).Text = " " Then Text14(Val(AccountType) - 1).Text = "?": Sendkeys "{TAB}"
End Sub
Private Sub Text14_Validate(Index As Integer, Cancel As Boolean)
    Dim SearchString As String
    SearchString = FixQuote(Text14(Val(AccountType) - 1).Text)
    If rstStateList.RecordCount = 0 Then
       DisplayError ("No Record in State Master")
       Cancel = True
       Exit Sub
    Else
       rstStateList.MoveFirst
    End If
    rstStateList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstStateList.EOF Then
       SelectionType = "S"
       StateCode = ""
       Call LoadSelectionList(rstStateList, "List of States...", "Name")
       SearchOrder = 0
       Call DisplaySelectionList(Text14(Val(AccountType) - 1), StateCode)
       Call CloseForm(FrmSelectionList)
       If CheckEmpty(Text14(Val(AccountType) - 1).Text, False) Then Text14(Val(AccountType) - 1).Text = "?"
       If RTrim(StateCode) <> "" Then Sendkeys "{TAB}"
       Cancel = True
    Else
       StateCode = rstStateList.Fields("Code").Value
    End If
End Sub
Private Sub Text10_Change(Index As Integer)
    If Text10(Val(AccountType) - 1).Text = " " Then Text10(Val(AccountType) - 1).Text = "?": Sendkeys "{TAB}"
End Sub
Private Sub Text10_Validate(Index As Integer, Cancel As Boolean)
    Dim SearchString As String
    SearchString = FixQuote(Text10(Val(AccountType) - 1).Text)
    If rstAccountGroupList.RecordCount = 0 Then
       DisplayError ("No Record in Account Group Master")
       Cancel = True
       Exit Sub
    Else
       rstAccountGroupList.MoveFirst
    End If
    rstAccountGroupList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstAccountGroupList.EOF Then
       SelectionType = "S"
       AccountGroupCode = ""
       Call LoadSelectionList(rstAccountGroupList, "List of Account Groups...", "Name")
       SearchOrder = 0
       Call DisplaySelectionList(Text10(Val(AccountType) - 1), AccountGroupCode)
       Call CloseForm(FrmSelectionList)
       If CheckEmpty(Text10(Val(AccountType) - 1).Text, False) Then Text10(Val(AccountType) - 1).Text = "?"
       If RTrim(AccountGroupCode) <> "" Then Sendkeys "{TAB}"
       Cancel = True
    Else
       AccountGroupCode = rstAccountGroupList.Fields("Code").Value
    End If
End Sub
Private Sub Text9_Validate(Index As Integer, Cancel As Boolean)
    If InStr(1, "01_04_07_08", AccountType) <> 0 Then Exit Sub
    If rstAccountChild.RecordCount = 0 Then
        Call AddRecord(rstAccountChild)
        Call DataGrid2_KeyDown(Index, vbKeyE, vbCtrlMask)
    End If
End Sub
Private Sub chkRound_Validate(Cancel As Boolean)
    If AccountType = "05" Then
        If rstAccountChild.RecordCount = 0 Then
            Call AddRecord(rstAccountChild)
            Call DataGrid2_KeyDown(5, vbKeyE, vbCtrlMask)
        End If
    End If
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstAccountList.EOF Then
        If rstAccountChild.State = adStateOpen Then rstAccountChild.Close
        Exit Sub
    End If
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstAccountMaster.State = adStateOpen Then rstAccountMaster.Close
    rstAccountMaster.Open "Select *,(Select Name From GeneralMaster Where Type=56 And Code=State) As StateName From AccountMaster Where Code = '" & FixQuote(rstAccountList.Fields("Code").Value) & "'", cnAccountMaster, adOpenKeyset, adLockOptimistic
    If rstAccountMaster.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields()
    Text2(Val(AccountType) - 1).Text = ""
    Text3(Val(AccountType) - 1).Text = ""
    Text4(Val(AccountType) - 1).Text = ""
    Text5(Val(AccountType) - 1).Text = ""
    Text6(Val(AccountType) - 1).Text = ""
    Text7(Val(AccountType) - 1).Text = ""
    Text8(Val(AccountType) - 1).Text = ""
    Text9(Val(AccountType) - 1).Text = ""
    If AccountType = "01" Then Text10(Val(AccountType) - 1).Text = ""
    If AccountType = "01" Then Text14(Val(AccountType) - 1).Text = ""
    Text11(Val(AccountType) - 1).Text = ""
    Text12(Val(AccountType) - 1).Text = ""
    Text13(Val(AccountType) - 1).Text = ""
    chkRound.Value = 0
    MhRealInput1.Text = "0.00"
    MhRealInput2.Text = "0.00"
    MhRealInput3.Text = "0.00"
    MhRealInput4.Text = "0.00"
    MhRealInput5.Text = "0.00"
    MhRealInput6.Text = "0.00"
    txtNotes = ""
    TDBNumber1.Value = "0.00"
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
End Sub
Private Sub LoadFields()
    If rstAccountMaster.EOF Or rstAccountMaster.BOF Then Exit Sub
    Text2(Val(AccountType) - 1).Text = rstAccountMaster.Fields("Name").Value
    Text3(Val(AccountType) - 1).Text = rstAccountMaster.Fields("PrintName").Value
    Text4(Val(AccountType) - 1).Text = rstAccountMaster.Fields("Address1").Value
    Text5(Val(AccountType) - 1).Text = rstAccountMaster.Fields("Address2").Value
    Text6(Val(AccountType) - 1).Text = rstAccountMaster.Fields("Address3").Value
    Text7(Val(AccountType) - 1).Text = rstAccountMaster.Fields("Address4").Value
    Text8(Val(AccountType) - 1).Text = rstAccountMaster.Fields("Phone").Value
    Text12(Val(AccountType) - 1).Text = rstAccountMaster.Fields("Mobile").Value
    Text13(Val(AccountType) - 1).Text = rstAccountMaster.Fields("Alias").Value
    Text9(Val(AccountType) - 1).Text = rstAccountMaster.Fields("TIN").Value
    Text11(Val(AccountType) - 1).Text = rstAccountMaster.Fields("EMail").Value
    AccountGroupCode = rstAccountMaster.Fields("Group").Value
    txtNotes.Text = rstAccountMaster.Fields("Notes").Value
    TDBNumber1.Value = rstAccountMaster.Fields("Opening").Value
    
    If AccountType = "01" Then
        If rstAccountGroupList.RecordCount > 0 Then rstAccountGroupList.MoveFirst
        rstAccountGroupList.Find "[Code] = '" & AccountGroupCode & "'"
        If Not rstAccountGroupList.EOF Then Text10(Val(AccountType) - 1).Text = rstAccountGroupList.Fields("Col0").Value
    End If
    If AccountType = "05" Then chkRound.Value = IIf(rstAccountMaster.Fields("RoundOffQty").Value, 1, 0)
    If AccountType = "08" Then Call LoadMaterialList(rstAccountMaster.Fields("Code").Value)
    Call LoadRateList(rstAccountMaster.Fields("Code").Value)
    StateCode = rstAccountMaster.Fields("State").Value
    Text14(Val(AccountType) - 1).Text = rstAccountMaster.Fields("StateName").Value
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstAccountMaster.RecordCount = 0 Then Exit Sub
    If InStr(1, "01_09", AccountType) = 0 Then
        If rstAccountChild.State = adStateClosed Then SSTab1.Tab = 0: Exit Sub
    End If
    If rstAccountMaster.State = adStateOpen Then rstAccountMaster.Close
    rstAccountMaster.CursorLocation = adUseServer
    rstAccountMaster.Open "Select * From AccountMaster Where Code = '" & FixQuote(rstAccountList.Fields("Code").Value) & "'", cnAccountMaster, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstAccountMaster.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text2(Val(AccountType) - 1).SetFocus
    blnRecordExist = True
    cnAccountMaster.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    If rstAccountMaster.EOF Or rstAccountMaster.BOF Then Exit Sub
    If Not blnRecordExist Then
        rstAccountMaster.Fields("Code").Value = GenerateCode(cnAccountMaster, "SELECT MAX(Code) FROM AccountMaster", 6, "0")
        rstAccountMaster.Fields("CreatedBy").Value = UserCode
        rstAccountMaster.Fields("CreatedOn").Value = Now()
        rstAccountMaster.Fields("Recordstatus").Value = "N"
    Else
        rstAccountMaster.Fields("ModifiedBy").Value = UserCode
        rstAccountMaster.Fields("ModifiedOn").Value = Now()
        rstAccountMaster.Fields("Recordstatus").Value = "M"
    End If
    rstAccountMaster.Fields("Name").Value = Trim(Text2(Val(AccountType) - 1).Text)
    rstAccountMaster.Fields("PrintName").Value = Trim(Text3(Val(AccountType) - 1).Text)
    rstAccountMaster.Fields("Address1").Value = Trim(Text4(Val(AccountType) - 1).Text)
    rstAccountMaster.Fields("Address2").Value = Trim(Text5(Val(AccountType) - 1).Text)
    rstAccountMaster.Fields("Address3").Value = Trim(Text6(Val(AccountType) - 1).Text)
    rstAccountMaster.Fields("Address4").Value = Trim(Text7(Val(AccountType) - 1).Text)
    rstAccountMaster.Fields("Phone").Value = Trim(Text8(Val(AccountType) - 1).Text)
    rstAccountMaster.Fields("Mobile").Value = Trim(Text12(Val(AccountType) - 1).Text)
    rstAccountMaster.Fields("Alias").Value = Trim(Text13(Val(AccountType) - 1).Text)
    rstAccountMaster.Fields("Group").Value = AccountGroupCode
    rstAccountMaster.Fields("TIN").Value = Trim(Text9(Val(AccountType) - 1).Text)
    rstAccountMaster.Fields("EMail").Value = Trim(Text11(Val(AccountType) - 1).Text)
    rstAccountMaster.Fields("PrintStatus").Value = "N"
    rstAccountMaster.Fields("Notes").Value = txtNotes.Text
    rstAccountMaster.Fields("Opening").Value = TDBNumber1.Value
    rstAccountMaster.Fields("State").Value = StateCode
    If AccountType = "04" Then
        If rstAccountChild.RecordCount = 0 Then Call AddRecord(rstAccountChild)
        rstAccountChild.Fields("NegativeOnePcRate").Value = Format(Val(MhRealInput1.Text), "0.00")
        rstAccountChild.Fields("NegativeCutPcRate").Value = Format(Val(MhRealInput2.Text), "0.00")
        rstAccountChild.Fields("NegativePastingRate").Value = Format(Val(MhRealInput3.Text), "0.00")
        rstAccountChild.Fields("PositiveOnePcRate").Value = Format(Val(MhRealInput4.Text), "0.00")
        rstAccountChild.Fields("PositiveCutPcRate").Value = Format(Val(MhRealInput5.Text), "0.00")
        rstAccountChild.Fields("PositivePastingRate").Value = Format(Val(MhRealInput6.Text), "0.00")
    ElseIf AccountType = "05" Then
        rstAccountMaster.Fields("RoundOffQty").Value = chkRound.Value
    End If
End Sub
Private Sub AddToList()
    On Error Resume Next
    rstAccountList.MoveFirst
    rstAccountList.Find "[Code] = '" & rstAccountMaster.Fields("Code").Value & "'"
    If rstAccountList.EOF Then
       rstAccountList.AddNew
       rstAccountList.Fields("Code").Value = rstAccountMaster.Fields("Code").Value
    End If
    rstAccountList.Fields("Name").Value = rstAccountMaster.Fields("Name").Value
    rstAccountList.Update
    rstAccountList.Sort = SortOrder & " Asc"
    rstAccountList.Find "[Code] = '" & rstAccountMaster.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2(Val(AccountType) - 1).Text, False) Then
       SSTab1.Tab = Val(AccountType)
       Text2(Val(AccountType) - 1).SetFocus
       CheckMandatoryFields = True
    ElseIf CheckDuplicate(cnAccountMaster, "AccountMaster", "Code", "Name", Trim(Text2(Val(AccountType) - 1).Text), rstAccountMaster.Fields("Code").Value, False) Then
       SSTab1.Tab = Val(AccountType)
       Text2(Val(AccountType) - 1).SetFocus
       CheckMandatoryFields = True
       SSTab1.Tab = Val(AccountType)
    ElseIf CheckEmpty(Text3(Val(AccountType) - 1).Text, False) Then
       SSTab1.Tab = Val(AccountType)
       Text3(Val(AccountType) - 1).SetFocus
       CheckMandatoryFields = True
    ElseIf CheckItem() Then
       SSTab1.Tab = Val(AccountType + 2)
       fpSpread1.SetFocus
       CheckMandatoryFields = True
    End If
    If AccountType = "01" Then
        If CheckEmpty(Text10(Val(AccountType) - 1).Text, False) Then
            SSTab1.Tab = 1
            Text10(Val(AccountType) - 1).SetFocus
            CheckMandatoryFields = True
        ElseIf Not CheckExists(Text10(Val(AccountType) - 1), "Col0", rstAccountGroupList, AccountGroupCode) Then
            SSTab1.Tab = 1
            Text10(Val(AccountType) - 1).SetFocus
            CheckMandatoryFields = True
        End If
    End If
End Function
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Name" Then rstAccountList.Filter = "[Name] Like '%" & SrchText & "%'"
End Sub
Private Function CheckRef() As Boolean
    On Error GoTo ErrorHandler
    If rstCheckRef.State = adStateOpen Then rstCheckRef.Close
    rstCheckRef.Open "Select BookPrinter From BookPOParent Where BookPrinter = '" & rstAccountList.Fields("Code").Value & "'", cnAccountMaster, adOpenKeyset, adLockReadOnly
    If rstCheckRef.RecordCount > 0 Then CheckRef = True: Exit Function
    If rstCheckRef.State = adStateOpen Then rstCheckRef.Close
    rstCheckRef.Open "Select TitlePrinter From BookPOParent Where TitlePrinter = '" & rstAccountList.Fields("Code").Value & "'", cnAccountMaster, adOpenKeyset, adLockReadOnly
    If rstCheckRef.RecordCount > 0 Then CheckRef = True: Exit Function
    If rstCheckRef.State = adStateOpen Then rstCheckRef.Close
    rstCheckRef.Open "Select Laminator From BookPOParent Where Laminator = '" & rstAccountList.Fields("Code").Value & "'", cnAccountMaster, adOpenKeyset, adLockReadOnly
    If rstCheckRef.RecordCount > 0 Then CheckRef = True: Exit Function
    If rstCheckRef.State = adStateOpen Then rstCheckRef.Close
    rstCheckRef.Open "Select Binder From BookPOParent Where Binder = '" & rstAccountList.Fields("Code").Value & "'", cnAccountMaster, adOpenKeyset, adLockReadOnly
    If rstCheckRef.RecordCount > 0 Then CheckRef = True
    Exit Function
ErrorHandler:
    CheckRef = True
End Function
Private Sub Timer1_Timer()
    On Error Resume Next
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then
       Timer1.Enabled = False
       ShowProgressInStatusBar False
    End If
End Sub
Private Sub LoadRateList(ByVal strAccountCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    
    If rstAccountChild.State = adStateOpen Then rstAccountChild.Close
    If AccountType = "04" Then
        rstAccountChild.Open "SELECT * FROM AccountChild04 WHERE Code = '" & strAccountCode & "'", cnAccountMaster, adOpenKeyset, adLockOptimistic
        If Not (rstAccountChild.EOF Or rstAccountChild.BOF) Then
            MhRealInput1.Text = Format(Val(rstAccountChild.Fields("NegativeOnePcRate").Value), "0.00")
            MhRealInput2.Text = Format(Val(rstAccountChild.Fields("NegativeCutPcRate").Value), "0.00")
            MhRealInput3.Text = Format(Val(rstAccountChild.Fields("NegativePastingRate").Value), "0.00")
            MhRealInput4.Text = Format(Val(rstAccountChild.Fields("PositiveOnePcRate").Value), "0.00")
            MhRealInput5.Text = Format(Val(rstAccountChild.Fields("PositiveCutPcRate").Value), "0.00")
            MhRealInput6.Text = Format(Val(rstAccountChild.Fields("PositivePastingRate").Value), "0.00")
        End If
    ElseIf AccountType = "05" Then
        rstAccountChild.Open "Select M2.[Name] As SizeName, M1.* From AccountChild05 M1, GeneralMaster M2 Where M1.[Size] = M2.Code And M1.Code = '" & strAccountCode & "' Order By M2.Name", cnAccountMaster, adOpenKeyset, adLockOptimistic
    ElseIf AccountType = "07" Then
        rstAccountChild.Open "SELECT M1.Name As OperationName,M2.Name As CalcModeName,M3.Name As SizeName,P.* FROM ((AccountChild07 P INNER JOIN GeneralMaster M1 ON P.LaminationType=M1.Code) INNER JOIN GeneralMaster M2 ON P.CalcMode=M2.Code) LEFT JOIN GeneralMaster M3 ON P.[Size]=M3.Code WHERE P.Code='" & strAccountCode & "' ORDER BY M1.Name,M3.Name,Range,Rate", cnAccountMaster, adOpenKeyset, adLockOptimistic
    ElseIf AccountType = "08" Then
        rstAccountChild.Open "Select M2.[Name] As SizeName,M3.[Name] As BindingTypeName,M1.* From AccountChild08 M1,GeneralMaster M2,GeneralMaster M3 Where M1.[Size] = M2.Code And M1.BindingType = M3.Code And M3.Type = '6' And M1.Code = '" & strAccountCode & "' Order By M2.Name, M3.Name", cnAccountMaster, adOpenKeyset, adLockOptimistic
    End If
    rstAccountChild.ActiveConnection = Nothing
    If InStr(1, "01_04_09", AccountType) = 0 Then Set DataGrid2(Val(AccountType) - 1).DataSource = rstAccountChild
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Rate List")
End Sub
Private Function UpdateRateList(ByVal ActionType As String) As Boolean
    On Error GoTo ErrorHandler
    UpdateRateList = True
    If (ActionType = "D" And (Not blnRecordExist)) Or InStr(1, "01_09", AccountType) > 0 Then Exit Function
    If ActionType <> "I" Then
        cnAccountMaster.Execute "DELETE FROM AccountChild" & AccountType & " WHERE Code = '" & rstAccountMaster.Fields("Code").Value & "'"
    Else
        If AccountType = "04" Then
            cnAccountMaster.Execute "Insert Into AccountChild04 Values ('" & rstAccountMaster.Fields("Code").Value & "'," & Val(MhRealInput1.Text) & "," & Val(MhRealInput2.Text) & "," & Val(MhRealInput3.Text) & "," & Val(MhRealInput4.Text) & "," & Val(MhRealInput5.Text) & "," & Val(MhRealInput6.Text) & ")"
        ElseIf AccountType = "05" Then
            cnAccountMaster.Execute "Insert Into AccountChild05 Values ('" & rstAccountMaster.Fields("Code").Value & "','" & rstAccountChild.Fields("Size").Value & "'," & Val(rstAccountChild.Fields("Range1").Value) & "," & Val(rstAccountChild.Fields("Range2").Value) & "," & Val(rstAccountChild.Fields("Range4").Value) & "," & Val(rstAccountChild.Fields("Range6").Value) & "," & _
                                      Val(rstAccountChild.Fields("PrintRate1").Value) & "," & Val(rstAccountChild.Fields("PrintRate2").Value) & "," & Val(rstAccountChild.Fields("PrintRate4").Value) & "," & Val(rstAccountChild.Fields("PrintRate6").Value) & "," & Val(rstAccountChild.Fields("PSPlateRate1").Value) & "," & Val(rstAccountChild.Fields("PSPlateRate2").Value) & "," & Val(rstAccountChild.Fields("PSPlateRate4").Value) & "," & Val(rstAccountChild.Fields("PSPlateRate6").Value) & "," & Val(rstAccountChild.Fields("DeepatchPlateRate1").Value) & "," & Val(rstAccountChild.Fields("DeepatchPlateRate2").Value) & "," & Val(rstAccountChild.Fields("DeepatchPlateRate4").Value) & "," & Val(rstAccountChild.Fields("DeepatchPlateRate6").Value) & "," & _
                                      Val(rstAccountChild.Fields("WipeonPlateRate1").Value) & "," & Val(rstAccountChild.Fields("WipeonPlateRate2").Value) & "," & Val(rstAccountChild.Fields("WipeonPlateRate4").Value) & "," & Val(rstAccountChild.Fields("WipeonPlateRate6").Value) & "," & Val(rstAccountChild.Fields("CTPPlateRate1").Value) & "," & Val(rstAccountChild.Fields("CTPPlateRate2").Value) & "," & Val(rstAccountChild.Fields("CTPPlateRate4").Value) & "," & Val(rstAccountChild.Fields("CTPPlateRate6").Value) & "," & Val(rstAccountChild.Fields("PaperWastageRate1").Value) & "," & Val(rstAccountChild.Fields("PaperWastageRate2").Value) & "," & Val(rstAccountChild.Fields("PaperWastageRate4").Value) & "," & Val(rstAccountChild.Fields("PaperWastageRate6").Value) & "," & Val(rstAccountChild.Fields("PaperWastageMin1").Value) & "," & Val(rstAccountChild.Fields("PaperWastageMin2").Value) & "," & Val(rstAccountChild.Fields("PaperWastageMin4").Value) & "," & Val(rstAccountChild.Fields("PaperWastageMin6").Value) & ")"
        ElseIf AccountType = "07" Then
            cnAccountMaster.Execute "INSERT INTO AccountChild07 VALUES ('" & rstAccountMaster.Fields("Code").Value & "','" & rstAccountChild.Fields("Size").Value & "','" & rstAccountChild.Fields("LaminationType").Value & "','" & rstAccountChild.Fields("CalcMode").Value & "'," & Val(rstAccountChild.Fields("Rate").Value) & "," & Val(rstAccountChild.Fields("Range").Value) & ")"
        ElseIf AccountType = "08" Then
            cnAccountMaster.Execute "Insert Into AccountChild08 Values ('" & rstAccountMaster.Fields("Code").Value & "','" & rstAccountChild.Fields("BindingType").Value & "','" & rstAccountChild.Fields("Size").Value & "'," & Val(rstAccountChild.Fields("Range04").Value) & "," & Val(rstAccountChild.Fields("Range06").Value) & "," & Val(rstAccountChild.Fields("Range08").Value) & "," & Val(rstAccountChild.Fields("Range12").Value) & "," & Val(rstAccountChild.Fields("Range16").Value) & "," & Val(rstAccountChild.Fields("Range24").Value) & "," & Val(rstAccountChild.Fields("Range32").Value) & "," & Val(rstAccountChild.Fields("Range64").Value) & "," & _
                                      Val(rstAccountChild.Fields("FormStitchRate04").Value) & "," & Val(rstAccountChild.Fields("FormStitchRate06").Value) & "," & Val(rstAccountChild.Fields("FormStitchRate08").Value) & "," & Val(rstAccountChild.Fields("FormStitchRate12").Value) & "," & Val(rstAccountChild.Fields("FormStitchRate16").Value) & "," & Val(rstAccountChild.Fields("FormStitchRate24").Value) & "," & Val(rstAccountChild.Fields("FormStitchRate32").Value) & "," & Val(rstAccountChild.Fields("FormStitchRate64").Value) & "," & _
                                      Val(rstAccountChild.Fields("FormPasteRate04").Value) & "," & Val(rstAccountChild.Fields("FormPasteRate06").Value) & "," & Val(rstAccountChild.Fields("FormPasteRate08").Value) & "," & Val(rstAccountChild.Fields("FormPasteRate12").Value) & "," & Val(rstAccountChild.Fields("FormPasteRate16").Value) & "," & Val(rstAccountChild.Fields("FormPasteRate24").Value) & "," & Val(rstAccountChild.Fields("FormPasteRate32").Value) & "," & Val(rstAccountChild.Fields("FormPasteRate64").Value) & "," & _
                                      Val(rstAccountChild.Fields("FormFoldRate04").Value) & "," & Val(rstAccountChild.Fields("FormFoldRate06").Value) & "," & Val(rstAccountChild.Fields("FormFoldRate08").Value) & "," & Val(rstAccountChild.Fields("FormFoldRate12").Value) & "," & Val(rstAccountChild.Fields("FormFoldRate16").Value) & "," & Val(rstAccountChild.Fields("FormFoldRate24").Value) & "," & Val(rstAccountChild.Fields("FormFoldRate32").Value) & "," & Val(rstAccountChild.Fields("FormFoldRate64").Value) & "," & _
                                      Val(rstAccountChild.Fields("Rate/Book04").Value) & "," & Val(rstAccountChild.Fields("Rate/Book06").Value) & "," & Val(rstAccountChild.Fields("Rate/Book08").Value) & "," & Val(rstAccountChild.Fields("Rate/Book12").Value) & "," & Val(rstAccountChild.Fields("Rate/Book16").Value) & "," & Val(rstAccountChild.Fields("Rate/Book24").Value) & "," & Val(rstAccountChild.Fields("Rate/Book32").Value) & "," & Val(rstAccountChild.Fields("Rate/Book64").Value) & "," & Val(rstAccountChild.Fields("PktPackRate").Value) & "," & Val(rstAccountChild.Fields("BoxPackRate").Value) & "," & Val(rstAccountChild.Fields("Cartage/Box").Value) & ")"
        End If
    End If
    Exit Function
ErrorHandler:
    UpdateRateList = False
End Function
Private Sub DataGrid2_DblClick(Index As Integer)
    Call DataGrid2_KeyDown(Index, vbKeyE, vbCtrlMask)
End Sub
Private Sub DataGrid2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim DelRec As Boolean
    If Shift = vbCtrlMask And KeyCode = vbKeyE Then
        If rstAccountChild.RecordCount = 0 Then KeyCode = 0: Exit Sub
        If AccountType = "05" Then
            Set FrmAccountChild05.rstAccountChild = rstAccountChild
            Set FrmAccountChild05.rstSizeList = rstSizeGroupList
            FrmAccountChild05.AccountName = Trim(Text2(Val(AccountType) - 1).Text)
            On Error Resume Next
            Load FrmAccountChild05
            If Err.Number <> 364 Then FrmAccountChild05.Show vbModal
            On Error GoTo 0
        ElseIf AccountType = "07" Then
            Set FrmAccountChild07.rstAccountChild = rstAccountChild
            FrmAccountChild07.AccountName = Trim(Text2(Val(AccountType) - 1).Text)
            On Error Resume Next
            Load FrmAccountChild07
            If Err.Number <> 364 Then FrmAccountChild07.Show vbModal
            On Error GoTo 0
        ElseIf AccountType = "08" Then
            Set FrmAccountChild08.rstAccountChild = rstAccountChild
            Set FrmAccountChild08.rstSizeList = rstSizeGroupList
            Set FrmAccountChild08.rstBindingTypeList = rstBindingTypeList
            FrmAccountChild08.AccountName = Trim(Text2(Val(AccountType) - 1).Text)
            On Error Resume Next
            Load FrmAccountChild08
            If Err.Number <> 364 Then FrmAccountChild08.Show vbModal
            On Error GoTo 0
        End If
        KeyCode = 0
        If AccountType <> "07" Then
            If CheckEmpty(rstAccountChild.Fields("Size").Value, False) Then DelRec = True
        Else
            If CheckEmpty(rstAccountChild.Fields("LaminationType").Value, False) Then DelRec = True
        End If
        If DelRec Then
            rstAccountChild.Delete
            rstAccountChild.MoveNext
            If rstAccountChild.RecordCount > 0 Then rstAccountChild.MoveFirst
        ElseIf rstAccountChild.AbsolutePosition = rstAccountChild.RecordCount Then
            Call DataGrid2_KeyDown(Index, vbKeyA, vbCtrlMask)
        End If
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
        Sendkeys "^"
        Call AddRecord(rstAccountChild)
        Call DataGrid2_KeyDown(Index, vbKeyE, vbCtrlMask)
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyD Then
        If rstAccountChild.RecordCount = 0 Then Exit Sub
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            Set DataGrid2(Index).DataSource = Nothing
            rstAccountChild.Delete
            rstAccountChild.MoveNext
            Set DataGrid2(Index).DataSource = rstAccountChild
            DataGrid2(Index).SetFocus
        End If
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS And Toolbar1.Buttons.Item(4).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
    ElseIf Shift = 0 And KeyCode = vbKeyReturn Then
        Text2(Val(AccountType) - 1).SetFocus
        KeyCode = 0
    ElseIf Shift = vbShiftMask And KeyCode = vbKeyTab Then
       Text11(Val(AccountType) - 1).SetFocus
    End If
End Sub
Private Sub DataGrid2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim menusel As String
    If Button = vbRightButton Then
        menusel = DisplayPopupMenu(Me.hwnd)
        Select Case menusel
            Case 1
                Call DataGrid2_KeyDown(Index, vbKeyA, vbCtrlMask)
            Case 2
                Call DataGrid2_KeyDown(Index, vbKeyE, vbCtrlMask)
            Case 3
                Call DataGrid2_KeyDown(Index, vbKeyD, vbCtrlMask)
        End Select
    End If
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyD Then
        Dim Imported As Variant
        fpSpread1.GetText 5, fpSpread1.ActiveRow, Imported
        If Imported = "Y" Then Exit Sub
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            fpSpread1.DeleteRows fpSpread1.ActiveRow, 1
            fpSpread1.SetFocus
        End If
    ElseIf Shift = 0 And KeyCode = vbKeyF5 Then
        Call RefreshDropDownList("R")
    End If
End Sub
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim ActiveCellVal As Variant, Category As Variant
    fpSpread1.GetText Col, Row, ActiveCellVal
    If CheckEmpty(ActiveCellVal, False) Then Cancel = True: Exit Sub
    fpSpread1.GetText 1, Row, Category
    If Col = 1 Then
        fpSpread1.Col = 2
        fpSpread1.TypeComboBoxList = IIf(Category = "Outsource Item", OutsourceItem, IIf(Category = "Paper", Paper, IIf(Category = "Repair Book", RepairBook, IIf(Category = "Fresh Book", FreshBook, Title))))
    ElseIf Col = 2 Then
        If Category = "Outsource Item" Then
           If rstOutsourceItemList.RecordCount > 0 Then rstOutsourceItemList.MoveFirst
           rstOutsourceItemList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstOutsourceItemList.EOF Then fpSpread1.SetText 4, Row, rstOutsourceItemList.Fields("NCode").Value
        ElseIf Category = "Paper" Then
           If rstPaperList.RecordCount > 0 Then rstPaperList.MoveFirst
           rstPaperList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstPaperList.EOF Then fpSpread1.SetText 4, Row, rstPaperList.Fields("NCode").Value
        ElseIf Category = "Repair Book" Then
           If rstRepairBookList.RecordCount > 0 Then rstRepairBookList.MoveFirst
           rstRepairBookList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstRepairBookList.EOF Then fpSpread1.SetText 4, Row, rstRepairBookList.Fields("NCode").Value
        Else
           If rstFreshBookList.RecordCount > 0 Then rstFreshBookList.MoveFirst
           rstFreshBookList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstFreshBookList.EOF Then fpSpread1.SetText 4, Row, rstFreshBookList.Fields("NCode").Value
        End If
    End If
End Sub
Private Sub fpSpread1_BeforeEditMode(ByVal Col As Long, ByVal Row As Long, ByVal UserAction As FPSpreadADO.BeforeEditModeActionConstants, CursorPos As Variant, Cancel As Variant)
    Dim Imported As Variant
    fpSpread1.GetText 5, fpSpread1.ActiveRow, Imported
    If Imported = "Y" Then Cancel = True
End Sub
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Function CheckItem() As Boolean
    Dim i As Integer, Item As Variant, Category As Variant
    CheckItem = False
    For i = 1 To fpSpread1.DataRowCnt
        fpSpread1.SetActiveCell 1, i
        fpSpread1.GetText 4, i, Item
        fpSpread1.GetText 1, i, Category
        If Category = "Outsource Item" Then
            If Left(Item, 1) <> "1" Then CheckItem = True
        ElseIf Category = "Paper" Then
            If Left(Item, 1) <> "2" Then CheckItem = True
        ElseIf Category = "Repair Book" Then
            If Left(Item, 1) <> "4" Then CheckItem = True
        Else
            If Left(Item, 1) <> "3" And Left(Item, 1) <> "5" Then CheckItem = True
        End If
        If CheckItem Then DisplayError "Data mismatch in row #" & Trim(Str(i)): Exit For
    Next
End Function
Private Sub LoadMaterialList(ByVal strAccountCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    If rstAccountChild.State = adStateOpen Then rstAccountChild.Close
    If DatabaseType = "MS SQL" Then
        rstAccountChild.Open "SELECT Category,CASE WHEN Category='1' THEN (SELECT Name FROM OutsourceItemMaster WHERE Code=T.Item) WHEN Category='2' THEN (SELECT LTRIM(M.Name)+' (UOM : '+LTRIM(C.Name)+')' As Name FROM PaperMaster M INNER JOIN GeneralMaster C ON M.UOM=C.Code WHERE M.Code=T.Item) ELSE (SELECT Name FROM BookMaster WHERE Code=T.Item) END AS ItemName,OpBal,Category+Item As ItemCode,Imported FROM AccountChild0801 T WHERE T.Code='" & strAccountCode & "' ORDER BY Category", cnAccountMaster, adOpenKeyset, adLockReadOnly
    Else
        rstAccountChild.Open "SELECT Category,IIF(Category='1',(SELECT Name FROM OutsourceItemMaster WHERE Code=T.Item),IIF(Category='2',(SELECT TRIM(M.Name)+' (UOM : '+TRIM(C.Name)+')' As Name FROM PaperMaster M INNER JOIN GeneralMaster C ON M.UOM=C.Code WHERE M.Code=T.Item),(SELECT Name FROM BookMaster WHERE Code=T.Item))) AS ItemName,OpBal,Category+Item As ItemCode,Imported FROM AccountChild0801 T WHERE T.Code='" & strAccountCode & "' ORDER BY Category", cnAccountMaster, adOpenKeyset, adLockReadOnly
    End If
    rstAccountChild.ActiveConnection = Nothing
    If rstAccountChild.RecordCount > 0 Then rstAccountChild.MoveFirst
    i = 0
    Do While Not rstAccountChild.EOF
        i = i + 1
        With fpSpread1
            .SetText 1, i, IIf(rstAccountChild.Fields("Category").Value = "1", "Outsource Item", IIf(rstAccountChild.Fields("Category").Value = "2", "Paper", IIf(rstAccountChild.Fields("Category").Value = "3", "Fresh Book", IIf(rstAccountChild.Fields("Category").Value = "4", "Repair Book", "Title"))))
            .Col = 2
            .TypeComboBoxList = IIf(rstAccountChild.Fields("Category").Value = "1", OutsourceItem, IIf(rstAccountChild.Fields("Category").Value = "2", Paper, IIf(rstAccountChild.Fields("Category").Value = "4", RepairBook, IIf(rstAccountChild.Fields("Category").Value = "3", FreshBook, Title))))
            .SetText 2, i, rstAccountChild.Fields("ItemName").Value
            .SetText 3, i, Val(rstAccountChild.Fields("OpBal").Value)
            .SetText 4, i, rstAccountChild.Fields("ItemCode").Value
            .SetText 5, i, rstAccountChild.Fields("Imported").Value
        End With
        rstAccountChild.MoveNext
    Loop
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Material List")
End Sub
Private Function UpdateMaterialList(ByVal ActionType As String) As Boolean
    Dim CellVal(1 To 3) As Variant
    On Error GoTo ErrorHandler
    UpdateMaterialList = True
    If ActionType = "D" And (Not blnRecordExist) Then Exit Function
    If ActionType <> "I" Then
        cnAccountMaster.Execute "Delete From AccountChild0801 Where Code = '" & rstAccountMaster.Fields("Code").Value & "' AND Imported='N'"
    Else
        With fpSpread1
            .GetText 1, .ActiveRow, CellVal(1)
            .GetText 3, .ActiveRow, CellVal(2)
            .GetText 4, .ActiveRow, CellVal(3)
        End With
        cnAccountMaster.Execute "Insert Into AccountChild0801 Values ('" & rstAccountMaster.Fields("Code").Value & "','" & IIf(CellVal(1) = "Outsource Item", "1", IIf(CellVal(1) = "Paper", "2", IIf(CellVal(1) = "Fresh Book", "3", IIf(CellVal(1) = "Repair Book", "4", "5")))) & "','" & Right(CellVal(3), 6) & "'," & Val(CellVal(2)) & ",'N')"
    End If
    Exit Function
ErrorHandler:
    UpdateMaterialList = False
End Function
Private Sub RefreshDropDownList(ByVal xType As String)
    If xType = "R" Then
        RefreshData rstOutsourceItemList: RefreshData rstPaperList: RefreshData rstFreshBookList: RefreshData rstRepairBookList
        OutsourceItem = "": Paper = "": FreshBook = "": RepairBook = "": Title = ""
    End If
    Do While Not rstOutsourceItemList.EOF
        If OutsourceItem = "" Then OutsourceItem = rstOutsourceItemList.Fields("Name").Value Else OutsourceItem = OutsourceItem + Chr$(9) + rstOutsourceItemList.Fields("Name").Value
        rstOutsourceItemList.MoveNext
    Loop
    Do While Not rstPaperList.EOF
        If Paper = "" Then Paper = rstPaperList.Fields("Name").Value Else Paper = Paper + Chr$(9) + rstPaperList.Fields("Name").Value
        rstPaperList.MoveNext
    Loop
    rstFreshBookList.Filter = "[Board]='000000'"
    Do While Not rstFreshBookList.EOF
        If FreshBook = "" Then FreshBook = rstFreshBookList.Fields("Name").Value Else FreshBook = FreshBook + Chr$(9) + rstFreshBookList.Fields("Name").Value
        rstFreshBookList.MoveNext
    Loop
    rstFreshBookList.Filter = "[Board]<>'000000'"
    Do While Not rstFreshBookList.EOF
        If Title = "" Then Title = rstFreshBookList.Fields("Name").Value Else Title = Title + Chr$(9) + rstFreshBookList.Fields("Name").Value
        rstFreshBookList.MoveNext
    Loop
    rstFreshBookList.Filter = adFilterNone
    Do While Not rstRepairBookList.EOF
        If RepairBook = "" Then RepairBook = rstRepairBookList.Fields("Name").Value Else RepairBook = RepairBook + Chr$(9) + rstRepairBookList.Fields("Name").Value
        rstRepairBookList.MoveNext
    Loop
End Sub
Private Sub DuplicateRecord()
    Dim TmpTbl As String
    TmpTbl = "T" & GetFileNameFromPath(GetTemporaryFileName()): TmpTbl = Left(TmpTbl, InStr(1, TmpTbl, ".", vbTextCompare) - 1)
    On Error GoTo ErrorHandler
    MdiMainMenu.MousePointer = vbHourglass
    Dim AccountCode As String, AccountName As String
    AccountCode = GenerateCode(cnAccountMaster, "SELECT MAX(Code) FROM AccountMaster", 6, "0")
    AccountName = Trim(Left(rstAccountList.Fields("Name").Value, 76)) + " (D)"
    cnAccountMaster.BeginTrans
    cnAccountMaster.Execute "SELECT * INTO [" & TmpTbl & "] FROM AccountMaster Where Code = '" & rstAccountList.Fields("Code").Value & "'"
    cnAccountMaster.Execute "UPDATE  [" & TmpTbl & "] SET Code='" & AccountCode & "',Name='" & AccountName & "',PrintName='" & AccountName & "'"
    cnAccountMaster.Execute "INSERT INTO AccountMaster SELECT * FROM " & TmpTbl
    cnAccountMaster.Execute "DROP TABLE " & TmpTbl
    cnAccountMaster.Execute "SELECT * INTO [" & TmpTbl & "] FROM AccountChild04 Where Code = '" & rstAccountList.Fields("Code").Value & "'"
    cnAccountMaster.Execute "UPDATE  [" & TmpTbl & "] SET Code='" & AccountCode & "'"
    cnAccountMaster.Execute "INSERT INTO AccountChild04 SELECT * FROM " & TmpTbl
    cnAccountMaster.Execute "DROP TABLE " & TmpTbl
    cnAccountMaster.Execute "SELECT * INTO [" & TmpTbl & "] FROM AccountChild05 Where Code = '" & rstAccountList.Fields("Code").Value & "'"
    cnAccountMaster.Execute "UPDATE  [" & TmpTbl & "] SET Code='" & AccountCode & "'"
    cnAccountMaster.Execute "INSERT INTO AccountChild05 SELECT * FROM " & TmpTbl
    cnAccountMaster.Execute "DROP TABLE " & TmpTbl
    cnAccountMaster.Execute "SELECT * INTO [" & TmpTbl & "] FROM AccountChild07 Where Code = '" & rstAccountList.Fields("Code").Value & "'"
    cnAccountMaster.Execute "UPDATE  [" & TmpTbl & "] SET Code='" & AccountCode & "'"
    cnAccountMaster.Execute "INSERT INTO AccountChild07 SELECT * FROM " & TmpTbl
    cnAccountMaster.Execute "DROP TABLE " & TmpTbl
    cnAccountMaster.Execute "SELECT * INTO [" & TmpTbl & "] FROM AccountChild08 Where Code = '" & rstAccountList.Fields("Code").Value & "'"
    cnAccountMaster.Execute "UPDATE  [" & TmpTbl & "] SET Code='" & AccountCode & "'"
    cnAccountMaster.Execute "INSERT INTO AccountChild08 SELECT * FROM " & TmpTbl
    cnAccountMaster.Execute "DROP TABLE " & TmpTbl
    cnAccountMaster.CommitTrans
    Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    Text1.Text = Trim(AccountName): Sendkeys "{END}"
    MdiMainMenu.MousePointer = vbNormal
    Call MsgBox("Successfully Duplicated the Record !", vbInformation, App.Title)
    Exit Sub
ErrorHandler:
    MdiMainMenu.MousePointer = vbNormal
    DisplayError ("Failed to Duplicate the Record")
    cnAccountMaster.RollbackTrans
End Sub
Private Sub SetMenuOptions(bVal As Boolean)
    MdiMainMenu.mnuAccountMaster.Enabled = bVal
    MdiMainMenu.mnuRateMaster.Enabled = bVal
    MdiMainMenu.mnuMaterialCentreMaster.Enabled = bVal
    MdiMainMenu.mnuDespatchManagement(1).Enabled = bVal
    MdiMainMenu.mnuDespatchManagement(2).Enabled = bVal
    MdiMainMenu.mnuDespatchManagement(3).Enabled = bVal
End Sub
