VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0F1F1508-C40A-101B-AD04-00AA00575482}#1.0#0"; "mhrinp32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBookMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Master"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11190
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
   ScaleHeight     =   9270
   ScaleWidth      =   11190
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   9255
      Left            =   0
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   0
      Width           =   11160
      _Version        =   65536
      _ExtentX        =   19685
      _ExtentY        =   16325
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
      Picture         =   "BookMaster.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   9045
         Left            =   120
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   120
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   15954
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
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
         TabPicture(0)   =   "BookMaster.frx":001C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Mh3dLabel1(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Mh3dLabel1(2)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "DataGrid1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Text1"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "BookMaster.frx":0038
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Mh3dFrame2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "btnNotes"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "txtNotes"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "&BOM"
         TabPicture(2)   =   "BookMaster.frx":0054
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Mh3dFrame3"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "&Editorial Components"
         TabPicture(3)   =   "BookMaster.frx":0070
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Mh3dFrame5"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).ControlCount=   1
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
            MaxLength       =   40
            MultiLine       =   -1  'True
            TabIndex        =   76
            ToolTipText     =   "Open Notes"
            Top             =   8580
            Visible         =   0   'False
            Width           =   5175
         End
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
            TabIndex        =   75
            Top             =   8580
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
            Left            =   720
            TabIndex        =   37
            Top             =   8565
            Width           =   5535
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   7995
            Left            =   120
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   450
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   14102
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BackColor       =   9164542
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
            ColumnCount     =   4
            BeginProperty Column00 
               DataField       =   "ItemGrp"
               Caption         =   "Group"
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
            BeginProperty Column02 
               DataField       =   "BusyCode"
               Caption         =   "Alias"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "ISBN"
               Caption         =   "ISBN"
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
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   1755.213
               EndProperty
               BeginProperty Column01 
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   5295.118
               EndProperty
               BeginProperty Column02 
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column03 
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   1755.213
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   6585
            Left            =   -74880
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   480
            Width           =   10695
            _Version        =   65536
            _ExtentX        =   18865
            _ExtentY        =   11615
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
            Picture         =   "BookMaster.frx":008C
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
               Left            =   7275
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   30
               ToolTipText     =   "Finish Size"
               Top             =   5835
               Width           =   3300
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
               Left            =   2040
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   2
               Top             =   740
               Width           =   8535
            End
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
               Left            =   2040
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   5
               ToolTipText     =   "Title Size"
               Top             =   1370
               Width           =   8535
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
               Left            =   7275
               MaxLength       =   255
               TabIndex        =   32
               Top             =   6155
               Width           =   3300
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
               Left            =   7275
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   11
               Top             =   3630
               Width           =   3300
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
               Left            =   2520
               MaxLength       =   40
               TabIndex        =   23
               Top             =   5210
               Width           =   2820
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
               Left            =   7275
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   10
               Top             =   3315
               Width           =   3300
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
               Left            =   2040
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   3
               Top             =   1050
               Width           =   3300
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
               Left            =   7275
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   24
               Top             =   5210
               Width           =   3300
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
               Left            =   7275
               MaxLength       =   17
               TabIndex        =   22
               Top             =   4895
               Width           =   3300
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
               Left            =   2040
               MaxLength       =   60
               TabIndex        =   1
               Top             =   420
               Width           =   8535
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
               Left            =   2040
               MaxLength       =   60
               TabIndex        =   0
               Top             =   100
               Width           =   8535
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   40
               Top             =   420
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
               Caption         =   " Print Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":00A8
               Picture         =   "BookMaster.frx":00C4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   41
               Top             =   100
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
               Caption         =   " Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":00E0
               Picture         =   "BookMaster.frx":00FC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   5325
               TabIndex        =   42
               Top             =   3630
               Width           =   1965
               _Version        =   65536
               _ExtentX        =   3466
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
               Caption         =   " Operation"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":0118
               Picture         =   "BookMaster.frx":0134
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Left            =   5325
               TabIndex        =   43
               Top             =   4895
               Width           =   1965
               _Version        =   65536
               _ExtentX        =   3466
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
               Caption         =   " ISBN"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":0150
               Picture         =   "BookMaster.frx":016C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
               Height          =   330
               Left            =   120
               TabIndex        =   44
               Top             =   740
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
               Caption         =   " Finish Size"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":0188
               Picture         =   "BookMaster.frx":01A4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Left            =   120
               TabIndex        =   45
               Top             =   4260
               Width           =   2415
               _Version        =   65536
               _ExtentX        =   4260
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
               Caption         =   " Total Binding Forms"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":01C0
               Picture         =   "BookMaster.frx":01DC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Left            =   120
               TabIndex        =   46
               Top             =   1680
               Width           =   10455
               _Version        =   65536
               _ExtentX        =   18441
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
               Caption         =   " Multi Sheet Printing Form Details"
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":01F8
               Picture         =   "BookMaster.frx":0214
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
               Height          =   330
               Left            =   120
               TabIndex        =   47
               Top             =   4890
               Width           =   2415
               _Version        =   65536
               _ExtentX        =   4260
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
               Caption         =   " Price/MRP"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":0230
               Picture         =   "BookMaster.frx":024C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
               Height          =   330
               Left            =   5325
               TabIndex        =   48
               Top             =   5210
               Width           =   1965
               _Version        =   65536
               _ExtentX        =   3466
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
               Picture         =   "BookMaster.frx":0268
               Picture         =   "BookMaster.frx":0284
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
               Height          =   330
               Left            =   5325
               TabIndex        =   49
               Top             =   3315
               Width           =   1965
               _Version        =   65536
               _ExtentX        =   3466
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
               Caption         =   " Binding Type"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":02A0
               Picture         =   "BookMaster.frx":02BC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel18 
               Height          =   330
               Left            =   5325
               TabIndex        =   50
               Top             =   1050
               Width           =   1965
               _Version        =   65536
               _ExtentX        =   3466
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
               Caption         =   " Form Type"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":02D8
               Picture         =   "BookMaster.frx":02F4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel19 
               Height          =   330
               Left            =   120
               TabIndex        =   51
               Top             =   3945
               Width           =   2415
               _Version        =   65536
               _ExtentX        =   4260
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
               Caption         =   " Pages "
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":0310
               Picture         =   "BookMaster.frx":032C
            End
            Begin MhinrelLib.MhRealInput MhRealInput7 
               Height          =   330
               Left            =   7275
               TabIndex        =   13
               TabStop         =   0   'False
               ToolTipText     =   "Forms"
               Top             =   3945
               Width           =   3300
               _Version        =   65536
               _ExtentX        =   5821
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
               FillColor       =   16777215
               MaxReal         =   9999999
               MinReal         =   0
               ReadOnly        =   -1  'True
               SpinChangeReal  =   0
               CaretColor      =   -2147483642
               DecimalPlaces   =   2
               VAlignment      =   2
               FocusSelect     =   -1  'True
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
               Height          =   330
               Left            =   120
               TabIndex        =   52
               Top             =   5205
               Width           =   2415
               _Version        =   65536
               _ExtentX        =   4260
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
               Picture         =   "BookMaster.frx":0348
               Picture         =   "BookMaster.frx":0364
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
               Height          =   330
               Left            =   2520
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   4260
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":0380
               Caption         =   "BookMaster.frx":03A0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":040C
               Keys            =   "BookMaster.frx":042A
               Spin            =   "BookMaster.frx":0474
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "##0"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "##0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999
               MinValue        =   0
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
               Height          =   330
               Left            =   2520
               TabIndex        =   21
               ToolTipText     =   "Printing Form"
               Top             =   4890
               Width           =   2820
               _Version        =   65536
               _ExtentX        =   4974
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":049C
               Caption         =   "BookMaster.frx":04BC
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":0528
               Keys            =   "BookMaster.frx":0546
               Spin            =   "BookMaster.frx":0590
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
               Format          =   "#########0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999999.99
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   233177093
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput15 
               Height          =   330
               Left            =   2520
               TabIndex        =   12
               TabStop         =   0   'False
               ToolTipText     =   "Pages"
               Top             =   3945
               Width           =   2820
               _Version        =   65536
               _ExtentX        =   4974
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":05B8
               Caption         =   "BookMaster.frx":05D8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":0644
               Keys            =   "BookMaster.frx":0662
               Spin            =   "BookMaster.frx":06AC
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###0"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "###0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   1
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1909915653
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
               Height          =   330
               Left            =   5325
               TabIndex        =   53
               Top             =   4260
               Width           =   1965
               _Version        =   65536
               _ExtentX        =   3466
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
               Caption         =   " Add-On Rates"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":06D4
               Picture         =   "BookMaster.frx":06F0
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
               Height          =   330
               Left            =   7275
               TabIndex        =   16
               TabStop         =   0   'False
               ToolTipText     =   "Book Printing"
               Top             =   4260
               Width           =   1650
               _Version        =   65536
               _ExtentX        =   2910
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":070C
               Caption         =   "BookMaster.frx":072C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":0798
               Keys            =   "BookMaster.frx":07B6
               Spin            =   "BookMaster.frx":0800
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
               Format          =   "#########0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999999.99
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   233177093
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput16 
               Height          =   330
               Left            =   8910
               TabIndex        =   17
               TabStop         =   0   'False
               ToolTipText     =   "Binding"
               Top             =   4260
               Width           =   1665
               _Version        =   65536
               _ExtentX        =   2937
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":0828
               Caption         =   "BookMaster.frx":0848
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":08B4
               Keys            =   "BookMaster.frx":08D2
               Spin            =   "BookMaster.frx":091C
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
               Format          =   "#########0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999999.99
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   233177093
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel26 
               Height          =   330
               Left            =   120
               TabIndex        =   54
               Top             =   3315
               Width           =   2415
               _Version        =   65536
               _ExtentX        =   4260
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
               Caption         =   " Single Sheet Form Plate Type"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":0944
               Picture         =   "BookMaster.frx":0960
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel27 
               Height          =   330
               Left            =   120
               TabIndex        =   55
               Top             =   3630
               Width           =   2415
               _Version        =   65536
               _ExtentX        =   4260
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
               Caption         =   " Single Sheet Form Color"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":097C
               Picture         =   "BookMaster.frx":0998
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput17 
               Height          =   330
               Left            =   2520
               TabIndex        =   8
               ToolTipText     =   "Front Color"
               Top             =   3630
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":09B4
               Caption         =   "BookMaster.frx":09D4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":0A40
               Keys            =   "BookMaster.frx":0A5E
               Spin            =   "BookMaster.frx":0AA8
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "0"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9
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
               Value           =   4
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput18 
               Height          =   330
               Left            =   3915
               TabIndex        =   9
               ToolTipText     =   "Back Color"
               Top             =   3630
               Width           =   1425
               _Version        =   65536
               _ExtentX        =   2514
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":0AD0
               Caption         =   "BookMaster.frx":0AF0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":0B5C
               Keys            =   "BookMaster.frx":0B7A
               Spin            =   "BookMaster.frx":0BC4
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "0"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9
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
            Begin FPSpreadADO.fpSpread fpSpread2 
               Height          =   1335
               Left            =   120
               TabIndex        =   6
               Top             =   1995
               Width           =   10455
               _Version        =   524288
               _ExtentX        =   18441
               _ExtentY        =   2355
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
               GridColor       =   33023
               MaxCols         =   7
               MaxRows         =   3
               OperationMode   =   2
               SpreadDesigner  =   "BookMaster.frx":0BEC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   120
               TabIndex        =   58
               Top             =   5520
               Width           =   2415
               _Version        =   65536
               _ExtentX        =   4260
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
               Caption         =   " Duplex Printing"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":13D8
               Picture         =   "BookMaster.frx":13F4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
               Height          =   330
               Left            =   120
               TabIndex        =   59
               Top             =   4575
               Width           =   2415
               _Version        =   65536
               _ExtentX        =   4260
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
               Caption         =   " Quantity/Packet"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":1410
               Picture         =   "BookMaster.frx":142C
            End
            Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame4 
               Height          =   330
               Left            =   2520
               TabIndex        =   60
               TabStop         =   0   'False
               Top             =   5520
               Width           =   2820
               _Version        =   65536
               _ExtentX        =   4974
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
               Picture         =   "BookMaster.frx":1448
               Begin VB.OptionButton Option2 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "No"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   1395
                  TabIndex        =   26
                  Top             =   60
                  Width           =   615
               End
               Begin VB.OptionButton Option1 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Yes"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   120
                  TabIndex        =   25
                  Top             =   60
                  Value           =   -1  'True
                  Width           =   585
               End
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
               Height          =   330
               Left            =   2520
               TabIndex        =   18
               ToolTipText     =   "Qty/Pkt"
               Top             =   4575
               Width           =   2820
               _Version        =   65536
               _ExtentX        =   4974
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":1464
               Caption         =   "BookMaster.frx":1484
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":14F0
               Keys            =   "BookMaster.frx":150E
               Spin            =   "BookMaster.frx":1558
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###0"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "###0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1909915653
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel20 
               Height          =   330
               Left            =   5325
               TabIndex        =   61
               Top             =   4575
               Width           =   1965
               _Version        =   65536
               _ExtentX        =   3466
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
               Caption         =   " Pkt/Box && Loose/Box"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":1580
               Picture         =   "BookMaster.frx":159C
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput6 
               Height          =   330
               Left            =   7275
               TabIndex        =   19
               ToolTipText     =   "Pkt/Box"
               Top             =   4575
               Width           =   1650
               _Version        =   65536
               _ExtentX        =   2910
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":15B8
               Caption         =   "BookMaster.frx":15D8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":1644
               Keys            =   "BookMaster.frx":1662
               Spin            =   "BookMaster.frx":16AC
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "#####0"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "#####0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   233177093
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput19 
               Height          =   330
               Left            =   3915
               TabIndex        =   15
               TabStop         =   0   'False
               ToolTipText     =   "Extra Forms"
               Top             =   4260
               Width           =   1425
               _Version        =   65536
               _ExtentX        =   2514
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":16D4
               Caption         =   "BookMaster.frx":16F4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":1760
               Keys            =   "BookMaster.frx":177E
               Spin            =   "BookMaster.frx":17C8
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "##0"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "##0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999
               MinValue        =   0
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel21 
               Height          =   330
               Left            =   5325
               TabIndex        =   62
               Top             =   6150
               Width           =   1965
               _Version        =   65536
               _ExtentX        =   3466
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
               Caption         =   " Remarks"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":17F0
               Picture         =   "BookMaster.frx":180C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel22 
               Height          =   330
               Left            =   5325
               TabIndex        =   63
               Top             =   5525
               Width           =   1965
               _Version        =   65536
               _ExtentX        =   3466
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
               Caption         =   " Royalty (%)"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":1828
               Picture         =   "BookMaster.frx":1844
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput10 
               Height          =   330
               Left            =   7275
               TabIndex        =   27
               TabStop         =   0   'False
               ToolTipText     =   "Binding"
               Top             =   5525
               Width           =   3300
               _Version        =   65536
               _ExtentX        =   5821
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":1860
               Caption         =   "BookMaster.frx":1880
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":18EC
               Keys            =   "BookMaster.frx":190A
               Spin            =   "BookMaster.frx":1954
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "#0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "#0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99.99
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1909915653
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput8 
               Height          =   330
               Left            =   8910
               TabIndex        =   20
               TabStop         =   0   'False
               ToolTipText     =   "Loose Qty/Box"
               Top             =   4575
               Width           =   1665
               _Version        =   65536
               _ExtentX        =   2937
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":197C
               Caption         =   "BookMaster.frx":199C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":1A08
               Keys            =   "BookMaster.frx":1A26
               Spin            =   "BookMaster.frx":1A70
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###0"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "###0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1909915653
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
               Height          =   330
               Left            =   120
               TabIndex        =   64
               Top             =   5835
               Width           =   2415
               _Version        =   65536
               _ExtentX        =   4260
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
               Caption         =   " Bill of Material"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":1A98
               Picture         =   "BookMaster.frx":1AB4
            End
            Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame6 
               Height          =   330
               Left            =   2520
               TabIndex        =   65
               TabStop         =   0   'False
               Top             =   5835
               Width           =   2820
               _Version        =   65536
               _ExtentX        =   4974
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
               Picture         =   "BookMaster.frx":1AD0
               Begin VB.OptionButton Option4 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Yes"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   120
                  TabIndex        =   28
                  Top             =   60
                  Width           =   585
               End
               Begin VB.OptionButton Option3 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "No"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   1395
                  TabIndex        =   29
                  Top             =   60
                  Value           =   -1  'True
                  Width           =   615
               End
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
               Left            =   4005
               TabIndex        =   66
               ToolTipText     =   "Finish Size"
               Top             =   2520
               Width           =   1485
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   330
               Left            =   5325
               TabIndex        =   69
               Top             =   5835
               Width           =   1965
               _Version        =   65536
               _ExtentX        =   3466
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
               Caption         =   " HSN Code"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":1AEC
               Picture         =   "BookMaster.frx":1B08
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
               Height          =   330
               Left            =   120
               TabIndex        =   70
               Top             =   1050
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
               Caption         =   " Multi Sheet Form Size"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":1B24
               Picture         =   "BookMaster.frx":1B40
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel23 
               Height          =   330
               Left            =   120
               TabIndex        =   71
               Top             =   1365
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
               Caption         =   " Single Sheet Form Size"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":1B5C
               Picture         =   "BookMaster.frx":1B78
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel24 
               Height          =   330
               Left            =   120
               TabIndex        =   72
               Top             =   6150
               Width           =   2415
               _Version        =   65536
               _ExtentX        =   4260
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
               Caption         =   " Weight"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":1B94
               Picture         =   "BookMaster.frx":1BB0
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput20 
               Height          =   330
               Left            =   2520
               TabIndex        =   31
               ToolTipText     =   "Printing Form"
               Top             =   6150
               Width           =   2820
               _Version        =   65536
               _ExtentX        =   4974
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":1BCC
               Caption         =   "BookMaster.frx":1BEC
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":1C58
               Keys            =   "BookMaster.frx":1C76
               Spin            =   "BookMaster.frx":1CC0
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "########0.000"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "########0.000"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999999999.999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   233373701
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel25 
               Height          =   330
               Left            =   5325
               TabIndex        =   77
               Top             =   3945
               Width           =   1965
               _Version        =   65536
               _ExtentX        =   3466
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
               Caption         =   " &Forms"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":1CE8
               Picture         =   "BookMaster.frx":1D04
            End
            Begin MSForms.ComboBox Combo7 
               Height          =   330
               Left            =   2520
               TabIndex        =   7
               Top             =   3315
               Width           =   2820
               VariousPropertyBits=   545282075
               BackColor       =   16777215
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "4974;582"
               MatchEntry      =   0
               ShowDropButtonWhen=   1
               SpecialEffect   =   0
               FontName        =   "Calibri"
               FontHeight      =   195
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox Combo1 
               Height          =   330
               Left            =   7275
               TabIndex        =   4
               Top             =   1050
               Width           =   3300
               VariousPropertyBits=   545282075
               BackColor       =   16777215
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "5821;582"
               MatchEntry      =   0
               ShowDropButtonWhen=   1
               SpecialEffect   =   0
               FontName        =   "Calibri"
               FontHeight      =   195
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame3 
            Height          =   7605
            Left            =   -74880
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   480
            Width           =   10695
            _Version        =   65536
            _ExtentX        =   18865
            _ExtentY        =   13414
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
            Picture         =   "BookMaster.frx":1D20
            Begin FPSpreadADO.fpSpread fpSpread1 
               Height          =   7380
               Left            =   120
               TabIndex        =   57
               Top             =   105
               Width           =   10455
               _Version        =   524288
               _ExtentX        =   18441
               _ExtentY        =   13018
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
               MaxCols         =   4
               MaxRows         =   100
               OperationMode   =   2
               SpreadDesigner  =   "BookMaster.frx":1D3C
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame5 
            Height          =   7605
            Left            =   -74880
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   480
            Width           =   10695
            _Version        =   65536
            _ExtentX        =   18865
            _ExtentY        =   13414
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
            Picture         =   "BookMaster.frx":2322
            Begin FPSpreadADO.fpSpread fpSpread3 
               Height          =   7380
               Left            =   120
               TabIndex        =   68
               Top             =   105
               Width           =   10455
               _Version        =   524288
               _ExtentX        =   18441
               _ExtentY        =   13017
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
               MaxCols         =   1
               MaxRows         =   100
               OperationMode   =   2
               SpreadDesigner  =   "BookMaster.frx":233E
            End
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   2
            Left            =   6240
            TabIndex        =   73
            Top             =   8565
            Width           =   4575
            _Version        =   65536
            _ExtentX        =   8070
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
            Caption         =   " Ctrl+A->Add  Ctrl+E->Edit  Ctrl+D->Delete  Ctrl+S->Save"
            Alignment       =   0
            FillColor       =   8421504
            TextColor       =   16777215
            Picture         =   "BookMaster.frx":2800
            Picture         =   "BookMaster.frx":281C
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   1
            Left            =   6600
            TabIndex        =   74
            Top             =   0
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
            Caption         =   "  F5-> Refresh-> F12-> Create Duplicate Item Master"
            Alignment       =   0
            FillColor       =   8421504
            TextColor       =   16777215
            Picture         =   "BookMaster.frx":2838
            Picture         =   "BookMaster.frx":2854
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   330
            Left            =   120
            TabIndex        =   38
            Top             =   8565
            Width           =   615
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   11190
      _ExtentX        =   19738
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
Attribute VB_Name = "FrmBookMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SL As Boolean 'Selection List
Public MasterCode As String  'Master to Modify
Public BookType As String
Dim cnBookMaster As New ADODB.Connection
Dim rstBookList As New ADODB.Recordset, rstBookMaster As New ADODB.Recordset, rstGroupList As New ADODB.Recordset, rstSizeList As New ADODB.Recordset, rstFinishSizeList As New ADODB.Recordset, rstBindingTypeList As New ADODB.Recordset, rstOperationList As New ADODB.Recordset, rstOutsourceItemList As New ADODB.Recordset, rstFreshBookList As New ADODB.Recordset, rstBookChild As New ADODB.Recordset, rstHSNCodeList As New ADODB.Recordset
Dim GroupCode As String, SizeCode As String, TitleSizeCode As String, FinishSizeCode As String, BindingTypeCode As String, LaminationTypeCode As String, TextSizeCode As String, HSNCode As String
Dim SortOrder As String
Dim PrevStr As String
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Dim OutsourceItem As String
Dim FreshBook As String
Dim EditMode As Boolean
Private Sub btnNotes_Click()
    frmNotes.NotesFlag = 2
    frmNotes.Label1.Caption = "Notes : " & Text2.Text
    frmNotes.Show (vbModal)
    Text2.SetFocus
End Sub
Private Sub Form_Load()
    If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
    If Not SL Then MasterCode = ""
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    Me.Caption = IIf(BookType = "F", "Item Master [FG]", "Item Master [UFG]")
    cnBookMaster.CursorLocation = adUseClient
    cnBookMaster.Open cnDatabase.ConnectionString
    rstBookList.Open "SELECT P.Name,BusyCode,Board,ISBN,C.Name As ItemGrp,P.Code FROM BookMaster P INNER JOIN GeneralMaster C ON P.[Group]=C.Code WHERE P.Type='" & BookType & "' ORDER BY P.Name", cnBookMaster, adOpenKeyset, adLockOptimistic
    LoadMasterList
    rstBookMaster.CursorLocation = adUseClient
    rstBookList.Filter = adFilterNone
    If rstBookList.RecordCount Then
        If CheckEmpty(MasterCode, False) Then
            rstBookList.MoveFirst
        Else
            rstBookList.MoveFirst
            rstBookList.Find "[Code]='" & MasterCode & "'"
        End If
    End If
    Set DataGrid1.DataSource = rstBookList
    BusySystemIndicator False
    SSTab1.Tab = 0
    SortOrder = "Name"
    If Not (rstBookList.EOF Or rstBookList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstBookList.ActiveConnection = Nothing
    Combo1.AddItem "2 Pages", 0
    Combo1.AddItem "4 Pages", 1
    Combo1.AddItem "6 Pages", 2
    Combo1.AddItem "8 Pages", 3
    Combo1.AddItem "12 Pages", 4
    Combo1.AddItem "16 Pages", 5
    Combo1.AddItem "24 Pages", 6
    Combo1.AddItem "32 Pages", 7
    Combo1.AddItem "64 Pages", 8
    Combo7.AddItem "Depatch", 0
    Combo7.AddItem "PS", 1
    Combo7.AddItem "Wipeon", 2
    Combo7.AddItem "CTP", 3
    Call RefreshDropDownList("A")
    fpSpread1.Col = 4
    fpSpread1.ColHidden = True
    SetButtonsForNoRecord
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Call CloseForm(FrmBookMaster)
End Sub
Private Sub Form_Activate()
    MdiMainMenu.mnuBook.Enabled = False
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
       If Not EditMode Then
            KeyCode = 0
        End If
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(1)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE And Toolbar1.Buttons.Item(2).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF8 And Toolbar1.Buttons.Item(3).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(3)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS And Toolbar1.Buttons.Item(4).Enabled Then
        If Not EditMode Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF12 Then
        If MsgBox("Are you sure to make a duplicate copy of the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then DuplicateRecord
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
                If SSTab1.Tab = 0 Then Me.Tag = "S": slCode = rstBookList.Fields("Code").Value: slName = rstBookList.Fields("Name").Value: KeyCode = 0: Unload Me: Exit Sub
            Else
                SSTab1.Tab = 1
                SSTab1.SetFocus
            End If
        Else
           If Me.ActiveControl.Name <> "fpSpread1" And Me.ActiveControl.Name <> "fpSpread2" And Me.ActiveControl.Name <> "fpSpread3" Then Sendkeys "{TAB}"
        End If
        If Me.ActiveControl.Name <> "fpSpread1" And Me.ActiveControl.Name <> "fpSpread2" And Me.ActiveControl.Name <> "fpSpread3" Then KeyCode = 0
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
    Call CloseRecordset(rstBookList)
    Call CloseRecordset(rstBookMaster)
    Call CloseRecordset(rstGroupList)
    Call CloseRecordset(rstSizeList)
    Call CloseRecordset(rstFinishSizeList)
    Call CloseRecordset(rstBindingTypeList)
    Call CloseRecordset(rstOperationList)
    Call CloseRecordset(rstHSNCodeList)
    Call CloseRecordset(rstBookChild)
    Call CloseRecordset(rstOutsourceItemList)
    Call CloseRecordset(rstFreshBookList)
    Call CloseConnection(cnBookMaster)
    ShowProgressInStatusBar False
    MdiMainMenu.mnuBook.Enabled = True
End Sub
Private Sub Text1_Change()
    If rstBookList.RecordCount = 0 Then Exit Sub
    rstBookList.MoveFirst
    If Len(Text1.Text) > 0 Then
        rstBookList.Filter = "[" & SortOrder & "] Like '%" & FixQuote(Text1.Text) & "%'"
        If rstBookList.EOF Then  'if Spelling mistake
            rstBookList.Filter = adFilterNone
            rstBookList.MoveFirst
            Beep
            DisplayError ("Spelling Error")
            Text1.Text = PrevStr
            Sendkeys "{End}"
        Else    'if Spelling alright
            PrevStr = Text1.Text
        End If
    Else
        rstBookList.Filter = adFilterNone
        rstBookList.MoveFirst
        Set DataGrid1.DataSource = rstBookList
        PrevStr = ""
    End If
    If Not (rstBookList.EOF Or rstBookList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    
    If rstBookList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstBookList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstBookList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstBookList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstBookList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstBookList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstBookList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstBookList
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
        If SSTab1.Tab >= 1 Then
            ViewRecord
        Else
            If Not (rstBookList.EOF Or rstBookList.BOF) Then
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
        If SSTab1.Tab = 1 Then
            Mh3dFrame2.Enabled = True
            Mh3dFrame3.Enabled = False
            Mh3dFrame5.Enabled = False
            Text2.SetFocus
        ElseIf SSTab1.Tab = 2 Then
            Mh3dFrame2.Enabled = False
            Mh3dFrame3.Enabled = True
            Mh3dFrame5.Enabled = False
            fpSpread1.SetFocus
        Else
            Mh3dFrame2.Enabled = False
            Mh3dFrame3.Enabled = False
            Mh3dFrame5.Enabled = True
            fpSpread3.SetFocus
        End If
    End If
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim HiLiteRecord As Boolean
    Dim UpdateFlag As Integer, i As Integer
     Dim CellVal As Variant
   
    If Button.Index = 1 Then
        If rstBookMaster.State = adStateOpen Then
           rstBookMaster.Close
        End If
        rstBookMaster.Open "Select * From BookMaster WHERE Code = ''", cnBookMaster, adOpenKeyset, adLockOptimistic
        ClearFields
        If AddRecord(rstBookMaster) Then
           Call SetButtons(False)
           SSTab1.Tab = 1
           Text2.SetFocus
           blnRecordExist = False
           cnBookMaster.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstBookList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstBookList.RecordCount = 0 Then Exit Sub
        If AllowMastersDeletion = 0 Or rstBookList.Fields("Board").Value = "000000" Then
            Call DisplayError("You don't have the rights to Delete this Master")
            Exit Sub
        End If
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            cnBookMaster.Execute "Delete From BookMaster WHERE Code = '" & rstBookList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstBookList.Delete
                rstBookList.MoveNext
                If rstBookList.RecordCount > 0 And rstBookList.EOF Then rstBookList.MoveLast
                Call UpdateUserAction("Book Master", "D", Trim(Text2.Text), cnBookMaster)
                ShowProgressInStatusBar True
                Timer1.Enabled = True
            Else
                DisplayError ("Failed to delete the record")
            End If
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
        If UpdateRecord(rstBookMaster) Then
            If UpdateMaterialList("D1") Then
                UpdateFlag = 1
                For i = 1 To fpSpread1.DataRowCnt
                    fpSpread1.SetActiveCell 3, i
                    fpSpread1.GetText 3, i, CellVal
                    If Val(CellVal) <> 0 Then
                        If Not UpdateMaterialList("I") Then
                            UpdateFlag = 0
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
        If UpdateFlag Then
            Call UpdateUserAction("Book Master", IIf(blnRecordExist, "M", "A"), Trim(Text2.Text), cnBookMaster)
            AddToList
            cnBookMaster.CommitTrans
            If rstBookMaster.State = adStateOpen Then rstBookMaster.Close
            rstBookMaster.CursorLocation = adUseClient
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
        If CancelRecordUpdate(rstBookMaster) Then
            cnBookMaster.RollbackTrans
            If rstBookMaster.State = adStateOpen Then rstBookMaster.Close
            rstBookMaster.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstBookList.ActiveConnection = cnBookMaster
        Do While Not RefreshRecord(rstBookList)
        Loop
        Set DataGrid1.DataSource = rstBookList
        rstBookList.ActiveConnection = Nothing
        rstGroupList.ActiveConnection = cnBookMaster
        Do While Not RefreshRecord(rstGroupList)
        Loop
        rstGroupList.ActiveConnection = Nothing
        rstSizeList.ActiveConnection = cnBookMaster
        Do While Not RefreshRecord(rstSizeList)
        Loop
        rstSizeList.ActiveConnection = Nothing
        rstFinishSizeList.ActiveConnection = cnBookMaster
        Do While Not RefreshRecord(rstFinishSizeList)
        Loop
        rstFinishSizeList.ActiveConnection = Nothing
        rstBindingTypeList.ActiveConnection = cnBookMaster
        Do While Not RefreshRecord(rstBindingTypeList)
        Loop
        rstBindingTypeList.ActiveConnection = Nothing
        rstOperationList.ActiveConnection = cnBookMaster
        Do While Not RefreshRecord(rstOperationList)
        Loop
        rstOperationList.ActiveConnection = Nothing
        rstHSNCodeList.ActiveConnection = cnBookMaster
        Do While Not RefreshRecord(rstHSNCodeList)
        Loop
        rstHSNCodeList.ActiveConnection = Nothing
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
        If rstBookList.RecordCount > 0 Then rstBookList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstBookList.RecordCount > 0 Then
           rstBookList.MovePrevious
           If rstBookList.BOF Then
              rstBookList.MoveNext
           End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstBookList.RecordCount > 0 Then
           rstBookList.MoveNext
           If rstBookList.EOF Then
              rstBookList.MovePrevious
           End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstBookList.RecordCount > 0 Then rstBookList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Call CloseForm(FrmBookMaster)
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstBookList.EOF Or rstBookList.BOF) Then
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
        rstBookList.Sort = "[" + SortOrder & "] Desc"
        AD = "Desc"
    Else
        rstBookList.Sort = "[" + SortOrder & "] Asc"
        AD = "Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstBookList.EOF Or rstBookList.BOF) Then
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
    Mh3dFrame2.Enabled = Not bVal
    Mh3dFrame3.Enabled = False
End Sub
Private Sub SetButtonsForNoRecord()
    If rstBookList.RecordCount = 0 Then
        Toolbar1.Buttons.Item(2).Enabled = False
        Toolbar1.Buttons.Item(3).Enabled = False
        Toolbar1.Buttons.Item(13).Enabled = False
        Toolbar1.Buttons.Item(14).Enabled = False
        Toolbar1.Buttons.Item(15).Enabled = False
        Toolbar1.Buttons.Item(16).Enabled = False
    End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    If rstBookMaster.EOF Or rstBookMaster.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(cnBookMaster, "BookMaster", "Code", "Name", Text2.Text, rstBookMaster.Fields("Code").Value, False) Then
        Cancel = True
    ElseIf CheckEmpty(Text3, False) Then
        Text3.Text = Text2.Text
    End If
End Sub
Private Sub Text4_Validate(Cancel As Boolean)
    If CheckEmpty(Text4.Text, False) Then Exit Sub
    If rstBookMaster.EOF Or rstBookMaster.BOF Then Exit Sub
    If CheckDuplicate(cnBookMaster, "BookMaster", "Code", "Isbn", Text4.Text, rstBookMaster.Fields("Code").Value, False) Then
        Cancel = True
    ElseIf Len(Text4.Text) = 13 Then
        If Not bVerifySum10(Text4.Text) Then Cancel = True
    ElseIf Len(Text4.Text) = 17 Then
        If Not bVerifySum13(Text4.Text) Then Cancel = True
    End If
End Sub
Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmGeneralMaster.SL = True
        FrmGeneralMaster.MasterType = "11"
        FrmGeneralMaster.MasterCode = FinishSizeCode
        Load FrmGeneralMaster
        If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
        On Error GoTo 0
        FinishSizeCode = slCode: Text5.Text = slName
        If Not CheckEmpty(FinishSizeCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    ElseIf KeyCode = vbKeyDelete Then
        FinishSizeCode = "": Text5.Text = ""
    End If
End Sub
Private Sub Text5_Validate(Cancel As Boolean)
    If CheckEmpty(Text5.Text, False) Then
        Cancel = True
    Else
        If rstBookChild.State = adStateOpen Then rstBookChild.Close
        rstBookChild.Open "SELECT 'Text Size: '+M1.Name+'|'+'Pgs/Ptg Form: '+IIF([Ups/Form]<10,'0','')+LTRIM([Ups/Form])+'|Pgs/Bdg Form: '+IIF([Ups/BdgForm]<10,'0','')+LTRIM([Ups/BdgForm])+'|Title Size: '+M2.Name As Col0,C.Code+C.[TextSize]+LTRIM([Ups/Form])+LTRIM([Ups/BdgForm])+C.TitleSize As Code,M1.Name As TextSizeName,[Ups/Form],[Ups/BdgForm],M2.Name As TitleSizeName,[TextSize],TitleSize FROM (FinishSizeChild C INNER JOIN GeneralMaster M1 ON C.[TextSize]=M1.Code) INNER JOIN GeneralMaster M2 ON C.TitleSize=M2.Code WHERE C.Code='" & FinishSizeCode & "' ORDER BY M1.Name,[Ups/Form]", cnBookMaster, adOpenKeyset, adLockReadOnly
        If rstBookChild.RecordCount = 0 Then Sendkeys "{TAB}": Exit Sub
        rstBookChild.MoveFirst
        SelectionType = "S"
        TextSizeCode = ""
        Call LoadSelectionList(rstBookChild, "List of Sizes...", "Name", "")
        SearchOrder = 0
        Call DisplaySelectionList(Text6, TextSizeCode)
        Call CloseForm(FrmSelectionList)
        If Not CheckEmpty(RTrim(TextSizeCode), False) Then
            rstBookChild.MoveFirst
            rstBookChild.Find "[Code]='" & TextSizeCode & "'"
            If CheckEmpty(Text9.Text, False) Then
                Text9.Text = rstBookChild.Fields("TextSizeName").Value: SizeCode = rstBookChild.Fields("TextSize").Value
            ElseIf Text9.Text <> rstBookChild.Fields("TextSizeName").Value Then
                If MsgBox("Variation in Current (" & Text9.Text & ") and Master (" & rstBookChild.Fields("TextSizeName").Value & ") Text Size !!! Change Text Size?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then Text9.Text = rstBookChild.Fields("TextSizeName").Value: SizeCode = rstBookChild.Fields("TextSize").Value
            End If
            If CheckEmpty(Text14.Text, False) Then
                Text14.Text = rstBookChild.Fields("TitleSizeName").Value: TitleSizeCode = rstBookChild.Fields("TitleSize").Value
            ElseIf Text14.Text <> rstBookChild.Fields("TitleSizeName").Value Then
                If MsgBox("Variation in Current (" & Text14.Text & ") and Master (" & rstBookChild.Fields("TitleSizeName").Value & ") Title Size !!! Change Title Size?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then Text14.Text = rstBookChild.Fields("TitleSizeName").Value: TitleSizeCode = rstBookChild.Fields("TitleSize").Value
            End If
            If Val(Combo1.Text) <> Val(rstBookChild.Fields("Ups/Form").Value) Then
                If MsgBox("Variation in Current (" & Combo1.Text & ") and Master (" & Trim(rstBookChild.Fields("Ups/Form").Value) & ") Ups/Form !!! Change Ups/Form?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then Combo1.Text = Trim(rstBookChild.Fields("Ups/Form").Value) & " Pages"
            End If
        End If
    End If
End Sub
Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmGeneralMaster.SL = True
        FrmGeneralMaster.MasterType = "1"
        FrmGeneralMaster.MasterCode = SizeCode
        Load FrmGeneralMaster
        If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
        On Error GoTo 0
        SizeCode = slCode: Text9.Text = slName
        If Not CheckEmpty(SizeCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    ElseIf KeyCode = vbKeyDelete Then
        SizeCode = "": Text9.Text = ""
    End If
End Sub
Private Sub Text9_Validate(Cancel As Boolean)
    If CheckEmpty(Text9.Text, False) Then Cancel = True
End Sub
Private Sub Text14_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmGeneralMaster.SL = True
        FrmGeneralMaster.MasterType = "1"
        FrmGeneralMaster.MasterCode = TitleSizeCode
        Load FrmGeneralMaster
        If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
        On Error GoTo 0
        TitleSizeCode = slCode: Text14.Text = slName
        If Not CheckEmpty(TitleSizeCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    ElseIf KeyCode = vbKeyDelete Then
        TitleSizeCode = "": Text14.Text = ""
    End If
End Sub
Private Sub Text10_KeyDown(KeyCode As Integer, Shift As Integer)    'Binding Type
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmGeneralMaster.SL = True
        FrmGeneralMaster.MasterType = "6"
        FrmGeneralMaster.MasterCode = BindingTypeCode
        Load FrmGeneralMaster
        If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
        On Error GoTo 0
        BindingTypeCode = slCode: Text10.Text = slName
        If Not CheckEmpty(BindingTypeCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    ElseIf KeyCode = vbKeyDelete Then
        BindingTypeCode = "": Text10.Text = ""
    End If
End Sub
Private Sub Text12_KeyDown(KeyCode As Integer, Shift As Integer)    'Operation
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmGeneralMaster.SL = True
        FrmGeneralMaster.MasterType = "7"
        FrmGeneralMaster.MasterCode = LaminationTypeCode
        Load FrmGeneralMaster
        If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
        On Error GoTo 0
        LaminationTypeCode = slCode: Text12.Text = slName
        If Not CheckEmpty(LaminationTypeCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    ElseIf KeyCode = vbKeyDelete Then
        LaminationTypeCode = "": Text12.Text = ""
    End If
End Sub
Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmGeneralMaster.SL = True
        FrmGeneralMaster.MasterType = "5"
        FrmGeneralMaster.MasterCode = GroupCode
        Load FrmGeneralMaster
        If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
        On Error GoTo 0
        GroupCode = slCode: Text8.Text = slName
        If Not CheckEmpty(GroupCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    ElseIf KeyCode = vbKeyDelete Then
        GroupCode = "": Text8.Text = ""
    End If
End Sub
Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer) 'HSN Code
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmGeneralMaster.SL = True
        FrmGeneralMaster.MasterType = "18"
        FrmGeneralMaster.MasterCode = HSNCode
        Load FrmGeneralMaster
        If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
        On Error GoTo 0
        HSNCode = slCode: Text7.Text = slName
        If Not CheckEmpty(HSNCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    ElseIf KeyCode = vbKeyDelete Then
        HSNCode = "": Text7.Text = ""
    End If
End Sub
Private Sub Text7_Validate(Cancel As Boolean)
    If CheckEmpty(Text7.Text, False) Then Cancel = True
End Sub
Private Sub Text8_Validate(Cancel As Boolean)
    If CheckEmpty(Text8.Text, False) Then Cancel = True
End Sub
Private Sub Combo1_Click()
    Dim Pages As Variant
    Dim Forms As Double, TotalPages As Long, TotalForms As Double, i As Integer
    TotalPages = 0: TotalForms = 0
    With fpSpread2
        For i = 1 To .DataRowCnt
            .GetText 2, i, Pages
            If Val(Pages) > 0 Then
                TotalPages = TotalPages + Pages
                Forms = Val(Pages) / Val(Combo1.Text)
                TotalForms = TotalForms + Forms
                .GetText 5, i, Pages   'F/B Forms
                If Val(Pages) = 0 Then .SetText 5, i, Int(Forms / 2) * 2
                .GetText 6, i, Pages   'W/T Forms
                If Val(Pages) = 0 Then .SetText 6, i, Int(Forms) - Int(Forms / 2) * 2
                .SetText 7, i, Forms
                Forms = Forms - Int(Forms)
                .GetText 3, i, Pages   '¼ Forms
                If Val(Pages) = 0 Then .SetText 3, i, IIf(Forms = 0.25, 1, IIf(Forms = 0.75, 1, IIf(Forms = 0.375, 1, IIf(Forms = 0.875, 1, 0))))
                .GetText 4, i, Pages   '½ Forms
                If Val(Pages) = 0 Then .SetText 4, i, IIf(Forms = 0.5, 1, IIf(Forms = 0.75, 1, IIf(Forms = 0.625, 1, IIf(Forms = 0.875, 1, IIf(Forms = (5 / 6), 1, 0)))))
            End If
        Next
    End With
    MhRealInput15.Value = TotalPages
    MhRealInput7.Text = Format(TotalForms, "0.00")
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstBookList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstBookMaster.State = adStateOpen Then
       rstBookMaster.Close
    End If
    rstBookMaster.Open "Select * From BookMaster WHERE Code = '" & FixQuote(rstBookList.Fields("Code").Value) & "'", cnBookMaster, adOpenKeyset, adLockOptimistic
    If rstBookMaster.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields()
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
    fpSpread2.ClearRange 1, 1, fpSpread2.MaxCols, fpSpread2.MaxRows, True
    fpSpread3.ClearRange 1, 1, fpSpread3.MaxCols, fpSpread3.MaxRows, True
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text11.Text = ""
    Text8.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text9.Text = ""
    Text14.Text = ""
    Text10.Text = ""
    Text12.Text = ""
    Text7.Text = "" 'HSN Code
    Text13.Text = ""
    MhRealInput1.Text = "0.00"
    MhRealInput3.Text = "0.00"
    MhRealInput16.Text = "0.00"
    MhRealInput17.Text = "4"
    MhRealInput18.Text = "0"
    MhRealInput4.Text = "0"
    MhRealInput19.Text = "0"
    MhRealInput20.Value = 0
    MhRealInput7.Text = 0#
    MhRealInput15.Text = "0"
    MhRealInput5.Text = "0"
    MhRealInput8.Text = "0"
    MhRealInput6.Text = "0"
    MhRealInput10.Text = "0.00"
    Option1.Value = True
    Option2.Value = False
    Option3.Value = True
    Option4.Value = False
    Combo1.ListIndex = 3
    Combo7.ListIndex = 3
    GroupCode = "": BindingTypeCode = "": LaminationTypeCode = "": HSNCode = "": FinishSizeCode = "": SizeCode = "": TitleSizeCode = ""
End Sub
Private Sub LoadFields()
    If rstBookMaster.EOF Or rstBookMaster.BOF Then Exit Sub
    Text2.Text = rstBookMaster.Fields("Name").Value
    Text3.Text = rstBookMaster.Fields("PrintName").Value
    Text4.Text = rstBookMaster.Fields("Isbn").Value
    Text11.Text = rstBookMaster.Fields("BusyCode").Value
    MhRealInput1.Text = Format(Val(rstBookMaster.Fields("Price").Value), "0.00")
    GroupCode = rstBookMaster.Fields("Group").Value
    If rstGroupList.RecordCount > 0 Then rstGroupList.MoveFirst
    rstGroupList.Find "[Code] = '" & GroupCode & "'"
    If Not rstGroupList.EOF Then Text8.Text = rstGroupList.Fields("Col0").Value
    SizeCode = rstBookMaster.Fields("Size").Value
    rstSizeList.MoveFirst
    rstSizeList.Find "[Code] = '" & SizeCode & "'"
    Text9.Text = rstSizeList.Fields("Col0").Value
    TitleSizeCode = rstBookMaster.Fields("TitleSize").Value
    rstSizeList.MoveFirst
    rstSizeList.Find "[Code] = '" & TitleSizeCode & "'"
    If Not rstSizeList.EOF Then Text14.Text = rstSizeList.Fields("Col0").Value
    FinishSizeCode = rstBookMaster.Fields("FinishSize").Value
    rstFinishSizeList.MoveFirst
    rstFinishSizeList.Find "[Code] = '" & FinishSizeCode & "'"
    Text5.Text = rstFinishSizeList.Fields("Col0").Value
    BindingTypeCode = rstBookMaster.Fields("BindingType").Value
    If rstBindingTypeList.RecordCount > 0 Then rstBindingTypeList.MoveFirst
    rstBindingTypeList.Find "[Code] = '" & BindingTypeCode & "'"
    If Not rstBindingTypeList.EOF Then Text10.Text = rstBindingTypeList.Fields("Col0").Value
    LaminationTypeCode = rstBookMaster.Fields("LaminationType").Value
    If rstOperationList.RecordCount > 0 Then rstOperationList.MoveFirst
    rstOperationList.Find "[Code] = '" & LaminationTypeCode & "'"
    If Not rstOperationList.EOF Then Text12.Text = rstOperationList.Fields("Col0").Value
    HSNCode = rstBookMaster.Fields("HSNCode").Value
    If rstHSNCodeList.RecordCount > 0 Then rstHSNCodeList.MoveFirst
    rstHSNCodeList.Find "[Code] = '" & HSNCode & "'"
    If Not rstHSNCodeList.EOF Then Text7.Text = rstHSNCodeList.Fields("Col0").Value
    MhRealInput3.Text = Format(Val(rstBookMaster.Fields("AddOnRate01").Value), "0.00")
    MhRealInput16.Text = Format(Val(rstBookMaster.Fields("AddOnRate02").Value), "0.00")
    Combo1.ListIndex = Choose(Val(rstBookMaster.Fields("FormType").Value), 3, 5, 1, 4, 6, 7, 8, 2, 0)
    MhRealInput4.Text = Format(Val(rstBookMaster.Fields("BindingForms01").Value), "0")
    MhRealInput19.Text = Format(Val(rstBookMaster.Fields("BindingForms02").Value), "0")
    MhRealInput20.Value = Val(rstBookMaster.Fields("Weight").Value)
    MhRealInput15.Text = Format(Val(rstBookMaster.Fields("Pages").Value), "0")
    MhRealInput7.Text = Format(Val(rstBookMaster.Fields("Forms").Value), "0.00")
    fpSpread2.SetText 1, 1, IIf(rstBookMaster.Fields("OneColorPlateType").Value = "1", "Deepatch", IIf(rstBookMaster.Fields("OneColorPlateType").Value = "2", "PS", IIf(rstBookMaster.Fields("OneColorPlateType").Value = "3", "Wipeon", "CTP")))
    fpSpread2.SetText 2, 1, Val(rstBookMaster.Fields("OneColorPages").Value)
    fpSpread2.SetText 3, 1, Val(rstBookMaster.Fields("OneColor¼Forms").Value)
    fpSpread2.SetText 4, 1, Val(rstBookMaster.Fields("OneColor½Forms").Value)
    fpSpread2.SetText 5, 1, Val(rstBookMaster.Fields("OneColor1F/BForms").Value)
    fpSpread2.SetText 6, 1, Val(rstBookMaster.Fields("OneColor1W/TForms").Value)
    fpSpread2.SetText 7, 1, Val(rstBookMaster.Fields("OneColorForms").Value)
    fpSpread2.SetText 1, 2, IIf(rstBookMaster.Fields("TwoColorPlateType").Value = "1", "Deepatch", IIf(rstBookMaster.Fields("TwoColorPlateType").Value = "2", "PS", IIf(rstBookMaster.Fields("TwoColorPlateType").Value = "3", "Wipeon", "CTP")))
    fpSpread2.SetText 2, 2, Val(rstBookMaster.Fields("TwoColorPages").Value)
    fpSpread2.SetText 3, 2, Val(rstBookMaster.Fields("TwoColor¼Forms").Value)
    fpSpread2.SetText 4, 2, Val(rstBookMaster.Fields("TwoColor½Forms").Value)
    fpSpread2.SetText 5, 2, Val(rstBookMaster.Fields("TwoColor1F/BForms").Value)
    fpSpread2.SetText 6, 2, Val(rstBookMaster.Fields("TwoColor1W/TForms").Value)
    fpSpread2.SetText 7, 2, Val(rstBookMaster.Fields("TwoColorForms").Value)
    fpSpread2.SetText 1, 3, IIf(rstBookMaster.Fields("FourColorPlateType").Value = "1", "Deepatch", IIf(rstBookMaster.Fields("FourColorPlateType").Value = "2", "PS", IIf(rstBookMaster.Fields("FourColorPlateType").Value = "3", "Wipeon", "CTP")))
    fpSpread2.SetText 2, 3, Val(rstBookMaster.Fields("FourColorPages").Value)
    fpSpread2.SetText 3, 3, Val(rstBookMaster.Fields("FourColor¼Forms").Value)
    fpSpread2.SetText 4, 3, Val(rstBookMaster.Fields("FourColor½Forms").Value)
    fpSpread2.SetText 5, 3, Val(rstBookMaster.Fields("FourColor1F/BForms").Value)
    fpSpread2.SetText 6, 3, Val(rstBookMaster.Fields("FourColor1W/TForms").Value)
    fpSpread2.SetText 7, 3, Val(rstBookMaster.Fields("FourColorForms").Value)
    Combo7.ListIndex = Val(rstBookMaster.Fields("TitlePlateType").Value) - 1
    MhRealInput17.Text = Format(Val(rstBookMaster.Fields("TitleFrontColor").Value), "0")
    MhRealInput18.Text = Format(Val(rstBookMaster.Fields("TitleBackColor").Value), "0")
    MhRealInput5.Text = Format(Val(rstBookMaster.Fields("Qty/Pkt").Value), "0")
    MhRealInput8.Text = Format(Val(rstBookMaster.Fields("LooseQty/Box").Value), "0")
    MhRealInput6.Text = Format(Val(rstBookMaster.Fields("Pkt/Box").Value), "0")
    MhRealInput10.Text = Format(Val(rstBookMaster.Fields("Royalty").Value), "0.00")
    Option1.Value = IIf(rstBookMaster.Fields("DuplexPrinting").Value = "Y", True, False)
    Option2.Value = IIf(rstBookMaster.Fields("DuplexPrinting").Value = "N", True, False)
    Option3.Value = IIf(rstBookMaster.Fields("Board").Value = "", True, False)
    Option4.Value = IIf(rstBookMaster.Fields("Board").Value = "000000", True, False)
    Text13.Text = rstBookMaster.Fields("Narration").Value
    txtNotes.Text = rstBookMaster.Fields("Notes").Value
    Call LoadMaterialList(rstBookMaster.Fields("Code").Value)
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstBookMaster.RecordCount = 0 Then Exit Sub
    If rstBookMaster.State = adStateOpen Then rstBookMaster.Close
    rstBookMaster.CursorLocation = adUseServer
    rstBookMaster.Open "Select * From BookMaster WHERE Code = '" & FixQuote(rstBookList.Fields("Code").Value) & "'", cnBookMaster, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstBookMaster.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text2.SetFocus
    blnRecordExist = True
    cnBookMaster.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    Dim Fld As Variant
    If rstBookMaster.EOF Or rstBookMaster.BOF Then Exit Sub
    If Not blnRecordExist Then
        rstBookMaster.Fields("Code").Value = GenerateCode(cnBookMaster, "Select Max(Code) From BookMaster", 6, "0")
        rstBookMaster.Fields("CreatedBy").Value = UserCode
        rstBookMaster.Fields("CreatedOn").Value = Now()
        rstBookMaster.Fields("Recordstatus").Value = "N"
    Else
        rstBookMaster.Fields("ModifiedBy").Value = UserCode
        rstBookMaster.Fields("ModifiedOn").Value = Now()
        rstBookMaster.Fields("Recordstatus").Value = "M"
    End If
    rstBookMaster.Fields("Name").Value = Trim(Text2.Text)
    rstBookMaster.Fields("PrintName").Value = Trim(Text3.Text)
    rstBookMaster.Fields("Isbn").Value = Trim(Text4.Text)
    rstBookMaster.Fields("BusyCode").Value = Trim(Text11.Text)
    rstBookMaster.Fields("Price").Value = Val(MhRealInput1.Text)
    rstBookMaster.Fields("Group").Value = GroupCode
    rstBookMaster.Fields("Size").Value = SizeCode
    rstBookMaster.Fields("TitleSize").Value = TitleSizeCode
    rstBookMaster.Fields("FinishSize").Value = FinishSizeCode
    rstBookMaster.Fields("BindingType").Value = BindingTypeCode
    rstBookMaster.Fields("LaminationType").Value = LaminationTypeCode
    rstBookMaster.Fields("HSNCode").Value = HSNCode
    rstBookMaster.Fields("AddOnRate01").Value = Val(MhRealInput3.Text)
    rstBookMaster.Fields("AddOnRate02").Value = Val(MhRealInput16.Text)
    rstBookMaster.Fields("FormType").Value = Choose(Combo1.ListIndex + 1, 9, 3, 8, 1, 4, 2, 5, 6, 7)
    rstBookMaster.Fields("BindingForms01").Value = Val(MhRealInput4.Text)
    rstBookMaster.Fields("BindingForms02").Value = Val(MhRealInput19.Text)
    rstBookMaster.Fields("Weight").Value = MhRealInput20.Value
    rstBookMaster.Fields("Pages").Value = Val(MhRealInput15.Text)
    rstBookMaster.Fields("Forms").Value = Val(MhRealInput7.Text)
    fpSpread2.GetText 1, 1, Fld
    rstBookMaster.Fields("OneColorPlateType").Value = IIf(Trim(Fld) = "Deepatch", "1", IIf(Trim(Fld) = "PS", "2", IIf(Trim(Fld) = "Wipeon", "3", "4")))
    fpSpread2.GetText 2, 1, Fld
    rstBookMaster.Fields("OneColorPages").Value = Val(Fld)
    fpSpread2.GetText 3, 1, Fld
    rstBookMaster.Fields("OneColor¼Forms").Value = Val(Fld)
    fpSpread2.GetText 4, 1, Fld
    rstBookMaster.Fields("OneColor½Forms").Value = Val(Fld)
    fpSpread2.GetText 5, 1, Fld
    rstBookMaster.Fields("OneColor1F/BForms").Value = Val(Fld)
    fpSpread2.GetText 6, 1, Fld
    rstBookMaster.Fields("OneColor1W/TForms").Value = Val(Fld)
    fpSpread2.GetText 7, 1, Fld
    rstBookMaster.Fields("OneColorForms").Value = Val(Fld)
    fpSpread2.GetText 1, 2, Fld
    rstBookMaster.Fields("TwoColorPlateType").Value = IIf(Trim(Fld) = "Deepatch", "1", IIf(Trim(Fld) = "PS", "2", IIf(Trim(Fld) = "Wipeon", "3", "4")))
    fpSpread2.GetText 2, 2, Fld
    rstBookMaster.Fields("TwoColorPages").Value = Val(Fld)
    fpSpread2.GetText 3, 2, Fld
    rstBookMaster.Fields("TwoColor¼Forms").Value = Val(Fld)
    fpSpread2.GetText 4, 2, Fld
    rstBookMaster.Fields("TwoColor½Forms").Value = Val(Fld)
    fpSpread2.GetText 5, 2, Fld
    rstBookMaster.Fields("TwoColor1F/BForms").Value = Val(Fld)
    fpSpread2.GetText 6, 2, Fld
    rstBookMaster.Fields("TwoColor1W/TForms").Value = Val(Fld)
    fpSpread2.GetText 7, 2, Fld
    rstBookMaster.Fields("TwoColorForms").Value = Val(Fld)
    fpSpread2.GetText 1, 3, Fld
    rstBookMaster.Fields("FourColorPlateType").Value = IIf(Trim(Fld) = "Deepatch", "1", IIf(Trim(Fld) = "PS", "2", IIf(Trim(Fld) = "Wipeon", "3", "4")))
    fpSpread2.GetText 2, 3, Fld
    rstBookMaster.Fields("FourColorPages").Value = Val(Fld)
    fpSpread2.GetText 3, 3, Fld
    rstBookMaster.Fields("FourColor¼Forms").Value = Val(Fld)
    fpSpread2.GetText 4, 3, Fld
    rstBookMaster.Fields("FourColor½Forms").Value = Val(Fld)
    fpSpread2.GetText 5, 3, Fld
    rstBookMaster.Fields("FourColor1F/BForms").Value = Val(Fld)
    fpSpread2.GetText 6, 3, Fld
    rstBookMaster.Fields("FourColor1W/TForms").Value = Val(Fld)
    fpSpread2.GetText 7, 3, Fld
    rstBookMaster.Fields("FourColorForms").Value = Val(Fld)
    rstBookMaster.Fields("TitlePlateType").Value = Trim(Str(Combo7.ListIndex + 1))
    rstBookMaster.Fields("TitleFrontColor").Value = Val(MhRealInput17.Text)
    rstBookMaster.Fields("TitleBackColor").Value = Val(MhRealInput18.Text)
    rstBookMaster.Fields("Qty/Pkt").Value = Val(MhRealInput5.Text)
    rstBookMaster.Fields("LooseQty/Box").Value = Val(MhRealInput8.Text)
    rstBookMaster.Fields("Pkt/Box").Value = Val(MhRealInput6.Text)
    rstBookMaster.Fields("Royalty").Value = Val(MhRealInput10.Text)
    rstBookMaster.Fields("DuplexPrinting").Value = IIf(Option1.Value, "Y", "N")
    rstBookMaster.Fields("Board").Value = IIf(Option4.Value, "000000", "")
    rstBookMaster.Fields("Narration").Value = Trim(Text13.Text)
    rstBookMaster.Fields("Type").Value = BookType
    rstBookMaster.Fields("PrintStatus").Value = "N"
    rstBookMaster.Fields("Notes").Value = txtNotes.Text
End Sub
Private Sub AddToList()
    On Error Resume Next
    rstBookList.MoveFirst
    rstBookList.Find "[Code] = '" & rstBookMaster.Fields("Code").Value & "'"
    If rstBookList.EOF Then rstBookList.AddNew:               rstBookList.Fields("Code").Value = rstBookMaster.Fields("Code").Value
    rstBookList.Fields("Name").Value = rstBookMaster.Fields("Name").Value
    rstBookList.Fields("BusyCode").Value = rstBookMaster.Fields("BusyCode").Value
    rstBookList.Fields("ItemGrp").Value = Text8.Text
    rstBookList.Update
    rstBookList.Sort = SortOrder & " Asc"
    rstBookList.Find "[Code] = '" & rstBookMaster.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
        SSTab1.Tab = 1
        Text2.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckDuplicate(cnBookMaster, "BookMaster", "Code", "Name", Text2.Text, rstBookMaster.Fields("Code").Value, False) Then
        SSTab1.Tab = 1
        Text2.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(Text3.Text, False) Then
        SSTab1.Tab = 1
        Text3.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(Text8.Text, False) Then
        SSTab1.Tab = 1
        Text8.SetFocus
        CheckMandatoryFields = True
    ElseIf Not CheckExists(Text8, "Col0", rstGroupList, GroupCode) Then
        SSTab1.Tab = 1
        Text8.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(Text9.Text, False) Then
        SSTab1.Tab = 1
        Text9.SetFocus
        CheckMandatoryFields = True
    ElseIf Not CheckExists(Text9, "Col0", rstSizeList, SizeCode) Then
        SSTab1.Tab = 1
        Text9.SetFocus
        CheckMandatoryFields = True
    ElseIf Not CheckExists(Text5, "Col0", rstFinishSizeList, FinishSizeCode) Then
        SSTab1.Tab = 1
        Text5.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(Text7.Text, False) Then
        SSTab1.Tab = 1: Text7.SetFocus: CheckMandatoryFields = True
    Else
        If Not CheckEmpty(Text4.Text, False) Then
            If CheckDuplicate(cnBookMaster, "BookMaster", "Code", "Isbn", Text4.Text, rstBookMaster.Fields("Code").Value, False) Then
                SSTab1.Tab = 1
                Text4.SetFocus
                CheckMandatoryFields = True
            End If
        End If
        If Not CheckEmpty(Text10.Text, False) Then
            If Not CheckExists(Text10, "Col0", rstBindingTypeList, BindingTypeCode) Then
                SSTab1.Tab = 1
                Text10.SetFocus
                CheckMandatoryFields = True
            End If
        End If
        If Not CheckEmpty(Text12.Text, False) Then
            If Not CheckExists(Text12, "Col0", rstOperationList, LaminationTypeCode) Then SSTab1.Tab = 1: Text12.SetFocus: CheckMandatoryFields = True
        End If
        If CheckForms() Then
            SSTab1.Tab = 1
            fpSpread2.SetFocus
            CheckMandatoryFields = True
        End If
        If CheckItem() Then SSTab1.Tab = 2: fpSpread1.SetFocus: CheckMandatoryFields = True: Exit Function
    End If
End Function
Private Sub Timer1_Timer()
    On Error Resume Next
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then
       Timer1.Enabled = False
       ShowProgressInStatusBar False
    End If
End Sub
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Name" Then
        rstBookList.Filter = "[Name] Like '%" & SrchText & "%'"
    End If
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyD Then
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
    If ActiveCellVal = "" Then Cancel = True: Exit Sub
    fpSpread1.GetText 1, Row, Category
    If Col = 1 Then
        fpSpread1.Col = 2
        fpSpread1.TypeComboBoxList = IIf(Category = "BOM", OutsourceItem, FreshBook)
    ElseIf Col = 2 Then
        If Category = "BOM" Then
           If rstOutsourceItemList.RecordCount > 0 Then rstOutsourceItemList.MoveFirst
           rstOutsourceItemList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstOutsourceItemList.EOF Then fpSpread1.SetText 4, Row, rstOutsourceItemList.Fields("NCode").Value
        Else
           If rstFreshBookList.RecordCount > 0 Then rstFreshBookList.MoveFirst
           rstFreshBookList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstFreshBookList.EOF Then fpSpread1.SetText 4, Row, rstFreshBookList.Fields("NCode").Value
        End If
    End If
End Sub
Private Sub fpSpread2_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim ActiveCellVal As Variant, Forms As Variant
    Dim i As Integer, TotalForms As Double
    fpSpread2.GetText Col, Row, ActiveCellVal
    If ActiveCellVal = "" Then Cancel = True: Exit Sub
    If Col = 2 Then
        Combo1_Click
    ElseIf Col = 3 Or Col = 4 Or Col = 5 Or Col = 6 Then   'Calculate Binding Forms
        TotalForms = 0
        For i = 1 To 3
            fpSpread2.GetText 3, i, Forms
            TotalForms = TotalForms + Forms
            fpSpread2.GetText 4, i, Forms
            TotalForms = TotalForms + Forms
            fpSpread2.GetText 5, i, Forms
            If Combo1.ListIndex <= 3 Then
                Forms = Val(Forms) / 2
                Forms = Int(Forms) + IIf(Val(Forms) = Int(Val(Forms)), 0, 1)
            End If
            TotalForms = TotalForms + Forms
            fpSpread2.GetText 6, i, Forms
            TotalForms = TotalForms + Forms
        Next
        MhRealInput4.Text = Format(TotalForms, "0")
    End If
End Sub
Private Function CheckItem() As Boolean
    Dim i As Integer
    Dim Item As Variant, Category As Variant
    CheckItem = False
    For i = 1 To fpSpread1.DataRowCnt
        fpSpread1.SetActiveCell 1, i
        fpSpread1.GetText 1, i, Category
        fpSpread1.GetText 4, i, Item
        If Category = "BOM" Then
            If Left(Item, 1) <> "1" Then CheckItem = True
        Else
            If Left(Item, 1) <> "3" Then CheckItem = True
        End If
        If CheckItem Then DisplayError "Data mismatch in row #" & Trim(Str(i)): Exit For
    Next
End Function
Private Function CheckForms() As Boolean
    Dim i As Integer
    Dim Pages As Variant, Forms¼ As Variant, Forms½ As Variant, Forms1FB As Variant, Forms1WT As Variant, TotalForms As Variant
    
    CheckForms = False
    
    For i = 1 To fpSpread2.DataRowCnt
        fpSpread2.SetActiveCell 1, i
        fpSpread2.GetText 2, i, Pages
        fpSpread2.GetText 7, i, TotalForms
        If Pages / Val(Combo1.Text) <> TotalForms Then
            CheckForms = True
        End If
        If Not CheckForms Then
            fpSpread2.GetText 3, i, Forms¼
            fpSpread2.GetText 4, i, Forms½
            fpSpread2.GetText 5, i, Forms1FB
            fpSpread2.GetText 6, i, Forms1WT
            If Val(Forms¼) * 0.25 + Val(Forms½) * 0.5 + Val(Forms1FB) + Val(Forms1WT) <> TotalForms Then
                CheckForms = True
            End If
        End If
        If CheckForms Then
            DisplayError "Data mismatch in row #" & Trim(Str(i))
            Exit For
        End If
    Next
    If Not CheckForms Then
        TotalForms = 0
        For i = 1 To 3
            fpSpread2.GetText 3, i, Forms¼
            fpSpread2.GetText 4, i, Forms½
            fpSpread2.GetText 5, i, Forms1FB
            If Combo1.ListIndex <= 3 Then
                Forms1FB = Val(Forms1FB) / 2
                Forms1FB = Int(Forms1FB) + IIf(Val(Forms1FB) = Int(Val(Forms1FB)), 0, 1)
            End If
            fpSpread2.GetText 6, i, Forms1WT
            TotalForms = TotalForms + Val(Forms¼) + Val(Forms½) + Val(Forms1FB) + Val(Forms1WT)
        Next
        If Val(MhRealInput4.Text) <> TotalForms Then
            DisplayError "Printing & Binding Forms Mismatch"
            CheckForms = True
        End If
    End If
End Function
Private Sub LoadMaterialList(ByVal strBookCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    If rstBookChild.State = adStateOpen Then rstBookChild.Close
    If DatabaseType = "MS SQL" Then
        rstBookChild.Open "SELECT Category,CASE WHEN Category='1' THEN (SELECT Name FROM OutsourceItemMaster WHERE Code=T.Item) ELSE (SELECT Name FROM BookMaster WHERE Code=T.Item) END AS ItemName,Quantity,Category+Item As ItemCode FROM BookChild01 T WHERE T.Code='" & strBookCode & "' ORDER BY Category", cnBookMaster, adOpenKeyset, adLockReadOnly
    Else
        rstBookChild.Open "SELECT Category,IIF(Category='1',(SELECT Name FROM OutsourceItemMaster WHERE Code=T.Item),(SELECT Name FROM BookMaster WHERE Code=T.Item)) As ItemName,Quantity,Category+Item As ItemCode FROM BookChild01 T WHERE Code='" & strBookCode & "' ORDER BY Category", cnBookMaster, adOpenKeyset, adLockReadOnly
    End If
    rstBookChild.ActiveConnection = Nothing
    If rstBookChild.RecordCount > 0 Then rstBookChild.MoveFirst
    i = 0
    Do While Not rstBookChild.EOF
        i = i + 1
        With fpSpread1
            .SetText 1, i, IIf(rstBookChild.Fields("Category").Value = "1", "BOM", "FG")
            .Col = 2
            .TypeComboBoxList = IIf(rstBookChild.Fields("Category").Value = "1", OutsourceItem, FreshBook)
            .SetText 2, i, rstBookChild.Fields("ItemName").Value
            .SetText 3, i, Val(rstBookChild.Fields("Quantity").Value)
            .SetText 4, i, rstBookChild.Fields("ItemCode").Value
        End With
        rstBookChild.MoveNext
    Loop
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load BOM List")
End Sub
Private Function UpdateMaterialList(ByVal ActionType As String) As Boolean
    Dim CellVal(1 To 3) As Variant
    On Error GoTo ErrorHandler
    UpdateMaterialList = True
    If Left(ActionType, 1) = "D" And (Not blnRecordExist) Then Exit Function
    If ActionType = "D1" Then
        cnBookMaster.Execute "Delete From BookChild01 WHERE Code = '" & rstBookMaster.Fields("Code").Value & "'"
    ElseIf ActionType = "I" Then
        With fpSpread1
            .GetText 1, .ActiveRow, CellVal(1)
            .GetText 3, .ActiveRow, CellVal(2)
            .GetText 4, .ActiveRow, CellVal(3)
        End With
        cnBookMaster.Execute "Insert Into BookChild01 Values ('" & rstBookMaster.Fields("Code").Value & "','" & IIf(CellVal(1) = "BOM", "1", "3") & "','" & Right(CellVal(3), 6) & "'," & Val(CellVal(2)) & ")"
    End If
    Exit Function
ErrorHandler:
    UpdateMaterialList = False
End Function
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Sub fpSpread2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Sub RefreshDropDownList(ByVal xType As String)
    If xType = "R" Then
        rstOutsourceItemList.ActiveConnection = cnBookMaster
        Do While Not RefreshRecord(rstOutsourceItemList): Loop
        rstOutsourceItemList.ActiveConnection = Nothing
        rstFreshBookList.ActiveConnection = cnBookMaster
        Do While Not RefreshRecord(rstFreshBookList): Loop
        rstFreshBookList.ActiveConnection = Nothing
        OutsourceItem = "": FreshBook = ""
    End If
    Do While Not rstOutsourceItemList.EOF
        If OutsourceItem = "" Then OutsourceItem = rstOutsourceItemList.Fields("Name").Value Else OutsourceItem = OutsourceItem + Chr$(9) + rstOutsourceItemList.Fields("Name").Value
        rstOutsourceItemList.MoveNext
    Loop
    Do While Not rstFreshBookList.EOF
        If FreshBook = "" Then FreshBook = rstFreshBookList.Fields("Name").Value Else FreshBook = FreshBook + Chr$(9) + rstFreshBookList.Fields("Name").Value
        rstFreshBookList.MoveNext
    Loop
End Sub
Private Sub DuplicateRecord()
    Dim TmpTbl As String
    TmpTbl = "T" & GetFileNameFromPath(GetTemporaryFileName()): TmpTbl = Left(TmpTbl, InStr(1, TmpTbl, ".", vbTextCompare) - 1)
    On Error GoTo ErrorHandler
    MdiMainMenu.MousePointer = vbHourglass
    Dim BookCode As String, BookName As String
    BookCode = GenerateCode(cnBookMaster, "SELECT MAX(Code) FROM BookMaster", 6, "0")
    BookName = Trim(Left(rstBookList.Fields("Name").Value, IIf(Len(rstBookList.Fields("Name").Value) > 56, 56, 60))) + " (D)"
    cnBookMaster.BeginTrans
    cnBookMaster.Execute "SELECT * INTO [" & TmpTbl & "] FROM BookMaster WHERE Code = '" & rstBookList.Fields("Code").Value & "'"
    cnBookMaster.Execute "UPDATE  [" & TmpTbl & "] SET Code='" & BookCode & "',Name='" & BookName & "',PrintName='" & BookName & "'"
    cnBookMaster.Execute "INSERT INTO BookMaster SELECT * FROM " & TmpTbl
    cnBookMaster.Execute "DROP TABLE " & TmpTbl
    cnBookMaster.Execute "SELECT * INTO [" & TmpTbl & "] FROM BookChild01 WHERE Code = '" & rstBookList.Fields("Code").Value & "'"
    cnBookMaster.Execute "UPDATE  [" & TmpTbl & "] SET Code='" & BookCode & "'"
    cnBookMaster.Execute "INSERT INTO BookChild01 SELECT * FROM " & TmpTbl
    cnBookMaster.Execute "DROP TABLE " & TmpTbl
    cnBookMaster.CommitTrans
    Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    Text1.Text = Trim(BookName): Sendkeys "{END}"
    MdiMainMenu.MousePointer = vbNormal
    Call MsgBox("Successfully Duplicated the Record !", vbInformation, App.Title)
    Exit Sub
ErrorHandler:
    MdiMainMenu.MousePointer = vbNormal
    DisplayError ("Failed to Duplicate the Record")
    cnBookMaster.RollbackTrans
End Sub
Private Sub LoadMasterList()
    If rstGroupList.State = adStateOpen Then rstGroupList.Close
    rstGroupList.Open "SELECT Name As Col0, Code FROM GeneralMaster WHERE Type = '5' ORDER BY Name", cnBookMaster, adOpenKeyset, adLockReadOnly
    rstGroupList.ActiveConnection = Nothing
    If rstSizeList.State = adStateOpen Then rstSizeList.Close
    rstSizeList.Open "SELECT Name As Col0, Code FROM GeneralMaster WHERE Type = '1' ORDER BY Name", cnBookMaster, adOpenKeyset, adLockReadOnly
    rstSizeList.ActiveConnection = Nothing
    If rstFinishSizeList.State = adStateOpen Then rstFinishSizeList.Close
    rstFinishSizeList.Open "SELECT Name As Col0, Code FROM GeneralMaster WHERE Type = '11' ORDER BY Name", cnBookMaster, adOpenKeyset, adLockReadOnly
    rstFinishSizeList.ActiveConnection = Nothing
    If rstBindingTypeList.State = adStateOpen Then rstBindingTypeList.Close
    rstBindingTypeList.Open "SELECT Name As Col0, Code FROM GeneralMaster WHERE Type = '6' ORDER BY Name", cnBookMaster, adOpenKeyset, adLockReadOnly
    rstBindingTypeList.ActiveConnection = Nothing
    If rstOperationList.State = adStateOpen Then rstOperationList.Close
    rstOperationList.Open "SELECT Name As Col0, Code FROM GeneralMaster WHERE Type = '7' ORDER BY Name", cnBookMaster, adOpenKeyset, adLockReadOnly
    rstOperationList.ActiveConnection = Nothing
    If rstHSNCodeList.State = adStateOpen Then rstHSNCodeList.Close
    rstHSNCodeList.Open "SELECT Name As Col0, Code FROM GeneralMaster WHERE Type= '18' ORDER BY Name", cnBookMaster, adOpenKeyset, adLockReadOnly
    rstHSNCodeList.ActiveConnection = Nothing
    If rstOutsourceItemList.State = adStateOpen Then rstOutsourceItemList.Close
    rstOutsourceItemList.Open "SELECT Name,'1'+Code As NCode FROM OutsourceItemMaster ORDER BY Name", cnBookMaster, adOpenKeyset, adLockOptimistic
    rstOutsourceItemList.ActiveConnection = Nothing
    If rstFreshBookList.State = adStateOpen Then rstFreshBookList.Close
    rstFreshBookList.Open "SELECT Name,'3'+Code As NCode FROM BookMaster WHERE Type='F' AND Board='000000' ORDER BY Name", cnBookMaster, adOpenKeyset, adLockOptimistic
    rstFreshBookList.ActiveConnection = Nothing
End Sub
