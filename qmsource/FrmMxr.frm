VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmMxr 
   ClientHeight    =   2655
   ClientLeft      =   510
   ClientTop       =   630
   ClientWidth     =   7215
   ControlBox      =   0   'False
   Icon            =   "FrmMxr.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   177
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   481
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "ThisProgramIsDedicatedToMyCatsBeavisAndButtheadAndTigerKittyAndGrayKittyBecauseTheyDidSomeOfTheTyping"
   Visible         =   0   'False
   Begin VB.CheckBox ChkMute 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   18
      Left            =   3270
      TabIndex        =   121
      TabStop         =   0   'False
      Top             =   2235
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox ChkMute 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   17
      Left            =   3090
      TabIndex        =   120
      TabStop         =   0   'False
      Top             =   2235
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Timer TmrRapid 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   5940
      Top             =   2235
   End
   Begin VB.CheckBox ChkOnOff 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6930
      TabIndex        =   118
      Top             =   495
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Timer TmrAutoMute 
      Interval        =   10000
      Left            =   5535
      Top             =   2235
   End
   Begin VB.Timer TmrRefresh 
      Interval        =   600
      Left            =   5130
      Top             =   2235
   End
   Begin VB.Timer TmrKillMe 
      Enabled         =   0   'False
      Interval        =   333
      Left            =   4725
      Top             =   2235
   End
   Begin VB.PictureBox PicControl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Index           =   1
      Left            =   480
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   113
      TabStop         =   0   'False
      Top             =   15
      Width           =   150
   End
   Begin VB.Timer TmrLightsOut 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   4320
      Top             =   2235
   End
   Begin VB.CheckBox ChkMute 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   16
      Left            =   2910
      TabIndex        =   114
      TabStop         =   0   'False
      Top             =   2235
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox PicControl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Index           =   0
      Left            =   30
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   111
      TabStop         =   0   'False
      Top             =   15
      Width           =   150
   End
   Begin VB.PictureBox PicTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   210
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   112
      TabStop         =   0   'False
      Top             =   15
      Width           =   225
   End
   Begin VB.PictureBox PicGang 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1725
      Index           =   0
      Left            =   30
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   315
      Begin VB.PictureBox PicSpecial 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H0000FFFF&
         Height          =   75
         Left            =   75
         ScaleHeight     =   5
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   10
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   30
         Visible         =   0   'False
         Width           =   150
         Begin VB.Shape Shape1 
            BorderColor     =   &H0000A000&
            Height          =   75
            Left            =   0
            Top             =   0
            Width           =   150
         End
      End
      Begin VB.PictureBox PicVol 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   1140
         Index           =   0
         Left            =   0
         ScaleHeight     =   76
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   330
         Width           =   315
         Begin ComctlLib.Slider SldrVol 
            Height          =   1365
            Index           =   0
            Left            =   -150
            TabIndex        =   2
            Top             =   -105
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   2408
            _Version        =   327682
            Orientation     =   1
            LargeChange     =   6554
            SmallChange     =   655
            Max             =   65535
            SelStart        =   65535
            TickStyle       =   2
            TickFrequency   =   10923
            Value           =   65535
         End
      End
      Begin VB.Shape shpFocus 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         DrawMode        =   11  'Not Xor Pen
         Height          =   315
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   315
      End
      Begin VB.Shape ShpMute 
         BackColor       =   &H0000FF00&
         BorderColor     =   &H0000FF00&
         Height          =   165
         Index           =   0
         Left            =   45
         Shape           =   4  'Rounded Rectangle
         Top             =   1515
         Width           =   270
      End
      Begin VB.Label lblVol 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   0
         Left            =   45
         TabIndex        =   68
         Top             =   1515
         Width           =   270
      End
      Begin VB.Image ImgIcon 
         Height          =   330
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.PictureBox PicGang 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1725
      Index           =   1
      Left            =   390
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   315
      Begin VB.PictureBox PicVol 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   1140
         Index           =   1
         Left            =   0
         ScaleHeight     =   76
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   330
         Width           =   315
         Begin ComctlLib.Slider SldrVol 
            Height          =   1365
            Index           =   1
            Left            =   -150
            TabIndex        =   5
            Top             =   -105
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   2408
            _Version        =   327682
            Orientation     =   1
            LargeChange     =   6554
            SmallChange     =   655
            Max             =   65535
            SelStart        =   65535
            TickStyle       =   2
            TickFrequency   =   10923
            Value           =   65535
         End
      End
      Begin VB.Shape shpFocus 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         DrawMode        =   11  'Not Xor Pen
         Height          =   315
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   315
      End
      Begin VB.Shape ShpMute 
         BackColor       =   &H0000FF00&
         BorderColor     =   &H0000FF00&
         Height          =   165
         Index           =   1
         Left            =   45
         Top             =   1515
         Width           =   270
      End
      Begin VB.Label lblVol 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   1
         Left            =   45
         TabIndex        =   69
         Top             =   1515
         Width           =   270
      End
      Begin VB.Image ImgIcon 
         Height          =   330
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.PictureBox PicGang 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1725
      Index           =   2
      Left            =   750
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   315
      Begin VB.PictureBox PicVol 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   1140
         Index           =   2
         Left            =   0
         ScaleHeight     =   76
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   330
         Width           =   315
         Begin ComctlLib.Slider SldrVol 
            Height          =   1365
            Index           =   2
            Left            =   -150
            TabIndex        =   8
            Top             =   -105
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   2408
            _Version        =   327682
            Orientation     =   1
            LargeChange     =   6554
            SmallChange     =   655
            Max             =   65535
            SelStart        =   65535
            TickStyle       =   2
            TickFrequency   =   10923
            Value           =   65535
         End
      End
      Begin VB.Shape shpFocus 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         DrawMode        =   11  'Not Xor Pen
         Height          =   315
         Index           =   2
         Left            =   0
         Top             =   0
         Width           =   315
      End
      Begin VB.Shape ShpMute 
         BackColor       =   &H0000FF00&
         BorderColor     =   &H0000FF00&
         Height          =   165
         Index           =   2
         Left            =   45
         Top             =   1515
         Width           =   270
      End
      Begin VB.Label lblVol 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   2
         Left            =   45
         TabIndex        =   70
         Top             =   1515
         Width           =   270
      End
      Begin VB.Image ImgIcon 
         Height          =   330
         Index           =   2
         Left            =   0
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.PictureBox PicGang 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1725
      Index           =   3
      Left            =   1110
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   315
      Begin VB.PictureBox PicVol 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   1140
         Index           =   3
         Left            =   0
         ScaleHeight     =   76
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   330
         Width           =   315
         Begin ComctlLib.Slider SldrVol 
            Height          =   1365
            Index           =   3
            Left            =   -150
            TabIndex        =   11
            Top             =   -105
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   2408
            _Version        =   327682
            Orientation     =   1
            LargeChange     =   6554
            SmallChange     =   655
            Max             =   65535
            SelStart        =   65535
            TickStyle       =   2
            TickFrequency   =   10923
            Value           =   65535
         End
      End
      Begin VB.Shape shpFocus 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         DrawMode        =   11  'Not Xor Pen
         Height          =   315
         Index           =   3
         Left            =   0
         Top             =   0
         Width           =   315
      End
      Begin VB.Shape ShpMute 
         BackColor       =   &H0000FF00&
         BorderColor     =   &H0000FF00&
         Height          =   165
         Index           =   3
         Left            =   45
         Top             =   1515
         Width           =   270
      End
      Begin VB.Label lblVol 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   3
         Left            =   45
         TabIndex        =   71
         Top             =   1515
         Width           =   270
      End
      Begin VB.Image ImgIcon 
         Height          =   330
         Index           =   3
         Left            =   0
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.PictureBox PicGang 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1725
      Index           =   4
      Left            =   1470
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   315
      Begin VB.PictureBox PicVol 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   1140
         Index           =   4
         Left            =   0
         ScaleHeight     =   76
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   330
         Width           =   315
         Begin ComctlLib.Slider SldrVol 
            Height          =   1365
            Index           =   4
            Left            =   -150
            TabIndex        =   14
            Top             =   -105
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   2408
            _Version        =   327682
            Orientation     =   1
            LargeChange     =   6554
            SmallChange     =   655
            Max             =   65535
            SelStart        =   65535
            TickStyle       =   2
            TickFrequency   =   10923
            Value           =   65535
         End
      End
      Begin VB.Shape shpFocus 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         DrawMode        =   11  'Not Xor Pen
         Height          =   315
         Index           =   4
         Left            =   0
         Top             =   0
         Width           =   315
      End
      Begin VB.Shape ShpMute 
         BackColor       =   &H0000FF00&
         BorderColor     =   &H0000FF00&
         Height          =   165
         Index           =   4
         Left            =   45
         Top             =   1515
         Width           =   270
      End
      Begin VB.Label lblVol 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   4
         Left            =   45
         TabIndex        =   72
         Top             =   1515
         Width           =   270
      End
      Begin VB.Image ImgIcon 
         Height          =   330
         Index           =   4
         Left            =   0
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.PictureBox PicGang 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1725
      Index           =   5
      Left            =   1830
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   315
      Begin VB.PictureBox PicVol 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   1140
         Index           =   5
         Left            =   0
         ScaleHeight     =   76
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   330
         Width           =   315
         Begin ComctlLib.Slider SldrVol 
            Height          =   1365
            Index           =   5
            Left            =   -150
            TabIndex        =   17
            Top             =   -105
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   2408
            _Version        =   327682
            Orientation     =   1
            LargeChange     =   6554
            SmallChange     =   655
            Max             =   65535
            SelStart        =   65535
            TickStyle       =   2
            TickFrequency   =   10923
            Value           =   65535
         End
      End
      Begin VB.Shape shpFocus 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         DrawMode        =   11  'Not Xor Pen
         Height          =   315
         Index           =   5
         Left            =   0
         Top             =   0
         Width           =   315
      End
      Begin VB.Shape ShpMute 
         BackColor       =   &H0000FF00&
         BorderColor     =   &H0000FF00&
         Height          =   165
         Index           =   5
         Left            =   45
         Top             =   1515
         Width           =   270
      End
      Begin VB.Label lblVol 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   5
         Left            =   45
         TabIndex        =   73
         Top             =   1515
         Width           =   270
      End
      Begin VB.Image ImgIcon 
         Height          =   330
         Index           =   5
         Left            =   0
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.PictureBox PicGang 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1725
      Index           =   6
      Left            =   2190
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   315
      Begin VB.PictureBox PicVol 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   1140
         Index           =   6
         Left            =   0
         ScaleHeight     =   76
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   330
         Width           =   315
         Begin ComctlLib.Slider SldrVol 
            Height          =   1365
            Index           =   6
            Left            =   -150
            TabIndex        =   20
            Top             =   -105
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   2408
            _Version        =   327682
            Orientation     =   1
            LargeChange     =   6554
            SmallChange     =   655
            Max             =   65535
            SelStart        =   65535
            TickStyle       =   2
            TickFrequency   =   10923
            Value           =   65535
         End
      End
      Begin VB.Shape shpFocus 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         DrawMode        =   11  'Not Xor Pen
         Height          =   315
         Index           =   6
         Left            =   0
         Top             =   0
         Width           =   315
      End
      Begin VB.Shape ShpMute 
         BackColor       =   &H0000FF00&
         BorderColor     =   &H0000FF00&
         Height          =   165
         Index           =   6
         Left            =   45
         Top             =   1515
         Width           =   270
      End
      Begin VB.Label lblVol 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   6
         Left            =   45
         TabIndex        =   74
         Top             =   1515
         Width           =   270
      End
      Begin VB.Image ImgIcon 
         Height          =   330
         Index           =   6
         Left            =   0
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.PictureBox PicGang 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1725
      Index           =   7
      Left            =   2550
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   315
      Begin VB.PictureBox PicVol 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   1140
         Index           =   7
         Left            =   0
         ScaleHeight     =   76
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   330
         Width           =   315
         Begin ComctlLib.Slider SldrVol 
            Height          =   1365
            Index           =   7
            Left            =   -150
            TabIndex        =   23
            Top             =   -105
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   2408
            _Version        =   327682
            Orientation     =   1
            LargeChange     =   6554
            SmallChange     =   655
            Max             =   65535
            SelStart        =   65535
            TickStyle       =   2
            TickFrequency   =   10923
            Value           =   65535
         End
      End
      Begin VB.Shape shpFocus 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         DrawMode        =   11  'Not Xor Pen
         Height          =   315
         Index           =   7
         Left            =   0
         Top             =   0
         Width           =   315
      End
      Begin VB.Shape ShpMute 
         BackColor       =   &H0000FF00&
         BorderColor     =   &H0000FF00&
         Height          =   165
         Index           =   7
         Left            =   45
         Top             =   1515
         Width           =   270
      End
      Begin VB.Label lblVol 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   7
         Left            =   45
         TabIndex        =   75
         Top             =   1515
         Width           =   270
      End
      Begin VB.Image ImgIcon 
         Height          =   330
         Index           =   7
         Left            =   0
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.PictureBox PicGang 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1725
      Index           =   8
      Left            =   2910
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   315
      Begin VB.PictureBox PicVol 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   1140
         Index           =   8
         Left            =   0
         ScaleHeight     =   76
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   330
         Width           =   315
         Begin ComctlLib.Slider SldrVol 
            Height          =   1365
            Index           =   8
            Left            =   -150
            TabIndex        =   26
            Top             =   -105
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   2408
            _Version        =   327682
            Orientation     =   1
            LargeChange     =   6554
            SmallChange     =   655
            Max             =   65535
            SelStart        =   65535
            TickStyle       =   2
            TickFrequency   =   10923
            Value           =   65535
         End
      End
      Begin VB.Shape shpFocus 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         DrawMode        =   11  'Not Xor Pen
         Height          =   315
         Index           =   8
         Left            =   0
         Top             =   0
         Width           =   315
      End
      Begin VB.Shape ShpMute 
         BackColor       =   &H0000FF00&
         BorderColor     =   &H0000FF00&
         Height          =   165
         Index           =   8
         Left            =   45
         Top             =   1515
         Width           =   270
      End
      Begin VB.Label lblVol 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   8
         Left            =   45
         TabIndex        =   76
         Top             =   1515
         Width           =   270
      End
      Begin VB.Image ImgIcon 
         Height          =   330
         Index           =   8
         Left            =   0
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.PictureBox PicGang 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1725
      Index           =   9
      Left            =   3270
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   315
      Begin VB.PictureBox PicVol 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   1140
         Index           =   9
         Left            =   0
         ScaleHeight     =   76
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   330
         Width           =   315
         Begin ComctlLib.Slider SldrVol 
            Height          =   1365
            Index           =   9
            Left            =   -150
            TabIndex        =   29
            Top             =   -105
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   2408
            _Version        =   327682
            Orientation     =   1
            LargeChange     =   6554
            SmallChange     =   655
            Max             =   65535
            SelStart        =   65535
            TickStyle       =   2
            TickFrequency   =   10923
            Value           =   65535
         End
      End
      Begin VB.Shape shpFocus 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         DrawMode        =   11  'Not Xor Pen
         Height          =   315
         Index           =   9
         Left            =   0
         Top             =   0
         Width           =   315
      End
      Begin VB.Shape ShpMute 
         BackColor       =   &H0000FF00&
         BorderColor     =   &H0000FF00&
         Height          =   165
         Index           =   9
         Left            =   45
         Top             =   1515
         Width           =   270
      End
      Begin VB.Label lblVol 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   9
         Left            =   45
         TabIndex        =   77
         Top             =   1515
         Width           =   270
      End
      Begin VB.Image ImgIcon 
         Height          =   330
         Index           =   9
         Left            =   0
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.PictureBox PicGang 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1725
      Index           =   10
      Left            =   3630
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   315
      Begin VB.PictureBox PicVol 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   1140
         Index           =   10
         Left            =   0
         ScaleHeight     =   76
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   330
         Width           =   315
         Begin ComctlLib.Slider SldrVol 
            Height          =   1365
            Index           =   10
            Left            =   -150
            TabIndex        =   32
            Top             =   -105
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   2408
            _Version        =   327682
            Orientation     =   1
            LargeChange     =   6554
            SmallChange     =   655
            Max             =   65535
            SelStart        =   65535
            TickStyle       =   2
            TickFrequency   =   10923
            Value           =   65535
         End
      End
      Begin VB.Shape shpFocus 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         DrawMode        =   11  'Not Xor Pen
         Height          =   315
         Index           =   10
         Left            =   0
         Top             =   0
         Width           =   315
      End
      Begin VB.Shape ShpMute 
         BackColor       =   &H0000FF00&
         BorderColor     =   &H0000FF00&
         Height          =   165
         Index           =   10
         Left            =   45
         Top             =   1515
         Width           =   270
      End
      Begin VB.Label lblVol 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   10
         Left            =   45
         TabIndex        =   78
         Top             =   1515
         Width           =   270
      End
      Begin VB.Image ImgIcon 
         Height          =   330
         Index           =   10
         Left            =   0
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.PictureBox PicGang 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1725
      Index           =   11
      Left            =   3990
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   315
      Begin VB.PictureBox PicVol 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   1140
         Index           =   11
         Left            =   0
         ScaleHeight     =   76
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   330
         Width           =   315
         Begin ComctlLib.Slider SldrVol 
            Height          =   1365
            Index           =   11
            Left            =   -150
            TabIndex        =   35
            Top             =   -105
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   2408
            _Version        =   327682
            Orientation     =   1
            LargeChange     =   6554
            SmallChange     =   655
            Max             =   65535
            SelStart        =   65535
            TickStyle       =   2
            TickFrequency   =   10923
            Value           =   65535
         End
      End
      Begin VB.Shape shpFocus 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         DrawMode        =   11  'Not Xor Pen
         Height          =   315
         Index           =   11
         Left            =   0
         Top             =   0
         Width           =   315
      End
      Begin VB.Shape ShpMute 
         BackColor       =   &H0000FF00&
         BorderColor     =   &H0000FF00&
         Height          =   165
         Index           =   11
         Left            =   45
         Top             =   1515
         Width           =   270
      End
      Begin VB.Label lblVol 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   11
         Left            =   45
         TabIndex        =   79
         Top             =   1515
         Width           =   270
      End
      Begin VB.Image ImgIcon 
         Height          =   330
         Index           =   11
         Left            =   0
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.PictureBox PicGang 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1725
      Index           =   12
      Left            =   4350
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   315
      Begin VB.PictureBox PicVol 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   1140
         Index           =   12
         Left            =   0
         ScaleHeight     =   76
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   330
         Width           =   315
         Begin ComctlLib.Slider SldrVol 
            Height          =   1365
            Index           =   12
            Left            =   -150
            TabIndex        =   38
            Top             =   -105
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   2408
            _Version        =   327682
            Orientation     =   1
            LargeChange     =   6554
            SmallChange     =   655
            Max             =   65535
            SelStart        =   65535
            TickStyle       =   2
            TickFrequency   =   10923
            Value           =   65535
         End
      End
      Begin VB.Shape shpFocus 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         DrawMode        =   11  'Not Xor Pen
         Height          =   315
         Index           =   12
         Left            =   0
         Top             =   0
         Width           =   315
      End
      Begin VB.Shape ShpMute 
         BackColor       =   &H0000FF00&
         BorderColor     =   &H0000FF00&
         Height          =   165
         Index           =   12
         Left            =   45
         Top             =   1515
         Width           =   270
      End
      Begin VB.Label lblVol 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   12
         Left            =   45
         TabIndex        =   80
         Top             =   1515
         Width           =   270
      End
      Begin VB.Image ImgIcon 
         Height          =   330
         Index           =   12
         Left            =   0
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.PictureBox PicGang 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1725
      Index           =   13
      Left            =   4710
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   315
      Begin VB.PictureBox PicVol 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   1140
         Index           =   13
         Left            =   0
         ScaleHeight     =   76
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   330
         Width           =   315
         Begin ComctlLib.Slider SldrVol 
            Height          =   1365
            Index           =   13
            Left            =   -150
            TabIndex        =   41
            Top             =   -105
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   2408
            _Version        =   327682
            Orientation     =   1
            LargeChange     =   6554
            SmallChange     =   655
            Max             =   65535
            SelStart        =   65535
            TickStyle       =   2
            TickFrequency   =   10923
            Value           =   65535
         End
      End
      Begin VB.Shape shpFocus 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         DrawMode        =   11  'Not Xor Pen
         Height          =   315
         Index           =   13
         Left            =   0
         Top             =   0
         Width           =   315
      End
      Begin VB.Shape ShpMute 
         BackColor       =   &H0000FF00&
         BorderColor     =   &H0000FF00&
         Height          =   165
         Index           =   13
         Left            =   45
         Top             =   1515
         Width           =   270
      End
      Begin VB.Label lblVol 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   13
         Left            =   45
         TabIndex        =   81
         Top             =   1515
         Width           =   270
      End
      Begin VB.Image ImgIcon 
         Height          =   330
         Index           =   13
         Left            =   0
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.PictureBox PicGang 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1725
      Index           =   14
      Left            =   5070
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   315
      Begin VB.PictureBox PicVol 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   1140
         Index           =   14
         Left            =   0
         ScaleHeight     =   76
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   330
         Width           =   315
         Begin ComctlLib.Slider SldrVol 
            Height          =   1365
            Index           =   14
            Left            =   -150
            TabIndex        =   44
            Top             =   -105
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   2408
            _Version        =   327682
            Orientation     =   1
            LargeChange     =   6554
            SmallChange     =   655
            Max             =   65535
            SelStart        =   65535
            TickStyle       =   2
            TickFrequency   =   10923
            Value           =   65535
         End
      End
      Begin VB.Shape shpFocus 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         DrawMode        =   11  'Not Xor Pen
         Height          =   315
         Index           =   14
         Left            =   0
         Top             =   0
         Width           =   315
      End
      Begin VB.Shape ShpMute 
         BackColor       =   &H0000FF00&
         BorderColor     =   &H0000FF00&
         Height          =   165
         Index           =   14
         Left            =   45
         Top             =   1515
         Width           =   270
      End
      Begin VB.Label lblVol 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   14
         Left            =   45
         TabIndex        =   82
         Top             =   1515
         Width           =   270
      End
      Begin VB.Image ImgIcon 
         Height          =   330
         Index           =   14
         Left            =   0
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.PictureBox PicGang 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1725
      Index           =   15
      Left            =   5430
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   315
      Begin VB.PictureBox PicVol 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   1140
         Index           =   15
         Left            =   0
         ScaleHeight     =   76
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   330
         Width           =   315
         Begin ComctlLib.Slider SldrVol 
            Height          =   1365
            Index           =   15
            Left            =   -150
            TabIndex        =   47
            Top             =   -105
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   2408
            _Version        =   327682
            Orientation     =   1
            LargeChange     =   6554
            SmallChange     =   655
            Max             =   65535
            SelStart        =   65535
            TickStyle       =   2
            TickFrequency   =   10923
            Value           =   65535
         End
      End
      Begin VB.Shape shpFocus 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         DrawMode        =   11  'Not Xor Pen
         Height          =   315
         Index           =   15
         Left            =   0
         Top             =   0
         Width           =   315
      End
      Begin VB.Shape ShpMute 
         BackColor       =   &H0000FF00&
         BorderColor     =   &H0000FF00&
         Height          =   165
         Index           =   15
         Left            =   45
         Top             =   1515
         Width           =   270
      End
      Begin VB.Label lblVol 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   15
         Left            =   45
         TabIndex        =   83
         Top             =   1515
         Width           =   270
      End
      Begin VB.Image ImgIcon 
         Height          =   330
         Index           =   15
         Left            =   0
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.PictureBox PicGang 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1725
      Index           =   16
      Left            =   5790
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   315
      Begin VB.PictureBox PicVol 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   1140
         Index           =   16
         Left            =   0
         ScaleHeight     =   76
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   330
         Width           =   315
         Begin ComctlLib.Slider SldrVol 
            Height          =   1365
            Index           =   16
            Left            =   -150
            TabIndex        =   50
            Top             =   -105
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   2408
            _Version        =   327682
            Orientation     =   1
            LargeChange     =   6554
            SmallChange     =   655
            Max             =   65535
            SelStart        =   65535
            TickStyle       =   2
            TickFrequency   =   10923
            Value           =   65535
         End
      End
      Begin VB.Shape shpFocus 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         DrawMode        =   11  'Not Xor Pen
         Height          =   315
         Index           =   16
         Left            =   0
         Top             =   0
         Width           =   315
      End
      Begin VB.Shape ShpMute 
         BackColor       =   &H0000FF00&
         BorderColor     =   &H0000FF00&
         Height          =   165
         Index           =   16
         Left            =   45
         Top             =   1515
         Width           =   270
      End
      Begin VB.Label lblVol 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   16
         Left            =   45
         TabIndex        =   84
         Top             =   1515
         Width           =   270
      End
      Begin VB.Image ImgIcon 
         Height          =   330
         Index           =   16
         Left            =   0
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.PictureBox PicGang 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1725
      Index           =   17
      Left            =   6150
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   101
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   315
      Begin VB.PictureBox PicVol 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   1140
         Index           =   17
         Left            =   0
         ScaleHeight     =   76
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   330
         Width           =   315
         Begin ComctlLib.Slider SldrVol 
            Height          =   1365
            Index           =   17
            Left            =   -150
            TabIndex        =   103
            Top             =   -105
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   2408
            _Version        =   327682
            Orientation     =   1
            LargeChange     =   6554
            SmallChange     =   655
            Max             =   65535
            SelStart        =   65535
            TickStyle       =   2
            TickFrequency   =   10923
            Value           =   65535
         End
      End
      Begin VB.Shape ShpMute 
         BackColor       =   &H0000FF00&
         BorderColor     =   &H0000FF00&
         Height          =   165
         Index           =   17
         Left            =   45
         Top             =   1515
         Width           =   270
      End
      Begin VB.Image ImgIcon 
         Height          =   330
         Index           =   17
         Left            =   0
         Top             =   0
         Width           =   315
      End
      Begin VB.Label lblVol 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   17
         Left            =   45
         TabIndex        =   104
         Top             =   1515
         Width           =   270
      End
      Begin VB.Shape shpFocus 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         DrawMode        =   11  'Not Xor Pen
         Height          =   315
         Index           =   17
         Left            =   0
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.PictureBox PicGang 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1725
      Index           =   18
      Left            =   6510
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   315
      Begin VB.PictureBox PicVol 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   1140
         Index           =   18
         Left            =   0
         ScaleHeight     =   76
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   106
         TabStop         =   0   'False
         Top             =   330
         Width           =   315
         Begin ComctlLib.Slider SldrVol 
            Height          =   1365
            Index           =   18
            Left            =   -150
            TabIndex        =   107
            Top             =   -105
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   2408
            _Version        =   327682
            Orientation     =   1
            LargeChange     =   6554
            SmallChange     =   655
            Max             =   65535
            SelStart        =   65535
            TickStyle       =   2
            TickFrequency   =   10923
            Value           =   65535
         End
      End
      Begin VB.Shape ShpMute 
         BackColor       =   &H0000FF00&
         BorderColor     =   &H0000FF00&
         Height          =   165
         Index           =   18
         Left            =   45
         Top             =   1515
         Width           =   270
      End
      Begin VB.Image ImgIcon 
         Height          =   330
         Index           =   18
         Left            =   0
         Top             =   0
         Width           =   315
      End
      Begin VB.Label lblVol 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   18
         Left            =   45
         TabIndex        =   108
         Top             =   1515
         Width           =   270
      End
      Begin VB.Shape shpFocus 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         DrawMode        =   11  'Not Xor Pen
         Height          =   315
         Index           =   18
         Left            =   0
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.CheckBox ChkMute 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   15
      Left            =   2730
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   2235
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox ChkMute 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   14
      Left            =   2550
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   2235
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox ChkMute 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   13
      Left            =   2370
      TabIndex        =   98
      TabStop         =   0   'False
      Top             =   2235
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox ChkMute 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   12
      Left            =   2190
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   2235
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox ChkMute 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   11
      Left            =   2010
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   2235
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox ChkMute 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   10
      Left            =   1830
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   2235
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox ChkMute 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   1650
      TabIndex        =   94
      TabStop         =   0   'False
      Top             =   2235
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox ChkMute 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   1470
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   2235
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox ChkMute 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   1290
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   2235
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox ChkMute 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   1110
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   2235
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox ChkMute 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   930
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   2235
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox ChkMute 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   750
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   2235
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox ChkMute 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   570
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   2235
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox ChkMute 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   390
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   2235
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox ChkMute 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   210
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   2235
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox ChkMute 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   30
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   2235
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Timer TmrRepeat 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3915
      Top             =   2235
   End
   Begin VB.Timer TmrRepeatDelay 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3510
      Top             =   2235
   End
   Begin VB.PictureBox PicSkin 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1530
      Left            =   6420
      Picture         =   "FrmMxr.frx":030A
      ScaleHeight     =   102
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   351
      TabIndex        =   115
      TabStop         =   0   'False
      Top             =   1860
      Visible         =   0   'False
      Width           =   5265
   End
   Begin VB.Image ImgIconRes 
      Height          =   330
      Index           =   16
      Left            =   5820
      Picture         =   "FrmMxr.frx":1A80C
      Top             =   1860
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Selector 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   7035
      TabIndex        =   117
      Top             =   195
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Selector 
      BackColor       =   &H80000016&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   6870
      TabIndex        =   116
      Top             =   195
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image TrayIconRes 
      Height          =   240
      Index           =   1
      Left            =   6915
      Picture         =   "FrmMxr.frx":1A996
      Stretch         =   -1  'True
      Top             =   1110
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image TrayIconRes 
      Height          =   240
      Index           =   0
      Left            =   6915
      Picture         =   "FrmMxr.frx":1ACA0
      Stretch         =   -1  'True
      Top             =   780
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Shape shpMuteKey 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF80FF&
      DrawMode        =   3  'Not Merge Pen
      Height          =   165
      Index           =   0
      Left            =   6900
      Shape           =   4  'Rounded Rectangle
      Top             =   1650
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Shape shpMuteKey 
      BackColor       =   &H00404000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFF00&
      DrawMode        =   3  'Not Merge Pen
      Height          =   165
      Index           =   1
      Left            =   6900
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label LblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   18
      Left            =   3315
      LinkTimeout     =   0
      TabIndex        =   110
      Top             =   2475
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label LblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   17
      Left            =   3135
      LinkTimeout     =   0
      TabIndex        =   109
      Top             =   2475
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image ImgIconRes 
      Height          =   330
      Index           =   15
      Left            =   5460
      Picture         =   "FrmMxr.frx":1AFAA
      Top             =   1860
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image ImgIconRes 
      Height          =   330
      Index           =   14
      Left            =   5100
      Picture         =   "FrmMxr.frx":1B134
      Top             =   1860
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image ImgIconRes 
      Height          =   330
      Index           =   13
      Left            =   4740
      Picture         =   "FrmMxr.frx":1B6F6
      Top             =   1860
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image ImgIconRes 
      Height          =   330
      Index           =   12
      Left            =   4380
      Picture         =   "FrmMxr.frx":1BCB8
      Stretch         =   -1  'True
      Top             =   1860
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image ImgIconRes 
      Height          =   330
      Index           =   11
      Left            =   4020
      Picture         =   "FrmMxr.frx":1BE42
      Top             =   1860
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image ImgIconRes 
      Height          =   330
      Index           =   10
      Left            =   3660
      Picture         =   "FrmMxr.frx":1BFCC
      Top             =   1860
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image ImgIconRes 
      Height          =   330
      Index           =   9
      Left            =   3300
      Picture         =   "FrmMxr.frx":1C156
      Top             =   1860
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image ImgIconRes 
      Height          =   330
      Index           =   8
      Left            =   2940
      Picture         =   "FrmMxr.frx":1C2E0
      Top             =   1860
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image ImgIconRes 
      Height          =   330
      Index           =   7
      Left            =   2580
      Picture         =   "FrmMxr.frx":1C46A
      Top             =   1860
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image ImgIconRes 
      Height          =   330
      Index           =   6
      Left            =   2220
      Picture         =   "FrmMxr.frx":1C5F4
      Top             =   1860
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image ImgIconRes 
      Height          =   330
      Index           =   5
      Left            =   1860
      Picture         =   "FrmMxr.frx":1C77E
      Top             =   1860
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image ImgIconRes 
      Height          =   330
      Index           =   4
      Left            =   1500
      Picture         =   "FrmMxr.frx":1C908
      Top             =   1860
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image ImgIconRes 
      Height          =   330
      Index           =   3
      Left            =   1140
      Picture         =   "FrmMxr.frx":1CA92
      Top             =   1860
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image ImgIconRes 
      Height          =   330
      Index           =   2
      Left            =   780
      Picture         =   "FrmMxr.frx":1CC1C
      Top             =   1860
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image ImgIconRes 
      Height          =   330
      Index           =   1
      Left            =   420
      Picture         =   "FrmMxr.frx":1CDA6
      Top             =   1860
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label LblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   75
      LinkTimeout     =   0
      TabIndex        =   67
      Top             =   2475
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label LblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   255
      LinkTimeout     =   0
      TabIndex        =   66
      Top             =   2475
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label LblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   435
      LinkTimeout     =   0
      TabIndex        =   65
      Top             =   2475
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label LblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   615
      LinkTimeout     =   0
      TabIndex        =   64
      Top             =   2475
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label LblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   795
      LinkTimeout     =   0
      TabIndex        =   63
      Top             =   2475
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label LblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   975
      LinkTimeout     =   0
      TabIndex        =   62
      Top             =   2475
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label LblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   1155
      LinkTimeout     =   0
      TabIndex        =   61
      Top             =   2475
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label LblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   1335
      LinkTimeout     =   0
      TabIndex        =   60
      Top             =   2475
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label LblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   8
      Left            =   1515
      LinkTimeout     =   0
      TabIndex        =   59
      Top             =   2475
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label LblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   1695
      LinkTimeout     =   0
      TabIndex        =   58
      Top             =   2475
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label LblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   10
      Left            =   1875
      LinkTimeout     =   0
      TabIndex        =   57
      Top             =   2475
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label LblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   11
      Left            =   2055
      LinkTimeout     =   0
      TabIndex        =   56
      Top             =   2475
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label LblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   12
      Left            =   2235
      LinkTimeout     =   0
      TabIndex        =   55
      Top             =   2475
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label LblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   13
      Left            =   2415
      LinkTimeout     =   0
      TabIndex        =   54
      Top             =   2475
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label LblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   14
      Left            =   2595
      LinkTimeout     =   0
      TabIndex        =   53
      Top             =   2475
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label LblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   15
      Left            =   2775
      LinkTimeout     =   0
      TabIndex        =   52
      Top             =   2475
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label LblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   16
      Left            =   2955
      LinkTimeout     =   0
      TabIndex        =   51
      Top             =   2475
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image ImgIconRes 
      Height          =   330
      Index           =   0
      Left            =   60
      Picture         =   "FrmMxr.frx":1CF30
      Top             =   1860
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Menu MnuPopUp 
      Caption         =   "Pop-Up"
      Visible         =   0   'False
      Begin VB.Menu MnuAboutQuickMixer 
         Caption         =   "&About Quick Mixer"
         Begin VB.Menu MnuHomepage 
            Caption         =   "&Quick Mixer Home-Page..."
         End
         Begin VB.Menu MnuEmail 
            Caption         =   "Send &Feedback..."
         End
         Begin VB.Menu MnuHelp 
            Caption         =   "&Help/About Quick Mixer..."
            Visible         =   0   'False
         End
         Begin VB.Menu MnuControlKeysInfo 
            Caption         =   "Control &Keys Info..."
         End
         Begin VB.Menu MnuResetDefault 
            Caption         =   "Reset to &Default Condition..."
            Visible         =   0   'False
         End
         Begin VB.Menu MnuAboutSep1 
            Caption         =   "-"
         End
         Begin VB.Menu MnuProduct 
            Caption         =   "&Device is"
         End
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSettings 
         Caption         =   "&General Settings"
         Begin VB.Menu MnuGSToolWin 
            Caption         =   "&Multi-Select..."
            Shortcut        =   {F6}
         End
         Begin VB.Menu sep7 
            Caption         =   "-"
         End
         Begin VB.Menu MnuAutoHide 
            Caption         =   "&Auto-Hide"
         End
         Begin VB.Menu MnuAlwaysOnTop 
            Caption         =   "Always &On-Top"
         End
         Begin VB.Menu MnuSnapSlidersEvenly 
            Caption         =   "Space Sliders &Evenly"
         End
         Begin VB.Menu MnuSingleSliderMode 
            Caption         =   "Single Slide&r Mode"
         End
         Begin VB.Menu MnuPointedSliders 
            Caption         =   "&Pointed Sliders"
         End
         Begin VB.Menu MnuShowGraduations 
            Caption         =   "Show &Graduations"
         End
         Begin VB.Menu MnuShowSkin 
            Caption         =   "Show S&kin"
         End
         Begin VB.Menu MnuShowToolTips 
            Caption         =   "Show &Tool-Tips"
         End
         Begin VB.Menu MnuShowTB 
            Caption         =   "Show Treb/&Bass Sliders"
            Visible         =   0   'False
         End
         Begin VB.Menu MnuReverseLogicTrebBass 
            Caption         =   "Re&verse Treb/Bass Logic"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu MnuTimedMute 
         Caption         =   "&Timed Mute..."
         Shortcut        =   {F7}
      End
      Begin VB.Menu MnuBackColor 
         Caption         =   "&Background Color..."
         Shortcut        =   {F8}
      End
      Begin VB.Menu MnuAudioProfiles 
         Caption         =   "&Audio Profiles..."
         Shortcut        =   {F4}
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuLaunch 
         Caption         =   "&Launch"
         Begin VB.Menu MnuWindowsVolumeControl 
            Caption         =   "&Windows Volume Control..."
            Visible         =   0   'False
         End
         Begin VB.Menu MnuMultimediaProperties 
            Caption         =   "&Multimedia Properties..."
         End
         Begin VB.Menu MnuSoundsProps 
            Caption         =   "S&ounds Properties..."
         End
         Begin VB.Menu MnuWindowsSoundRecorder 
            Caption         =   "Windows &Sound Recorder..."
            Visible         =   0   'False
         End
         Begin VB.Menu MnuWindowsCDPlayer 
            Caption         =   "Windows &CD Player..."
            Visible         =   0   'False
         End
         Begin VB.Menu MnuWindowsMediaPlayer 
            Caption         =   "Windows Media &Player..."
            Visible         =   0   'False
         End
         Begin VB.Menu MnuWinAmp 
            Caption         =   "Win&Amp..."
            Visible         =   0   'False
         End
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRestoreMixer 
         Caption         =   "&Show Mixer"
      End
      Begin VB.Menu MnuHideMixer 
         Caption         =   "&Hide Mixer"
      End
      Begin VB.Menu MnuSizeToWinAmp 
         Caption         =   "Size to Wi&nAmp"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuUnhideSliders 
         Caption         =   "."
         Visible         =   0   'False
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOnOff 
         Caption         =   "Advanced Control #&1"
         Shortcut        =   {F12}
         Visible         =   0   'False
      End
      Begin VB.Menu MnuMute 
         Caption         =   "&Mute"
         Shortcut        =   {F11}
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExitAndCloseQuickMixer 
         Caption         =   "&Exit && Close Quick Mixer"
      End
   End
End
Attribute VB_Name = "FrmMxr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any _
) As Long

Private Declare Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long _
) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long _
) As Long

'Note: ALL constants here that represent width or height are PIXEL amounts!
Private Const HTCAPTION = 2, WM_NCLBUTTONDOWN = &HA1, WM_SYSCOMMAND = &H112
Private Const SWP_NOMOVE = &H2, SWP_NOSIZE = &H1, HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2, SC_MOVE = &HF010&, SC_RESTORE = &HF120&
Private Const SC_SIZE = &HF000&, SW_SHOW = 5, LANG_FILE = "\Language.txt"
Private Const ZERO = 0, ONE = 1, TWO = 2, THREE = 3, CENT = 100, DECA = 10, TOTALGRADS = 11
Private Const RAPID_UPDATE_RATE = 20, NORMAL_UPDATE_RATE = 250, SLOW_UPDATE_RATE = 3000 'milliseconds ( 1/1000 seconds )
Private Const PC1LEFTOFFORMWIDTH = 12, MINTITLEBARFORMWIDTH = 28, SKIN = "\skin.bmp"
Private Const PCTOP = 1, HIDECTRL = -25, PCLEFTCUTOFF = 13, ONEPERCENT = 655, SLIDERMAX = 65535
Private Const PCBRIGHT = &H80000016, PCDARK = &H80000010, COLORFLAG = &H1&
Private Const ON_COLOR = vbGreen, OFF_COLOR = &HC000&, MINSIZE = 43, BACKWARDS = -1
Private Const HALFGANGWIDTH = 10.5, MINSLIDERHEIGHT = 12, MUTEWIDTH = 21, MUTELEFT = 0
Private Const VOLCONTAINERSHORTOFFORMHEIGHT = 44, VOLCONTROLSHORTOFFORMHEIGHT = 29
Private Const MUTESHORTOFFORMHEIGHT = 20, CENTERINGFACTOR = 2, BIGNUMBER = 10000
Private Const ICONMAIN = 0, ICONWAVE = 1, ICONMIDI = 2, ICONCD = 3, ICONLINE = 4, ICONMIC = 5
Private Const ICONSPKR = 6, ICONVOICE = 7, ICONDICT = 8, ICONSPEECH = 9, ICONAUX = 10, ICONTELEPHONE = 16
Private Const ICONMODEM = 11, ICONVIDEO = 12, ICONTREBLE = 13, ICONBASS = 14
Private Const ICONUNKNOWN = 15, SLIDERLARGECHANGE = 6553, GANGWIDTH = 21, GRADMARGIN = 5
Private Const GRADKNOBADJUST = 13, PICVOLTOP = 22, WINAMPWIDTH = 275, WINAMPHEIGHT = 116
Private Const LEFTSELECTOROFFSET = -22, RIGHTSELECTOROFFSET = 11, BUTTONFACECOLOR = &H8000000F
Private Const PATHTOSNDVOL32 = "c:\windows\sndvol32.exe", PATHTOWINAMP = "c:\program files\winamp\winamp.exe"
Private Const PATHTOSNDREC32 = "c:\windows\sndrec32.exe", PATHTOCDPLAYER = "c:\windows\cdplayer.exe"
Private Const PATHTOMEDIAPLAYER2 = "c:\Program Files\Microsoft Media Player\mplayer2.exe"
Private Const HOMEPAGE = "http://www.geocities.com/quickaudiomixer/index.html", PLURAL = "s"
Private Const PATHTONOTEPAD = "c:\windows\notepad.exe", HELPFILENAME = "\QMixer.txt"
Private Const PATHTOWORDPAD = "c:\program files\accessories\wordpad.exe", DOCFILENAME = "\QMixer.doc"
Private Const INIMESSAGE = "This is the Quick Mixer initialization file. It can safely be deleted anytime."
'                   WARNING!!! -- INIMESSAGE CANNOT CONTAIN COMMAS OR OTHER LINE SEPERATORS!!!

Private ButtonFlag%, IndexFlag%, bw%, IconClickFlag As Boolean, VisibleControls%, NukeIni As Boolean
Private LangMuteWord$, LangReadoutWord$, LangMixerMenuWord$, LangHideMixerWord$, LangMoveWord$
Private LangAdvCtrlWord$, LangQuickMixerWord$

Private Sub Form_Activate()
    'We only ever want to do this ONCE when the mixer first starts running, otherwise we would be
    'doing this again each time we moved to a sister-form of frmmxer and then returned, which would
    'cause a stack-overflow with the tray-icon code, and a rather spectacular crash.
    Static Already_Did_This As Boolean
    If Already_Did_This Then Exit Sub
    Already_Did_This = True
    'activate active-focus window-following
    OldWindowProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf NewWindowProc)
    'puts the form's icon in the tray, makes it red if main-volume is muted
    AddToTray Me, MnuPopUp
    'see if some external apps and/or an external skin and/or language and/or helpfile/s exist
    If FileExists(PATHTOSNDVOL32) Then MnuWindowsVolumeControl.Visible = True
    If FileExists(PATHTOWINAMP) Then MnuWinAmp.Visible = True: MnuSizeToWinAmp.Visible = True
    If FileExists(App.Path & SKIN) Then Set PicSkin.Picture = LoadPicture(App.Path & SKIN)
    If FileExists(PATHTOSNDREC32) Then MnuWindowsSoundRecorder.Visible = True
    If FileExists(PATHTOCDPLAYER) Then MnuWindowsCDPlayer.Visible = True
    If FileExists(PATHTOMEDIAPLAYER2) Then MnuWindowsMediaPlayer.Visible = True
    If FileExists(PATHTONOTEPAD) And FileExists(App.Path & HELPFILENAME) Then MnuHelp.Visible = True
    If FileExists(PATHTOWORDPAD) And FileExists(App.Path & DOCFILENAME) Then MnuHelp.Visible = True
    If FileExists(App.Path & "\qmixer.html") Then MnuHelp.Visible = True
    If FileExists(INI) Then MnuResetDefault.Visible = True
    'ensure correct menu options are visible, enabled, and checked.
    'also perfoms some other tasks relating to initial setup of various settings.
    If MixerVisible Then MnuRestoreMixer.Enabled = False Else MnuHideMixer.Enabled = False
    If ShowGrads Then Me.AutoRedraw = True: MnuShowGraduations.Checked = vbChecked: ShowSkin = False
    If SnapSliders Then MnuSnapSlidersEvenly.Checked = vbChecked
    If AutoHide Then MnuAutoHide.Checked = vbChecked
    If OnTop Then
        MnuAlwaysOnTop.Checked = vbChecked
        SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    End If
    If ShowSkin Then MnuShowSkin.Checked = vbChecked: ShowGrads = False
    If PointedSliders Then MnuPointedSliders.Checked = vbChecked
    If SingleSliderMode Then MnuSingleSliderMode.Checked = vbChecked: SSMode
    If ReverseLogicTrebBass Then MnuReverseLogicTrebBass.Checked = vbChecked
    If ShowTB Then MnuShowTB.Checked = vbChecked
    If TBSupport Then MnuShowTB.Visible = True: MnuReverseLogicTrebBass.Visible = True
    If TMuteFlag Or TUnMuteFlag Then MnuTimedMute.Checked = vbChecked
    If ChkOnOff.Enabled = True Then MnuOnOff.Visible = True: PicSpecial.Visible = True
    'set private(form) and public(tool window) text to English by default
    LangMuteWord = "Mute": LangReadoutWord = "Readout": LangMixerMenuWord = "Mixer Menu"
    LangHideMixerWord = "Hide Mixer": LangMoveWord = "Move": LangAdvCtrlWord = "Advanced Control #1"
    LangQuickMixerWord = "Quick Mixer": LangActivatedWord = "Activated!": LangDeactivatedWord = "Deactivated"
    LangRedWord = "Red": LangGreenWord = "Green": LangBlueWord = "Blue": LangOkWord = "OK"
    LangTimedMuteWord = "Timed Mute": LangTimedUnmuteWord = "Timed Unmute": LangBackcolorWord = "Back Color"
    LangDefaultWord = "&Default": LangCustomWord = "&Custom"
    'if language.txt exists, import it and 'ding'
    If FileExists(App.Path & LANG_FILE) Then Import_Language: Beep
    If ToolTips Then MnuShowToolTips.Checked = vbChecked: Tips True
    MnuAboutQuickMixer.Caption = MnuAboutQuickMixer.Caption & " " & App.Major & "." & App.Minor & "." & App.Revision
    MnuProduct.Caption = MnuProduct.Caption & " " & ProductName & "..."
    'further modifications
    Me.BackColor = MixerBackColor
    If Me.BackColor <> BUTTONFACECOLOR Then MnuBackColor.Checked = vbChecked
    Dim k%
    'put some things that may have wandered on the form a bit in their proper places
    For k = ZERO To PicGang.UBound
        ShpMute(k).Left = MUTELEFT
        ShpMute(k).Width = MUTEWIDTH
        lblVol(k).Left = MUTELEFT - ONE
        lblVol(k).Width = MUTEWIDTH + ONE
        PicGang(k).Height = BIGNUMBER
        SldrVol(k).Max = SLIDERMAX
        SldrVol(k).LargeChange = SLIDERLARGECHANGE
        SldrVol(k).SmallChange = ONEPERCENT
        If PointedSliders Then SldrVol(k).TickStyle = sldTopLeft
    Next k
    For k = ZERO To MaxSources + TWO
        'set the sliders to the source settings
        SldrVol(k).Value = MixerState(k).MxrVol
        'set the mutes likewise
        If MixerState(k).MxrMute <> ZERO Then
            ChkMute(k).Value = vbChecked
            With ShpMute(k)
                .BackColor = shpMuteKey(ONE).BackColor
                .BorderColor = shpMuteKey(ONE).BorderColor
                .BackStyle = shpMuteKey(ONE).BackStyle
                .DrawMode = shpMuteKey(ONE).DrawMode
                .Shape = shpMuteKey(ONE).Shape
            End With
        Else
            ChkMute(k).Value = vbUnchecked
            With ShpMute(k)
                .BackColor = shpMuteKey(ZERO).BackColor
                .BorderColor = shpMuteKey(ZERO).BorderColor
                .BackStyle = shpMuteKey(ZERO).BackStyle
                .DrawMode = shpMuteKey(ZERO).DrawMode
                .Shape = shpMuteKey(ZERO).Shape
            End With
        End If
        'Note: The first PicGang container (zero) is for the main volume, it's a DESTINATION, not a SOURCE!
        'show the PicGang container of the current source
        If SldrVol(k).Enabled = True Then PicGang(k).Visible = True: VisibleControls = k 'if a slider works, make it visible.
        lblVol(k).Caption = CENT - Int(SldrVol(k).Value / ONEPERCENT) 'update the volume percentage labels.
        Call PickIcon(k) ' Tries to give each slider the right icon based on the name supplied by the mixer-line.
    Next k
    Call Form_Resize 'size everything to the form Somewhat messy the first time around, so the mixer is off-screen.
    'set focus if possible, the mixer is visible but off screen...
    If SldrVol(CurrentFocus).Enabled = True And Me.Visible = True Then SldrVol(CurrentFocus).SetFocus
    'if the mixer is to start with the deafault audio-profile, this will load it.
    Dim i%, p$
    If StartWithProfile Then
        MnuAudioProfiles.Checked = True
        i = InStr(Profile(ZERO), "•")
        p = Right(Profile(ZERO), Len(Profile(ZERO)) - i)
        For k = ZERO To MaxSources + 2
            SldrVol(k).Value = SLIDERMAX - (ONEPERCENT * Val(Mid(p, (k * 6) + 2, 3)))
            If Val(Mid(p, (k * 6) + 6, 1)) = 1 Then ChkMute(k).Value = vbChecked Else ChkMute(k).Value = vbUnchecked
        Next k
    End If
    'If the mixer opens while the main volume is muted or not, this puts the correct icon, red or black, in the tray.
    Call Update_Tray_Icon
    'if the mixer was closed while hidden, then open it hidden.
    If MixerVisible = False Then Me.Hide
    'if the mixer was in single slider mode when closed, place it in single slider mode and restore and non-standard width.
    If ISSM Then Call MnuSingleSliderMode_Click: Me.Width = IssmWidth
    'if the mixer was last visible, then let's see it. Remember it is now visible, but off screen.
    'this saves everyone from seeing a very chaotic initial resize of all the controls that really thrashes the form.
    Me.Left = Me.Left - HACKHIDE
    If Not DontShowInfo Then FrmInfo.Show 'First-time help for special mouse-clicks used in the program.
End Sub

Private Sub ReInsert()
'this allows us to re-enter from a sister-form like timed-mute-un-mute and only do the bits we need to do.
'it's much like the form_activate routine, only it does not do the stuff that doing again would be bad.
'this is also done if the number of displayed sliders needs to change, as in the show-treble/bass-sliders option.
Dim k%
    For k = ZERO To MaxSources + TWO
        'set the sliders to the source settings
        SldrVol(k).Value = MixerState(k).MxrVol
        'set the mutes likewise
        If MixerState(k).MxrMute <> ZERO Then
            ChkMute(k).Value = vbChecked
            With ShpMute(k)
                .BackColor = shpMuteKey(ONE).BackColor
                .BorderColor = shpMuteKey(ONE).BorderColor
                .BackStyle = shpMuteKey(ONE).BackStyle
                .DrawMode = shpMuteKey(ONE).DrawMode
                .Shape = shpMuteKey(ONE).Shape
            End With
        Else
            ChkMute(k).Value = vbUnchecked
            With ShpMute(k)
                .BackColor = shpMuteKey(ZERO).BackColor
                .BorderColor = shpMuteKey(ZERO).BorderColor
                .BackStyle = shpMuteKey(ZERO).BackStyle
                .DrawMode = shpMuteKey(ZERO).DrawMode
                .Shape = shpMuteKey(ZERO).Shape
            End With
        End If
        'show the PicGang container of the current source
        'Note: The first PicGang container is for the main volume, it's a DESTINATION, not a SOURCE!
        If SldrVol(k).Enabled = True Then PicGang(k).Visible = True: VisibleControls = k Else PicGang(k).Visible = False
        lblVol(k).Caption = CENT - Int(SldrVol(k).Value / ONEPERCENT)
        'Call PickIcon(k)
    Next k
    Call Form_Resize 'size everything to the form
    If SldrVol(CurrentFocus).Enabled = True And SldrVol(CurrentFocus).Visible And Me.Visible = True Then SldrVol(CurrentFocus).SetFocus
End Sub

Private Sub UpdateMixer()
    'This keeps the sliders and mutes updated if the user changes their system-wide status in another application.
    Dim k%, m
    Static Oldm
    m = MixerState(ZERO).MxrMute
    If m = Oldm Then GoTo SkipTrayMuteUpdate
    If m <> ZERO Then MnuMute.Checked = vbChecked Else MnuMute.Checked = vbUnchecked
    Update_Tray_Icon
    Oldm = m
SkipTrayMuteUpdate:
    If MixerState(ZERO).MxrOnOff <> ZERO Then
        ChkOnOff.Value = vbChecked
        MnuOnOff.Checked = vbChecked
        PicSpecial.BackColor = ON_COLOR
    Else
        ChkOnOff.Value = vbUnchecked
        MnuOnOff.Checked = vbUnchecked
        PicSpecial.BackColor = OFF_COLOR
    End If
    For k = ZERO To VisibleControls
        SldrVol(k).Value = MixerState(k).MxrVol
        If MixerState(k).MxrMute <> ZERO Then
            ChkMute(k).Value = vbChecked
            With ShpMute(k)
                .BackColor = shpMuteKey(ONE).BackColor
                .BorderColor = shpMuteKey(ONE).BorderColor
                .BackStyle = shpMuteKey(ONE).BackStyle
                .DrawMode = shpMuteKey(ONE).DrawMode
                .Shape = shpMuteKey(ONE).Shape
            End With
        Else
            ChkMute(k).Value = vbUnchecked
            With ShpMute(k)
                .BackColor = shpMuteKey(ZERO).BackColor
                .BorderColor = shpMuteKey(ZERO).BorderColor
                .BackStyle = shpMuteKey(ZERO).BackStyle
                .DrawMode = shpMuteKey(ZERO).DrawMode
                .Shape = shpMuteKey(ZERO).Shape
            End With
        End If
        lblVol(k).Caption = CENT - Int(SldrVol(k).Value / ONEPERCENT)
    Next k
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'User pressed a key -- this sees keys before controls do
    'Form's key-preview mode is set to TRUE.
    Dim k%
    Select Case Shift
        Case 0 'No SHIFT, ALT, or CTRL
            If KeyCode = vbKeyF4 Then Call MnuAudioProfiles_Click
            If KeyCode = vbKeyF5 And FileExists(PATHTOWINAMP) Then Call MnuWinAmp_Click
            If KeyCode = vbKeyF9 And FileExists(PATHTOSNDVOL32) Then Call MnuWindowsVolumeControl_Click
            If KeyCode = vbKeyF6 Then Call MnuGSToolWin_Click
            If KeyCode = vbKeyF7 Then Call MnuTimedMute_Click
            If KeyCode = vbKeyF8 Then Call MnuBackColor_Click
            If KeyCode = vbKeyF10 Then Call PicControl_MouseUp(ZERO, vbLeftButton, ZERO, ZERO, ZERO)
            If KeyCode = vbKeySpace Then Call PicControl_MouseUp(ONE, vbLeftButton, ZERO, ZERO, ZERO)
            If KeyCode = vbKeyF11 Then Call MnuMute_Click
            If KeyCode = vbKeyF12 And ChkOnOff.Enabled = True Then
                Call MnuOnOff_Click
                TmrRefresh.Enabled = False: TmrRefresh.Interval = RAPID_UPDATE_RATE: TmrRefresh.Enabled = True
                TmrRapid.Enabled = False: TmrRapid.Enabled = True
            End If
            If KeyCode = vbKeyNumpad4 Then Me.Left = Me.Left - Screen.TwipsPerPixelX
            If KeyCode = vbKeyNumpad6 Then Me.Left = Me.Left + Screen.TwipsPerPixelX
            If KeyCode = vbKeyNumpad8 Then Me.Top = Me.Top - Screen.TwipsPerPixelY
            If KeyCode = vbKeyNumpad2 Then Me.Top = Me.Top + Screen.TwipsPerPixelY
            If KeyCode = vbKeyNumpad1 Then Me.Left = Me.Left - Screen.TwipsPerPixelX: Me.Top = Me.Top + Screen.TwipsPerPixelY
            If KeyCode = vbKeyNumpad7 Then Me.Left = Me.Left - Screen.TwipsPerPixelX: Me.Top = Me.Top - Screen.TwipsPerPixelY
            If KeyCode = vbKeyNumpad9 Then Me.Left = Me.Left + Screen.TwipsPerPixelX: Me.Top = Me.Top - Screen.TwipsPerPixelY
            If KeyCode = vbKeyNumpad3 Then Me.Left = Me.Left + Screen.TwipsPerPixelX: Me.Top = Me.Top + Screen.TwipsPerPixelY
        Case 1 'With SHIFT
            If KeyCode = vbKeyF4 Then Call MnuExitAndCloseQuickMixer_Click        'SHIFT+F4 = Close
        Case 2 'With CTRL
            If KeyCode = vbKeyR Then Call MnuSingleSliderMode_Click
            If KeyCode = vbKeyG Then Call MnuShowGraduations_Click
            If KeyCode = vbKeyK Then Call MnuShowSkin_Click
            If KeyCode = vbKeyO Then Call MnuAlwaysOnTop_Click
            If KeyCode = vbKeyA Then Call MnuAutoHide_Click
            If KeyCode = vbKeyE Then Call MnuSnapSlidersEvenly_Click
            If KeyCode = vbKeyP Then Call MnuPointedSliders_Click
            If KeyCode = vbKeyT Then Call MnuShowToolTips_Click
            If KeyCode = vbKeyB And TBSupport Then Call MnuShowTB_Click
            If KeyCode = vbKeyV And TBSupport Then Call mnuReverseLogicTrebBass_Click
            If KeyCode = vbKeyNumpad4 Then Me.Left = Me.Left - Screen.TwipsPerPixelX * DECA
            If KeyCode = vbKeyNumpad6 Then Me.Left = Me.Left + Screen.TwipsPerPixelX * DECA
            If KeyCode = vbKeyNumpad8 Then Me.Top = Me.Top - Screen.TwipsPerPixelY * DECA
            If KeyCode = vbKeyNumpad2 Then Me.Top = Me.Top + Screen.TwipsPerPixelY * DECA
            If KeyCode = vbKeyNumpad1 Then Me.Left = Me.Left - Screen.TwipsPerPixelX * DECA: Me.Top = Me.Top + Screen.TwipsPerPixelY * DECA
            If KeyCode = vbKeyNumpad7 Then Me.Left = Me.Left - Screen.TwipsPerPixelX * DECA: Me.Top = Me.Top - Screen.TwipsPerPixelY * DECA
            If KeyCode = vbKeyNumpad9 Then Me.Left = Me.Left + Screen.TwipsPerPixelX * DECA: Me.Top = Me.Top - Screen.TwipsPerPixelY * DECA
            If KeyCode = vbKeyNumpad3 Then Me.Left = Me.Left + Screen.TwipsPerPixelX * DECA: Me.Top = Me.Top + Screen.TwipsPerPixelY * DECA
        Case 4 'With ALT
            If KeyCode = vbKeyLeft Then Me.Left = Me.Left - Screen.TwipsPerPixelX
            If KeyCode = vbKeyRight Then Me.Left = Me.Left + Screen.TwipsPerPixelX
            If KeyCode = vbKeyUp Then Me.Top = Me.Top - Screen.TwipsPerPixelY
            If KeyCode = vbKeyDown Then Me.Top = Me.Top + Screen.TwipsPerPixelY
            If KeyCode = vbKey0 Or KeyCode = vbKeyNumpad0 Then DoProfile 0
            If KeyCode = vbKey1 Or KeyCode = vbKeyNumpad1 Then DoProfile 1
            If KeyCode = vbKey2 Or KeyCode = vbKeyNumpad2 Then DoProfile 2
            If KeyCode = vbKey3 Or KeyCode = vbKeyNumpad3 Then DoProfile 3
            If KeyCode = vbKey4 Or KeyCode = vbKeyNumpad4 Then DoProfile 4
            If KeyCode = vbKey5 Or KeyCode = vbKeyNumpad5 Then DoProfile 5
            If KeyCode = vbKey6 Or KeyCode = vbKeyNumpad6 Then DoProfile 6
            If KeyCode = vbKey7 Or KeyCode = vbKeyNumpad7 Then DoProfile 7
            If KeyCode = vbKey8 Or KeyCode = vbKeyNumpad8 Then DoProfile 8
            If KeyCode = vbKey9 Or KeyCode = vbKeyNumpad9 Then DoProfile 9
        End Select
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Drag the mixer by the form
    If Button <> vbLeftButton Then Exit Sub
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_MouseMove(Button%, Shift%, X!, Y!)
    'Extinguish mouse-over illumination of the control-boxes near the title-bar
    Call Darkness
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MixerVisible And Button = vbRightButton Then Me.PopupMenu MnuPopUp
    If Button = vbMiddleButton Then Beep
End Sub

Public Sub Form_Paint()
    'This tiles the skin on the form's background
    If Not ShowSkin Then Exit Sub
    Dim wid!, hgt!, X!, Y!
    wid = PicSkin.ScaleWidth
    hgt = PicSkin.ScaleHeight
    Y = ZERO
    Do While Y < ScaleHeight
        X = ZERO
        Do While X < ScaleWidth
            PaintPicture PicSkin.Picture, X, Y, wid, hgt
            X = X + wid
        Loop
        Y = Y + hgt
    Loop
End Sub

Public Sub Form_Resize()
    Dim msw%, msh%, mw%, mh%, k%, w%, c!, l!, uhc$: Static ReadyFlag As Boolean, TimesThrough%
    Static swac$, DoneIt As Boolean: If Not DoneIt Then swac = MnuSizeToWinAmp.Caption: DoneIt = True
    If Not ReadyFlag Then TimesThrough = TimesThrough + ONE: If TimesThrough = 20 Then ReadyFlag = True
    If InvisibleControls > ZERO Then MnuUnhideSliders.Visible = True Else MnuUnhideSliders.Visible = False
    uhc = "&Un-Hide " & InvisibleControls & " Hidden Slider": If InvisibleControls > ONE Then uhc = uhc & PLURAL
    MnuUnhideSliders.Caption = uhc: msw = Me.ScaleWidth: msh = Me.ScaleHeight: mw = Me.Width: mh = Me.Height
    bw = (mw - (msw * Screen.TwipsPerPixelX)) \ Screen.TwipsPerPixelX 'Combined width of two window boarders.
    Select Case SingleSliderMode
        Case True
            If msw >= GANGWIDTH * (VisibleControls - InvisibleControls + CENTERINGFACTOR) Then Call MnuSingleSliderMode_Click: GoTo DivertPoint1
            If msw <= MINSIZE Then
                msw = MINSIZE
                Me.Width = (msw + bw) * Screen.TwipsPerPixelX
DivertPoint1:
            End If
        Case False
            If msw <= MINSIZE Then Call MnuSingleSliderMode_Click: GoTo DivertPoint2
            If msw <= GANGWIDTH * (VisibleControls - InvisibleControls + CENTERINGFACTOR) Then
                msw = GANGWIDTH * (VisibleControls - InvisibleControls + CENTERINGFACTOR)
                Me.Width = (msw + bw) * Screen.TwipsPerPixelX
DivertPoint2:
            End If
    End Select
    If msh <= MINSIZE Then msh = MINSIZE: Me.Height = (msh + bw) * Screen.TwipsPerPixelY
    k = msw - MINTITLEBARFORMWIDTH
    If k > ZERO Then PicTitle.Width = k: PicTitle.Top = PCTOP Else PicTitle.Top = HIDECTRL
    k = msw - PC1LEFTOFFORMWIDTH: If k > PCLEFTCUTOFF Then PicControl(ONE).Left = k
    c = msw / (VisibleControls - InvisibleControls + CENTERINGFACTOR)
    If SingleSliderMode Then c = (msw / CENTERINGFACTOR)
    If SnapSliders Or SingleSliderMode Then c = Int(c)
    l = c + ONE
    If ReadyFlag Then DoEvents
    For k = ZERO To VisibleControls
        If SingleSliderMode Then
            c = c + ONE: l = c
            k = CurrentFocus
            SSMode
            Selector(ZERO).Left = c + LEFTSELECTOROFFSET
            Selector(ONE).Left = c + RIGHTSELECTOROFFSET
        End If
        If HideSlider(k) Then PicGang(k).Visible = False
        PicGang(k).Left = Int(l - HALFGANGWIDTH): If Not HideSlider(k) Then l = l + c
        w = msh - VOLCONTAINERSHORTOFFORMHEIGHT
        If w < MINSLIDERHEIGHT Then
            PicVol(k).Left = HIDECTRL:  GoTo Skip
        Else
            If PicVol(k).Left = HIDECTRL Then PicVol(k).Left = ZERO
        End If
        PicVol(k).Height = w
        SldrVol(k).Height = msh - VOLCONTROLSHORTOFFORMHEIGHT
Skip:
        ShpMute(k).Top = msh - MUTESHORTOFFORMHEIGHT
        lblVol(k).Top = msh - MUTESHORTOFFORMHEIGHT
        If ReadyFlag Then DoEvents
        If SingleSliderMode Then Exit For
    Next k
    If mw = WINAMPWIDTH * Screen.TwipsPerPixelX And mh = WINAMPHEIGHT * Screen.TwipsPerPixelY Then
        MnuSizeToWinAmp.Caption = swac & "  (×2)"
    Else
        MnuSizeToWinAmp.Caption = swac & "  (×1)"
    End If
    If Not ShowGrads Then Exit Sub
    Me.Cls
    If w < MINSLIDERHEIGHT Then Exit Sub
    Dim t!, b!, hi!, cc!, ll%, lr%, ccount% 'From this point down, draws the graduations.
    ll = PicGang(ZERO).Left - GRADMARGIN
    If SingleSliderMode Then ll = PicGang(CurrentFocus).Left - GRADMARGIN
    For k = PicGang.UBound To PicGang.lbound Step BACKWARDS
        If PicGang(k).Visible Then Exit For
    Next k
    lr = PicGang(k).Left + GANGWIDTH + GRADMARGIN
    If SingleSliderMode Then lr = PicGang(CurrentFocus).Left + GANGWIDTH + GRADMARGIN
    t = PICVOLTOP + GRADKNOBADJUST
    b = t + PicVol(CurrentFocus).Height
    hi = (b - t) / DECA
    For cc = t To b Step hi - ONE
        ccount = ccount + ONE
        If ccount < (TOTALGRADS + ONE) Then
            Me.Line (ll, Int(cc + ONE))-(lr, Int(cc + ONE)), PCDARK
            Me.Line (ll, Int(cc + TWO))-(lr, Int(cc + TWO)), PCBRIGHT
        End If
    Next cc
End Sub

Private Sub ImgIcon_MouseDown(Idx%, Button%, Shift%, X!, Y!)
    If Shift = vbShiftMask Then Call SldrVol_MouseUp(Idx, Button, Shift, X, Y): Exit Sub
    'If Idx = CurrentFocus Then ImgIcon(Idx).Top = 1: ImgIcon(Idx).Left = 1: DoEvents
    'this is the mouse-button-hold-down-mode of adjusting volume using the device-icons
    IconClickFlag = True:  TmrRepeatDelay.Enabled = True: ButtonFlag = Button: IndexFlag = Idx
End Sub

Private Sub Darkness()
    'extinguish mouse-over illuminaiton of the control-boxes
    Dim k%
    If TmrLightsOut.Enabled = False Then Exit Sub
    TmrLightsOut.Enabled = False
    For k = ZERO To PicControl.UBound
        PicControl(k).BackColor = PCDARK
    Next k
End Sub

Private Sub ImgIcon_MouseUp(Idx%, Button%, Shift%, X!, Y!)
    If Shift = vbShiftMask Then Exit Sub
    'changes the focus when user clicks device-icon or if that control already has focus adjusts the volume
    'up or down, depending on the mouse-button, one percent.
     'If Idx = CurrentFocus Then ImgIcon(Idx).Top = 0: ImgIcon(Idx).Left = 0
    If TmrRepeatDelay.Enabled = True Then TmrRepeatDelay.Enabled = False: TmrRepeat.Enabled = False
    If Idx <> CurrentFocus Then SldrVol(Idx).SetFocus: Exit Sub
    If Button = vbLeftButton Then SldrVol(Idx).Value = SldrVol(Idx).Value - ONEPERCENT
    If Button = vbRightButton Then SldrVol(Idx).Value = SldrVol(Idx).Value + ONEPERCENT
End Sub

Private Sub lblVol_Click(Idx%)
    'toggles the mute when the user clicks it
    If ChkMute(Idx).Value = vbUnchecked Then
        ChkMute(Idx).Value = vbChecked
    Else
        ChkMute(Idx).Value = vbUnchecked
    End If
    Call ChkMute_Click(Idx)
    TmrRefresh.Enabled = False: TmrRefresh.Interval = RAPID_UPDATE_RATE: TmrRefresh.Enabled = True
    TmrRapid.Enabled = False: TmrRapid.Enabled = True
End Sub

Public Sub MnuAlwaysOnTop_Click()
'toggles on-top
    OnTop = Not OnTop
    If OnTop Then
        SetWindowPos _
            hwnd, _
            HWND_TOPMOST, _
            ZERO, _
            ZERO, _
            ZERO, _
            ZERO, _
            SWP_NOMOVE + SWP_NOSIZE
        MnuAlwaysOnTop.Checked = vbChecked
    Else
        SetWindowPos _
            hwnd, _
            HWND_NOTOPMOST, _
            ZERO, _
            ZERO, _
            ZERO, _
            ZERO, _
            SWP_NOMOVE + SWP_NOSIZE
        MnuAlwaysOnTop.Checked = vbUnchecked
    End If
    FrmSettings.DoChecks
    If BackToSettings Then FrmSettings.SetFocus
End Sub


Private Sub MnuAudioProfiles_Click()
    If FrmAudioProfiles.Visible = True Then GoTo NoAudioProfilesReposition
    FrmAudioProfiles.Left = FrmAudioProfilesLeft
    FrmAudioProfiles.Top = FrmAudioProfilesTop
NoAudioProfilesReposition:
    FrmAudioProfiles.Show
End Sub

Public Sub MnuBackColor_Click()
    'Show Back Color Tool-Window
    If FrmColor.Visible = True Then GoTo NoColorReposition
    FrmColor.Left = FrmColorLeft
    FrmColor.Top = FrmColorTop
NoColorReposition:
    FrmColor.Show
End Sub

Public Sub MnuControlKeysInfo_Click()
FrmControlKeysInfo.Show
End Sub

Public Sub MnuEmail_Click()
   Dim r As Long, sc$
   sc = "mailto:particle@neo.rr.com%20(Evan%20Edwards)?subject=Quick%20Mixer%20Feedback   Version = " & App.Major & "." & App.Minor & "." & App.Revision & "   Soundcard = " & ProductName
   r = ShellExecute( _
        Me.hwnd, _
        "Open", _
        sc, _
        vbNullString, _
        App.Path, _
        vbNormalFocus)
End Sub

Private Sub MnuExitAndCloseQuickMixer_Click()
    'the user wants to end execution of the mixer, now.
    'get out of the mixer-update (timer controled) loop, or we'll crash big-time.
    TmrRapid.Enabled = False
    TmrRefresh.Enabled = False
    'turn off all the timers while we are at it.
    TmrRepeatDelay.Enabled = False
    TmrRepeat.Enabled = False
    TmrLightsOut.Enabled = False
    TmrAutoMute.Enabled = False
    'this waits twice as long as it needs to to make sure the update-loop has ceased
    'and sends us to the form_unload routine in a round-about but pratical way.
    TmrKillMe.Enabled = True
End Sub

Private Sub MnuGSToolWin_Click()
    'Show Multi-select General Settings Tool-Window
    If FrmSettings.Visible = True Then GoTo NoSettingsReposition
    FrmSettings.Left = FrmSettingsLeft
    FrmSettings.Top = FrmSettingsTop
NoSettingsReposition:
    FrmSettings.Show
End Sub

Public Sub MnuHelp_Click()
    'User clicked HELP on the menu
    If FileExists("c:\program files\qmixer\qmixer.html") Then
        ShellExecute _
        hwnd, _
        "open", _
        "qmixer.html", _
        vbNullString, _
        vbNullString, _
        SW_SHOW
        Exit Sub
    End If
    If FileExists(PATHTOWORDPAD) And FileExists(App.Path & DOCFILENAME) Then
        Shell PATHTOWORDPAD & " " & Chr$(34) & App.Path & DOCFILENAME & Chr$(34), vbNormalFocus
    Else
        If FileExists(PATHTONOTEPAD) And FileExists(App.Path & HELPFILENAME) Then
            Shell PATHTONOTEPAD & " " & Chr$(34) & App.Path & HELPFILENAME & Chr$(34), vbNormalFocus
        End If
    End If
End Sub

Private Sub MnuHideMixer_Click()
    'user clicked HIDE MIXER on the menu
    Me.Hide
    MixerVisible = False
    MnuRestoreMixer.Enabled = True
    MnuHideMixer.Enabled = False
End Sub

Public Sub MnuHomepage_Click()
    'takes user to homepage in his/her default browser from the menu - cool.
    ShellExecute _
        hwnd, _
        "open", _
        HOMEPAGE, _
        vbNullString, _
        vbNullString, _
        SW_SHOW
End Sub

Private Sub MnuMultimediaProperties_Click()
    'user clicked Multimedia Properties... in launch menu.
    Dim dblReturn As Double
    dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,0", 5)
End Sub

Private Sub MnuMute_Click()
    TmrRefresh.Enabled = False: TmrRefresh.Interval = NORMAL_UPDATE_RATE: TmrRefresh.Enabled = True
    'user clicked mute on the menu
    Call lblVol_Click(ZERO)
End Sub

Private Sub MnuOnOff_Click()
    If ChkOnOff.Value = vbChecked Then ChkOnOff.Value = vbUnchecked Else ChkOnOff.Value = vbChecked
End Sub

Public Sub MnuPointedSliders_Click()
    'toggles pointed slider mode from the menu
    Dim k%
    PointedSliders = Not PointedSliders
    If PointedSliders Then
        MnuPointedSliders.Checked = vbChecked
        For k = ZERO To PicGang.UBound
            SldrVol(k).TickStyle = sldTopLeft
        Next k
    Else
        MnuPointedSliders.Checked = vbUnchecked
        For k = ZERO To PicGang.UBound
            SldrVol(k).TickStyle = sldBoth
        Next k
    End If
    FrmSettings.DoChecks
End Sub

Private Sub MnuProduct_Click()
    'user clicked DEVICE: menu-item
    MnuMultimediaProperties_Click
End Sub

Private Sub MnuResetDefault_Click()
    Dim r
    r = MsgBox("WARNING: This will reset Quick Mixer to all of it's original-default settings." & vbCrLf & vbCrLf & "This can solve problems with crashes if bad data is somehow saved to the initialization file." & vbCrLf & "All audio profiles, timed mute settings, and other settings will be reset to their defaults if you proceed." & vbCrLf & "Quick Mixer will also immediately close if you proceed. You will have to re-start it manually." & vbCrLf & vbCrLf & "Proceed now?", 308, "Reset Quick Mixer to all original-default settings?")
    If r = vbYes Then
        NukeIni = True
        MnuExitAndCloseQuickMixer_Click
    End If
End Sub

Public Sub MnuRestoreMixer_Click()
    'user clicked SHOW MIXER on menu -OR- left-clicked tray-icon
    Me.Show
    MixerVisible = True
    MnuHideMixer.Enabled = True
    MnuRestoreMixer.Enabled = False
    TmrRefresh.Enabled = False: TmrRefresh.Interval = NORMAL_UPDATE_RATE: TmrRefresh.Enabled = True
End Sub

Public Sub mnuReverseLogicTrebBass_Click()
    'user clicked reverse treble bass logic on the menu
    ReverseLogicTrebBass = Not ReverseLogicTrebBass
    If ReverseLogicTrebBass Then MnuReverseLogicTrebBass.Checked = vbChecked Else MnuReverseLogicTrebBass.Checked = vbUnchecked
    FrmSettings.DoChecks
End Sub

Public Sub MnuShowGraduations_Click()
    'toggles graduations on/off from menu
    ShowGrads = Not ShowGrads
    If ShowGrads Then
        MnuShowGraduations.Checked = vbChecked
        ShowSkin = False: MnuShowSkin.Checked = vbUnchecked
        Me.AutoRedraw = True
        Form_Resize
    Else
        MnuShowGraduations.Checked = vbUnchecked
        Me.Cls
        Me.AutoRedraw = False
    End If
    FrmSettings.DoChecks
End Sub

Public Sub MnuShowSkin_Click()
    'toggles skin on/off from menu
    ShowSkin = Not ShowSkin
    If ShowSkin Then
        If FileExists(App.Path & SKIN) Then Set PicSkin.Picture = LoadPicture(App.Path & SKIN)
        MnuShowSkin.Checked = True
        ShowGrads = False: MnuShowGraduations.Checked = vbUnchecked
        Me.Cls
        Me.AutoRedraw = False
        Form_Paint
        Form_Resize
    Else
    MnuShowSkin.Checked = False
    Me.Cls
    End If
    FrmSettings.DoChecks
End Sub

Public Sub MnuShowTB_Click()
    Dim w%
    'toggles the display of treble and bass sliders from the menu if they are available
    ShowTB = Not ShowTB
    If ShowTB Then w = 1 Else w = -1
    If HideSlider(MaxSources + ONE) Then InvisibleControls = InvisibleControls + w
    If HideSlider(MaxSources + TWO) Then InvisibleControls = InvisibleControls + w
    ChangeFlag = True
    GetMixerInfo
    ReInsert
    If ShowTB Then
        MnuShowTB.Checked = vbChecked
    Else
        MnuShowTB.Checked = vbUnchecked
    End If
    FrmSettings.DoChecks
    If BackToSettings Then FrmSettings.SetFocus
End Sub

Public Sub MnuShowToolTips_Click()
    'toggles tool-tips on/off from the menu
    ToolTips = Not ToolTips
    If ToolTips Then MnuShowToolTips.Checked = vbChecked Else MnuShowToolTips.Checked = vbUnchecked
    Tips ToolTips
    FrmSettings.DoChecks
End Sub

Public Sub MnuSingleSliderMode_Click()
    'toggles single-slider-mode on/off from the menu
    SingleSliderMode = Not SingleSliderMode
    If SingleSliderMode Then
        MnuSingleSliderMode.Checked = vbChecked: SSMode: Me.Width = (43 + bw) * Screen.TwipsPerPixelX
    Else
        Selector(ZERO).Visible = False: Selector(ONE).Visible = False
        MnuSingleSliderMode.Checked = vbUnchecked: Me.Width = (50 + bw) * Screen.TwipsPerPixelX
    End If
    FrmSettings.DoChecks
End Sub

Private Sub MnuSizeToWinAmp_Click()
    If SingleSliderMode Then MnuSingleSliderMode_Click
    'sizes the mixer to the same size as WinAmp, alternates between 1x and 2x WinAmp sizes - this was clicked on menu
    If FrmMxr.Width = WINAMPWIDTH * Screen.TwipsPerPixelX And FrmMxr.Height = WINAMPHEIGHT * Screen.TwipsPerPixelY Then
        FrmMxr.Width = (WINAMPWIDTH * TWO) * Screen.TwipsPerPixelX
        FrmMxr.Height = (WINAMPHEIGHT * TWO) * Screen.TwipsPerPixelY
    Else
        FrmMxr.Width = WINAMPWIDTH * Screen.TwipsPerPixelX
        FrmMxr.Height = WINAMPHEIGHT * Screen.TwipsPerPixelY
    End If
End Sub

Public Sub MnuSnapSlidersEvenly_Click()
    'toggles SPACE SLIDERS EVENLY mode on/off from the menu
    SnapSliders = Not SnapSliders
    If SnapSliders Then MnuSnapSlidersEvenly.Checked = vbChecked Else MnuSnapSlidersEvenly.Checked = vbUnchecked
    Form_Resize
    FrmSettings.DoChecks
End Sub

Private Sub MnuSoundsProps_Click()
    'user clicked SOUNDS PROPERTIES on the menu
    Dim dblReturn As Double
    dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl @1", 5)
End Sub

Public Sub MnuTimedMute_Click()
    'Show Timed Mute Tool-Window
    If FrmTimer.Visible = True Then GoTo NoReposition
    FrmTimer.Left = FrmTimerLeft
    FrmTimer.Top = FrmTimerTop
NoReposition:
    FrmTimer.Show
End Sub

Private Sub MnuUnhideSliders_Click()
    Dim k%
    For k = 0 To 18
        HideSlider(k) = False
    Next k
    InvisibleControls = 0
    For k = 0 To VisibleControls
        PicGang(k).Visible = True
    Next
    Form_Resize
End Sub

Private Sub MnuWinAmp_Click()
    'user clicked WinAmp on the Launch menu
    Shell PATHTOWINAMP, vbNormalFocus
End Sub

Private Sub MnuWindowsCDPlayer_Click()
    'user clicked Windows CD Player on the Launch menu
    Shell PATHTOCDPLAYER, vbNormalFocus
End Sub

Private Sub MnuWindowsMediaPlayer_Click()
    'user clicked Windows Media Player on the Launch menu
    Shell PATHTOMEDIAPLAYER2, vbNormalFocus
End Sub

Private Sub MnuWindowsSoundRecorder_Click()
    'user clicked Windows Sound Recorder on the launch menu
    Shell PATHTOSNDREC32, vbNormalFocus
End Sub

Private Sub MnuWindowsVolumeControl_Click()
    'user clicked Windows Volume Control on the Launch menu
    Shell PATHTOSNDVOL32, vbNormalFocus
End Sub

Private Sub PicControl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'user just clicked a control-box, return fake-focus to correct device-icon
    If SldrVol(CurrentFocus).Visible = True Then SldrVol(CurrentFocus).SetFocus
End Sub

Private Sub PicControl_MouseMove(Idx%, Button%, Shift%, X!, Y!)
    'this illuminates the control-boxes on mouse-over
    Dim k%
    For k = ZERO To ONE
        If Idx = k Then
            PicControl(k).BackColor = PCBRIGHT
            TmrLightsOut.Enabled = False
            TmrLightsOut.Enabled = True
        Else
            PicControl(k).BackColor = PCDARK
        End If
    Next k
End Sub

Private Sub PicControl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 'user just FINISHED clicking a control-box, act on it
    If SldrVol(CurrentFocus).Visible = True Then SldrVol(CurrentFocus).SetFocus
    Call Darkness
    Select Case Index
        Case Is = ZERO: Me.PopupMenu MnuPopUp
        Case Is = ONE: Call MnuHideMixer_Click
    End Select
End Sub

Private Sub PicSpecial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SldrVol(CurrentFocus).Visible = True Then SldrVol(CurrentFocus).SetFocus
    TmrRefresh.Enabled = False: TmrRefresh.Interval = RAPID_UPDATE_RATE: TmrRefresh.Enabled = True
    TmrRapid.Enabled = False: TmrRapid.Enabled = True
    Call MnuOnOff_Click
End Sub

Private Sub PicTitle_MouseDown(Button%, Shift%, X!, Y!)
    'drag form from the fake title-bar
    SldrVol(CurrentFocus).SetFocus
    If Button <> vbLeftButton Then Exit Sub
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub PicTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'if the mouse is here the control-boxes should not be illuminated
    Call Darkness
End Sub

Private Sub PicTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'this lets the user have the menu if he/she right-clicks the title bar, non-standard but nice.
    'if the user double-clicks the title-bar, it will load the default audio profile.
    'The first mouseUp event in the double-click is suppressed for left-button due to proc in MouseDown and is interpreted as being a single-cick. (Phew!)
    If Button = vbRightButton Then Call PicControl_MouseUp(ZERO, Button, Shift, X, Y) Else If Shift <> vbShiftMask Then DoProfile 0 Else FrmInfo.Show
    'if the user SHIFT-DoubleClicks the title-bar the first-use-info form (FrmInfo) is shown.
End Sub

Private Sub Selector_Click(Index As Integer)
    'this rotates through all the sliders when in Single Slider Mode
    'this procedure executes when the user clicks the small arrow buttons which only
    'appear in Single Slider Mode
RepeatStep:
    If Index = ZERO Then
        If CurrentFocus > ZERO Then CurrentFocus = CurrentFocus - ONE Else CurrentFocus = VisibleControls
    Else
        If CurrentFocus < VisibleControls Then CurrentFocus = CurrentFocus + ONE Else CurrentFocus = ZERO
    End If
    If HideSlider(CurrentFocus) Then GoTo RepeatStep
    If SldrVol(CurrentFocus).Visible = True Then SldrVol(CurrentFocus).SetFocus
    Form_Resize
End Sub

Private Sub SldrVol_Change(Idx%)
    'user is clicking a slider
   If IconClickFlag Then GoTo SkipRapidUpdate
   TmrRefresh.Enabled = False: TmrRefresh.Interval = RAPID_UPDATE_RATE: TmrRefresh.Enabled = True
   TmrRapid.Enabled = False: TmrRapid.Enabled = True
SkipRapidUpdate:
    AdjustOutput Idx
    IconClickFlag = False
End Sub

Private Sub SldrVol_GotFocus(Idx%)
    'a slider got focus in any way
    shpFocus(Idx).DrawMode = vbMergePen
    CurrentFocus = Idx
    If SingleSliderMode Then Form_Resize
End Sub

Private Sub SldrVol_LostFocus(Idx%)
    'a slider lost focus in any way
    shpFocus(Idx).DrawMode = vbNop
End Sub

Private Sub SldrVol_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift = vbShiftMask And VisibleControls - InvisibleControls > ONE Then
        HideSlider(Index) = True
        InvisibleControls = InvisibleControls + ONE
        Form_Resize
    End If
End Sub

Private Sub SldrVol_Scroll(Idx%)
    'user is scrolling a slider
    AdjustOutput Idx
End Sub

Private Sub AdjustOutput(Idx%)
    'this is what we do whenever the user changes the volume in any way
    Dim FaderVol&
    Dim hmem&
    Dim MCDMono As MIXERCONTROLDETAILS
    Dim MCDStereo As MIXERCONTROLDETAILS
    'see if it is stereo
    If MixerState(Idx).MxrChannels = TWO Then
        'get the volume of the fader that is being adjusted
        If Idx > MaxSources And ReverseLogicTrebBass Then
            FaderVol = SldrVol(Idx).Value
        Else
            FaderVol = SLIDERMAX - SldrVol(Idx).Value
        End If
        'set both channels (since this is stereo) to the volume of that fader
        MixerState(Idx).MxrRightVol = FaderVol
        MixerState(Idx).MxrLeftVol = FaderVol
        'prepare the MCD struct
        MCDStereo.cbDetails = 4  'four bytes, the size of a long
        MCDStereo.cbStruct = 24 '
        MCDStereo.dwControlID = MixerState(Idx).MxrVolID 'use the ID of the current fader
        MCDStereo.item = ZERO 'no 'advanced' items, just the fader itself
        MCDStereo.cChannels = TWO 'stereo has two channels
        'allocate a small amount of memory
        hmem = GlobalAlloc(&H40, 8)
        MCDStereo.paDetails = GlobalLock(hmem)
        'copy the values of the left and right channels to the struct
        CopyPtrFromStruct MCDStereo.paDetails, MixerState(Idx).MxrRightVol, 8
        CopyPtrFromStruct MCDStereo.paDetails, MixerState(Idx).MxrLeftVol, 8
        'finalize the settings
        mixerSetControlDetails hMixer, MCDStereo, MIXER_SETCONTROLDETAILSF_VALUE
        'tidy up memory
        GlobalUnlock hmem
        GlobalFree hmem
    Else
        'set the mono channel (since this is mono) to the volume of the fader that is being ajusted
        If Idx > MaxSources And ReverseLogicTrebBass Then
            MixerState(Idx).MxrVol = SldrVol(Idx).Value
        Else
            MixerState(Idx).MxrVol = SLIDERMAX - SldrVol(Idx).Value
        End If
        'prepare the MCD struct
        MCDMono.cbDetails = Len(MixerState(Idx).MxrVol)
        MCDMono.cbStruct = Len(MCDMono)
        MCDMono.dwControlID = MixerState(Idx).MxrVolID
        MCDMono.item = ZERO
        MCDMono.cChannels = ONE 'mono has but one channel
        'allocate a small amount of memory
        hmem = GlobalAlloc(&H40, 4)
        MCDMono.paDetails = GlobalLock(hmem)
        'copy the value of the one channel to the struct
        CopyPtrFromStruct MCDMono.paDetails, MixerState(Idx).MxrVol, 4
        'finalize the setting
        mixerSetControlDetails hMixer, MCDMono, MIXER_SETCONTROLDETAILSF_VALUE
        'tidy up memory
        GlobalUnlock hmem
        GlobalFree hmem
    End If
    'update the volume readout
    lblVol(Idx).Caption = CENT - Int(SldrVol(Idx).Value / ONEPERCENT)
End Sub

Private Sub ChkMute_Click(Idx%)
    'a mute-toggle has been clicked or actuated in some way
    Dim hmem&
    'save setting of the mute control which was clicked
    MixerState(Idx).MxrMute = ChkMute(Idx).Value
    'prepare the MCD struct
    MCD.cbStruct = Len(MCD) 'overall struct size
    MCD.dwControlID = MixerState(Idx).MxrMuteID 'control ID of the mute control which was clicked
    MCD.cbDetails = 4 'four bytes, the size of a long
    MCD.cChannels = ONE 'mute has but one channel
    MCD.item = ZERO 'no items
    'allocate a small amount of memory
    hmem = GlobalAlloc(&H40, 4)
    'consign the memory to MCD.paDetails
    MCD.paDetails = GlobalLock(hmem)
    'copy the value of the clicked mute to MCD.paDetails
    CopyPtrFromStruct MCD.paDetails, MixerState(Idx).MxrMute, 4
    'finalize the mute setting
    mixerSetControlDetails hMixer, MCD, MIXER_SETCONTROLDETAILSF_VALUE
    'tidy up memory
    GlobalUnlock hmem
    GlobalFree hmem
End Sub

Private Sub ChkOnOff_Click()
    Dim hmem&
    'save setting of the ON/OFF control which was clicked
    MixerState(ZERO).MxrOnOff = ChkOnOff.Value
    'prepare the MCD struct
    MCD.cbStruct = Len(MCD) 'overall struct size
    MCD.dwControlID = MixerState(ZERO).MxrOnOffID 'control ID of the mute control which was clicked
    MCD.cbDetails = 4 'four bytes, the size of a long
    MCD.cChannels = ONE 'ON/OFF has but one channel
    MCD.item = ZERO 'no items
    'allocate a small amount of memory
    hmem = GlobalAlloc(&H40, 4)
    'consign the memory to MCD.paDetails
    MCD.paDetails = GlobalLock(hmem)
    'copy the value of the clicked mute to MCD.paDetails
    CopyPtrFromStruct MCD.paDetails, MixerState(ZERO).MxrOnOff, 4
    'finalize the mute setting
    mixerSetControlDetails hMixer, MCD, MIXER_SETCONTROLDETAILSF_VALUE
    'tidy up memory
    GlobalUnlock hmem
    GlobalFree hmem
End Sub

Private Sub Form_Unload(Cancel%)
    'kill off the last running timer
    TmrKillMe.Enabled = False
    'unload any sister-forms that may loaded
    Unload FrmTimer
    Unload FrmSettings
    Unload FrmColor
    Unload FrmControlKeysInfo
    Unload FrmAudioProfiles
    Unload FrmInfo
    Useage = Useage + 1 'recorde number of uses in ini file.
    If NukeIni = True And FileExists(INI) Then Kill INI: GoTo FinalizeNuke
    Dim f%, k%
    'write to .ini file
    On Error GoTo badwriteini
    f = FreeFile
    Open INI For Output As #f
    Print #f, INIMESSAGE
    Print #f, Me.Left
    Print #f, Me.Top
    Print #f, Me.Width
    Print #f, Me.Height
    Print #f, MixerVisible
    Print #f, ShowGrads
    Print #f, SnapSliders
    Print #f, AutoHide
    Print #f, OnTop
    Print #f, ToolTips
    Print #f, Useage
    Print #f, ShowSkin
    Print #f, PointedSliders
    Print #f, SingleSliderMode
    Print #f, CurrentFocus
    Print #f, ReverseLogicTrebBass
    Print #f, ShowTB
    Print #f, FrmTimerLeft
    Print #f, FrmTimerTop
    Print #f, MuteHour
    Print #f, MuteMinute
    Print #f, UnMuteHour
    Print #f, UnMuteMinute
    Print #f, TMuteFlag
    Print #f, TUnMuteFlag
    Print #f, MixerBackColor
    Print #f, FrmSettingsLeft
    Print #f, FrmSettingsTop
    Print #f, FrmColorLeft
    Print #f, FrmColorTop
    Print #f, FrmAudioProfilesLeft
    Print #f, FrmAudioProfilesTop
    Print #f, Profile(0)
    Print #f, Profile(1)
    Print #f, Profile(2)
    Print #f, Profile(3)
    Print #f, Profile(4)
    Print #f, Profile(5)
    Print #f, Profile(6)
    Print #f, Profile(7)
    Print #f, Profile(8)
    Print #f, Profile(9)
    Print #f, StartWithProfile
    For k = 0 To 18
    Print #f, HideSlider(k)
    Next k
    Print #f, InvisibleControls
    Print #f, DontShowInfo
badwriteini:
    Close #f
FinalizeNuke:
    mixerClose hMixer 'important - close the mixer
    RemoveFromTray 'important - remove tray icon
End Sub

Private Sub PickIcon(Idx%)
    'selects icons based on the name of each device
    Dim Compare$, k%
    Compare = UCase(Left$(LblName(Idx), THREE))
    k = ICONUNKNOWN
    Select Case Compare
        Case Is = "VOL": k = ICONMAIN
        Case Is = "MAI": k = ICONMAIN
        Case Is = "MAS": k = ICONMAIN
        Case Is = "WAV": k = ICONWAVE
        Case Is = "MID": k = ICONMIDI
        Case Is = "SYN": k = ICONMIDI
        Case Is = "CD ": k = ICONCD
        Case Is = "CD": k = ICONCD
        Case Is = "LIN": k = ICONLINE
        Case Is = "MIC": k = ICONMIC
        Case Is = "PC ": k = ICONSPKR
        Case Is = "PC": k = ICONSPKR
        Case Is = "INT": k = ICONSPKR
        Case Is = "VOI": k = ICONVOICE
        Case Is = "DIC": k = ICONDICT
        Case Is = "SPE": k = ICONSPEECH
        Case Is = "AUX": k = ICONAUX
        Case Is = "MOD": k = ICONMODEM
        Case Is = "VID": k = ICONVIDEO
        Case Is = "TEL": k = ICONTELEPHONE
        Case Is = "PHO": k = ICONTELEPHONE
    End Select
    If Idx = MaxSources + ONE Then k = ICONTREBLE
    If Idx = MaxSources + TWO Then k = ICONBASS
    ImgIcon(Idx).Picture = ImgIconRes(k).Picture
End Sub

Private Sub Tips(Status As Boolean)
    'turns the tool-tips and tray-tip on/off
    Dim k%
    If Status = True Then
        For k = ZERO To MaxSources + TWO
            ImgIcon(k).ToolTipText = LblName(k).Caption
            SldrVol(k).ToolTipText = LblName(k).Caption
            lblVol(k).ToolTipText = LblName(k).Caption & " " & LangReadoutWord
            If ShpMute(k).Visible Then lblVol(k).ToolTipText = lblVol(k).ToolTipText & "/" & LangMuteWord
        Next k
        PicControl(ZERO).ToolTipText = LangMixerMenuWord
        PicControl(ONE).ToolTipText = LangHideMixerWord
        PicTitle.ToolTipText = LangMoveWord
        PicSpecial.ToolTipText = LangAdvCtrlWord
    Else
        For k = ZERO To MaxSources + TWO
            ImgIcon(k).ToolTipText = ""
            SldrVol(k).ToolTipText = ""
            lblVol(k).ToolTipText = ""
        Next k
        PicControl(ZERO).ToolTipText = ""
        PicControl(ONE).ToolTipText = ""
        PicTitle.ToolTipText = ""
        PicSpecial.ToolTipText = ""
    End If
    Update_Tray_Icon 'takes care of tool-tip for tray icon
End Sub

Private Sub TmrAutoMute_Timer()
    'this checks ever 10 seconds to see if it's time to mute or un-mute the master volume
    'according to the auto-mute and auto-un-mute times if they are activated
    Dim t$, th%, tm%
    t = Time$: th = Val(Left$(t, TWO)): tm = Val(Mid$(t, 4, TWO))
    If TMuteFlag Then
        If th = MuteHour And tm = MuteMinute Then ChkMute(ZERO).Value = vbChecked: ChkMute_Click ZERO
    End If
    If TUnMuteFlag Then
        If th = UnMuteHour And tm = UnMuteMinute Then ChkMute(ZERO).Value = vbUnchecked: ChkMute_Click ZERO
    End If
End Sub

Private Sub TmrKillMe_Timer()
    'this waites 'till the mixer-update timer-loop is surely finished and sends us to form_unload to finish up and terminate execution
    Unload Me
End Sub

Private Sub TmrLightsOut_Timer()
    'the control-box has been illumined for long enough (3 seconds)
    Call Darkness
End Sub

Private Sub TmrRapid_Timer()
    TmrRapid.Enabled = False
End Sub

Private Sub TmrRefresh_Timer()
   'This is really the main loop, driven by the timer of course, just two sub calls
    GetMixerInfo
    UpdateMixer
    If Not TmrRapid.Enabled And MixerVisible Then TmrRefresh.Enabled = False: TmrRefresh.Interval = NORMAL_UPDATE_RATE: TmrRefresh.Enabled = True
    If Not TmrRapid.Enabled And Not MixerVisible Then TmrRefresh.Enabled = False: TmrRefresh.Interval = SLOW_UPDATE_RATE: TmrRefresh.Enabled = True
End Sub

Private Sub TmrRepeat_Timer()
    'this is for the mouse-hold-down mode of adjusting volume with the device-icons
    If IndexFlag <> CurrentFocus Then SldrVol(IndexFlag).SetFocus
    If ButtonFlag = vbLeftButton Then SldrVol(IndexFlag).Value = SldrVol(IndexFlag).Value - ONEPERCENT
    If ButtonFlag = vbRightButton Then SldrVol(IndexFlag).Value = SldrVol(IndexFlag).Value + ONEPERCENT
End Sub

Private Sub TmrRepeatDelay_Timer()
    'this adds a small delay to the hold-down mode so we can still use single clicks to make fine volume adjustments
    TmrRepeat.Enabled = True
End Sub

Public Sub MnuAutoHide_Click()
    'toggles Auto-Hide mode from the menu
    AutoHide = Not AutoHide
    If AutoHide Then
        MnuAutoHide.Checked = vbChecked
    Else
        MnuAutoHide.Checked = vbUnchecked
    End If
    FrmSettings.DoChecks
    If Not AutoHide Then MnuRestoreMixer_Click
    If BackToSettings Then FrmSettings.SetFocus
End Sub

Private Sub Update_Tray_Icon()
    'puts the right icon in the tray, red-icon when master-volume is muted
    If MixerState(ZERO).MxrMute = False Then
        Me.Icon = TrayIconRes(ZERO).Picture
    Else
        Me.Icon = TrayIconRes(ONE).Picture
    End If
    'RemoveFromTray
    'AddToTray Me, MnuPopUp
    SetTrayIcon Me.Icon
    If ToolTips Then SetTrayTip LangQuickMixerWord
End Sub

Private Sub SSMode()
    'hides all but the currently selected slider while in Single Slider Mode
    Dim k%
    For k = ZERO To VisibleControls
        If k <> CurrentFocus Then PicGang(k).Left = HIDECTRL
    Next k
    Selector(ZERO).Visible = True: Selector(ONE).Visible = True
End Sub

Private Sub Import_Language()
    Dim f%, r As Variant
    'try to read the .ini file
    On Error GoTo badreadlang
    f = FreeFile
    Open App.Path & LANG_FILE For Input As #f
    Input #f, r: Debug.Print r: 'reads the lang-file message line, does nothing with it
    Input #f, r: Debug.Print r: 'reads the lang-file message line, does nothing with it
    Input #f, r: Debug.Print r: 'reads the lang-file message line, does nothing with it
    Input #f, r: Debug.Print r: MnuAboutQuickMixer.Caption = Convert(r)
    Input #f, r: Debug.Print r: MnuHomepage.Caption = Convert(r)
    Input #f, r: Debug.Print r: MnuHelp.Caption = Convert(r)
    Input #f, r: Debug.Print r: MnuProduct.Caption = Convert(r)
    Input #f, r: Debug.Print r: MnuSettings.Caption = Convert(r): FrmSettings.Caption = Convert(r)
    Input #f, r: Debug.Print r: MnuGSToolWin.Caption = Convert(r) 'Multi-select
    Input #f, r: Debug.Print r: MnuAutoHide.Caption = Convert(r): FrmSettings.Check1(0).Caption = Convert(r)
    Input #f, r: Debug.Print r: MnuAlwaysOnTop.Caption = Convert(r): FrmSettings.Check1(1).Caption = Convert(r)
    Input #f, r: Debug.Print r: MnuSnapSlidersEvenly.Caption = Convert(r): FrmSettings.Check1(2).Caption = Convert(r)
    Input #f, r: Debug.Print r: MnuSingleSliderMode.Caption = Convert(r): FrmSettings.Check1(3).Caption = Convert(r)
    Input #f, r: Debug.Print r: MnuPointedSliders.Caption = Convert(r): FrmSettings.Check1(4).Caption = Convert(r)
    Input #f, r: Debug.Print r: MnuShowGraduations.Caption = Convert(r): FrmSettings.Check1(5).Caption = Convert(r)
    Input #f, r: Debug.Print r: MnuShowSkin.Caption = Convert(r): FrmSettings.Check1(6).Caption = Convert(r)
    Input #f, r: Debug.Print r: MnuShowToolTips.Caption = Convert(r): FrmSettings.Check1(7).Caption = Convert(r)
    Input #f, r: Debug.Print r: MnuShowTB.Caption = Convert(r): FrmSettings.Check1(8).Caption = Convert(r)
    Input #f, r: Debug.Print r: MnuReverseLogicTrebBass.Caption = Convert(r): FrmSettings.Check1(9).Caption = Convert(r)
    Input #f, r: Debug.Print r: 'MnuAdvancedSettings.Caption = Convert(r) 'menu level option eliminated
    Input #f, r: Debug.Print r: MnuTimedMute.Caption = Convert(r)
    Input #f, r: Debug.Print r: MnuBackColor.Caption = Convert(r)
    Input #f, r: Debug.Print r: MnuLaunch.Caption = Convert(r)
    Input #f, r: Debug.Print r: MnuWindowsVolumeControl.Caption = Convert(r)
    Input #f, r: Debug.Print r: MnuMultimediaProperties.Caption = Convert(r)
    Input #f, r: Debug.Print r: MnuSoundsProps.Caption = Convert(r)
    Input #f, r: Debug.Print r: MnuWindowsSoundRecorder.Caption = Convert(r)
    Input #f, r: Debug.Print r: MnuWindowsCDPlayer.Caption = Convert(r)
    Input #f, r: Debug.Print r: MnuWindowsMediaPlayer.Caption = Convert(r)
    Input #f, r: Debug.Print r: MnuWinAmp.Caption = Convert(r)
    Input #f, r: Debug.Print r: MnuRestoreMixer.Caption = Convert(r)
    Input #f, r: Debug.Print r: MnuHideMixer.Caption = Convert(r)
    Input #f, r: Debug.Print r: MnuSizeToWinAmp.Caption = Convert(r)
    Input #f, r: Debug.Print r: MnuOnOff.Caption = Convert(r): LangAdvCtrlWord = Convert(r)
    Input #f, r: Debug.Print r: MnuMute.Caption = Convert(r): LangMuteWord = Convert(r)
    Input #f, r: Debug.Print r: 'MnuQuiet.Caption = Convert(r) 'menu option deleted
    Input #f, r: Debug.Print r: 'MnuZeroAllSliders.Caption = Convert(r) 'menu option deleted
    Input #f, r: Debug.Print r: 'MnuCancelThisMenu.Caption = Convert(r) 'menu option deleted
    Input #f, r: Debug.Print r: MnuExitAndCloseQuickMixer.Caption = Convert(r)
    Input #f, r: Debug.Print r: FrmSettings.Command1.Caption = Convert(r): FrmTimer.Command1.Caption = Convert(r): FrmColor.Command1.Caption = Convert(r): LangOkWord = Convert(r)
    Input #f, r: Debug.Print r: FrmTimer.Caption = Convert(r): FrmTimer.Label4(0).Caption = Convert(r): LangTimedMuteWord = Convert(r)
    Input #f, r: Debug.Print r: FrmTimer.Label4(1).Caption = Convert(r): LangTimedUnmuteWord = Convert(r)
    Input #f, r: Debug.Print r: LangActivatedWord = Convert(r)
    Input #f, r: Debug.Print r: LangDeactivatedWord = Convert(r)
    Input #f, r: Debug.Print r: FrmColor.Caption = Convert(r): LangBackcolorWord = Convert(r)
    Input #f, r: Debug.Print r: FrmColor.Option1(1).Caption = Convert(r): LangDefaultWord = Convert(r)
    Input #f, r: Debug.Print r: FrmColor.Option1(0).Caption = Convert(r): LangCustomWord = Convert(r)
    Input #f, r: Debug.Print r: LangRedWord = Convert(r)
    Input #f, r: Debug.Print r: LangGreenWord = Convert(r)
    Input #f, r: Debug.Print r: LangBlueWord = Convert(r)
    Input #f, r: Debug.Print r: LangMixerMenuWord = Convert(r)
    Input #f, r: Debug.Print r: LangMoveWord = Convert(r)
    Input #f, r: Debug.Print r: LangHideMixerWord = Convert(r)
    Input #f, r: Debug.Print r: LblName(MaxSources + 1).Caption = Convert(r)
    Input #f, r: Debug.Print r: LblName(MaxSources + 2).Caption = Convert(r)
    Input #f, r: Debug.Print r: LangReadoutWord = Convert(r)
    Input #f, r: Debug.Print r: LangQuickMixerWord = Convert(r)
badreadlang:
    Close #f
    If ToolTips Then Call Tips(True)
End Sub

Private Function Convert(k As Variant)
    Dim s%, l%
    s = InStr(k, "{") + ONE
    l = InStr(k, "}") - s
    Convert = Mid$(k, s, l)
    s = InStr(Convert, "@")
    If s > ZERO Then Mid$(Convert, s, ONE) = "&"
End Function

Private Sub DoProfile(Idx%)
    Dim p$, i%, k%
    i = InStr(Profile(Idx), "•")
    p = Right(Profile(Idx), Len(Profile(Idx)) - i)
    For k = ZERO To MaxSources + 2
        SldrVol(k).Value = SLIDERMAX - (ONEPERCENT * Val(Mid(p, (k * 6) + 2, 3)))
        If Val(Mid(p, (k * 6) + 6, 1)) = 1 Then ChkMute(k).Value = vbChecked Else ChkMute(k).Value = vbUnchecked
    Next k
End Sub
