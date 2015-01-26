VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmColor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Back Color"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2250
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   174
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1830
      Left            =   75
      TabIndex        =   10
      Top             =   315
      Width           =   2130
      Begin ComctlLib.Slider SldrRGB 
         Height          =   1275
         Index           =   0
         Left            =   315
         TabIndex        =   4
         Top             =   465
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   2249
         _Version        =   327682
         Orientation     =   1
         Min             =   -255
         Max             =   0
         SelStart        =   -192
         TickStyle       =   3
         TickFrequency   =   26
         Value           =   -192
      End
      Begin ComctlLib.Slider SldrRGB 
         Height          =   1275
         Index           =   2
         Left            =   1455
         TabIndex        =   8
         Top             =   465
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   2249
         _Version        =   327682
         Orientation     =   1
         Min             =   -255
         Max             =   0
         SelStart        =   -192
         TickStyle       =   3
         TickFrequency   =   26
         Value           =   -192
      End
      Begin ComctlLib.Slider SldrRGB 
         Height          =   1275
         Index           =   1
         Left            =   885
         TabIndex        =   6
         Top             =   465
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   2249
         _Version        =   327682
         Orientation     =   1
         Min             =   -255
         Max             =   0
         SelStart        =   -192
         TickStyle       =   3
         TickFrequency   =   26
         Value           =   -192
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "&Red"
         ForeColor       =   &H000000FF&
         Height          =   210
         Index           =   0
         Left            =   270
         TabIndex        =   3
         Top             =   210
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "&Green"
         ForeColor       =   &H00008000&
         Height          =   210
         Index           =   1
         Left            =   840
         TabIndex        =   5
         Top             =   210
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "&Blue"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   2
         Left            =   1410
         TabIndex        =   7
         Top             =   210
         Width           =   465
      End
   End
   Begin VB.PictureBox PicColor 
      Height          =   270
      Left            =   975
      ScaleHeight     =   210
      ScaleWidth      =   255
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   45
      Width           =   315
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Caption         =   "&Default"
      Height          =   210
      Index           =   1
      Left            =   75
      TabIndex        =   1
      Top             =   90
      Width           =   810
   End
   Begin VB.OptionButton Option1 
      Caption         =   "&Custom"
      Height          =   210
      Index           =   0
      Left            =   1350
      TabIndex        =   2
      Top             =   90
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   75
      TabIndex        =   0
      Top             =   2220
      Width           =   2115
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   -105
      Top             =   1065
   End
End
Attribute VB_Name = "FrmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ColorR, ColorG, ColorB, ColorAntiLoopFlag As Boolean

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.Caption = LangBackcolorWord
    Command1.Caption = LangOkWord
    Option1(0).Caption = LangCustomWord
    Option1(1).Caption = LangDefaultWord
    
    
    Dim FormColor, ColorWorkString$
    FormColor = FrmMxr.BackColor
    PicColor.BackColor = FormColor
    If FormColor <> &H8000000F Then SldrRGB(0).Enabled = True: SldrRGB(1).Enabled = True: SldrRGB(2).Enabled = True
    If FormColor = &H8000000F Then Option1(1).Value = True: Exit Sub Else Option1(0).Value = True
    ColorWorkString = Right$("000000" & Hex$(FormColor), 6): ColorR = Val("&h" & Right$(ColorWorkString, 2))
    ColorWorkString = Right$("000000" & Hex$(FormColor), 6): ColorB = Val("&h" & Left$(ColorWorkString, 2))
    ColorWorkString = Right$("000000" & Hex$(FormColor), 6): ColorG = Val("&h" & Mid$(ColorWorkString, 3, 2))
    SldrRGB(0).Value = ColorR * -1: SldrRGB(1).Value = ColorG * -1: SldrRGB(2).Value = ColorB * -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    If Me.Visible = False Then Exit Sub 'Don't save window position if window isn't visible!
    FrmColorLeft = FrmColor.Left
    FrmColorTop = FrmColor.Top
End Sub

Private Sub Label1_Click(Index As Integer)
    SldrRGB(Index).SetFocus
End Sub

Private Sub Option1_Click(Index As Integer)
    Dim k%
    Select Case Index
        Case 0 'CUSTOM 'These are backwards on the form, so it's a little confusing...
            SldrRGB(0).Enabled = True: SldrRGB(1).Enabled = True: SldrRGB(2).Enabled = True
            Label1(0).Enabled = True: Label1(1).Enabled = True: Label1(2).Enabled = True
            Label1(0).Caption = "&" & LangRedWord: Label1(1).Caption = "&" & LangGreenWord: Label1(2).Caption = "&" & LangBlueWord
            For k = 0 To 2: SldrRGB_Scroll (k): Next k
            FrmMxr.MnuBackColor.Checked = vbChecked
        Case 1 'DEFAULT
            SldrRGB(0).Enabled = False: SldrRGB(1).Enabled = False: SldrRGB(2).Enabled = False
            Label1(0).Enabled = False: Label1(1).Enabled = False: Label1(2).Enabled = False
            Label1(0).Caption = LangRedWord: Label1(1).Caption = LangGreenWord: Label1(2).Caption = LangBlueWord
            PicColor.BackColor = &H8000000F: FrmMxr.BackColor = &H8000000F: MixerBackColor = &H8000000F
            FrmMxr.MnuBackColor.Checked = vbUnchecked
            ColorAntiLoopFlag = True
            SldrRGB(0).Value = -192
            ColorAntiLoopFlag = True
            SldrRGB(1).Value = -192
            ColorAntiLoopFlag = True
            SldrRGB(2).Value = -192
    End Select
    ColorAntiLoopFlag = False: FrmMxr.Form_Resize: FrmMxr.Form_Paint
End Sub

Private Sub SldrRGB_Change(Index As Integer)
    SldrRGB_Scroll Index
End Sub

Private Sub SldrRGB_Scroll(Index As Integer)
    If ColorAntiLoopFlag Then Exit Sub
    Select Case Index
        Case 0: ColorR = Abs(SldrRGB(Index).Value)
        Case 1: ColorG = Abs(SldrRGB(Index).Value)
        Case 2: ColorB = Abs(SldrRGB(Index).Value)
    End Select
     PicColor.BackColor = RGB(ColorR, ColorG, ColorB)
     FrmMxr.BackColor = RGB(ColorR, ColorG, ColorB)
     MixerBackColor = RGB(ColorR, ColorG, ColorB)
     Timer1.Enabled = False 'This will wait until 1/2 second AFTER the user is done sliding the scroller and
     Timer1.Enabled = True  'refresh things in case skin or graduations are turned on!
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    FrmMxr.Form_Resize: FrmMxr.Form_Paint
End Sub
