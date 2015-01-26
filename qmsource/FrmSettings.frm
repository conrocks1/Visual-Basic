VERSION 5.00
Begin VB.Form FrmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "General Settings"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   174
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   75
      TabIndex        =   0
      Top             =   2220
      Width           =   2115
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Re&verse Treb/Bass Logic"
      Height          =   195
      Index           =   9
      Left            =   75
      TabIndex        =   10
      Top             =   1950
      Width           =   2520
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show Treb/&Bass Sliders"
      Height          =   195
      Index           =   8
      Left            =   75
      TabIndex        =   9
      Top             =   1740
      Width           =   2520
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show &Tool-Tips"
      Height          =   195
      Index           =   7
      Left            =   75
      TabIndex        =   8
      Top             =   1530
      Width           =   2520
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show S&kin"
      Height          =   195
      Index           =   6
      Left            =   75
      TabIndex        =   7
      Top             =   1320
      Width           =   2520
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show &Graduations"
      Height          =   195
      Index           =   5
      Left            =   75
      TabIndex        =   6
      Top             =   1110
      Width           =   2520
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Pointed Sliders"
      Height          =   195
      Index           =   4
      Left            =   75
      TabIndex        =   5
      Top             =   900
      Width           =   2520
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Single Slide&r Mode"
      Height          =   195
      Index           =   3
      Left            =   75
      TabIndex        =   4
      Top             =   690
      Width           =   2520
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Space Sliders &Evenly"
      Height          =   195
      Index           =   2
      Left            =   75
      TabIndex        =   3
      Top             =   480
      Width           =   2520
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Always &On-Top"
      Height          =   195
      Index           =   1
      Left            =   75
      TabIndex        =   2
      Top             =   270
      Width           =   2520
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Auto-Hide"
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   1
      Top             =   60
      Width           =   2520
   End
End
Attribute VB_Name = "FrmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private AntiLoop As Boolean

Private Sub Check1_Click(Index As Integer)
    If AntiLoop Then Exit Sub
    BackToSettings = True
    Select Case Index
        Case 0: FrmMxr.MnuAutoHide_Click: If AutoHide Then FrmMxr.Hide: MixerVisible = False: FrmMxr.MnuRestoreMixer.Enabled = True: FrmMxr.MnuHideMixer.Enabled = False
        Case 1: FrmMxr.MnuAlwaysOnTop_Click
        Case 2: FrmMxr.MnuSnapSlidersEvenly_Click
        Case 3: FrmMxr.MnuSingleSliderMode_Click
        Case 4: FrmMxr.MnuPointedSliders_Click
        Case 5: FrmMxr.MnuShowGraduations_Click
        Case 6: FrmMxr.MnuShowSkin_Click
        Case 7: FrmMxr.MnuShowToolTips_Click
        Case 8: FrmMxr.MnuShowTB_Click
        Case 9: FrmMxr.mnuReverseLogicTrebBass_Click
    End Select
   BackToSettings = False
End Sub

Private Sub Check1_GotFocus(Index As Integer)
    Check1(Index).BackColor = &H8000000D
    Check1(Index).ForeColor = &H8000000E
    Dim z%
    For z = 0 To Check1.UBound
        If z <> Index Then
            Check1_LostFocus z
        End If
    Next z
End Sub

Private Sub Check1_LostFocus(Index As Integer)
    Check1(Index).BackColor = &H8000000F
    Check1(Index).ForeColor = &H80000012
End Sub

Private Sub Check1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim z%
    For z = 0 To Check1.UBound
        If z = Index Then
            Check1(z).BackColor = &H8000000D
            Check1(z).ForeColor = &H8000000E
        Else
            Check1(z).BackColor = &H8000000F
            Check1(z).ForeColor = &H80000012
        End If
    Next z
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub Form_Activate()
    Dim t$, i%, f$, o$
    
    t = FrmMxr.MnuSettings.Caption
    For i = 1 To Len(t)
    f = Mid$(t, i, 1): If f <> "&" Then o = o & f
    Next i: Me.Caption = o
    Command1.Caption = LangOkWord
    Check1(0).Caption = FrmMxr.MnuAutoHide.Caption
    Check1(1).Caption = FrmMxr.MnuAlwaysOnTop.Caption
    Check1(2).Caption = FrmMxr.MnuSnapSlidersEvenly.Caption
    Check1(3).Caption = FrmMxr.MnuSingleSliderMode.Caption
    Check1(4).Caption = FrmMxr.MnuPointedSliders.Caption
    Check1(5).Caption = FrmMxr.MnuShowGraduations.Caption
    Check1(6).Caption = FrmMxr.MnuShowSkin.Caption
    Check1(7).Caption = FrmMxr.MnuShowToolTips.Caption
    Check1(8).Caption = FrmMxr.MnuShowTB.Caption
    Check1(9).Caption = FrmMxr.MnuReverseLogicTrebBass.Caption
    If Not TBSupport Then Check1(8).Visible = False: Check1(9).Visible = False
    DoChecks
End Sub

Public Sub DoChecks()
    Dim k%, ZB As Boolean
    For k = 0 To Check1.UBound
        Select Case k
            Case 0: ZB = AutoHide
            Case 1: ZB = OnTop
            Case 2: ZB = SnapSliders
            Case 3: ZB = SingleSliderMode
            Case 4: ZB = PointedSliders
            Case 5: ZB = ShowGrads
            Case 6: ZB = ShowSkin
            Case 7: ZB = ToolTips
            Case 8: ZB = ShowTB
            Case 9: ZB = ReverseLogicTrebBass
        End Select
    AntiLoop = True
    If ZB Then Check1(k).Value = vbChecked Else Check1(k).Value = vbUnchecked
    AntiLoop = False
    Next k
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim k%
    For k = 0 To Check1.UBound
        Check1(k).BackColor = &H8000000F
        Check1(k).ForeColor = &H80000012
    Next k
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.Visible = False Then Exit Sub
    FrmSettingsLeft = FrmSettings.Left
    FrmSettingsTop = FrmSettings.Top
End Sub
