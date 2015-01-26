VERSION 5.00
Begin VB.Form FrmAudioProfiles 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Audio Profiles"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4470
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   233
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   298
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Start QM with default profile"
      Height          =   435
      Left            =   3150
      TabIndex        =   2
      ToolTipText     =   "Upon starting, Quick Mixer will load the Default Profile if this is checked."
      Top             =   3060
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   330
      Left            =   75
      TabIndex        =   0
      Top             =   3120
      Width           =   2115
   End
   Begin VB.CommandButton Command2 
      Caption         =   "su"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2295
      TabIndex        =   1
      ToolTipText     =   "Help with saving, loading, and renaming profiles."
      Top             =   3120
      Width           =   765
   End
   Begin VB.TextBox TxtName 
      BackColor       =   &H00FFF2F2&
      Height          =   285
      Index           =   9
      Left            =   1230
      TabIndex        =   31
      Top             =   2760
      Width           =   2565
   End
   Begin VB.TextBox TxtName 
      BackColor       =   &H00FFF2F2&
      Height          =   285
      Index           =   8
      Left            =   1230
      TabIndex        =   28
      Top             =   2460
      Width           =   2565
   End
   Begin VB.TextBox TxtName 
      BackColor       =   &H00FFF2F2&
      Height          =   285
      Index           =   7
      Left            =   1230
      TabIndex        =   25
      Top             =   2160
      Width           =   2565
   End
   Begin VB.TextBox TxtName 
      BackColor       =   &H00FFF2F2&
      Height          =   285
      Index           =   6
      Left            =   1230
      TabIndex        =   22
      Top             =   1860
      Width           =   2565
   End
   Begin VB.TextBox TxtName 
      BackColor       =   &H00FFF2F2&
      Height          =   285
      Index           =   5
      Left            =   1230
      TabIndex        =   19
      Top             =   1560
      Width           =   2565
   End
   Begin VB.TextBox TxtName 
      BackColor       =   &H00FFF2F2&
      Height          =   285
      Index           =   4
      Left            =   1230
      TabIndex        =   16
      Top             =   1260
      Width           =   2565
   End
   Begin VB.TextBox TxtName 
      BackColor       =   &H00FFF2F2&
      Height          =   285
      Index           =   3
      Left            =   1230
      TabIndex        =   13
      Top             =   960
      Width           =   2565
   End
   Begin VB.TextBox TxtName 
      BackColor       =   &H00FFF2F2&
      Height          =   285
      Index           =   2
      Left            =   1230
      TabIndex        =   10
      Top             =   660
      Width           =   2565
   End
   Begin VB.TextBox TxtName 
      BackColor       =   &H00FFF2F2&
      Height          =   285
      Index           =   1
      Left            =   1230
      TabIndex        =   7
      Top             =   360
      Width           =   2565
   End
   Begin VB.TextBox TxtName 
      BackColor       =   &H00F1F0FF&
      Height          =   285
      Index           =   0
      Left            =   1230
      TabIndex        =   4
      Top             =   60
      Width           =   2565
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      Height          =   270
      Index           =   9
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2760
      Width           =   585
   End
   Begin VB.CommandButton CmdUse 
      Caption         =   "Load"
      Height          =   270
      Index           =   9
      Left            =   615
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2760
      Width           =   585
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      Height          =   270
      Index           =   8
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2460
      Width           =   585
   End
   Begin VB.CommandButton CmdUse 
      Caption         =   "Load"
      Height          =   270
      Index           =   8
      Left            =   615
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2460
      Width           =   585
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      Height          =   270
      Index           =   7
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2160
      Width           =   585
   End
   Begin VB.CommandButton CmdUse 
      Caption         =   "Load"
      Height          =   270
      Index           =   7
      Left            =   615
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2160
      Width           =   585
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      Height          =   270
      Index           =   6
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1860
      Width           =   585
   End
   Begin VB.CommandButton CmdUse 
      Caption         =   "Load"
      Height          =   270
      Index           =   6
      Left            =   615
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1860
      Width           =   585
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      Height          =   270
      Index           =   5
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1560
      Width           =   585
   End
   Begin VB.CommandButton CmdUse 
      Caption         =   "Load"
      Height          =   270
      Index           =   5
      Left            =   615
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1560
      Width           =   585
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      Height          =   270
      Index           =   4
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1260
      Width           =   585
   End
   Begin VB.CommandButton CmdUse 
      Caption         =   "Load"
      Height          =   270
      Index           =   4
      Left            =   615
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1260
      Width           =   585
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      Height          =   270
      Index           =   3
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   960
      Width           =   585
   End
   Begin VB.CommandButton CmdUse 
      Caption         =   "Load"
      Height          =   270
      Index           =   3
      Left            =   615
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   960
      Width           =   585
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      Height          =   270
      Index           =   2
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   660
      Width           =   585
   End
   Begin VB.CommandButton CmdUse 
      Caption         =   "Load"
      Height          =   270
      Index           =   2
      Left            =   615
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   660
      Width           =   585
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      Height          =   270
      Index           =   1
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   360
      Width           =   585
   End
   Begin VB.CommandButton CmdUse 
      Caption         =   "Load"
      Height          =   270
      Index           =   1
      Left            =   615
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   360
      Width           =   585
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      Height          =   270
      Index           =   0
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   60
      Width           =   585
   End
   Begin VB.CommandButton CmdUse 
      Caption         =   "Load"
      Height          =   270
      Index           =   0
      Left            =   615
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   60
      Width           =   585
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3555
      Left            =   -90
      TabIndex        =   40
      Top             =   -45
      Width           =   4830
      Begin VB.Label Label1 
         Caption         =   "ALT+0"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   50
         Top             =   150
         Width           =   510
      End
      Begin VB.Label Label1 
         Caption         =   "ALT+1"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   49
         Top             =   450
         Width           =   510
      End
      Begin VB.Label Label1 
         Caption         =   "ALT+2"
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   48
         Top             =   750
         Width           =   510
      End
      Begin VB.Label Label1 
         Caption         =   "ALT+3"
         Height          =   225
         Index           =   3
         Left            =   150
         TabIndex        =   47
         Top             =   1050
         Width           =   510
      End
      Begin VB.Label Label1 
         Caption         =   "ALT+4"
         Height          =   225
         Index           =   4
         Left            =   150
         TabIndex        =   46
         Top             =   1350
         Width           =   510
      End
      Begin VB.Label Label1 
         Caption         =   "ALT+5"
         Height          =   225
         Index           =   5
         Left            =   150
         TabIndex        =   45
         Top             =   1650
         Width           =   510
      End
      Begin VB.Label Label1 
         Caption         =   "ALT+6"
         Height          =   225
         Index           =   6
         Left            =   150
         TabIndex        =   44
         Top             =   1950
         Width           =   510
      End
      Begin VB.Label Label1 
         Caption         =   "ALT+7"
         Height          =   225
         Index           =   7
         Left            =   150
         TabIndex        =   43
         Top             =   2250
         Width           =   510
      End
      Begin VB.Label Label1 
         Caption         =   "ALT+8"
         Height          =   225
         Index           =   8
         Left            =   150
         TabIndex        =   42
         Top             =   2550
         Width           =   510
      End
      Begin VB.Label Label1 
         Caption         =   "ALT+9"
         Height          =   225
         Index           =   9
         Left            =   150
         TabIndex        =   41
         Top             =   2850
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3150
      Left            =   60
      TabIndex        =   33
      Top             =   615
      Width           =   4470
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000010&
         Height          =   900
         Index           =   4
         Left            =   15
         Top             =   1935
         Width           =   4350
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000010&
         Height          =   900
         Index           =   2
         Left            =   15
         Top             =   975
         Width           =   4350
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000010&
         Height          =   900
         Index           =   0
         Left            =   15
         Top             =   15
         Width           =   4350
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000016&
         Height          =   900
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   4350
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmAudioProfiles.frx":0000
         Height          =   630
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   4125
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Saving Profiles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   915
         TabIndex        =   38
         Top             =   30
         Width           =   2580
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmAudioProfiles.frx":00AB
         Height          =   630
         Index           =   3
         Left            =   120
         TabIndex        =   37
         Top             =   1200
         Width           =   4125
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Loading Profiles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   915
         TabIndex        =   36
         Top             =   990
         Width           =   2580
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmAudioProfiles.frx":0149
         Height          =   630
         Index           =   4
         Left            =   120
         TabIndex        =   35
         Top             =   2160
         Width           =   4125
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Renaming Profiles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   915
         TabIndex        =   34
         Top             =   1950
         Width           =   2580
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000016&
         Height          =   900
         Index           =   5
         Left            =   0
         Top             =   1920
         Width           =   4350
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000016&
         Height          =   900
         Index           =   3
         Left            =   0
         Top             =   960
         Width           =   4350
      End
   End
End
Attribute VB_Name = "FrmAudioProfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MAX_CHARS = 35
Dim ProfileAntiLoop As Boolean

Private Sub Check1_Click()
    If ProfileAntiLoop Then Exit Sub
    StartWithProfile = Not StartWithProfile
    If StartWithProfile Then FrmMxr.MnuAudioProfiles.Checked = True Else FrmMxr.MnuAudioProfiles.Checked = False
    Command1.SetFocus
End Sub

Private Sub CmdSave_Click(Idx%)
    Command1.SetFocus
    If TxtName(Idx).Text = "" Then If Idx = 0 Then TxtName(Idx) = "Default Profile" Else TxtName(Idx) = "User Profile " & Idx
    If Len(TxtName(Idx)) > MAX_CHARS Then TxtName(Idx).Text = Left(TxtName(Idx), MAX_CHARS)
    If InStr("-.+0123456789", Left(TxtName(Idx), 1)) Then TxtName(Idx) = "_" & TxtName(Idx)
    Profile(Idx) = TxtName(Idx).Text & "•"
    Dim k%, v$, m%
    For k = 0 To MaxSources + 2
        v = FrmMxr.lblVol(k).Caption
        If Len(v) < 3 Then v = "0" & v
        If Len(v) < 3 Then v = "0" & v
        Profile(Idx) = Profile(Idx) & "·" & v
        If MixerState(k).MxrMute <> 0 Then m = 1 Else m = 0
        Profile(Idx) = Profile(Idx) & "·" & m
    Next k
    For k = 192 To 255
        CmdSave(Idx).BackColor = RGB(k, 192, 192)
        DoEvents
        CmdSave(Idx).BackColor = RGB(k, 192, 192)
        DoEvents
    Next k
    For k = 255 To 192 Step -1
        CmdSave(Idx).BackColor = RGB(k, 192, 192)
        DoEvents
        CmdSave(Idx).BackColor = RGB(k, 192, 192)
        DoEvents
    Next k
    CmdSave(Idx).BackColor = &H8000000F
End Sub

Private Sub CmdUse_Click(Idx%)
    Command1.SetFocus
    Dim p$, i%, k%
    i = InStr(Profile(Idx), "•")
    p = Right(Profile(Idx), Len(Profile(Idx)) - i)
    For k = 0 To MaxSources + 2
        FrmMxr.SldrVol(k).Value = 65535 - (655 * Val(Mid(p, (k * 6) + 2, 3)))
        If Val(Mid(p, (k * 6) + 6, 1)) = 1 Then FrmMxr.ChkMute(k).Value = vbChecked Else FrmMxr.ChkMute(k).Value = vbUnchecked
    Next k
    For k = 192 To 255
        CmdUse(Idx).BackColor = RGB(192, k, 192)
        DoEvents
        CmdUse(Idx).BackColor = RGB(192, k, 192)
        DoEvents
    Next k
    For k = 255 To 192 Step -1
        CmdUse(Idx).BackColor = RGB(192, k, 192)
        DoEvents
        CmdUse(Idx).BackColor = RGB(192, k, 192)
        DoEvents
    Next k
    CmdUse(Idx).BackColor = &H8000000F
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Command1.SetFocus
    Dim w%
    Command2.Enabled = False
    If Command2.Caption = "su" Then
        For w = 1 To 97
            Me.Height = Me.Height + (Screen.TwipsPerPixelY * 2)
            DoEvents
            Frame1.Top = Frame1.Top + 2
            DoEvents
        Next
        Command2.Caption = "st"
    Else
        For w = 1 To 97
            Me.Height = Me.Height - (Screen.TwipsPerPixelY * 2)
            DoEvents
            Frame1.Top = Frame1.Top - 2
            DoEvents
        Next
        Command2.Caption = "su"
    End If
    Command2.Enabled = True
End Sub

Private Sub Form_Activate()
    Me.Height = 3825
    Frame1.Top = 41
    Command2.Caption = "su"
    Command1.Caption = LangOkWord
    ProfileAntiLoop = True
    If StartWithProfile Then Check1.Value = vbChecked Else Check1.Value = vbUnchecked
    ProfileAntiLoop = False
    Dim k%, i%
    For k = 0 To 9
        i = InStr(Profile(k), "•") - 1
        TxtName(k).Text = Left(Profile(k), i)
    Next k
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift <> 4 Then Exit Sub
    If KeyCode = vbKey0 Or KeyCode = vbKeyNumpad0 Then CmdUse_Click (0)
    If KeyCode = vbKey1 Or KeyCode = vbKeyNumpad1 Then CmdUse_Click (1)
    If KeyCode = vbKey2 Or KeyCode = vbKeyNumpad2 Then CmdUse_Click (2)
    If KeyCode = vbKey3 Or KeyCode = vbKeyNumpad3 Then CmdUse_Click (3)
    If KeyCode = vbKey4 Or KeyCode = vbKeyNumpad4 Then CmdUse_Click (4)
    If KeyCode = vbKey5 Or KeyCode = vbKeyNumpad5 Then CmdUse_Click (5)
    If KeyCode = vbKey6 Or KeyCode = vbKeyNumpad6 Then CmdUse_Click (6)
    If KeyCode = vbKey7 Or KeyCode = vbKeyNumpad7 Then CmdUse_Click (7)
    If KeyCode = vbKey8 Or KeyCode = vbKeyNumpad8 Then CmdUse_Click (8)
    If KeyCode = vbKey9 Or KeyCode = vbKeyNumpad9 Then CmdUse_Click (9)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.Visible = False Then Exit Sub
    FrmAudioProfilesLeft = FrmAudioProfiles.Left
    FrmAudioProfilesTop = FrmAudioProfiles.Top
End Sub

Private Sub TxtName_KeyPress(Idx%, KeyAscii%)
    If InStr(",·•|\^&:;", Chr$(KeyAscii)) Then KeyAscii = 0: Beep: Effect Idx: Exit Sub
    If KeyAscii = 34 Or KeyAscii = 13 Or KeyAscii = 10 Then KeyAscii = 0: Beep: Effect Idx: Exit Sub
    If Len(TxtName(Idx).Text) > MAX_CHARS Then KeyAscii = 8: Beep: Effect Idx: Exit Sub
End Sub

Private Sub Effect(Idx%)
    Dim c, k%
    c = TxtName(Idx).BackColor
    For k = 255 To 0 Step -1
        TxtName(Idx).BackColor = RGB(255, k, k)
        DoEvents
    Next k
    For k = 0 To 255
        TxtName(Idx).BackColor = RGB(255, k, k)
        DoEvents
    Next k
    TxtName(Idx).BackColor = c
End Sub
