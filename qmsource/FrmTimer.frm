VERSION 5.00
Begin VB.Form FrmTimer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Timed Mute"
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
   Begin VB.Frame Frame1 
      Height          =   1080
      Index           =   1
      Left            =   75
      TabIndex        =   10
      Top             =   1065
      Width           =   2130
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   420
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   540
         Width           =   300
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000016&
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   300
         End
      End
      Begin VB.CheckBox ChkActivated 
         Caption         =   "Deactivated"
         Height          =   240
         Index           =   1
         Left            =   825
         TabIndex        =   6
         Top             =   420
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   90
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   540
         Width           =   300
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000016&
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   300
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   675
         Index           =   2
         Left            =   90
         Max             =   23
         TabIndex        =   4
         Top             =   300
         Value           =   23
         Width           =   300
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   675
         Index           =   3
         Left            =   420
         Max             =   59
         TabIndex        =   5
         Top             =   300
         Value           =   59
         Width           =   300
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   105
         Index           =   3
         Left            =   510
         Shape           =   3  'Circle
         Top             =   150
         Width           =   105
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   105
         Index           =   2
         Left            =   180
         Shape           =   3  'Circle
         Top             =   150
         Width           =   105
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Timed Unmute"
         Height          =   225
         Index           =   1
         Left            =   840
         TabIndex        =   8
         Top             =   165
         Width           =   1155
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000010&
         BorderWidth     =   2
         Height          =   240
         Index           =   1
         Left            =   922
         Top             =   750
         Width           =   1020
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         Caption         =   "12:34 PM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   945
         TabIndex        =   14
         Top             =   765
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000016&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   375
         TabIndex        =   12
         Top             =   540
         Width           =   45
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000016&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000016&
         BorderStyle     =   0  'Transparent
         Height          =   750
         Index           =   1
         Left            =   60
         Top             =   270
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1080
      Index           =   0
      Left            =   75
      TabIndex        =   9
      Top             =   -15
      Width           =   2130
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   90
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   540
         Width           =   300
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000016&
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   300
         End
      End
      Begin VB.CheckBox ChkActivated 
         Caption         =   "Deactivated"
         Height          =   240
         Index           =   0
         Left            =   825
         TabIndex        =   3
         Top             =   420
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   420
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   540
         Width           =   300
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000016&
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   300
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   675
         Index           =   0
         Left            =   90
         Max             =   23
         TabIndex        =   1
         Top             =   300
         Value           =   23
         Width           =   300
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   675
         Index           =   1
         Left            =   420
         Max             =   59
         TabIndex        =   2
         Top             =   300
         Value           =   59
         Width           =   300
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   105
         Index           =   1
         Left            =   510
         Shape           =   3  'Circle
         Top             =   150
         Width           =   105
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   105
         Index           =   0
         Left            =   180
         Shape           =   3  'Circle
         Top             =   150
         Width           =   105
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Timed Mute"
         Height          =   225
         Index           =   0
         Left            =   840
         TabIndex        =   7
         Top             =   165
         Width           =   1155
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000010&
         BorderWidth     =   2
         Height          =   240
         Index           =   0
         Left            =   922
         Top             =   750
         Width           =   1020
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         Caption         =   "12:34 PM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   945
         TabIndex        =   13
         Top             =   765
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000016&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   375
         TabIndex        =   18
         Top             =   540
         Width           =   45
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000016&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000016&
         BorderStyle     =   0  'Transparent
         Height          =   750
         Index           =   0
         Left            =   60
         Top             =   270
         Width           =   705
      End
   End
End
Attribute VB_Name = "FrmTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ZERO = 0, ONE = 1, TWO = 2, THREE = 3, TWELVE = 12, DECA = 10
Private Const MUTECOLOR = &HC0&, UNMUTECOLOR = &HC000&, DEACTIVATEDCOLOR = &H80000010
Private Const HOURS = 23, MINUTES = 59, PM = " PM", AM = " AM", SEP = ":", LZ = "0"

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    
    Me.Caption = LangTimedMuteWord
    Label4(0).Caption = LangTimedMuteWord
    Label4(1).Caption = LangTimedUnmuteWord
    Command1.Caption = LangOkWord
    
    
    VScroll1(ZERO).Value = HOURS - MuteHour
    VScroll1(ONE).Value = MINUTES - MuteMinute
    VScroll1(TWO).Value = HOURS - UnMuteHour
    VScroll1(THREE).Value = MINUTES - UnMuteMinute
    If TMuteFlag Then
        ChkActivated(ZERO).Value = vbChecked
        ChkActivated(ZERO).Caption = LangActivatedWord
        Shape2(ZERO).BorderColor = MUTECOLOR
    Else
        ChkActivated(ZERO).Value = vbUnchecked
        ChkActivated(ZERO).Caption = LangDeactivatedWord
        Shape2(ZERO).BorderColor = DEACTIVATEDCOLOR
    End If
    If TUnMuteFlag Then
        ChkActivated(ONE).Value = vbChecked
        ChkActivated(ONE).Caption = LangActivatedWord
        Shape2(ONE).BorderColor = UNMUTECOLOR
    Else
        ChkActivated(ONE).Value = vbUnchecked
        ChkActivated(ONE).Caption = LangDeactivatedWord
        Shape2(ONE).BorderColor = DEACTIVATEDCOLOR
    End If
    Dim k%
    For k = ZERO To THREE
        VScroll1_Change k
    Next k
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.Visible = False Then Exit Sub
    FrmTimerLeft = FrmTimer.Left
    FrmTimerTop = FrmTimer.Top
End Sub

Private Sub Label1_Click(Index As Integer)
    Beep
End Sub

Private Sub Label3_Click(Index As Integer)
    Label4_Click Index
End Sub

Private Sub Label4_Click(Index As Integer)
    If ChkActivated(Index).Value = vbChecked Then ChkActivated(Index).Value = vbUnchecked Else ChkActivated(Index).Value = vbChecked
    ChkActivated_Click Index
End Sub

Private Sub VScroll1_Change(Index As Integer)
    Dim a%, b$
    a = VScroll1(Index).Max - VScroll1(Index).Value
    b = Right$(Str$(a), Len(Str$(a)) - ONE)
    If Len(b) < TWO Then b = LZ & b
    Label1(Index).Caption = b
    Select Case Index
        Case Is = ZERO: MuteHour = Val(b)
        Case Is = ONE: MuteMinute = Val(b)
        Case Is = TWO: UnMuteHour = Val(b)
        Case Is = THREE: UnMuteMinute = Val(b)
    End Select
    Dim mh%, mm%, mp$, ms$, uh%, um%, up$, us$
    mh = MuteHour
    If mh = ZERO Then mh = TWELVE
    If mh > TWELVE Then
        mh = mh - TWELVE
        mp = PM
    Else
        mp = AM
    End If
    If MuteHour = TWELVE Then mp = PM
    mm = MuteMinute
    If mm < DECA Then ms = SEP & LZ Else ms = SEP
    uh = UnMuteHour
    If uh = ZERO Then uh = TWELVE
    If uh > TWELVE Then
        uh = uh - TWELVE
        up = PM
    Else
        up = AM
    End If
    If UnMuteHour = TWELVE Then up = PM
    um = UnMuteMinute
    If um < DECA Then us = SEP & LZ Else us = SEP
    Label3(ZERO).Caption = mh & ms & mm & mp
    Label3(ONE).Caption = uh & us & um & up
End Sub

Private Sub ChkActivated_Click(Index As Integer)
    If ChkActivated(Index).Value = vbChecked Then
        ChkActivated(Index).Caption = LangActivatedWord
    Else
        ChkActivated(Index).Caption = LangDeactivatedWord
    End If
    Select Case Index
        Case Is = ZERO
            If ChkActivated(Index).Value = vbChecked Then
                TMuteFlag = True
                Shape2(ZERO).BorderColor = MUTECOLOR
            Else
                TMuteFlag = False
                Shape2(ZERO).BorderColor = DEACTIVATEDCOLOR
            End If
        Case Is = ONE
            If ChkActivated(Index).Value = vbChecked Then
                TUnMuteFlag = True
                Shape2(ONE).BorderColor = UNMUTECOLOR
            Else
                TUnMuteFlag = False
                Shape2(ONE).BorderColor = DEACTIVATEDCOLOR
            End If
    End Select
    If TMuteFlag Or TUnMuteFlag Then FrmMxr.MnuTimedMute.Checked = vbChecked Else FrmMxr.MnuTimedMute.Checked = vbUnchecked
End Sub

Private Sub VScroll1_GotFocus(Index As Integer)
Shape3(Index).BackColor = &HFF00&
End Sub

Private Sub VScroll1_LostFocus(Index As Integer)
Shape3(Index).BackColor = &HC000&
End Sub
