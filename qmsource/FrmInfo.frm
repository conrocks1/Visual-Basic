VERSION 5.00
Begin VB.Form FrmInfo 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   253
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1770
      Left            =   195
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "FrmInfo.frx":0000
      Top             =   1005
      Width           =   4305
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&List all control-keys"
      Height          =   345
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   2925
      Width           =   1545
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Feedback"
      Height          =   345
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   3375
      Width           =   1035
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   45
      Top             =   30
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Home &Page"
      Height          =   345
      Index           =   2
      Left            =   1260
      TabIndex        =   2
      Top             =   3375
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Help"
      Height          =   345
      Index           =   1
      Left            =   2400
      TabIndex        =   1
      Top             =   3375
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   345
      Index           =   0
      Left            =   3540
      TabIndex        =   0
      Top             =   3375
      Width           =   1035
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Don't show this message again."
      Height          =   255
      Left            =   1905
      TabIndex        =   5
      Top             =   2970
      Width           =   2595
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000016&
      Height          =   1875
      Index           =   1
      Left            =   120
      Top             =   945
      Width           =   4440
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000010&
      Height          =   1875
      Index           =   0
      Left            =   135
      Top             =   960
      Width           =   4440
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000016&
      Height          =   330
      Index           =   1
      Left            =   1770
      Top             =   2925
      Width           =   2790
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000010&
      Height          =   330
      Index           =   0
      Left            =   1785
      Top             =   2940
      Width           =   2790
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "<Generated Version Info>"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1035
      TabIndex        =   7
      Top             =   390
      Width           =   2655
   End
   Begin VB.Image Image2 
      Height          =   435
      Left            =   540
      Picture         =   "FrmInfo.frx":01F0
      Top             =   45
      Width           =   3645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "I m p o r t a n t     I n f o r m a t i o n"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   217
      TabIndex        =   6
      Top             =   600
      Width           =   4290
   End
End
Attribute VB_Name = "FrmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AntiLoop As Boolean

Private Sub Check1_Click()
    If AntiLoop Then Exit Sub
    Timer1.Enabled = False: Timer1.Enabled = True
    If Check1.Value = vbChecked Then DontShowInfo = True Else DontShowInfo = False
End Sub

Private Sub Command1_Click(Index As Integer)
    Timer1.Enabled = False
    Select Case Index
        Case 0: Me.Visible = False
        Case 1: Call FrmMxr.MnuHelp_Click
        Case 2: Call FrmMxr.MnuHomepage_Click
        Case 3: Call FrmMxr.MnuEmail_Click
        Case 4: Call FrmMxr.MnuControlKeysInfo_Click
    End Select
End Sub

Private Sub Form_Load()
    Label3.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision & " by Evan Edwards."
    AntiLoop = True
    If DontShowInfo Then Check1.Value = vbChecked Else Check1.Value = vbUnchecked
    AntiLoop = False
    Timer1.Enabled = True
End Sub



Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = False: Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Me.Visible = False
End Sub
