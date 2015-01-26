VERSION 5.00
Begin VB.Form FrmControlKeysInfo 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Quick Mixer's Control Keys                                                          (Press ESCAPE to exit)"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8295
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Monotype.com"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   157
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   553
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar VScroll1 
      Height          =   2355
      Left            =   8040
      Max             =   47
      TabIndex        =   1
      Top             =   0
      Width           =   270
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   12000
      Left            =   150
      TabIndex        =   0
      Top             =   0
      Width           =   7770
   End
End
Attribute VB_Name = "FrmControlKeysInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Or KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Or KeyCode = vbKeyX Then Unload Me
End Sub

Private Sub Form_Load()
Label1.Caption = "TAB ................ Rotates the selected slider and device-icon forward." & vbCrLf
Label1.Caption = Label1.Caption & "SHIFT+TAB .......... Rotates the selected slider and device-icon backward." & vbCrLf
Label1.Caption = Label1.Caption & "UP ARROW ........... Turns up volume of selected slider." & vbCrLf
Label1.Caption = Label1.Caption & "DOWN ARROW ......... Turns down volume of selected slider." & vbCrLf
Label1.Caption = Label1.Caption & "PAGE UP ............ Turns up volume of selected slider quickly." & vbCrLf
Label1.Caption = Label1.Caption & "PAGE DOWN .......... Turns down volume of selected slider quickly." & vbCrLf
Label1.Caption = Label1.Caption & "ESC ................ Escapes from the menu without doing anything." & vbCrLf
Label1.Caption = Label1.Caption & "SPACEBAR ........... Hides Quick Mixer." & vbCrLf
Label1.Caption = Label1.Caption & "SHIFT+F4 ........... Close and exit Quick Mixer." & vbCrLf
Label1.Caption = Label1.Caption & "F4 ................. Audio profiles tool-window." & vbCrLf
Label1.Caption = Label1.Caption & "F5 ................. Runs WinAmp if installed in default location." & vbCrLf
Label1.Caption = Label1.Caption & "F6 ................. General settings multi-select tool-window." & vbCrLf
Label1.Caption = Label1.Caption & "F7 ................. Timed mute/unmute tool-window." & vbCrLf
Label1.Caption = Label1.Caption & "F8 ................. Background color tool-window." & vbCrLf
Label1.Caption = Label1.Caption & "F9 ................. Runs SNDVOL32.EXE if installed in default location." & vbCrLf
Label1.Caption = Label1.Caption & "F10 ................ Pops up menu where mouse is located." & vbCrLf
Label1.Caption = Label1.Caption & "F11 ................ Toggle master mute on/off." & vbCrLf
Label1.Caption = Label1.Caption & "F12 ................ Toggle special-control on/off. eg. stereo enhancement" & vbCrLf
Label1.Caption = Label1.Caption & "CONTROL+A .......... Toggle Auto-Hide mode on/off." & vbCrLf
Label1.Caption = Label1.Caption & "CONTROL+O .......... Toggle On-Top mode on/off." & vbCrLf
Label1.Caption = Label1.Caption & "CONTROL+E .......... Toggle Space Sliders Evenly mode on/off." & vbCrLf
Label1.Caption = Label1.Caption & "CONTROL+R .......... Toggle Single Slider Mode on/off." & vbCrLf
Label1.Caption = Label1.Caption & "CONTROL+P .......... Toggle Pointed Sliders mode on/off." & vbCrLf
Label1.Caption = Label1.Caption & "CONTROL+G .......... Toggle Show Graduations mode on/off." & vbCrLf
Label1.Caption = Label1.Caption & "CONTROL+K .......... Toggle Show Skin mode on/off." & vbCrLf
Label1.Caption = Label1.Caption & "CONTROL+T .......... Toggle Show Tool Tips mode on/off." & vbCrLf
Label1.Caption = Label1.Caption & "CONTROL+B .......... Toggle Show Treble/Bass Sliders mode on/off." & vbCrLf
Label1.Caption = Label1.Caption & "CONTROL+V .......... Toggle Reverse Treble/Bass Logic mode on/off." & vbCrLf

Label1.Caption = Label1.Caption & "ALT+0 .............. Load default audio profile ............ These   " & vbCrLf
Label1.Caption = Label1.Caption & "ALT+1 .............. Load user audio profile number 1 ...... also    " & vbCrLf
Label1.Caption = Label1.Caption & "ALT+2 .............. Load user audio profile number 2 ...... work    " & vbCrLf
Label1.Caption = Label1.Caption & "ALT+3 .............. Load user audio profile number 3 ...... with    " & vbCrLf
Label1.Caption = Label1.Caption & "ALT+4 .............. Load user audio profile number 4 ...... the ALT " & vbCrLf
Label1.Caption = Label1.Caption & "ALT+5 .............. Load user audio profile number 5 ...... NUMPAD  " & vbCrLf
Label1.Caption = Label1.Caption & "ALT+6 .............. Load user audio profile number 6 ...... keys,   " & vbCrLf
Label1.Caption = Label1.Caption & "ALT+7 .............. Load user audio profile number 7 ...... zero    " & vbCrLf
Label1.Caption = Label1.Caption & "ALT+8 .............. Load user audio profile number 1 ...... through " & vbCrLf
Label1.Caption = Label1.Caption & "ALT+9 .............. Load user audio profile number 9 ...... nine.   " & vbCrLf
Label1.Caption = Label1.Caption & "ALT+UP ARROW ....... Move Quick Mixer up one pixel." & vbCrLf
Label1.Caption = Label1.Caption & "ALT+DOWN ARROW ..... Move Quick Mixer down one pixel." & vbCrLf
Label1.Caption = Label1.Caption & "ALT+LEFT ARROW ..... Move Quick Mixer left one pixel." & vbCrLf
Label1.Caption = Label1.Caption & "ALT+RIGHT ARROW .... Move Quick Mixer right one pixel." & vbCrLf
Label1.Caption = Label1.Caption & "NUMPAD 1 ........... Move Quick Mixer down and to the left one pixel." & vbCrLf
Label1.Caption = Label1.Caption & "NUMPAD 2 ........... Move Quick Mixer down one pixel." & vbCrLf
Label1.Caption = Label1.Caption & "NUMPAD 3 ........... Move Quick Mixer down and to the right one pixel." & vbCrLf
Label1.Caption = Label1.Caption & "NUMPAD 4 ........... Move Quick Mixer to the left one pixel." & vbCrLf
Label1.Caption = Label1.Caption & "NUMPAD 6 ........... Move Quick Mixer to the right one pixel." & vbCrLf
Label1.Caption = Label1.Caption & "NUMPAD 7 ........... Move Quick Mixer up and to the left one pixel." & vbCrLf
Label1.Caption = Label1.Caption & "NUMPAD 8 ........... Move Quick Mixer up one pixel." & vbCrLf
Label1.Caption = Label1.Caption & "NUMPAD 9 ........... Move Quick Mixer up and to the right one pixel." & vbCrLf
Label1.Caption = Label1.Caption & "CONTROL+NUMPAD 1 ... Move Quick Mixer down and to the left ten pixels." & vbCrLf
Label1.Caption = Label1.Caption & "CONTROL+NUMPAD 2 ... Move Quick Mixer down ten pixels." & vbCrLf
Label1.Caption = Label1.Caption & "CONTROL+NUMPAD 3 ... Move Quick Mixer down and to the right ten pixels." & vbCrLf
Label1.Caption = Label1.Caption & "CONTROL+NUMPAD 4 ... Move Quick Mixer to the left ten pixels." & vbCrLf
Label1.Caption = Label1.Caption & "CONTROL+NUMPAD 6 ... Move Quick Mixer to the right ten pixels." & vbCrLf
Label1.Caption = Label1.Caption & "CONTROL+NUMPAD 7 ... Move Quick Mixer up and to the left ten pixels." & vbCrLf
Label1.Caption = Label1.Caption & "CONTROL+NUMPAD 8 ... Move Quick Mixer up ten pixels." & vbCrLf
Label1.Caption = Label1.Caption & "CONTORL+NUMPAD 9 ... Move Quick Mixer up and to the right ten pixels."
End Sub
Private Sub VScroll1_Change()
VScroll1_Scroll
End Sub
Private Sub VScroll1_Scroll()
Label1.Top = VScroll1.Value * -12
End Sub
