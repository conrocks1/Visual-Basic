Attribute VB_Name = "ModMain"
Option Explicit

'Active app sensing section
Public OldWindowProc As Long

'These two functions are also used by the tray functionality section
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_WNDPROC = (-4) 'GWL_WNDPROC also used in tray functionality section
Public Const WM_ACTIVATE = &H6
Public Const WA_ACTIVE = 1
Public Const WA_CLICKACTIVE = 2
Public Const WA_INACTIVE = 0

'Add tray functionality
Public TrayOldWindowProc As Long
Public TheForm As Form
Public TheMenu As Menu
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Const WM_USER = &H400
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONUP = &H208
Public Const WM_RBUTTONUP = &H205
Public Const TRAY_CALLBACK = (WM_USER + 1001&)
Public Const GWL_USERDATA = (-21)
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIF_MESSAGE = &H1
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private TheData As NOTIFYICONDATA

'used for getting details about any specific mixer control...  Level, mute, 'advanced' controls, whatever...
Public MCD As MIXERCONTROLDETAILS

Private ML As MIXERLINE

'this type makes coding simpler
Type MIXERSETTINGS
     MxrChannels As Long    ' Indicates Whether A Line Is Mono Or Stereo.
     MxrLeftVol As Long       ' Left Volume Value (Balance).
     MxrRightVol As Long     ' Right Volume Value (Balance).
     MxrVol As Long             ' Fader Volume.
     MxrVolID As Long         ' Fader Control ID.
     MxrMute As Long          ' Mute Status.
     MxrMuteID As Long      ' Mute Control ID.
     MxrPeakID As Long      ' Peak Meter ID. 'Not presently implemented, but there.
     MxrOnOff As Long           'On/off status
     MxrOnOffID As Long        ' On/Off controls like stereo-enhancement, digital audio only, etc...
End Type

'make a public array of this type, dimension it later on...
Public MixerState() As MIXERSETTINGS

'for playing with memory
Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr&, ByVal cb&)
Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr&, struct As Any, ByVal cb&)
Declare Function GlobalAlloc& Lib "kernel32" (ByVal wFlags&, ByVal dwBytes&)
Declare Function GlobalFree& Lib "kernel32" (ByVal hmem&)
Declare Function GlobalLock& Lib "kernel32" (ByVal hmem&)
Declare Function GlobalUnlock& Lib "kernel32" (ByVal hmem&)

Public hMixer& 'handle of the mixer
Public ProductName$ 'mixer's product-name
Public MaxSources& 'number of SOURCES available
Private Destinations& 'number of DESTINATIONS available

'various setting toggles. program-wide variables, and flags
Public MixerVisible As Boolean, ShowGrads As Boolean, SnapSliders As Boolean, AutoHide As Boolean
Public OnTop As Boolean, ToolTips As Boolean, ShowSkin As Boolean, PointedSliders As Boolean
Public SingleSliderMode As Boolean, CurrentFocus%, ReverseLogicTrebBass As Boolean, ShowTB As Boolean
Public ChangeFlag As Boolean, TBSupport As Boolean, FrmTimerLeft%, FrmTimerTop%, Useage%
Public MuteHour%, MuteMinute%, UnMuteHour%, UnMuteMinute%, TMuteFlag As Boolean, TUnMuteFlag As Boolean
Public MixerBackColor&, FrmSettingsLeft%, FrmSettingsTop%, FrmColorLeft%, FrmColorTop%, FrmAudioProfilesLeft%, FrmAudioProfilesTop%
Public StartWithProfile As Boolean

Public Const AppName = "Quick Mixer" 'Application name, works better than app.name
'this puts the mixer off-screen while it does some messy initial resizing
Public Const HACKHIDE = 100000 'This constant represents a distance in the screen coordinate system and is not in pixels.
Private Const DEFAULTHEIGHT = 116, DEFAULTWIDTH = 460 'a good default size for the mixer
Private BeenThere As Boolean ' pares down processing a bit in main (TmrRefresh) loop... Doesn't hurt...
Public BackToSettings As Boolean
Public LangActivatedWord$, LangDeactivatedWord$, LangRedWord$, LangGreenWord$, LangBlueWord$
Public LangOkWord$, LangTimedMuteWord$, LangTimedUnmuteWord$, LangBackcolorWord$
Public LangDefaultWord$, LangCustomWord$
Public Profile$(0 To 9)
Public INI$, ISSM As Boolean, IssmWidth%, HideSlider(0 To 18) As Boolean, InvisibleControls%, DontShowInfo As Boolean

Private Sub Main()

    #If Win32 Then
        'these three conditions must be met in order for the mixer to run
        If Not MixerPresent Then End 'life sucks, no mixer
        If Not OpenMixer Then End 'life sucks, can't open the mixer
        If Not GetDeviceCapabilities Then End 'life sucks, can't talk with the mixer
        'This IS a 32-bit OS, there IS a mixer, it CAN be opened, and we CAN communicate with it.  So, let's start!
        'establish DEFAULTS in case an initialization profile has not yet been recorded.
        FrmMxr.Left = (Screen.Width - (DEFAULTWIDTH * Screen.TwipsPerPixelX)) / 2
        FrmMxr.Top = (Screen.Height - (DEFAULTHEIGHT * Screen.TwipsPerPixelY)) / 2
        FrmMxr.Width = DEFAULTWIDTH * Screen.TwipsPerPixelX
        FrmMxr.Height = DEFAULTHEIGHT * Screen.TwipsPerPixelY
        FrmTimer.Left = (Screen.Width - (FrmTimer.Width)) / 2
        FrmTimer.Top = (Screen.Height - (FrmTimer.Height)) / 2
        FrmSettings.Left = (Screen.Width - (FrmSettings.Width)) / 2
        FrmSettings.Top = (Screen.Height - (FrmSettings.Height)) / 2
        FrmColor.Left = (Screen.Width - (FrmColor.Width)) / 2
        FrmColor.Top = (Screen.Height - (FrmColor.Height)) / 2
        FrmAudioProfiles.Left = (Screen.Width - (FrmAudioProfiles.Width)) / 2
        FrmAudioProfiles.Top = (Screen.Height - (FrmAudioProfiles.Height)) / 2
        'establish DEFAULTS for setting toggles
        MixerVisible = True: ShowGrads = False: SnapSliders = False: AutoHide = False
        OnTop = False: ToolTips = True: ShowSkin = False: PointedSliders = False
        SingleSliderMode = False: ReverseLogicTrebBass = False: ShowTB = False
        FrmTimerLeft = FrmTimer.Left: FrmTimerTop = FrmTimer.Top
        MuteHour = 0: MuteMinute = 0: UnMuteHour = 0: UnMuteMinute = 0: TMuteFlag = False: TUnMuteFlag = False
        MixerBackColor = &H8000000F: FrmSettingsLeft = FrmSettings.Left: FrmSettingsTop = FrmSettings.Top
        FrmColorLeft = FrmColor.Left: FrmColorTop = FrmColor.Top
        FrmAudioProfilesLeft = FrmAudioProfiles.Left: FrmAudioProfilesTop = FrmAudioProfiles.Top
        Dim k%
        Profile(0) = "Default Profile暦000򈆤00򈆤00򈆤00򈆤00򈆤00򈆤00򈆤00򈆤00򈆤00򈆤00򈆤00򈆤00򈆤00򈆤00򈆤00򈆤00򈆤00򈆤00�0"
        For k = 1 To 9
        Profile(k) = "User Profile " & k & "暦000򈆤00򈆤00򈆤00򈆤00򈆤00򈆤00򈆤00򈆤00򈆤00򈆤00򈆤00򈆤00򈆤00򈆤00򈆤00򈆤00򈆤00򈆤00�0"
        Next k
        StartWithProfile = False
        INI = App.Path & "\QMixer.ini"
        Dim f%, r As Variant
        'try to read the .ini file
        On Error GoTo badreadini
        f = FreeFile
        Open INI For Input As #f
        Input #f, r 'reads the ini-file message line, does nothing with it
        Input #f, r: FrmMxr.Left = r
        Input #f, r: FrmMxr.Top = r
        Input #f, r: FrmMxr.Width = r: IssmWidth = r
        Input #f, r: FrmMxr.Height = r
        Input #f, r: MixerVisible = r
        Input #f, r: ShowGrads = r
        Input #f, r: SnapSliders = r
        Input #f, r: AutoHide = r
        Input #f, r: OnTop = r
        Input #f, r: ToolTips = r
        Input #f, r: Useage = r
        Input #f, r: ShowSkin = r
        Input #f, r: PointedSliders = r
        Input #f, r: SingleSliderMode = r: ISSM = r
        Input #f, r: CurrentFocus = r
        Input #f, r: ReverseLogicTrebBass = r
        Input #f, r: ShowTB = r
        Input #f, r: FrmTimerLeft = r
        Input #f, r: FrmTimerTop = r
        Input #f, r: MuteHour = r
        Input #f, r: MuteMinute = r
        Input #f, r: UnMuteHour = r
        Input #f, r: UnMuteMinute = r
        Input #f, r: TMuteFlag = r
        Input #f, r: TUnMuteFlag = r
        Input #f, r: MixerBackColor = r
        Input #f, r: FrmSettingsLeft = r
        Input #f, r: FrmSettingsTop = r
        Input #f, r: FrmColorLeft = r
        Input #f, r: FrmColorTop = r
        Input #f, r: FrmAudioProfilesLeft = r
        Input #f, r: FrmAudioProfilesTop = r
        Input #f, r: Profile(0) = r
        Input #f, r: Profile(1) = r
        Input #f, r: Profile(2) = r
        Input #f, r: Profile(3) = r
        Input #f, r: Profile(4) = r
        Input #f, r: Profile(5) = r
        Input #f, r: Profile(6) = r
        Input #f, r: Profile(7) = r
        Input #f, r: Profile(8) = r
        Input #f, r: Profile(9) = r
        Input #f, r: StartWithProfile = r
        For k = 0 To 18
        Input #f, r: HideSlider(k) = r
        Next k
        Input #f, r: InvisibleControls = r
        Input #f, r: DontShowInfo = r
badreadini:
        Close #f
        BeenThere = False 'for some items in the loop that only need processed one time
        GetMixerInfo 'find out all the mixer settings
        FrmMxr.Left = FrmMxr.Left + HACKHIDE 'hide the mixer off-screen, it's about to do a messy initial resize...
        FrmMxr.Show 'load-n-show the mixer form (off screen) and execution jumps right to FrmMxr_Form_Activate...
    #Else
        'life sucks, you don't have a 32-bit OS
        End
    #End If
End Sub

Private Function MixerPresent() As Boolean
    Dim msg$
    'mixerGetNumDevs API will check to see if there is a mixer
    If mixerGetNumDevs() Then
        MixerPresent = True 'life is great, we have a mixer.
    Else
        'life sucks, no mixer, this program is useless so stop running
        msg = AppName & " was unable to find a mixer-line on this computer."
        msg = msg & vbCrLf
        msg = msg & "Perhaps there is no sound-card, or it's drivers are not installed or configured properly."
        msg = msg & vbCrLf & vbCrLf
        msg = msg & AppName & ", being an audio-mixer, cannot run without a mixer-line present."
        MsgBox msg, vbCritical, AppName & " - Error."
    End If
End Function

Private Function OpenMixer() As Boolean
    Dim msg$
    'try to open the mixer
    'if it can be opened, then the global hMixer variable will be set to it's handle
    If mixerOpen(hMixer, 0, 0, 0, 0) = 0 Then
        OpenMixer = True   'life is grand, the mixer opened
    Else
        'life sucks, the mixer wouldn't open, this program is useless so stop running
        msg = AppName & " was not able to open the existing mixer-line."
        msg = msg & vbCrLf & vbCrLf
        msg = msg & "Closing..."
         MsgBox msg, vbCritical, AppName & " - Error."
    End If
End Function

Private Function GetDeviceCapabilities() As Boolean
    Dim msg$
    Dim MxrCaps As MIXERCAPS 'mixer capabilities struct
    'query the mixer's capabilities
    If mixerGetDevCaps(0, MxrCaps, Len(MxrCaps)) = 0 Then
        'only need to know the number of destinations and the product-name
        Destinations = MxrCaps.cDestinations - 1
        ProductName = Left(MxrCaps.szPname, InStr(MxrCaps.szPname, vbNullChar) - 1)
        'life is great, it worked
        GetDeviceCapabilities = True
    Else
        'life sucks, could not get the mixer's capabilities, stop running
        msg = "Cannot read mixer-line capabilities."
        msg = msg & vbCrLf & vbCrLf
        msg = msg & "Closing..."
        MsgBox msg, vbCritical, AppName & " - Error."
    End If
End Function

Public Sub GetMixerInfo()

    'Purpose: Scans The Destination's Until The Speaker's Are Found.
    'Then, Information About All Sources Connected To The Speaker's
    'Are Saved Into The "MixerState" Array For Use In The Main Form
    Dim Dst&, Src& ' Destination And Source Counter's.
    Dim ControlID& ' ID Of A Given Control.
    For Dst = 0 To Destinations
        ' Prep The MIXERLINE Structure.
        ML.cbStruct = Len(ML)
        ML.dwDestination = Dst
        'Get Destination Line Info.
        mixerGetLineInfo hMixer, ML, MIXER_GETLINEINFOF_DESTINATION
        'Was The Component Type The Speaker's?
        If ML.dwComponentType = MIXERLINE_COMPONENTTYPE_DST_SPEAKERS Then
            'How Many Item's Are Connected To The Speaker's?
            'gotta draw the line somewhere 16 is probably enough sources
            If ML.cConnections > 16 Then
                ML.cConnections = 16
                MaxSources = 16
            Else
                MaxSources = ML.cConnections  'Less Than 16.
            End If
            'Re-Dimension The "MixerState" Array.
            'Note: The Array Is Zero Based, Element Zero Is For The Master Voume
            'The Remaining Elements Are For The Source's.
            ReDim MixerState(MaxSources + 2) 'Thank the great cosmic powers for ReDim!
            
            'MASTER VOLUME (DESTINATION)
            'Save The Number Of Channels For The Master Volume.
            MixerState(0).MxrChannels = ML.cChannels
            'Update The Name Label On The Main Form.
           If Not BeenThere Then FrmMxr.LblName(0).Caption = Left(ML.szName, InStr(ML.szName, vbNullChar) - 1)
            'Call The "GetControlID" Function So We Can Get The Control ID
            'Of The Master Volume.
            ControlID = GetControlID(ML.dwComponentType, MIXERCONTROL_CONTROLTYPE_VOLUME)
            If ControlID <> 0 Then
                'Prep The MCD Structure For The Master Volume Fader.
                With MCD
                    .cbDetails = 4  'Size Of A Long In Byte's.
                    .cbStruct = 24
                    .cChannels = ML.cChannels
                    .dwControlID = ControlID
                    .item = 0
                    .paDetails = VarPtr(MixerState(0).MxrVol)
                End With
                'Get The Master Volume Setting.
                mixerGetControlDetails hMixer, MCD, MIXER_GETCONTROLDETAILSF_VALUE
                'Track Bar Logic Is The Reverse Of Fader's On A Hardware Mixer
                'So Reverse The Value.
                MixerState(0).MxrVol = 65535 - MixerState(0).MxrVol
                'Save The Master Volume Control ID.
                MixerState(0).MxrVolID = MCD.dwControlID
            Else
                'Couldn't Get It, Disable The Fader.
                FrmMxr.SldrVol(0).Enabled = 0
            End If
            
            'MASTER TREBLE (DESTINATION)
            MixerState(MaxSources + 1).MxrChannels = ML.cChannels
            'Update The Name Label On The Main Form.
            If Not BeenThere Then FrmMxr.LblName(MaxSources + 1).Caption = "Treble"
            'Call The "GetControlID" Function So We Can Get The Control ID
            'Of The Master Volume.
            ControlID = GetControlID(ML.dwComponentType, MIXERCONTROL_CONTROLTYPE_TREBLE)
            If ControlID <> 0 Then
                'Prep The MCD Structure For The Master Volume Fader.
                With MCD
                    .cbDetails = 4  'Size Of A Long In Byte's.
                    .cbStruct = 24
                    .cChannels = ML.cChannels
                    .dwControlID = ControlID
                    .item = 0
                    .paDetails = VarPtr(MixerState(MaxSources + 1).MxrVol)
                End With
                'Get The Master Volume Setting.
                mixerGetControlDetails hMixer, MCD, MIXER_GETCONTROLDETAILSF_VALUE
                'Track Bar Logic Is The Reverse Of Fader's On A Hardware Mixer
                'So Reverse The Value.
                If ReverseLogicTrebBass Then
                    MixerState(MaxSources + 1).MxrVol = MixerState(MaxSources + 1).MxrVol
                Else
                    MixerState(MaxSources + 1).MxrVol = 65535 - MixerState(MaxSources + 1).MxrVol
                End If
                'Save The Master Volume Control ID.
                MixerState(MaxSources + 1).MxrVolID = MCD.dwControlID
                'FrmMxr.PicGang(MaxSources + 1).Visible = True
                If Not BeenThere Then FrmMxr.ShpMute(MaxSources + 1).Visible = False
                If Not BeenThere Then FrmMxr.lblVol(MaxSources + 1).BackColor = vbButtonFace
                If ChangeFlag Then FrmMxr.SldrVol(MaxSources + 1).Enabled = 1
                TBSupport = True
            Else
                'Couldn't Get It, Disable The Fader.
                FrmMxr.SldrVol(MaxSources + 1).Enabled = 0
            End If
            If Not ShowTB Then FrmMxr.SldrVol(MaxSources + 1).Enabled = False
            
            'MASTER BASS (DESTINATION)
            MixerState(MaxSources + 2).MxrChannels = ML.cChannels
            'Update The Name Label On The Main Form.
            If Not BeenThere Then FrmMxr.LblName(MaxSources + 2).Caption = "Bass"
            'Call The "GetControlID" Function So We Can Get The Control ID
            'Of The Master Volume.
            ControlID = GetControlID(ML.dwComponentType, MIXERCONTROL_CONTROLTYPE_BASS)
            If ControlID <> 0 Then
                'Prep The MCD Structure For The Master Volume Fader.
                With MCD
                    .cbDetails = 4  'Size Of A Long In Byte's.
                    .cbStruct = 24
                    .cChannels = ML.cChannels
                    .dwControlID = ControlID
                    .item = 0
                    .paDetails = VarPtr(MixerState(MaxSources + 2).MxrVol)
                End With
                'Get The Master Volume Setting.
                mixerGetControlDetails hMixer, MCD, MIXER_GETCONTROLDETAILSF_VALUE
                'Track Bar Logic Is The Reverse Of Fader's On A Hardware Mixer
                'So Reverse The Value.
                If ReverseLogicTrebBass Then
                    MixerState(MaxSources + 2).MxrVol = MixerState(MaxSources + 2).MxrVol
                Else
                    MixerState(MaxSources + 2).MxrVol = 65535 - MixerState(MaxSources + 2).MxrVol
                End If
                'Save The Master Volume Control ID.
                MixerState(MaxSources + 2).MxrVolID = MCD.dwControlID
                'FrmMxr.PicGang(MaxSources + 2).Visible = True
                If Not BeenThere Then FrmMxr.ShpMute(MaxSources + 2).Visible = False
                If Not BeenThere Then FrmMxr.lblVol(MaxSources + 2).BackColor = vbButtonFace
                If ChangeFlag Then FrmMxr.SldrVol(MaxSources + 2).Enabled = 1
                TBSupport = True
            Else
                'Couldn't Get It, Disable The Fader.
                FrmMxr.SldrVol(MaxSources + 2).Enabled = 0
            End If
            If Not ShowTB Then FrmMxr.SldrVol(MaxSources + 2).Enabled = False
            
            'MASTER MUTE (DESTINATION)
            'Call The "GetControlID" Function So We Can Get The Control ID
            'Of The Master Mute.
            ControlID = GetControlID(ML.dwComponentType, MIXERCONTROL_CONTROLTYPE_MUTE)
            If ControlID <> 0 Then
                'Prep The MCD Structure For The Master Mute.
                With MCD
                    .cbDetails = 4  'Size Of A Long In Byte's.
                    .cbStruct = Len(MCD)
                    .cChannels = 1 'Mute has but one channel.
                    .dwControlID = ControlID
                    .item = 0
                    .paDetails = VarPtr(MixerState(0).MxrMute)
                End With
                'Get The Master Mute Setting.
                mixerGetControlDetails hMixer, MCD, MIXER_GETCONTROLDETAILSF_VALUE
                'Save The Master Mute Control ID.
                MixerState(0).MxrMuteID = MCD.dwControlID
            Else
                'Couldn't Get It, Disable The Master Mute.
                If Not BeenThere Then FrmMxr.ChkMute(0).Enabled = 0
                If Not BeenThere Then FrmMxr.ShpMute(0).Visible = False
            End If
            
             'MASTER ON/OFF (DESTINATION)
            'Call The "GetControlID" Function So We Can Get The Control ID
            'Of The first Master ON/OFF.
            ControlID = GetControlID(ML.dwComponentType, MIXERCONTROL_CONTROLTYPE_ONOFF)
            If ControlID <> 0 Then
                'Prep The MCD Structure For The Master Mute.
               With MCD
                   .cbDetails = 4  'Size Of A Long In Byte's.
                    .cbStruct = Len(MCD)
                    .cChannels = 1 'ON/OFF has but one channel.
                    .dwControlID = ControlID
                    .item = 0
                    .paDetails = VarPtr(MixerState(0).MxrOnOff)
                End With
                'Get The Master Mute Setting.
                mixerGetControlDetails hMixer, MCD, MIXER_GETCONTROLDETAILSF_VALUE
                'Save The Master Mute Control ID.
                MixerState(0).MxrOnOffID = MCD.dwControlID
            Else
                'Couldn't Get It, Disable The Master Mute
                If Not BeenThere Then FrmMxr.ChkOnOff.Enabled = 0
            End If
            
            'ON TO THE SOURCES NOW...
            'Now That We've Found The Speakers And The Master Volume,
            'Let's Get The Source's...
            For Src = 0 To ML.cConnections - 1
                'Prep The "MIXERLINE" Struct For Source's.
                ML.cbStruct = Len(ML)
                ML.dwDestination = Dst
                ML.dwSource = Src
                'Get The Line Info For The Current Source.
                mixerGetLineInfo hMixer, ML, MIXER_GETLINEINFOF_SOURCE
                'Save The Channels Of The Source.
                MixerState(Src + 1).MxrChannels = ML.cChannels
                'Update The Name Label On The Main Form.
                If Not BeenThere Then FrmMxr.LblName(Src + 1).Caption = Left(ML.szName, InStr(ML.szName, vbNullChar) - 1)
                'Call The "GetControlID" Function So We Can Get The Control ID
                'Of The Current Source Volume.
                ControlID = GetControlID(ML.dwComponentType, MIXERCONTROL_CONTROLTYPE_VOLUME)
                If ControlID <> 0 Then
                    'Prep The MCD Structure For The Current Source Volume.
                    With MCD
                        .cbDetails = 4   'Size Of A Long In Byte's.
                        .cbStruct = Len(MCD)
                        .cChannels = ML.cChannels
                        .dwControlID = ControlID
                        .item = 0
                        .paDetails = VarPtr(MixerState(Src + 1).MxrVol)
                    End With
                    'Get The Current Source Volume Setting.
                    mixerGetControlDetails hMixer, MCD, MIXER_GETCONTROLDETAILSF_VALUE
                    'Save The Volume Setting.
                    MixerState(Src + 1).MxrVol = 65535 - MixerState(Src + 1).MxrVol
                    'Save The ID
                    MixerState(Src + 1).MxrVolID = MCD.dwControlID
                Else
                    'Couldn't Get It, So Disable The Control.
                    If Not BeenThere Then FrmMxr.SldrVol(Src + 1).Enabled = 0
                End If

               'MUTES FOR SOURCES
               'Call The "GetControlID" Function So We Can Get The Control ID
               'Of The Current Source Mute.
               ControlID = GetControlID(ML.dwComponentType, MIXERCONTROL_CONTROLTYPE_MUTE)
               If ControlID <> 0 Then
                    'Prep The MCD Structure For The Current Source Mute.
                    With MCD
                        .cbDetails = 4   'Size Of A Long In Byte's.
                        .cbStruct = Len(MCD)
                        .cChannels = 1 'Mutes have but one channel.
                        .dwControlID = ControlID
                        .item = 0
                        .paDetails = VarPtr(MixerState(Src + 1).MxrMute)
                    End With
                    'Get The Current Source Mute Setting.
                    mixerGetControlDetails hMixer, MCD, MIXER_GETCONTROLDETAILSF_VALUE
                    'Save The Mute Control ID.
                    MixerState(Src + 1).MxrMuteID = MCD.dwControlID
                Else
                    'Couldn't Get It, So Disable The Control.
                    If Not BeenThere Then FrmMxr.ChkMute(Src + 1).Enabled = 0
                    If Not BeenThere Then FrmMxr.ShpMute(Src + 1).Visible = False
                End If
            Next Src
           ' We Found The Destination That Is The Speaker's, So Exit The Outer Loop.
            Exit For
        End If
    Next Dst
    ChangeFlag = False
    BeenThere = True
End Sub

' *********************************************
'Returns The Requested Control ID.
' *********************************************
Public Function GetControlID&(ByVal ComponentType&, ByVal ControlType&)
   Dim hmem&
   Dim MC As MIXERCONTROL
   Dim MxrLine As MIXERLINE
   Dim MLC As MIXERLINECONTROLS
   'Prep The MxrLine Structure.
   MxrLine.cbStruct = Len(MxrLine)
   MxrLine.dwComponentType = ComponentType  'This Value Sent In.
   'Get The Line Info.
   If mixerGetLineInfo(hMixer, MxrLine, MIXER_GETLINEINFOF_COMPONENTTYPE) = 0 Then
        ' Prep The MLC Structure.
        MLC.cbStruct = Len(MLC)
        MLC.dwLineID = ML.dwLineID
        MLC.dwControl = ControlType     'This Value Sent In.
        MLC.cControls = 1
        MLC.cbmxctrl = Len(MC)
        hmem = GlobalAlloc(&H40, Len(MC))
        MLC.pamxctrl = GlobalLock(hmem)
        MC.cbStruct = Len(MC)
        'Get The Line Control.
        If mixerGetLineControls(hMixer, MLC, MIXER_GETLINECONTROLSF_ONEBYTYPE) = 0 Then
            'Copy The Data To The MC Structure.
            CopyStructFromPtr MC, MLC.pamxctrl, Len(MC)
            'Return The Control ID.
            GetControlID = MC.dwControlID
        End If
        GlobalUnlock hmem
        GlobalFree hmem
    End If
End Function

' *********************************************
'Act on application becoming the active or inactive window.
' *********************************************
Public Function NewWindowProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim w%
    If msg = WM_ACTIVATE Then
        If (wParam = WA_ACTIVE Or wParam = WA_CLICKACTIVE) Then
            FrmMxr.PicTitle.BackColor = vbActiveTitleBar
            MixerVisible = True
            FrmMxr.MnuRestoreMixer.Enabled = False
            FrmMxr.MnuHideMixer.Enabled = True
            For w = FrmMxr.shpFocus.lbound To FrmMxr.shpFocus.UBound
                FrmMxr.shpFocus(w).Visible = True
            Next w
        Else
            FrmMxr.PicTitle.BackColor = vbInactiveTitleBar
            For w = FrmMxr.shpFocus.lbound To FrmMxr.shpFocus.UBound
                FrmMxr.shpFocus(w).Visible = False
            Next w
            If AutoHide Then
                FrmMxr.Hide
                MixerVisible = False
                FrmMxr.MnuHideMixer.Enabled = False
                FrmMxr.MnuRestoreMixer.Enabled = True
            End If
        End If
    End If
    ' Send other messages to the original window proceedure.
    NewWindowProc = CallWindowProc(OldWindowProc, hwnd, msg, wParam, lParam)
End Function

' *********************************************
' Act on tray-icon clicks.
' *********************************************
Public Function TrayNewWindowProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If msg = TRAY_CALLBACK Then
        ' The user clicked on the tray icon.  Look for click events, namely which mouse-button was clicked.
        If lParam = WM_LBUTTONUP Then
            ' On left click, show the form.  MnuRestoreMixer_Click will also hide the mixer if it is visible.
            Call FrmMxr.MnuRestoreMixer_Click
            Exit Function
        End If
        If lParam = WM_RBUTTONUP Then
            ' On right click, show the pop-up menu.
            TheForm.PopupMenu TheMenu
            Exit Function
        End If
        If lParam = WM_MBUTTONUP Then
        'Put middle-click statements here, but remember, not all mice have a middle-button.
        Beep
        End If
    End If
    ' Send other messages to the original window proceedure.
    TrayNewWindowProc = CallWindowProc(TrayOldWindowProc, hwnd, msg, wParam, lParam)
End Function

' *********************************************
' Add the form's icon to the tray.
' *********************************************
Public Sub AddToTray(frm As Form, mnu As Menu)
    ' ShowInTaskbar must be set to False at design time because it is read-only at run time.
    ' Save the form and menu for later use.
    Set TheForm = frm
    Set TheMenu = mnu
    ' Install the new WindowProc.
    TrayOldWindowProc = SetWindowLong(frm.hwnd, GWL_WNDPROC, AddressOf TrayNewWindowProc)
    ' Install the form's icon in the tray.
    With TheData
        .uID = 0
        .hwnd = frm.hwnd
        .cbSize = Len(TheData)
        .hIcon = frm.Icon.Handle
        .uFlags = NIF_ICON
        .uCallbackMessage = TRAY_CALLBACK
        .uFlags = .uFlags Or NIF_MESSAGE
        .cbSize = Len(TheData)
    End With
    Shell_NotifyIcon NIM_ADD, TheData
End Sub

' *********************************************
' Remove the icon from the system tray.
' *********************************************
Public Sub RemoveFromTray()
    ' Remove the icon from the tray.
    With TheData
        .uFlags = 0
    End With
    Shell_NotifyIcon NIM_DELETE, TheData
    ' Restore the original window proc.
    SetWindowLong TheForm.hwnd, GWL_WNDPROC, TrayOldWindowProc
End Sub

'*********************************************
' Set a new tray tip.
' *********************************************
Public Sub SetTrayTip(tip As String)
    With TheData
        .szTip = tip & vbNullChar
        .uFlags = NIF_TIP
    End With
    Shell_NotifyIcon NIM_MODIFY, TheData
End Sub

' *********************************************
' Set a new tray icon.
' *********************************************
Public Sub SetTrayIcon(pic As Picture)
    ' Do nothing if the picture is not an icon.
    If pic.Type <> vbPicTypeIcon Then Exit Sub
    ' Update the tray icon.
    With TheData
        .hIcon = pic.Handle
        .uFlags = NIF_ICON
    End With
    Shell_NotifyIcon NIM_MODIFY, TheData
End Sub

' *********************************************
' Check for existance of any file by attempting to open it.
' *********************************************
Public Function FileExists(ByVal filename As String)
    Dim f%: f = FreeFile
    On Error GoTo FileDoesntExist
    Open filename For Input As #f: Close #f
    FileExists = True: Exit Function
FileDoesntExist:
    FileExists = False
End Function
