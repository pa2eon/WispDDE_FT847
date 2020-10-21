VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmMain 
   Caption         =   "WiSP DDE Client V.4.3.8 FT-847"
   ClientHeight    =   6390
   ClientLeft      =   1065
   ClientTop       =   735
   ClientWidth     =   6030
   ForeColor       =   &H00FF0000&
   Icon            =   "Wispdde.frx":0000
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Source
   LinkTopic       =   "WispDDE"
   ScaleHeight     =   6390
   ScaleWidth      =   6030
   Begin VB.TextBox AzRaw 
      Height          =   285
      Left            =   4440
      TabIndex        =   35
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Minus10KHzButton 
      Caption         =   "<<"
      Height          =   255
      Left            =   240
      TabIndex        =   34
      ToolTipText     =   "Decrease master Radio frequency in 10KHz."
      Top             =   3240
      Width           =   315
   End
   Begin VB.Timer RadioControlLoopTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   5100
   End
   Begin VB.TextBox DownlinkFreq 
      Height          =   285
      Left            =   1800
      TabIndex        =   30
      ToolTipText     =   "Downlink Radio Frequency (MHz)."
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox UplinkFreq 
      Height          =   285
      Left            =   240
      TabIndex        =   29
      ToolTipText     =   "Uplink Radio Frequency (MHz)."
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox DownlinkDDEFreq 
      Height          =   285
      HideSelection   =   0   'False
      Left            =   4440
      TabIndex        =   28
      ToolTipText     =   "Downlink frequency received from DDE or manually set."
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox UplinkDDEFreq 
      Height          =   285
      HideSelection   =   0   'False
      Left            =   4440
      TabIndex        =   27
      ToolTipText     =   "Uplink frequency received from DDE or manually set."
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Timer SliderTimer 
      Enabled         =   0   'False
      Left            =   3720
      Top             =   5520
   End
   Begin MSCommLib.MSComm MSComm3 
      Left            =   2880
      Top             =   5400
      _ExtentX        =   979
      _ExtentY        =   979
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox Satellite 
      ForeColor       =   &H00808080&
      Height          =   285
      HideSelection   =   0   'False
      Left            =   1080
      TabIndex        =   16
      ToolTipText     =   "Name of the satellite being tracked."
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton UpdateRadioButton 
      Caption         =   "Update radio"
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      ToolTipText     =   "Send data to radio NOW!"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton UpdateRotorButton 
      Caption         =   "Update rotor"
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      ToolTipText     =   "Send data to rotor interface NOW!"
      Top             =   600
      Width           =   1095
   End
   Begin VB.Timer RadioTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   5520
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   2160
      Top             =   5400
      _ExtentX        =   979
      _ExtentY        =   979
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1440
      Top             =   5400
      _ExtentX        =   979
      _ExtentY        =   979
      _Version        =   393216
      DTREnable       =   -1  'True
      InputMode       =   1
   End
   Begin VB.Frame Frame2 
      Caption         =   "Radio(s)"
      Height          =   2775
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   3135
      Begin VB.CommandButton Plus10KHzButton 
         Caption         =   ">>"
         Height          =   255
         Left            =   2700
         TabIndex        =   33
         ToolTipText     =   "Increase master Radio frequency in 10KHz."
         Top             =   1200
         Width           =   315
      End
      Begin MSComctlLib.ProgressBar DownlinkRSSI 
         Height          =   195
         Left            =   1680
         TabIndex        =   32
         ToolTipText     =   "Received Signal Strength Indicator"
         Top             =   2520
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
         Min             =   1e-4
         Max             =   255
      End
      Begin MSComctlLib.Slider Slider 
         Height          =   465
         Left            =   360
         TabIndex        =   31
         ToolTipText     =   "Controls master Radio frequency. Use mouse, arrow keys or PgUp/PgDown. Ctrl-S to focus."
         Top             =   1080
         Width           =   2415
         _ExtentX        =   4233
         _ExtentY        =   847
         _Version        =   393216
         LargeChange     =   100
         SmallChange     =   10
         Min             =   -1000
         Max             =   1000
         TickStyle       =   1
         TickFrequency   =   100
         TextPosition    =   1
      End
      Begin VB.OptionButton TrackRev 
         Caption         =   "&Rev."
         Height          =   255
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Set Reverse Frequency Tracking (double click for no tracking). Shortcut: Ctrl-R."
         Top             =   840
         Width           =   495
      End
      Begin VB.OptionButton TrackDir 
         Caption         =   "&Dir."
         Height          =   255
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Set Direct Frequency Tracking (double click for no tracking). Shortcut: Ctrl-D."
         Top             =   840
         Width           =   495
      End
      Begin VB.CheckBox SliderDownlink 
         Caption         =   "Check2"
         Height          =   195
         Left            =   2820
         TabIndex        =   23
         ToolTipText     =   "Set Downlink Radio as Master. Shortcut: Ctrl-W. Ctrl-Space toggles no Master."
         Top             =   840
         Width           =   195
      End
      Begin VB.CheckBox SliderUplink 
         Caption         =   "Check1"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Set Uplink Radio as Master. Shortcut: Ctrl-U. Ctrl-Space toggles no Master."
         Top             =   840
         Width           =   195
      End
      Begin VB.ComboBox DownlinkIndex 
         Height          =   315
         Left            =   1680
         TabIndex        =   19
         ToolTipText     =   "Shows automatically selected downlink radio or manually selects one."
         Top             =   2160
         Width           =   735
      End
      Begin VB.ComboBox UplinkIndex 
         Height          =   315
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Shows automatically selected uplink radio or manually selects one."
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox UplinkMode 
         Height          =   285
         HideSelection   =   0   'False
         Left            =   720
         TabIndex        =   13
         ToolTipText     =   "Uplink mode received from DDE or manually set."
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox DownlinkMode 
         Height          =   285
         HideSelection   =   0   'False
         Left            =   2280
         TabIndex        =   12
         ToolTipText     =   "Downlink mode received from DDE or manually set."
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "RSSI:"
         Height          =   195
         Left            =   1200
         TabIndex        =   26
         Top             =   2520
         Width           =   435
      End
      Begin VB.Label Label9 
         Caption         =   "Selected Radio:"
         Height          =   255
         Left            =   1680
         TabIndex        =   20
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Selected Radio:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Mode:"
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Do&wnlink:"
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Mode:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "&Uplink:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rotor"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1935
      Begin VB.CheckBox RotorAuto 
         Alignment       =   1  'Right Justify
         Caption         =   "Auto update"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "Enables automatic update of rotor interface values."
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Elevation 
         Height          =   285
         Left            =   960
         TabIndex        =   11
         ToolTipText     =   "Elevation angle received from DDE or manually set."
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Azimuth 
         Height          =   285
         Left            =   960
         TabIndex        =   10
         ToolTipText     =   "Azimuth angle received from DDE or manually set."
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Azimuth:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Elevation:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Timer DDEPollTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   5520
   End
   Begin VB.Label UplinkDDELabel 
      Caption         =   "Uplink DDE Freq:"
      Height          =   255
      Left            =   4440
      TabIndex        =   39
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label DnlinkDDELabel 
      Caption         =   "Dnlink DDE Freq:"
      Height          =   255
      Left            =   4440
      TabIndex        =   38
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label DDERawLabel 
      Caption         =   "Raw DDE String:"
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label AzRawLabel 
      Caption         =   "Raw Az:"
      Height          =   255
      Left            =   4440
      TabIndex        =   36
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      DrawMode        =   1  'Blackness
      X1              =   2640
      X2              =   2640
      Y1              =   1800
      Y2              =   2040
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   2640
      X2              =   2640
      Y1              =   960
      Y2              =   1200
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      DrawMode        =   1  'Blackness
      X1              =   2160
      X2              =   2640
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label4 
      Caption         =   "Satellite:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Label DDELabel 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      LinkItem        =   "Tracking"
      LinkTimeout     =   10000
      LinkTopic       =   "GSC|Tracking"
      TabIndex        =   0
      Top             =   6000
      Width           =   5655
   End
   Begin VB.Menu close 
      Caption         =   "Close"
   End
   Begin VB.Menu set 
      Caption         =   "Settings"
      Begin VB.Menu rotor 
         Caption         =   "Rotor"
      End
      Begin VB.Menu radio 
         Caption         =   "Radio"
      End
      Begin VB.Menu ddelink 
         Caption         =   "DDE Link"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu viewhelp 
         Caption         =   "View Help File"
      End
      Begin VB.Menu about 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WiSPDDE build history since 3.0.54
'
'3.0.54:    fixed receiver turn-off for IC-R8500 and
'           added Nova for Windows support (thanks Alan K.
'           Adamson). Not released
'3.0.56:    Enhanced Nova support to any number of satellites
'           (was just for WXSats). Released 3-sep-01.
'3.0.57:    Enhanced Satellites database by adding a sat enabling
'           flag. Corrected bug of opposite uplink doppler
'           for Nova for Windows.
'           Released 7-sep-01.
'3.0.58:    Added DDE support from SatPC32 (WiSPDDE acts as
'           DDE Server in this mode).
'           Added DDE support from Satscape.
'           Added TrakBox support for rotor control (no rig ctl.).
'           Changed application title to "wispdde" intead of
'           "WiSP DDE Client" in order to simplify DDE server
'           naming. Released oct-01
'3.0.59:    Fixed opposite doppler correction for WinOrbit.
'           Renamed Radio routines to a more "standard" fashion
'           to improve legibility.
'           Corrected Auto Update checkbox to not to update the
'           rotor fields with DDE data.
'           Added TS-711, TS-811 and FRG-9600 thanks to Michael-VK2JXI
'           Corrected bug that made integer rounding on freqs
'           from WinOrbit.
'4.0.0:     Added "Transparent Tuning" capability using a sample
'           pair of frequencies doppler compensated by the DDE
'           server to introduce doppler compensation into the
'           rig's frequencies. This is an approximation
'           since doppler effect is calculated at a slightly
'           different frequency than the actual freq of operation.
'           For example if doppler is calculated at 435.675MHz
'           (center of AO40 70cm uplink) and we TX at
'           435.800 (upper extreme of AO40 70cm uplink) there is
'           an error of 0.029% downwards in doppler compensation
'           (thats 1Hz error if total doppler drift is 5KHz).
'           Care must be taken to avoid operating far away from
'           the frequencies set on the orbital prediction program
'           or this error may increase significantly.
'           Added support for up/down-converters by storing
'           upconverter and downconverter local oscillators
'           frequencies for every radio.
'4.0.1:     Added support for 450 degrees rotators
'           Added VR-5000 receiver support.
'4.0.2:     Added radio control thru TrakBox
'           Corrected some bugs on radio activation routines.
'4.0.3:     Correct bugs in FT847 readback routines -by Mark
'           P. Grimes WA0TOP-.
'           Changed FT847 frame readback routine to return
'           as soon as a complete set of 5 bytes have been
'           received. Also changed VR5000 readback routine to
'           return as soon as 1 byte is received.
'           RSSI info from FT847 can still be read with VR5000
'           corresponding routine.
'           Changed RadioControlLoop routine to resist master
'           radio not responding by using latest good freq response
'           in such case.
'4.0.4:     Correct behaviour of transparent tuning while
'           radio not responding (ie due to user moving VFO
'           knob).
'4.0.5:     Eliminated error reporting pop-up windows when
'           radios not responding (ie due to user adjusting
'           freq. from VFO knob).
'4.0.6:     Corrected bug causing divide-by-zero error when no rotor
'           step size was set. Corrected mis-behaviour when
'           trying to open connection with unreachable radio.
'4.0.7:     Added split-mode capability with FT-100 radio.
'           Added FT-100 to transparent-tuning enabled radios
'4.1.0:     A new option in Radio settings to enable Split-Mode.
'           DownlinkFreqs will be set into VFO A and uplinkfreqs
'           into VFO B if Split-Mode is enabled for a particular
'           radio. Only FT-100 is enabled for this operation.
'4.1.1:     Added IC-706/746/275/475 to Split-Mode enabled radios.
'4.1.2:     Added COM-ports above 9 and fixed bugs in
'           TS2000 & TS790 routines -by Howard G6LVB-.
'           Changed distribution file wispdde.zip to include a
'           setup.exe program that installs WiSPDDE automatically
'           along with needed shared libraries.
'4.1.3:     Added code to prevent errors from loopback CI-V interfaces,
'           only relevant for home-made CI-V interfaces.
'4.1.4:     Changed DDE LinkMode of main window to 'Source'
'           in order to let other programs (like ARSWin) read
'           satellite name and position from WiSPDDE.
'4.1.5:     Changed DDE LinkTopic of main window to "WispDDE" to
'           make it compatible with ARSWin 2.1.
'           Any program can now retrieve information from any of
'           WiSPDDE's main window fields by just establishing a
'           DDE link to it.
'4.1.6:     Corrected LinkNotify DDE mode to update data from
'           satellite as soon as the orbit prediction app. changes
'           its DDE string. This is only active when query-period is
'           set to zero.
'4.1.7:     Corrected WXtrack application name for DDE link.
'           Added WaitOutBuffEmpty routine to facilitate elimination of
'           echoes on COM ports when using home-made CI-V interfaces
'           (only for Icom rigs).
'4.1.8:     Eliminated ARS as a DDE server. Now ARSWin acts as a client
'           of WiSPDDE in order to get antenna positions.
'4.1.9:     Disabled -nonexistant- filter setting for FRG9600
'           scanner. Corrected antennae park function for interfaces
'           using parallel ports.
'4.1.10:    Added 'memory' on Auto-rotor-update checkbox and Rev/Dir
'           selections.
'           Implemented CI-V frame destination identification in
'           order to filter-out echoed frames happening with home-made
'           CI-V interfaces. This should also enable other traffic
'           thru the CI-V bus other than rig<->PC (inter-rig transceive
'           function for example)...
'4.1.11:    Changed DDE configuration for SatPC32 to let WiSPDDE act as
'           DDE-Client just like WiSP-GSC etc. 2-12-04
'4.1.12:    Added IC910 (replica of IC821). Corrected some issues when
'           selecting a radio with no DDE link. Added keyboard shortcuts
'           for frecuency tuning/tracking controls. 2-13-04
'4.2.0:     Debuged Icom control routines and transparent tuning feature.
'           Also made some Icom routines more robust by tolerating
'           communications errors without interrupting normal operation.
'4.2.1:     Released version of WiSPDDE 4.2.
'4.2.2:     Eliminated Satellite-mode setting on satellite-ready rigs
'           to enable in-band VFOs tracking (ISS etc.).
'4.2.3:     IC910 band-inversion bug corrected. Still working on a bug with
'           TrakBox control.
'4.2.4:     Fixed corrupted readme.txt file. Added PORT95NT.ZIP file to
'           distribution to provide port control under NT-like Windows.
'           Added Orbitron to list of DDE servers. 9-18-04.
'4.2.5:     Corrected FRG9600 routines.9-19-04.
'4.2.6:     Fixed faulty rotor-control processing causing type-mismatch errors
'           and no auto-uptades.
'           Quit using old Val() function, changed to Cdbl() to make the program
'           locale-aware regarding numeric formats.
'4.2.9:     Fixed problems with numeric format received from WiSP in computers
'           with European locale settings.
'           Added decimal separator character configuration field in DDE-Settings.
'4.2.10:    Added FT-817 Split-Mode support.
'           Added Radio Control Delay-Time configuration field to Radio-Settings.
'4.2.11:    Added PTT status readback for FT817. Do not update freqs while PTT is ON
'           Added FT-897 to radio list (almost like 817 with FM-N mode added).
'           Corrected up/dnlink freq tracking to let slave band free when no tracking
'           is desired.
'4.3.0:     Added capability of logging events in radio control routines.
'           Corrected bug in TS-790 band setting commands.
'           Added capability of logging radio commands and replies from radio, for
'           FT-817 and FT-100.
'4.3.1:     Added reply timeout setting besides command delay for reliablility of
'           control routines.
'           Added split-mode operation for IC-821 and IC-910 radios.
'           Added configuration parameters written to log files.
'4.3.2:     Corrected DDE link update from server application when DDE interval is
'           set to zero.
'           Eliminated DoEvents while performing delays due to problems of
'           concurrency in radio control routines.
'4.3.3:     Added option in Satellites window (active only for Nova DDE reception) for
'           using uplink channel as second downlink channel. As requested by Damiano
'           Accurso from AlmaSat project.
'4.3.4:     Added FT-847 Split-Mode using Repeater-Offset.
'           Corrected bug in elevation retrieveing from Nova for Windows
'           DDE string, thanks to Dennis - W7KMV
'           Corrected bug in TS711/811 freq. readback.
'           EB2CTA: Added operation for AOR AR-3000A radio. S-Meter control
'           not implemented.
'           Added for TrakBox: 'Q' command before '7' command when initiating.
'4.3.5:     Added RC2800PX rotor controller.
'4.3.6:     Corrected bug in frequency setting routines
'           caused by Int() function, moved to Fix() function. Corrected some bugs
'           in RadioControlLoopTimer routine causing frequency errors.
'4.3.7:     Changed to use duplex-mode for in-band TX/RX operation, split-mode has
'           problems of freq update while TX.
'4.3.8:     Add CTCSS for FT-847 - 67 Hz for use with ISS / SO-50 - PA2EON
'



Private Sub about_Click()
    'display the 'about' window
    frmAbout1.Show
End Sub

Private Sub Close_Click()
Form_Unload (0)
End Sub
Sub UpdateRotor(Az, El)
'if stop position is South:
'make correction for the stop
'position **ONLY** if rotor is not a GS-232
If (frmRotor.RotorSouth.Value = 1) _
    And (frmRotor.RotorType.text <> "GS-232") Then
    If Az < 180 Then
        Az = Az + 180
    Else
        Az = Az - 180
    End If
End If

'adjust for offset angles:
Az = Az + Cdbl2(frmRotor.RotorAzOffset.text)
El = El + Cdbl2(frmRotor.RotorElOffset.text)

If frmRotor.Rotor450Deg = 0 Then
    If Az < 0 Then Az = Az + 360
    If Az > 360 Then Az = Az - 360
    
    If Az < 0 Then Az = 0
    If Az > 360 Then Az = 360
Else
    If Az < 0 Then Az = Az + 450
    If Az > 450 Then Az = Az - 450
    
    If Az < 0 Then Az = 0
    If Az > 450 Then Az = 450
End If

If El < 0 Then El = 0
If El > 180 Then El = 180

'update hardware devices only if last update is complete:
If RotorUpdateComplete And RotorAuto.Value Then
    Dim Port As Integer
    
    Select Case frmRotor.RotorType.text
    
        Case "TrakBox"
        'if rotor controller is TrakBox:
        'verify port is open...
        If RotorHandle% Then
            Call TBRotorSetAzEl(Az, El, RotorHandle%)
        End If
    
        Case "GS-232"
        'if rotor controller is GS-232 we send
        'command in this format.
        If RotorHandle% Then
            Call GS232RotorSetAzEl(Az, El, RotorHandle%)
        End If
    
        Case "RC2800PX"
        'if rotor controller is RC2800PX we send
        'elevation and azimuth positions individually
        If RotorHandle% Then
            Call RC2800PXRotorSetAz(Az, RotorHandle%)
            Call RC2800PXRotorSetEl(El, RotorHandle%)
        End If
    
        Case "EASYCOMM-I"
        'if rotor controller is EASYCOMM I we send
        'command in this format.
        If RotorHandle% Then
            Call EASYCOMMIRotorSetAzEl(Az, El, RotorHandle%)
        End If
    
        Case "CI-V"
        'if rotor controller is CI-V we send
        'command in this format.
        If RotorHandle% Then
            Call IC821RadioSetFreq(El * 1000 + Az, &HDD, 0, RotorHandle%)
        End If
    
        Case "FODTrack"
        'if rotor controller is FODTrack:
        'make the correspondence between LPT number
        'and port address
        Select Case frmRotor.RotorPort.text
            Case Is = "LPT1 (378)"
                Port = &H378
            Case Is = "LPT2 (278)"
                Port = &H278
            Case Is = "LPT3 (3BC)"
                Port = &H3BC
        End Select
        Call FODRotorSetAzEl(Az, El, Port)
    
        Case "IF-100"
        'if rotor controller is IF-100:
        'make the correspondence between LPT number
        'and port address
        Select Case frmRotor.RotorPort.text
            Case Is = "LPT1 (378)"
                Port = &H378
            Case Is = "LPT2 (278)"
                Port = &H278
            Case Is = "LPT3 (3BC)"
                Port = &H3BC
        End Select
        Call IF100RotorSetAzEl(Az, El, Port)
        
    End Select
End If
End Sub
Sub IF100RotorSetAzEl(Az, El, Port As Integer)
    DATO = &H1
    CLOCK = &H2
    TRACK = &H8
    'and finally scale data and send to
    'the port.
    If frmRotor.Rotor450Deg = 0 Then
        Az = CInt(Az * 0.70833)
    Else
        Az = CInt((Az + 45) * 0.56667)
    End If
    El = CInt(El * 1.4166)
    Value& = El * 256 + Az
    For f% = 1 To 16
        If (Value& And 32768) Then
            Call OutPort(Port, (TRACK Or DATO))
            Call OutPort(Port, (TRACK Or DATO Or CLOCK))
            Call OutPort(Port, (TRACK Or DATO))
        Else
            Call OutPort(Port, (TRACK))
            Call OutPort(Port, (TRACK Or CLOCK))
            Call OutPort(Port, (TRACK))
        End If
        Value& = (Value& * 2) And 65535
    Next
    Call OutPort(Port, CInt(TRACK))

End Sub

Sub FODRotorSetAzEl(Az, El, Port As Integer)
    'scale data for full 8 bits:
If frmRotor.Rotor450Deg = 0 Then
    Az = CInt(Az * 0.70833)
    Call OutPort(Port + 2, 3)
    Call OutPort(Port, CInt(Az))
    Call OutPort(Port + 2, 2)
    
    AzRaw.text = Str$(Az)
    
Else
    Az = CInt(Az * 0.56667)
    Call OutPort(Port + 2, 3)
    Call OutPort(Port, CInt(Az))
    Call OutPort(Port + 2, 2)
End If
    
    El = CInt(El * 1.4166)
    Call OutPort(Port + 2, 1)
    Call OutPort(Port, CInt(El))
    Call OutPort(Port + 2, 0)

End Sub
Sub GS232RotorSetAzEl(Az, El, Handle%)
    a$ = RTrim(LTrim(Str(CInt(Az))))
    a$ = String(3 - Len(a$), Asc("0")) + a$
    
    b$ = RTrim(LTrim(Str(CInt(El))))
    b$ = String(3 - Len(b$), Asc("0")) + b$
   
    Call WriteToPort("W" + a$ + " " + b$ + Chr$(&HD), RotorHandle%)
    
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RotorTimeOut
End Sub
Sub RC2800PXRotorSetAz(Az, Handle%)
    a$ = RTrim(LTrim(Str(CInt(Az))))
    'a$ = String(3 - Len(a$), Asc("0")) + a$
    
    Call WriteToPort("A" + a$ + Chr$(&HD), RotorHandle%)
    
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RotorTimeOut
End Sub
Sub RC2800PXRotorSetEl(El, Handle%)
    a$ = RTrim(LTrim(Str(CInt(El))))
    'a$ = String(3 - Len(a$), Asc("0")) + a$
    
    Call WriteToPort("E" + a$ + Chr$(&HD), RotorHandle%)
    
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RotorTimeOut
End Sub
Sub EASYCOMMIRotorSetAzEl(Az, El, Handle%)
    a$ = RTrim(LTrim(Str(CInt(Az * 10) / 10)))
    
    b$ = RTrim(LTrim(Str(CInt(El * 10) / 10)))
   
    s$ = "AZ" + a$ + " EL" + b$ + " " + Chr$(&HD) + Chr$(&HA)
    
    For f% = 1 To Len(s$)
        Call WriteToPort(Mid$(s$, f%, 1), Handle%)
        Call TBRotorPaceDown(RotorPaceDelaySecs)
    Next
        
    a = Timer
    Do
        
    Loop Until Abs(Timer - a) > RotorTimeOut
End Sub
'This routine is used to slow down the rate at with chars are
'sent to TB to prevent comms errors:
Sub TBRotorPaceDown(DelaySecs)
a = Timer
Do
    
Loop Until Abs(Timer - a) > DelaySecs
End Sub
Sub TBRotorSetAzEl(Az, El, Handle%)

RotorUpdateComplete = False

InString = ReadFromPort(Handle%)
'TrakBox does not accept Elev greater than 90 so:
If El > 90 Then
    El = 180 - El
    If Az >= 180 Then
        Az = Az - 180
    Else
        Az = Az + 180
    End If
End If
a$ = RTrim(LTrim(Str(CInt(Az))))
a$ = String(3 - Len(a$), Asc("0")) + a$

b$ = RTrim(LTrim(Str(CInt(El))))
b$ = String(2 - Len(b$), Asc("0")) + b$

s$ = "AZ" + a$ + " EL" + b$ + Chr$(&HD)
For f% = 1 To Len(s$)
    Call WriteToPort(Mid$(s$, f%, 1), Handle%)
    Call TBRotorPaceDown(RotorPaceDelaySecs)
Next
'wait until TB stops sending position reports
Do
    a$ = TBRotorReadFrame(Handle%)
Loop Until Cdbl2(PickWord(a$, 1)) <> 1020
'check that motion ended OK...
If Cdbl2(PickWord(a$, 1)) <> 1 Then
    Call frmMessage.ShowMessage("Comms. Error during" + Chr$(13) _
        + "TrakBox Set Az/El", 10)
End If

RotorUpdateComplete = True
End Sub
'This stes TrakBox into HostMode:
Sub TBRotorSetHost(Handle%)

RotorUpdateComplete = False

InString = ReadFromPort(Handle%)
retry = 3
Do
    Call WriteToPort("7", Handle%)
    Call TBRotorPaceDown(RotorPaceDelaySecs)
    'wait until TB sends OK prompt,
    'if no OK prompt received, repeat command up to 3 times:
    s$ = TBRotorReadFrame(Handle%)
    If Cdbl2(PickWord(s$, 1)) <> 1 Then
        retry = retry - 1
        If retry = 0 Then
            Call frmMessage.ShowMessage("Comms. Error during" + Chr$(13) _
                + "TrakBox Set Host Mode", 10)
        End If
    Else
        retry = 0
    End If
Loop Until retry = 0

RotorUpdateComplete = True
End Sub
'This stes TrakBox into TerminalMode:
Sub TBRotorSetTerminal(Handle%)
Call WriteToPort("Q", Handle%)
Call TBRotorPaceDown(RotorPaceDelaySecs)
Call WriteToPort(Chr$(&HD), Handle%)
End Sub
'TrakBox frame readback routine for rotor control
'receives chars from rotor port until "TrakBox>>" word is detected
'then parses the complete phrase received for valid answers.
'The output of this function is a string beginning with the decoded
'frame type code followed by the data available.
'the following type codes can be returned:
'code "00:" - frame not understood pass timeout period.
'code "01:" - acknowlegde frame (just "TrakBox>>" was received)
'code "02:" - trakbox reported an error.
'code "1020:" - trackbox reported rotor positions. Azimuth angle
'               follows and then elevation angle separated by
'               spaces.
Function TBRotorReadFrame(Handle%) As String
'This will hold the received frame
InBuff$ = ""
'initialize buffer pointer..
InBuffPtr% = 0
'this will indicate the position where the OK word is found.
'while it is -1 it means it haven't been found yet. after it
'is found we need to be sure that the reqd. number of bytes
'is received
found% = -1
'set timeout time...
a = Timer
Do
    'Wait until we receive bytes
    'or a time-out occurs
    Select Case Handle%
    Case 1
        Do
            
        Loop Until MSComm1.InBufferCount >= 1 Or Abs(Timer - a) > RotorTimeOut
        InString = MSComm1.Input
    Case 2
        Do
            
        Loop Until MSComm2.InBufferCount >= 1 Or Abs(Timer - a) > RotorTimeOut
        InString = MSComm2.Input
    Case 3
        Do
            
        Loop Until MSComm3.InBufferCount >= 1 Or Abs(Timer - a) > RotorTimeOut
        InString = MSComm3.Input
    End Select
    'if a byte was received we reset timer to wait for another
    'time-out period
    If Abs(Timer - a) < RotorTimeOut Then
        a = Timer
    End If
    'We pass the received bytes to the buffer
    For i% = 0 To LenB(InString) - 1
        InBuff$ = InBuff$ + Chr$(InString(i%))
    Next
    'we will examine the buffer searching for the OK string:
    f% = InStr(InBuff$, "Box>>")
Loop While f% = 0 And Abs(Timer - a) < RotorTimeOut
'if loop ended due to timeout:
If f% = 0 Then
    s$ = "00:"
    TBRotorReadFrame = s$
    Exit Function
Else
    'otherwise examine frame:
    a = InStr(LCase$(InBuff$), "error")
    If a Then
        s$ = "02:"
        TBRotorReadFrame = s$
        Exit Function
    End If
    a = InStr(InBuff$, "AZ=")
    If a Then
        'look for separator of raw data / angle value
        a = InStr(a, InBuff$, "/")
        Az = Firstnum(Mid$(InBuff$, a + 1))
        a = InStr(a, InBuff$, "EL=")
        If a Then
            'look for separator of raw data / angle value
            a = InStr(a, InBuff$, "/")
            El = Firstnum(Mid$(InBuff$, a + 1))
                        
            'we have all the info.
            s$ = "1020: " + Format(Az, "##0.0") + " " + Format(El, "##0.0")
            TBRotorReadFrame = s$
            Exit Function
        Else
            'non-valid frame:
            Call ErrorHandler("TrakBox comms.", 0)
        End If
    End If
    'if it is non of the obove types of frame -> just acknowledge
    s$ = "01:"
    TBRotorReadFrame = s$
End If
End Function
'Send downlink frequency to TB radio controller:
Sub TBRadioSetRXFreq(Freq#, Handle%)
InString = ReadFromPort(Handle%)
'TB needs 8 digits for the frequency
s$ = "FD" + Format(Freq# / 10, "00000000") + Chr$(&HD)
For f% = 1 To Len(s$)
    Call WriteToPort(Mid$(s$, f%, 1), Handle%)
    Call TBRotorPaceDown(RotorPaceDelaySecs)
Next
'expect ack from TB:
a$ = TBRadioReadFrame(Handle%)
If Cdbl2(PickWord(a$, 1)) <> 1 Then
    Call frmMessage.ShowMessage("Comms. Error during" + Chr$(13) _
        + "TrakBox Set RX Freq", 10)
End If
End Sub
'Send uplink frequency to TB radio controller:
Sub TBRadioSetTXFreq(Freq#, Handle%)
InString = ReadFromPort(Handle%)
'TB needs 8 digits for the frequency
s$ = "FU" + Format(Freq# / 10, "00000000") + Chr$(&HD)
For f% = 1 To Len(s$)
    Call WriteToPort(Mid$(s$, f%, 1), Handle%)
    Call TBRotorPaceDown(RotorPaceDelaySecs)
Next
'expect ack from TB:
a$ = TBRadioReadFrame(Handle%)
If Cdbl2(PickWord(a$, 1)) <> 1 Then
    Call frmMessage.ShowMessage("Comms. Error during" + Chr$(13) _
        + "TrakBox Set TX Freq", 10)
End If
End Sub
'Send downlink mode to TB radio controller:
Sub TBRadioSetRXMode(mode$, Handle%)
Select Case LCase$(mode$)
Case "fm", "fm-n", "fm-w"
    m$ = "FM"
Case "usb"
    m$ = "USB"
Case "lsb"
    m$ = "LSB"
Case Else
    m$ = ""
End Select

InString = ReadFromPort(Handle%)
s$ = "MD" + m$ + Chr$(&HD)
For f% = 1 To Len(s$)
    Call WriteToPort(Mid$(s$, f%, 1), Handle%)
    Call TBRotorPaceDown(RotorPaceDelaySecs)
Next
'expect ack from TB:
a$ = TBRadioReadFrame(Handle%)
If Cdbl2(PickWord(a$, 1)) <> 1 Then
    Call frmMessage.ShowMessage("Comms. Error during" + Chr$(13) _
        + "TrakBox Set RX Mode", 10)
End If
End Sub
'Send uplink mode to TB radio controller:
Sub TBRadioSetTXMode(mode$, Handle%)
Select Case LCase$(mode$)
Case "fm", "fm-n", "fm-w"
    m$ = "FM"
Case "usb"
    m$ = "USB"
Case "lsb"
    m$ = "LSB"
Case Else
    m$ = ""
End Select

InString = ReadFromPort(Handle%)
s$ = "MU" + m$ + Chr$(&HD)
For f% = 1 To Len(s$)
    Call WriteToPort(Mid$(s$, f%, 1), Handle%)
    Call TBRotorPaceDown(RotorPaceDelaySecs)
Next
'expect ack from TB:
a$ = TBRadioReadFrame(Handle%)
If Cdbl2(PickWord(a$, 1)) <> 1 Then
    Call frmMessage.ShowMessage("Comms. Error during" + Chr$(13) _
        + "TrakBox Set TX Mode", 10)
End If
End Sub
'TrakBox frame readback routine for radio control
'receives chars from rotor port until "TrakBox>>" word is detected
'then parses the complete phrase received for valid answers.
'The output of this function is a string beginning with the decoded
'frame type code followed by the data available.
'the following type codes can be returned:
'code "00:" - frame not understood pass timeout period.
'code "01:" - acknowlegde frame (just "TrakBox>>" was received)
'code "02:" - trakbox reported an error.
'code "1020:" - trackbox reported rotor positions. Azimuth angle
'               follows and then elevation angle separated by
'               spaces.
Function TBRadioReadFrame(Handle%) As String
'This will hold the received frame
InBuff$ = ""
'initialize buffer pointer..
InBuffPtr% = 0
'this will indicate the position where the OK word is found.
'while it is -1 it means it haven't been found yet. after it
'is found we need to be sure that the reqd. number of bytes
'is received
found% = -1
'set timeout time...
a = Timer
Do
    'Wait until we receive bytes
    'or a time-out occurs
    Select Case Handle%
    Case 1
        Do
            
        Loop Until MSComm1.InBufferCount >= 1 Or Abs(Timer - a) > RadioControlDelay(Handle%)
        InString = MSComm1.Input
    Case 2
        Do
            
        Loop Until MSComm2.InBufferCount >= 1 Or Abs(Timer - a) > RadioControlDelay(Handle%)
        InString = MSComm2.Input
    Case 3
        Do
            
        Loop Until MSComm3.InBufferCount >= 1 Or Abs(Timer - a) > RadioControlDelay(Handle%)
        InString = MSComm3.Input
    End Select
    'if a byte was received we reset timer to wait for another
    'time-out period
    If Abs(Timer - a) < RadioControlDelay(Handle%) Then
        a = Timer
    End If
    'We pass the received bytes to the buffer
    For i% = 0 To LenB(InString) - 1
        InBuff$ = InBuff$ + Chr$(InString(i%))
    Next
    'we will examine the buffer searching for the OK string:
    f% = InStr(InBuff$, "Box>>")
Loop While f% = 0 And Abs(Timer - a) < RadioControlDelay(Handle%)
'if loop ended due to timeout:
If f% = 0 Then
    s$ = "00:"
    TBRadioReadFrame = s$
    Exit Function
Else
    'otherwise examine frame:
    a = InStr(LCase$(InBuff$), "error")
    If a Then
        s$ = "02:"
        TBRadioReadFrame = s$
        Exit Function
    End If
    a = InStr(InBuff$, "AZ=")
    If a Then
        'look for separator of raw data / angle value
        a = InStr(a, InBuff$, "/")
        Az = Firstnum(Mid$(InBuff$, a + 1))
        a = InStr(a, InBuff$, "EL=")
        If a Then
            'look for separator of raw data / angle value
            a = InStr(a, InBuff$, "/")
            El = Firstnum(Mid$(InBuff$, a + 1))
                        
            'we have all the info.
            s$ = "1020: " + Format(Az, "##0.0") + " " + Format(El, "##0.0")
            TBRadioReadFrame = s$
            Exit Function
        Else
            'non-valid frame:
            Call ErrorHandler("TrakBox comms.", 0)
        End If
    End If
    'if it is non of the obove types of frame -> just acknowledge
    s$ = "01:"
    TBRadioReadFrame = s$
End If
End Function
'This function tries to update downlink radios' freq and return 1 if successfull
'otherwise return 0
'Takes freq number from Downlink field in main window.
Function UpdateDownlink() As Integer
'assume we updated fine...
UpdateDownlink = 1
If frmMessage.MessageTimer.Enabled = True Then
    UpdateDownlink = 0
    Exit Function
End If

' Check if logging is active
If frmRadio.CheckLog.Value Then
    f% = FreeFile
    Open "Radio_Log.txt" For Append As f%
    Print #f%, "Updating " + DownlinkModel$ + " downlink radio on stream" + Str(DownlinkHandle%) + ", " + Str(Time)
    Close f%
End If

'for PSK downlink, freq. setting
'is enabled only at the beginning of the
'pass.
'also if there is a message beign shown, no comms to radio:
If DownlinkHandle% <> 0 And RadioDownlinkEnabled = True Then
    Select Case DownlinkModel$
    
    Case "FT-100" ' G6LVB 16 Dec 2000
        'if splitmode then downlink is VFO A:
        If DownlinkSplit% = 1 Then
            Call FT100RadioSetVFOA(DownlinkHandle%)
        End If
        
        Call FT100RadioSetFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%)
        
        If DownlinkSplit% = 1 Then
             'only return radio to VFO B if uplink is the master
            'band and both bands of the radio are being used:
            If SliderUplink.Value = 1 And (UplinkHandle% = DownlinkHandle% _
                    And UplinkHandle% <> 0) Then
                Call FT100RadioSetVFOB(DownlinkHandle%)
            End If
        End If
        
   Case "FT-817", "FT-897"
        ' Update freq only if not TXing
        If FT817RadioReadPTT(DownlinkHandle%) <> "ON" Then
            Call FT817RadioSetFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%)
        Else
            ' Check if logging is active
            If frmRadio.CheckLog.Value Then
                f% = FreeFile
                Open "Radio_Log.txt" For Append As f%
                Print #f%, "Cannot update FT-817 PTT is ON" + ", " + Str(Time)
                Close f%
            End If
            
            UpdateDownlink = 0
        End If
    
    Case "AR-8000"
        Call AR8000RadioSetFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%)
    
    Case "AR-3000A" ' EB2CTA
        Call AR3000ARadioSetFreq(100000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%)
    
    Case "AR-5000"
        Call AR8000RadioSetFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%)
    
    Case "FRG-9600"
        Call FRG9600RadioSetFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%)
    
    Case "VR-5000"
        Call VR5000RadioSetMainFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%)
        DownlinkRSSI.Value = VR5000RadioReadRSSI(DownlinkHandle%)
    
    Case "PCR-1000"
        Call ICPCRRadioSet(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkMode.text, DownlinkFilter, DownlinkHandle%)
        DownlinkRSSI.Value = ICPCRRadioReadRSSI(DownlinkHandle%)
    
    Case "IC-821", "IC-970"
        'if splitmode then downlink is VFO A:
        If DownlinkSplit% = 1 Then
        Else
            'downlink band must be Sub in satellite mode
            Call IC821RadioSub(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
        End If
        
        If DownlinkSplit% = 0 Or DownlinkTXOn% = 0 Then
            If IC821RadioSetFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%) <> 1 Then
                ' Check if logging is active
                If frmRadio.CheckLog.Value Then
                    f% = FreeFile
                    Open "Radio_Log.txt" For Append As f%
                    Print #f%, "Error updating freq. of " + DownlinkModel$ + " downlink radio on stream" + Str(DownlinkHandle%) + ", " + Str(Time)
                    Close f%
                End If
                
                UpdateDownlink = 0
            End If
         End If
        
        'only return radio to main (uplink) band if uplink is the master
        'band and both bands of the radio are being used:
        If SliderUplink.Value = 1 And (UplinkHandle% = DownlinkHandle% _
            And UplinkHandle% <> 0) Then
            If DownlinkSplit% = 1 Then
            Else
                Call IC821RadioMain(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
            End If
        End If
    
    Case "IC-910"
        'if splitmode then downlink is VFO A:
        If DownlinkSplit% = 1 Then
        Else
            'downlink band must be Main for IC910
            Call IC821RadioMain(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
        End If
        
        If DownlinkSplit% = 0 Or DownlinkTXOn% = 0 Then
            If IC821RadioSetFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%) <> 1 Then
                ' Check if logging is active
                If frmRadio.CheckLog.Value Then
                    f% = FreeFile
                    Open "Radio_Log.txt" For Append As f%
                    Print #f%, "Error updating freq. of " + DownlinkModel$ + " downlink radio on stream" + Str(DownlinkHandle%) + ", " + Str(Time)
                    Close f%
                End If
                
                UpdateDownlink = 0
            End If
         End If
        'only return radio to main band if uplink is the master
        'band and both bands of the radio are being used:
        If SliderUplink.Value = 1 And (UplinkHandle% = DownlinkHandle% _
            And UplinkHandle% <> 0) Then
            If DownlinkSplit% = 1 Then
            Else
                Call IC821RadioSub(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
            End If
        End If
    
    Case "IC-275", "IC-475", "IC-746", "IC-706"
        
        If DownlinkSplit% = 0 Or DownlinkTXOn% = 0 Then
            If IC821RadioSetFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%) <> 1 Then
                ' Check if logging is active
                If frmRadio.CheckLog.Value Then
                    f% = FreeFile
                    Open "Radio_Log.txt" For Append As f%
                    Print #f%, "Error updating freq. of " + DownlinkModel$ + " downlink radio on stream" + Str(DownlinkHandle%) + ", " + Str(Time)
                    Close f%
                End If
                
                UpdateDownlink = 0
            End If
        End If
        
    Case "IC-R7000"
        If IC821RadioSetFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%) <> 1 Then
            ' Check if logging is active
            If frmRadio.CheckLog.Value Then
                f% = FreeFile
                Open "Radio_Log.txt" For Append As f%
                Print #f%, "Error updating freq. of " + DownlinkModel$ + " downlink radio on stream" + Str(DownlinkHandle%) + ", " + Str(Time)
                Close f%
            End If
            
            UpdateDownlink = 0
        End If
    
    Case "IC-R8500"
        If IC821RadioSetFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%) <> 1 Then
            ' Check if logging is active
            If frmRadio.CheckLog.Value Then
                f% = FreeFile
                Open "Radio_Log.txt" For Append As f%
                Print #f%, "Error updating freq. of " + DownlinkModel$ + " downlink radio on stream" + Str(DownlinkHandle%) + ", " + Str(Time)
                Close f%
            End If
            
            UpdateDownlink = 0
        End If
    
    Case "FT-847"
        If DownlinkSplit% = 1 Then
            'if downlink radio is set to split-mode
            'We will use Repeater Offset for uplink freq.
            'Call FT847RadioShiftPlus(DownlinkHandle%)
            Call FT847RadioSetMainFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%)
        Else
            Call FT847RadioSetRXFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%)
        End If
    
    Case "FT-736"
        Call FT736RadioSetRXFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%)
    
    Case "TM-D700", "TH-D7"
        'only Band B can handle UHF frequencies
        If Cdbl2(DownlinkDDEFreq.text) >= 300 Then
            Call TMD700RadioSetB(DownlinkHandle%)
        Else
            Call TMD700RadioSetA(DownlinkHandle%)
        End If
        Call TMD700RadioSetFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%)
    
    Case "TS-790"
        Call TS790RadioSetSub(DownlinkHandle%, DownlinkBidir%)
        Call TS790RadioSetVFOA(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%, DownlinkBidir%)
        DownlinkRSSI.Value = TS790RadioReadSubRSSI(DownlinkHandle%)
    
    Case "TS-2000" ' 2 March 2000 G6LVB TS-2000
'        Call TS790RadioSetSub(DownlinkHandle%, DownlinkBidir%)
        Call TS790RadioSetVFOA(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%, DownlinkBidir%)
        DownlinkRSSI.Value = TS2000RadioReadSubRSSI(DownlinkHandle%)
    
    Case "TS-711", "TS-811"
        Call TS711RadioSetVFOA(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%)
    
    Case "TrakBox"
        Call TBRadioSetRXFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%)
    
    End Select
Else
    ' Check if logging is active
    If frmRadio.CheckLog.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Not supported " + DownlinkModel$ + " downlink radio on stream" + Str(DownlinkHandle%) + ", " + Str(Time)
        Close f%
    End If
    
    UpdateDownlink = 0
End If

End Function
'This function tries to update uplink radios' freq and return 1 if successfull
'otherwise return 0
'Takes freq number from Uplink field in main window.
Function UpdateUplink() As Integer
'assume success for now...
UpdateUplink = 1
If frmMessage.MessageTimer.Enabled = True Then
    UpdateUplink = 0
    Exit Function
End If

' Check if logging is active
If frmRadio.CheckLog.Value Then
    f% = FreeFile
    Open "Radio_Log.txt" For Append As f%
    Print #f%, "Updating " + UplinkModel$ + " uplink radio on stream" + Str(UplinkHandle%) + ", " + Str(Time)
    Close f%
End If

If UplinkHandle% <> 0 Then
    Select Case UplinkModel$
    Case "IC-821", "IC-970"
        If UplinkSplit% = 1 Then
        Else
            'uplink band is Main-Band in satellite mode
            Call IC821RadioMain(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
        End If
        
        If UplinkSplit% = 1 Then
            'if downlink radio is set to split-mode
            'We will duplex mode
            If IC821RadioSetOffset(1000000# * ((Cdbl2(UplinkFreq.text) - UplinkLO#) - (Cdbl2(DownlinkFreq.text) - DownlinkLO#)) _
                   , UplinkCIVAddress%, UplinkBidir%, UplinkHandle%) <> 1 Then
               
               'If rig doesn't want to update offset it's probably on TX mode, we don't report this as error
               
                ' Check if logging is active
                
'                If frmRadio.CheckLog.Value Then
'                    f% = FreeFile
'                    Open "Radio_Log.txt" For Append As f%
'                    Print #f%, "Error updating " + UplinkModel$ + " uplink radio on stream" + Str(UplinkHandle%) + ", " + Str(Time)
'                    Close f%
'                End If
                
                UpdateUplink = 0
            End If
        Else
            If IC821RadioSetFreq(1000000# * (Cdbl2(UplinkFreq.text) - UplinkLO#), UplinkCIVAddress%, UplinkBidir%, UplinkHandle%) <> 1 Then
                ' Check if logging is active
                If frmRadio.CheckLog.Value Then
                    f% = FreeFile
                    Open "Radio_Log.txt" For Append As f%
                    Print #f%, "Error updating " + UplinkModel$ + " uplink radio on stream" + Str(UplinkHandle%) + ", " + Str(Time)
                    Close f%
                End If
                
                UpdateUplink = 0
            End If
        End If
        
        'only return radio to sub band if downlink is the master
        'band and both bands of the radio are being used:
        If SliderDownlink.Value = 1 And (UplinkHandle% = DownlinkHandle% _
            And DownlinkHandle% <> 0) Then
            
            'if splitmode then dnlink is VFO A:
            If UplinkSplit% = 1 Then
            Else
                'dnlink band is Sub-Band in satellite mode
                Call IC821RadioSub(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
            End If
        End If
    
    Case "IC-910"
        If DownlinkSplit% = 1 Then
        Else
            'uplink band is Sub-Band in satellite mode
            Call IC821RadioSub(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
        End If
        
        If UplinkSplit% = 1 Then
            'if downlink radio is set to split-mode
            'We will duplex mode
            If IC821RadioSetOffset(1000000# * ((Cdbl2(UplinkFreq.text) - UplinkLO#) - (Cdbl2(DownlinkFreq.text) - DownlinkLO#)) _
                   , UplinkCIVAddress%, UplinkBidir%, UplinkHandle%) <> 1 Then
               
               'If rig doesn't want to update offset it's probably on TX mode, we don't report this as error
               
                UpdateUplink = 0
            End If
        Else
            If IC821RadioSetFreq(1000000# * (Cdbl2(UplinkFreq.text) - UplinkLO#), UplinkCIVAddress%, UplinkBidir%, UplinkHandle%) <> 1 Then
                ' Check if logging is active
                If frmRadio.CheckLog.Value Then
                    f% = FreeFile
                    Open "Radio_Log.txt" For Append As f%
                    Print #f%, "Error updating " + UplinkModel$ + " uplink radio on stream" + Str(UplinkHandle%) + ", " + Str(Time)
                    Close f%
                End If
                
                UpdateUplink = 0
            End If
         End If
        'only return radio to sub band if downlink is the master
        'band and both bands of the radio are being used:
        If SliderDownlink.Value = 1 And (UplinkHandle% = DownlinkHandle% _
            And DownlinkHandle% <> 0) Then
        
            If DownlinkSplit% = 1 Then
            Else
                'dnlink band is Main-Band in satellite mode
                Call IC821RadioMain(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
            End If
        End If
    
    Case "IC-275", "IC-475", "IC-746", "IC-706"
        
        If UplinkSplit% = 1 Then
            'if downlink radio is set to split-mode
            'We will duplex mode
            If IC821RadioSetOffset(1000000# * ((Cdbl2(UplinkFreq.text) - UplinkLO#) - (Cdbl2(DownlinkFreq.text) - DownlinkLO#)) _
                   , UplinkCIVAddress%, UplinkBidir%, UplinkHandle%) <> 1 Then
               
               'If rig doesn't want to update offset it's probably on TX mode, we don't report this as error
               
                UpdateUplink = 0
            End If
        Else
            If IC821RadioSetFreq(1000000# * (Cdbl2(UplinkFreq.text) - UplinkLO#), UplinkCIVAddress%, UplinkBidir%, UplinkHandle%) <> 1 Then
                ' Check if logging is active
                If frmRadio.CheckLog.Value Then
                    f% = FreeFile
                    Open "Radio_Log.txt" For Append As f%
                    Print #f%, "Error updating " + UplinkModel$ + " uplink radio on stream" + Str(UplinkHandle%) + ", " + Str(Time)
                    Close f%
                End If
                
                UpdateUplink = 0
            End If
        End If
        
    Case "FT-847"
        If UplinkSplit% = 1 Then
            'if uplink radio is set to split-mode
            'We will use Repeater Offset to obtain correct TX freq.
            'Need to calculate difference between tx and rx freqs.
            a = 1000000# * ((Cdbl2(UplinkFreq.text) - UplinkLO#) - (Cdbl2(DownlinkFreq.text) - DownlinkLO#))
            If a >= 0 Then
                Call FT847RadioShiftPlus(UplinkHandle%)
            Else
                Call FT847RadioShiftMinus(UplinkHandle%)
                a = -a
            End If
            'FT847 truncates shift frequency at 10KHz increments
            'we anticipate this and round-off shifts to the nearest 10KHz multiple
            a = 10000# * (CLng(a / 10000#))
            Call FT847RadioSetShiftFreq(CLng(a), UplinkHandle%)
        Else
            Call FT847RadioSetTXFreq(1000000# * (Cdbl2(UplinkFreq.text) - UplinkLO#), UplinkHandle%)
        End If
        'Add PA2EON
        If UplinkCTCSS% = 1 Then
             Call FT847RadioCTCSSTXOn(UplinkHandle%)
        End If
    
    Case "FT-736"
        Call FT736RadioSetTXFreq(1000000# * (Cdbl2(UplinkFreq.text) - UplinkLO#), UplinkHandle%)
     
     Case "FT-100" ' G6LVB 16 Dec 2000
        'if splitmode, uplink is VFO B
        If UplinkSplit% = 1 Then
            Call FT100RadioSetVFOB(UplinkHandle%)
        End If
        
        Call FT100RadioSetFreq(1000000# * (Cdbl2(UplinkFreq.text) - UplinkLO#), UplinkHandle%)
        
        If UplinkSplit% = 1 Then
            'only return radio to VFO A if downlink is the master
            'band and both bands of the radio are being used:
            If SliderDownlink.Value = 1 And (UplinkHandle% = DownlinkHandle% _
                And DownlinkHandle% <> 0) Then
                Call FT100RadioSetVFOA(UplinkHandle%)
            End If
        End If
        
    Case "FT-817", "FT-897"
        ' Update freq only if not TXing
        If FT817RadioReadPTT(UplinkHandle%) <> "ON" Then
            'if splitmode then we need to toggle VFO:
            If UplinkSplit% = 1 Then
                Call FT817RadioToggleVFO(UplinkHandle%)
            End If
            
            Call FT817RadioSetFreq(1000000# * (Cdbl2(UplinkFreq.text) - UplinkLO#), UplinkHandle%)
            
            'if splitmode then we need to toggle VFO:
            If UplinkSplit% = 1 Then
                Call FT817RadioToggleVFO(UplinkHandle%)
            End If
        Else
            ' Check if logging is active
            If frmRadio.CheckLog.Value Then
                f% = FreeFile
                Open "Radio_Log.txt" For Append As f%
                Print #f%, "Couldn't update " + UplinkModel$ + " PTT is ON , " + Str(Time)
                Close f%
            End If
            
            UpdateUplink = 0
        End If
   
   Case "TM-D700", "TH-D7"
        If Cdbl2(UplinkDDEFreq.text) >= 300 Then
            Call TMD700RadioSetB(UplinkHandle%)
        Else
            Call TMD700RadioSetA(UplinkHandle%)
        End If
        Call TMD700RadioSetFreq(1000000# * (Cdbl2(UplinkFreq.text) - UplinkLO#), UplinkHandle%)
    
    Case "TS-790"
        Call TS790RadioSetMain(UplinkHandle%, UplinkBidir%)
        Call TS790RadioSetVFOA(1000000# * (Cdbl2(UplinkFreq.text) - UplinkLO#), UplinkHandle%, UplinkBidir%)
    
    Case "TS-2000" ' 2 March 2000 G6LVB TS-2000
'        Call TS790RadioSetMain(UplinkHandle%, UplinkBidir%)
        Call TS790RadioSetVFOB(1000000# * (Cdbl2(UplinkFreq.text) - UplinkLO#), UplinkHandle%, UplinkBidir%)
    
    Case "TS-711", "TS-811"
        ' use VFO B for TX if inband
        If (DownlinkHandle% = UplinkHandle%) Then
            Call TS711RadioSetVFOB(1000000# * (Cdbl2(UplinkFreq.text) - UplinkLO#), UplinkHandle%)
        Else
            Call TS711RadioSetVFOA(1000000# * (Cdbl2(UplinkFreq.text) - UplinkLO#), UplinkHandle%)
        End If
    
    Case "TrakBox"
        Call TBRadioSetTXFreq(1000000# * (Cdbl2(UplinkFreq.text) - UplinkLO#), UplinkHandle%)
    
    End Select
Else
    ' Check if logging is active
    If frmRadio.CheckLog.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Not supported " + UplinkModel$ + " uplink radio on stream" + Str(UplinkHandle%) + ", " + Str(Time)
        Close f%
    End If
    
    UpdateUplink = 0
End If

End Function
'Generic function that returns frequency of downlink radio
Function ReadDownlinkFreq() As Double

' Check if logging is active
If frmRadio.CheckLog.Value Then
    ff% = FreeFile
    Open "Radio_Log.txt" For Append As ff%
    Print #ff%, "Reading freq. of " + DownlinkModel$ + " downlink radio on stream" + Str(DownlinkHandle%) + ", " + Str(Time)
    Close ff%
End If

f# = 0
If DownlinkHandle% <> 0 Then
    Select Case DownlinkModel$
    Case "FT-817", "FT-897"
        f# = FT817RadioReadFreq(DownlinkHandle%)
        If f# <> 0 Then
            f# = f# + 1000000# * DownlinkLO#
        Else
            ' Check if logging is active
            If frmRadio.CheckLog.Value Then
                ff% = FreeFile
                Open "Radio_Log.txt" For Append As ff%
                Print #ff%, "Error reading freq. of " + DownlinkModel$ + " downlink radio, " + Str(Time)
                Close ff%
            End If
            
        End If
        
    Case "IC-821", "IC-970"
        'if splitmode then check if rig is on TX
        If DownlinkSplit% = 1 Then
            f# = IC821RadioReadOffset(DownlinkCIVAddress%, DownlinkHandle%)
            'if radio cannot set offset we assume it's on TX
            If IC821RadioSetOffset(f#, DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%) <> 1 Then
               DownlinkTXOn% = 1
            Else
               DownlinkTXOn% = 0
            End If
        Else
            'downlink band must be Sub in satellite mode
            Call IC821RadioSub(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
        End If
        
        If DownlinkTXOn% = 0 Then
            f# = IC821RadioReadFreq(DownlinkCIVAddress%, DownlinkHandle%)
            If f# <> 0 Then
                f# = f# + 1000000# * DownlinkLO#
            Else
                ' Check if logging is active
                If frmRadio.CheckLog.Value Then
                    ff% = FreeFile
                    Open "Radio_Log.txt" For Append As ff%
                    Print #ff%, "Error reading freq. of " + DownlinkModel$ + " downlink radio, " + Str(Time)
                    Close ff%
                End If
                
            End If
         Else
            f# = 0
         End If
        
        'only return radio to main band if uplink is the master
        'band and both bands of the radio are being used:
        If SliderUplink.Value = 1 And (UplinkHandle% = DownlinkHandle% _
            And UplinkHandle% <> 0) Then
            If DownlinkSplit% = 1 Then
            Else
                Call IC821RadioMain(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
            End If
        End If
    
    Case "IC-910"
        'if splitmode then check if rig is on TX
        If DownlinkSplit% = 1 Then
            f# = IC821RadioReadOffset(DownlinkCIVAddress%, DownlinkHandle%)
            'if radio cannot set offset we assume it's on TX
            If IC821RadioSetOffset(f#, DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%) <> 1 Then
               DownlinkTXOn% = 1
            Else
               DownlinkTXOn% = 0
            End If
        Else
            'downlink band must be Main in satellite mode
            Call IC821RadioMain(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
        End If
        
        If DownlinkTXOn% = 0 Then
            f# = IC821RadioReadFreq(DownlinkCIVAddress%, DownlinkHandle%)
            If f# <> 0 Then
                f# = f# + 1000000# * DownlinkLO#
            Else
                ' Check if logging is active
                If frmRadio.CheckLog.Value Then
                    ff% = FreeFile
                    Open "Radio_Log.txt" For Append As ff%
                    Print #ff%, "Error reading freq. of " + DownlinkModel$ + " downlink radio, " + Str(Time)
                    Close ff%
                End If
                
            End If
         Else
            f# = 0
         End If
        'only return radio to main band if uplink is the master
        'band and both bands of the radio are being used:
        If SliderUplink.Value = 1 And (UplinkHandle% = DownlinkHandle% _
            And UplinkHandle% <> 0) Then
            If DownlinkSplit% = 1 Then
                Call IC706RadioSetVFOB(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
            Else
                Call IC821RadioSub(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
            End If
        End If
    
    Case "IC-275", "IC-475", "IC-746", "IC-706"
        'if splitmode then check if rig is on TX
        If DownlinkSplit% = 1 Then
            f# = IC821RadioReadOffset(DownlinkCIVAddress%, DownlinkHandle%)
            'if radio cannot set offset we assume it's on TX
            If IC821RadioSetOffset(f#, DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%) <> 1 Then
               DownlinkTXOn% = 1
            Else
               DownlinkTXOn% = 0
            End If
        End If
        
        If DownlinkTXOn% = 0 Then
            f# = IC821RadioReadFreq(DownlinkCIVAddress%, DownlinkHandle%)
            If f# <> 0 Then
                f# = f# + 1000000# * DownlinkLO#
            Else
                ' Check if logging is active
                If frmRadio.CheckLog.Value Then
                    ff% = FreeFile
                    Open "Radio_Log.txt" For Append As ff%
                    Print #ff%, "Error reading freq. of " + DownlinkModel$ + " downlink radio, " + Str(Time)
                    Close ff%
                End If
                
            End If
         Else
            f# = 0
         End If
        
    Case "IC-R7000"
        f# = IC821RadioReadFreq(DownlinkCIVAddress%, DownlinkHandle%)
        If f# <> 0 Then
            f# = f# + 1000000# * DownlinkLO#
        Else
            ' Check if logging is active
            If frmRadio.CheckLog.Value Then
                ff% = FreeFile
                Open "Radio_Log.txt" For Append As ff%
                Print #ff%, "Error reading freq. of " + DownlinkModel$ + " downlink radio, " + Str(Time)
                Close ff%
            End If
            
        End If
    
    Case "IC-R8500"
        f# = IC821RadioReadFreq(DownlinkCIVAddress%, DownlinkHandle%)
        If f# <> 0 Then
            f# = f# + 1000000# * DownlinkLO#
        Else
            ' Check if logging is active
            If frmRadio.CheckLog.Value Then
                ff% = FreeFile
                Open "Radio_Log.txt" For Append As ff%
                Print #ff%, "Error reading freq. of " + DownlinkModel$ + " downlink radio, " + Str(Time)
                Close ff%
            End If
            
        End If
    
    Case "FT-847"
        If DownlinkSplit% = 1 Then
            'if splitmode then downlink is Main VFO:
            f# = FT847RadioReadMainFreq(DownlinkHandle%)
        Else
            f# = FT847RadioReadRXFreq(DownlinkHandle%)
        End If
        If f# <> 0 Then
            f# = f# + 1000000# * DownlinkLO#
        Else
            ' Check if logging is active
            If frmRadio.CheckLog.Value Then
                ff% = FreeFile
                Open "Radio_Log.txt" For Append As ff%
                Print #ff%, "Error reading freq. of " + DownlinkModel$ + " downlink radio, " + Str(Time)
                Close ff%
            End If
            
        End If
    
    Case "FT-100"
        If DownlinkSplit% = 1 Then
            'if splitmode then downlink is VFO A:
            Call FT100RadioSetVFOA(DownlinkHandle%)
        End If
        
        f# = FT100RadioReadFreq(DownlinkHandle%)
        If f# <> 0 Then
            f# = f# + 1000000# * DownlinkLO#
        Else
            ' Check if logging is active
            If frmRadio.CheckLog.Value Then
                ff% = FreeFile
                Open "Radio_Log.txt" For Append As ff%
                Print #ff%, "Error reading freq. of " + DownlinkModel$ + " downlink radio, " + Str(Time)
                Close ff%
            End If
            
        End If
        
        If DownlinkSplit% = 1 Then
            'only return radio to VFO B if uplink is the master
            'band and both bands of the radio are being used:
            If SliderUplink.Value = 1 And (UplinkHandle% = DownlinkHandle% _
                And UplinkHandle% <> 0) Then
                Call FT100RadioSetVFOB(DownlinkHandle%)
            End If
        End If
        
    Case "TS-790"
        Call TS790RadioSetSub(DownlinkHandle%, DownlinkBidir%)
        f# = TS790RadioReadVFOA(DownlinkHandle%)
        If f# <> 0 Then
            f# = f# + 1000000# * DownlinkLO#
        Else
            ' Check if logging is active
            If frmRadio.CheckLog.Value Then
                ff% = FreeFile
                Open "Radio_Log.txt" For Append As ff%
                Print #ff%, "Error reading freq. of " + DownlinkModel$ + " downlink radio, " + Str(Time)
                Close ff%
            End If
            
        End If
    
    Case "TS-2000"
'        Call TS790RadioSetSub(UplinkHandle%, UplinkBidir%)
        f# = TS790RadioReadVFOA(DownlinkHandle%)
        If f# <> 0 Then
            f# = f# + 1000000# * DownlinkLO#
        Else
            ' Check if logging is active
            If frmRadio.CheckLog.Value Then
                ff% = FreeFile
                Open "Radio_Log.txt" For Append As ff%
                Print #ff%, "Error reading freq. of " + DownlinkModel$ + " downlink radio, " + Str(Time)
                Close ff%
            End If
            
        End If
    
    Case "TS-711", "TS-811"
        f# = TS790RadioReadVFOA(DownlinkHandle%)
        If f# <> 0 Then
            f# = f# + 1000000# * DownlinkLO#
        Else
            ' Check if logging is active
            If frmRadio.CheckLog.Value Then
                ff% = FreeFile
                Open "Radio_Log.txt" For Append As ff%
                Print #ff%, "Error reading freq. of " + DownlinkModel$ + " downlink radio, " + Str(Time)
                Close ff%
            End If
            
        End If
    
    Case Else
            ' Check if logging is active
        If frmRadio.CheckLog.Value Then
            ff% = FreeFile
            Open "Radio_Log.txt" For Append As ff%
            Print #ff%, "Not supported reading freq. of " + DownlinkModel$ + " downlink radio, " + Str(Time)
            Close ff%
        End If
        
    End Select
End If
'If f# <> 0 Then ReadDownlinkFreq = f#
ReadDownlinkFreq = f#

End Function
'Generic function that returns frequency of uplink radio
Function ReadUplinkFreq() As Double

' Check if logging is active
If frmRadio.CheckLog.Value Then
    ff% = FreeFile
    Open "Radio_Log.txt" For Append As ff%
    Print #ff%, "Reading freq. of " + UplinkModel$ + " uplink radio on stream" + Str(UplinkHandle%) + ", " + Str(Time)
    Close ff%
End If

f# = 0
If UplinkHandle% <> 0 Then
    Select Case UplinkModel$
    
    Case "FT-100"
        If UplinkSplit% = 1 Then
            'uplink is VFO-B:
            Call FT100RadioSetVFOB(UplinkHandle%)
        End If
        
        f# = FT100RadioReadFreq(UplinkHandle%)
        If f# <> 0 Then
            f# = f# + 1000000# * UplinkLO#
        Else
            ' Check if logging is active
            If frmRadio.CheckLog.Value Then
                ff% = FreeFile
                Open "Radio_Log.txt" For Append As ff%
                Print #ff%, "Error reading freq. of " + UplinkModel$ + " uplink radio, " + Str(Time)
                Close ff%
            End If
            
        End If
        
        If UplinkSplit% = 1 Then
            'only return radio to VFO A if downlink is the master
            'band and both bands of the radio are being used:
            If SliderDownlink.Value = 1 And (UplinkHandle% = DownlinkHandle% _
                    And DownlinkHandle% <> 0) Then
                Call FT100RadioSetVFOA(UplinkHandle%)
            End If
        End If
        
    Case "FT-817", "FT-897"
        ' Update freq only if not TXing
        If FT817RadioReadPTT(UplinkHandle%) <> "ON" Then
            If UplinkSplit% = 1 Then
                'need to toggle VFOs to read uplink freq
                Call FT817RadioToggleVFO(UplinkHandle%)
            End If
            
            f# = FT817RadioReadFreq(UplinkHandle%)
            If f# <> 0 Then
                f# = f# + 1000000# * UplinkLO#
            Else
                ' Check if logging is active
                If frmRadio.CheckLog.Value Then
                    ff% = FreeFile
                    Open "Radio_Log.txt" For Append As ff%
                    Print #ff%, "Error reading freq. of " + UplinkModel$ + " uplink radio, " + Str(Time)
                    Close ff%
                End If
                
            End If
            
            If UplinkSplit% = 1 Then
                'toggle back to RX VFO
                Call FT817RadioToggleVFO(UplinkHandle%)
            End If
        Else
            ' Check if logging is active
            If frmRadio.CheckLog.Value Then
                ff% = FreeFile
                Open "Radio_Log.txt" For Append As ff%
                Print #ff%, "Couldn't read freq. of " + UplinkModel$ + " PTT is ON, " + Str(Time)
                Close ff%
            End If
            
        End If
        
    Case "IC-821", "IC-970"
        If UplinkSplit% = 1 Then
            'to get uplink freq in duplex se need to read offset and RX freq
            f# = IC821RadioReadOffset(UplinkCIVAddress%, UplinkHandle%)
            If UplinkDuplexPlus% = 0 Then
               'If duplex offset is negative
               f# = -f#
            End If
        Else
            'uplink band must be Main in satellite mode
            Call IC821RadioMain(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
            
            f# = 0
        End If
        
        f# = f# + IC821RadioReadFreq(UplinkCIVAddress%, UplinkHandle%)
        If f# <> 0 Then
            f# = f# + 1000000# * UplinkLO#
        Else
            ' Check if logging is active
            If frmRadio.CheckLog.Value Then
                ff% = FreeFile
                Open "Radio_Log.txt" For Append As ff%
                Print #ff%, "Error reading freq. of " + UplinkModel$ + " uplink radio, " + Str(Time)
                Close ff%
            End If
            
        End If
        'only return radio to main band if downlink is the master
        'band and both bands of the radio are being used:
        If SliderDownlink.Value = 1 And (UplinkHandle% = DownlinkHandle% _
            And DownlinkHandle% <> 0) Then
            If UplinkSplit% = 1 Then
            Else
                'dnlink band is Sub-Band in satellite mode
                Call IC821RadioSub(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
            End If
        End If
        
    Case "IC-910"
        If UplinkSplit% = 1 Then
            'to get uplink freq in duplex se need to read offset and RX freq
            f# = IC821RadioReadOffset(UplinkCIVAddress%, UplinkHandle%)
            If UplinkDuplexPlus% = 0 Then
               'If duplex offset is negative
               f# = -f#
            End If
        Else
            'uplink band must be sub in satellite mode
            Call IC821RadioSub(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
            
            f# = 0
        End If
        
        f# = f# + IC821RadioReadFreq(UplinkCIVAddress%, UplinkHandle%)
        If f# <> 0 Then
            f# = f# + 1000000# * UplinkLO#
        Else
            ' Check if logging is active
            If frmRadio.CheckLog.Value Then
                ff% = FreeFile
                Open "Radio_Log.txt" For Append As ff%
                Print #ff%, "Error reading freq. of " + UplinkModel$ + " uplink radio, " + Str(Time)
                Close ff%
            End If
            
        End If
        'only return radio to main band if downlink is the master
        'band and both bands of the radio are being used:
        If SliderDownlink.Value = 1 And (UplinkHandle% = DownlinkHandle% _
            And DownlinkHandle% <> 0) Then
            'if splitmode then dnlink is VFO A:
            If UplinkSplit% = 1 Then
            Else
                'dnlink band is main-Band in satellite mode
                Call IC821RadioMain(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
            End If
        End If
    
    Case "IC-275", "IC-475", "IC-746", "IC-706"
        If UplinkSplit% = 1 Then
            'to get uplink freq in duplex se need to read offset and RX freq
            f# = IC821RadioReadOffset(UplinkCIVAddress%, UplinkHandle%)
            If UplinkDuplexPlus% = 0 Then
               'If duplex offset is negative
               f# = -f#
            End If
        Else
            f# = 0
        End If
        
        f# = f# + IC821RadioReadFreq(UplinkCIVAddress%, UplinkHandle%)
        If f# <> 0 Then
            f# = f# + 1000000# * UplinkLO#
        Else
            ' Check if logging is active
            If frmRadio.CheckLog.Value Then
                ff% = FreeFile
                Open "Radio_Log.txt" For Append As ff%
                Print #ff%, "Error reading freq. of " + UplinkModel$ + " uplink radio, " + Str(Time)
                Close ff%
            End If
        End If
               
    Case "FT-847"
        If UplinkSplit% = 1 Then
            'uplink is Main VFO plus Repeater Offset, but we cannot
            'read offset... can do nothing..
            f# = 0
        Else
            f# = FT847RadioReadTXFreq(UplinkHandle%)
        End If
        If f# <> 0 Then
            f# = f# + 1000000# * UplinkLO#
        Else
            ' Check if logging is active
            If frmRadio.CheckLog.Value Then
                ff% = FreeFile
                Open "Radio_Log.txt" For Append As ff%
                Print #ff%, "Error reading freq. of " + UplinkModel$ + " uplink radio, " + Str(Time)
                Close ff%
            End If
            
        End If
    
    Case "TS-790"
        Call TS790RadioSetMain(UplinkHandle%, UplinkBidir%)
        f# = TS790RadioReadVFOA(DownlinkHandle%)
        If f# <> 0 Then
            f# = f# + 1000000# * UplinkLO#
        Else
            ' Check if logging is active
            If frmRadio.CheckLog.Value Then
                ff% = FreeFile
                Open "Radio_Log.txt" For Append As ff%
                Print #ff%, "Error reading freq. of " + UplinkModel$ + " uplink radio, " + Str(Time)
                Close ff%
            End If
            
        End If
    
    Case "TS-2000"
 '       Call TS790RadioSetMain(UplinkHandle%, UplinkBidir%)
        f# = TS790RadioReadVFOB(DownlinkHandle%)
        If f# <> 0 Then
            f# = f# + 1000000# * UplinkLO#
        Else
            ' Check if logging is active
            If frmRadio.CheckLog.Value Then
                ff% = FreeFile
                Open "Radio_Log.txt" For Append As ff%
                Print #ff%, "Error reading freq. of " + UplinkModel$ + " uplink radio, " + Str(Time)
                Close ff%
            End If
            
        End If
    
    Case "TS-711", "TS-811"
        ' use VFO B for TX if inband
        If (DownlinkHandle% = UplinkHandle%) Then
            f# = TS790RadioReadVFOB(DownlinkHandle%)
            If f# <> 0 Then
                f# = f# + 1000000# * UplinkLO#
            Else
                ' Check if logging is active
                If frmRadio.CheckLog.Value Then
                    ff% = FreeFile
                    Open "Radio_Log.txt" For Append As ff%
                    Print #ff%, "Error reading freq. of " + UplinkModel$ + " uplink radio, " + Str(Time)
                    Close ff%
                End If
                
            End If
        
        Else
            f# = TS790RadioReadVFOA(UplinkHandle%)
            If f# <> 0 Then
                f# = f# + 1000000# * UplinkLO#
            Else
                ' Check if logging is active
                If frmRadio.CheckLog.Value Then
                    ff% = FreeFile
                    Open "Radio_Log.txt" For Append As ff%
                    Print #ff%, "Error reading freq. of " + UplinkModel$ + " uplink radio, " + Str(Time)
                    Close ff%
                End If
                
            End If

        End If
    
    Case Else
        ' Check if logging is active
        If frmRadio.CheckLog.Value Then
            ff% = FreeFile
            Open "Radio_Log.txt" For Append As ff%
            Print #ff%, "Not supported reading freq. of " + UplinkModel$ + " uplink radio, " + Str(Time)
            Close ff%
        End If
            
    End Select
End If
If f# <> 0 Then ReadUplinkFreq = f#

End Function
Private Sub DDELabel_LinkNotify()
'This routine is triggered by the LinkNotify event of the DDE-link
'this event should be caused by the DDE server each time data is
'changed.
'this is for the case that automatic DDE link update is
'selected, timer1 will be disabled so the data-processing
'routine will be called each time the DDE info is modified
If DDEPollTimer.Enabled = False Then
    Call DDEPollTimer_Timer
End If
End Sub


Private Sub DownlinkIndex_Change()
' Check if logging is active
If frmRadio.CheckLog.Value Then
    f% = FreeFile
    Open "Radio_Log.txt" For Append As f%
    Print #f%, "Downlink Radio Changed to " + DownlinkIndex.text + ", " + Str(Time)
    Close f%
End If

If DownlinkHandle% <> 0 Then
'this is not completely accurate, but try not to park a radio that
'is being used by other band...
    If UplinkHandle% = DownlinkHandle% And DownlinkModel$ = UplinkModel$ Then
        If UplinkModel$ = "TS-2000" Then ' 13 Feb 2002 G6LVB Revert back to previous non-sat setting
            Call ParkDownlinkRadio
        End If
    Else
        Call ParkDownlinkRadio
    End If
End If
'put accesory ports to zero
If DownlinkAccPort% > 0 Then
    Call OutPort(DownlinkAccPort%, 0)
End If
'Terminate downlink radio control port
'only if same handle is not in use by Uplink radio or Rotor
If (DownlinkHandle% <> UplinkHandle%) And (DownlinkHandle% <> RotorHandle%) Then
    ClosePort (DownlinkHandle%)
End If
DownlinkHandle% = 0

'open the selected radio's ports and initialize public
'variables with comms settings
If Cdbl2(DownlinkIndex.text) <> 0 Then
    'if there is any error opening the port -> release donwlink handle:
    If OpenDownlinkPort <> 0 Then
        DownlinkHandle% = 0
        Call frmMessage.ShowMessage("Error opening" + Chr$(13) _
            + "Downlink Radio's Port", 10)
        
        ' Check if logging is active
        If frmRadio.CheckLog.Value Then
            f% = FreeFile
            Open "Radio_Log.txt" For Append As f%
            Print #f%, "Error opening downlink radio's port, " + Str(Time)
            Close f%
        End If
        
        Exit Sub
    End If
    'turn on radio and initialize freq & mode to dde values, if no
    'dde values then read current freq from radio
    If DownlinkHandle% <> 0 Then
        Call ActivateDownlinkRadio
    End If
    'if both radios use the same acc.port we OR the values,
    'otherwise they are left independent
    If DownlinkAccPort% = UplinkAccPort% Then
        If DownlinkAccPort% <> 0 Then
            DownlinkAccPortValue% = DownlinkAccPortValur% Or UplinkAccPortValue%
            Call OutPort(DownlinkAccPort%, CInt(DownlinkAccPortValue%))
        End If
    Else
        If DownlinkAccPort% <> 0 Then
            Call OutPort(DownlinkAccPort%, CInt(DownlinkAccPortValue%))
        End If
    End If
End If
End Sub
Private Sub DownlinkIndex_Click()
    Call DownlinkIndex_Change
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' check press of Ctrl-S for slider in-focus
If KeyCode = Asc("S") And (Shift And vbCtrlMask) Then
    Slider.SetFocus
' check Ctrl-R for reverse tracking
ElseIf KeyCode = Asc("R") And (Shift And vbCtrlMask) Then
    TrackRev.Value = Not (TrackRev.Value)
    TrackDir.Value = False
' check Ctrl-D for reverse tracking
ElseIf KeyCode = Asc("D") And (Shift And vbCtrlMask) Then
    TrackDir.Value = Not (TrackDir.Value)
    TrackRev.Value = False
' check Ctrl-U for uplink master
ElseIf KeyCode = Asc("U") And (Shift And vbCtrlMask) Then
    SliderDownlink.Value = 0
    SliderUplink.Value = Abs(1 - SliderUplink.Value)
' check Ctrl-W for downlink master
ElseIf KeyCode = Asc("W") And (Shift And vbCtrlMask) Then
    SliderDownlink.Value = Abs(1 - SliderDownlink.Value)
    SliderUplink.Value = 0
' check Ctrl-Space toggle beetween no master and previous selection
ElseIf KeyCode = Asc(" ") Then
    UpdateRadioButton.SetFocus
    If SliderDownlinkMemory% = 1 Or SliderUplinkMemory% = 1 Then
    
        SliderDownlink.Value = SliderDownlinkMemory%
        SliderUplink.Value = SliderUplinkMemory%
        
        SliderDownlinkMemory% = 0
        SliderUplinkMemory% = 0
    Else
        SliderDownlinkMemory% = SliderDownlink.Value
        SliderUplinkMemory% = SliderUplink.Value
        
        SliderDownlink.Value = 0
        SliderUplink.Value = 0
    End If

' check for Ctrl-Q
ElseIf KeyCode = Asc("Q") And (Shift And vbCtrlMask) Then
    Call Form_Terminate
End If
End Sub


Private Sub Minus10KHzButton_Click()
ButtonCorrection = ButtonCorrection - 10000#
End Sub

Private Sub Plus10KHzButton_Click()
ButtonCorrection = ButtonCorrection + 10000#
End Sub
Private Sub RadioControlLoopTimer_Timer()

'if pop-up message is displayed, no updates.
If frmMessage.MessageTimer.Enabled = True Then Exit Sub

'in order not to anidate...
RadioControlLoopTimer.Enabled = False

'if downlink is master:
If DownlinkHandle% <> 0 And SliderDownlink.Value = 1 _
    And DownlinkBidir% Then
    'read master radio freq., if radio answered then update R# with
    'the answer.
    a = ReadDownlinkFreq
    If a > 0 + 1000000# * DownlinkLO# Then
        RD# = a
        'if slider or buttons changed we need to update radio's freq
        'and then process as if user had changed it from radio
        a = Slider.Value - SliderCorrection
        If a <> 0 Or ButtonCorrection <> 0 Then
            RD# = RD# + a + ButtonCorrection
            b = Cdbl2(DownlinkFreq.text)
            DownlinkFreq.text = Str(RD# / 1000000#)
            Call UpdateDownlink
            'we have updated freq., need to restore field contents
            'as this will be used to calculate user intervention.
            DownlinkFreq.text = Str(b)
            SliderCorrection = Slider.Value
            ButtonCorrection = 0
        End If
        If Abs(Slider.Value) = 1000 Then
            Slider.Value = 0
            SliderCorrection = 0
        End If
        
        'check if user changed dnlink freq (master) from radio:
        If RD# <> Fix(1000000# * Cdbl2(DownlinkFreq.text)) Then
            'Activate flag changing downlink
            cd% = 1
            'before changing dnlinkcorrection check for variation to
            'be substracted/added from/to uplinkcorrection:
            If UplinkHandle% <> 0 And SliderUplink.Value = 0 Then
                If TrackRev.Value = True Then
                    UplinkCorrection = Fix(UplinkCorrection - CLng( _
                        RD# - 1000000# * Cdbl2(DownlinkFreq.text)))
                    uu% = 1
                ElseIf TrackDir.Value = True Then
                    UplinkCorrection = UplinkCorrection + CLng( _
                        RD# - 1000000# * Cdbl2(DownlinkFreq.text))
                    uu% = 1
                End If
            End If
            'Calculate final correction amount, this includes slider changes
            'and changes made with rig's VFO knob
            DownlinkCorrection = DownlinkCorrection + CLng(RD# - _
                1000000# * Cdbl2(DownlinkFreq.text))
            'Calculate final RX freq, this includes final correction amount and
            'freq received via DDE
            DownlinkFreq.text = Str$((1000000# * Cdbl2(DownlinkDDEFreq.text) _
                + DownlinkCorrection) / 1000000#)
            'Check if downlink freq on rig will need update
            If RD# <> (Cdbl2(DownlinkFreq.text) * 1000000#) Then ud% = 1
            'if uplinkcorrection was changed:
            If uu% And UplinkHandle% <> 0 Then
                UplinkFreq.text = Str$((1000000# * Cdbl2(UplinkDDEFreq.text) _
                    + UplinkCorrection) / 1000000#)
            End If
        'if user is not changing dnlink freq:
        Else
            cd% = 0
            'if some change due to track is still to be updated to rig
            If uu% <> 0 Then
                Call UpdateUplink
                uu% = 0
            End If
            'user cannot change uplink freq in split (duplex) mode
            If UplinkSplit% = 0 Then
               'check if user changed uplink freq (slave) from rig
               a = ReadUplinkFreq
               'check read is good
               If a > 0 + 1000000# * UplinkLO# Then
                   RU# = a
                   If RU# <> 1000000# * Cdbl2(UplinkFreq.text) Then
                       UplinkCorrection = UplinkCorrection + CLng(RU# - _
                           1000000# * Cdbl2(UplinkFreq.text))
                       UplinkFreq.text = Str$((1000000# * Cdbl2(UplinkDDEFreq.text) _
                           + UplinkCorrection) / 1000000#)
                   End If
               End If
            End If
        End If
    End If
End If

'if uplink is master:
If uu% = 0 And UplinkHandle% <> 0 And SliderUplink.Value = 1 _
    And UplinkBidir% And UplinkSplit% = 0 Then
    'read master radio freq., if radio answered then update R# with
    'the answer.
    a = ReadUplinkFreq
    'check if read was successful
    If a > 0 + 1000000# * UplinkLO# Then
        RU# = a
        'if slider or buttons changed we need to update radio's freq
        'and then process as if user had changed it from radio
        a = Slider.Value - SliderCorrection
        If a <> 0 Or ButtonCorrection <> 0 Then
            RU# = RU# + a + ButtonCorrection
            b = Cdbl2(UplinkFreq.text)
            UplinkFreq.text = Str(RU# / 1000000#)
            Call UpdateUplink
            UplinkFreq.text = Str(b)
            SliderCorrection = Slider.Value
            ButtonCorrection = 0
        End If
        If Abs(Slider.Value) = 1000 Then
            Slider.Value = 0
            SliderCorrection = 0
        End If
        'if user is correcting uplink freq (master) from rig:
        If RU# <> 1000000# * Cdbl2(UplinkFreq.text) Then
            'Activate changing uplink flag
            cu% = 1
            
            ud% = 0
            'before changing uplinkcorrection check for variation to
            'be substracted/added from/to dnlinkcorrection:
            If DownlinkHandle% <> 0 And SliderDownlink.Value = 0 Then
                If TrackRev.Value = True Then
                    DownlinkCorrection = DownlinkCorrection - CLng( _
                        RU# - 1000000# * Cdbl2(UplinkFreq.text))
                    ud% = 1
                ElseIf TrackDir.Value = True Then
                    DownlinkCorrection = DownlinkCorrection + CLng( _
                        RU# - 1000000# * Cdbl2(UplinkFreq.text))
                    ud% = 1
                End If
            End If
            UplinkCorrection = UplinkCorrection + CLng(RU# - _
                1000000# * Cdbl2(UplinkFreq.text))
            UplinkFreq.text = Str$((1000000# * Cdbl2(UplinkDDEFreq.text) _
                + UplinkCorrection) / 1000000#)
            'Check if uplink freq on rig will need update
            If RU# <> (Cdbl2(UplinkFreq.text) * 1000000#) Then uu% = 1
            'if dnlinkcorrection was changed:
            If ud% And DownlinkHandle% <> 0 Then
                DownlinkFreq.text = Str$((1000000# * Cdbl2(DownlinkDDEFreq.text) _
                    + DownlinkCorrection) / 1000000#)
            End If
        'if user is not changing uplink freq:
        Else
            cu% = 0
            'if some change due to track is still to be updated to rig
            If ud% <> 0 Then
                Call UpdateDownlink
                'dnlinkcorrection change flag:
                ud% = 0
            End If
            'check if user changed dnlink freq (slave) from rig
            a = ReadDownlinkFreq
            'check if read was successful:
            If a > 0 + 1000000# * DownlinkLO# Then
                RD# = a
                If RD# <> 1000000# * Cdbl2(DownlinkFreq.text) Then
                    DownlinkCorrection = DownlinkCorrection + CLng(RD# - _
                        1000000# * Cdbl2(DownlinkFreq.text))
                    DownlinkFreq.text = Str$((1000000# * Cdbl2(DownlinkDDEFreq.text) _
                        + DownlinkCorrection) / 1000000#)
                End If
            End If
        End If
    End If
End If

'perform doppler corrections if any...
If DownlinkHandle% <> 0 Then
    a = Str((DownlinkCorrection + 1000000# * Cdbl2(DownlinkDDEFreq.text)) / 1000000#)
    If (cu% = 0) And (ud% Or (Cdbl2(a) <> 0) And (Cdbl2(a) <> Cdbl2(DownlinkFreq.text))) Then
        'keep old freq value
        b = DownlinkFreq.text
        DownlinkFreq.text = a
        'if update is not successfull then return to old value (the one actually on the
        'radio)
        If UpdateDownlink <> 1 Then
            DownlinkFreq.text = b
        Else
            ud% = 0
        End If
    End If
End If

If UplinkHandle% <> 0 Then
    a = Str((UplinkCorrection + 1000000# * Cdbl2(UplinkDDEFreq.text)) / 1000000#)
    If (cd% = 0) And (uu% Or (Cdbl2(a) <> 0) And (Cdbl2(UplinkFreq.text) <> Cdbl2(a))) Then
        'keep old value
        b = UplinkFreq.text
        UplinkFreq.text = a
        'if apdate not possible return field value to the actual freq on the rig.
        If UpdateUplink <> 1 Then
            UplinkFreq.text = b
        Else
            uu% = 0
        End If
    End If
End If



'wait till its our time again...
RadioControlLoopTimer.Enabled = True
End Sub

Private Sub Slider_Change()
SliderTimer.Enabled = False
SliderTimer.Enabled = True
'If Abs(Slider.Value) = 1000 Then
'    Slider.Value = 0
'    SliderCorrection = 0
'End If
End Sub

Private Sub SliderDownlink_Click()
If SliderDownlink.Value = 1 Then SliderUplink.Value = 0
If SliderDownlink.Value = 1 And SliderUplink.Value = 1 Then
    TrackRev.Value = False
    TrackDir.Value = False
End If
If SliderDownlink.Value = 0 And SliderUplink.Value = 0 Then
    Slider.Enabled = False
Else
    Slider.Enabled = True
End If
End Sub

Private Sub SliderTimer_Timer()
Slider.Value = 0
SliderCorrection = 0
SliderTimer.Enabled = False
End Sub

Private Sub SliderUplink_Click()
If SliderUplink.Value = 1 Then SliderDownlink.Value = 0
If SliderDownlink.Value = 1 And SliderUplink.Value = 1 Then
    TrackRev.Value = False
    TrackDir.Value = False
End If
If SliderDownlink.Value = 0 And SliderUplink.Value = 0 Then
    Slider.Enabled = False
Else
    Slider.Enabled = True
End If
End Sub

Private Sub TrackDir_Click()
If SliderUplink.Value = 1 And SliderDownlink.Value = 1 Then
    SliderUplink.Value = 0
End If
End Sub

Private Sub TrackDir_DblClick()
If TrackDir.Value = True Then
    TrackDir.Value = False
End If

End Sub

Private Sub TrackRev_Click()
If SliderUplink.Value = 1 And SliderDownlink.Value = 1 Then
    SliderUplink.Value = 0
End If
End Sub

Private Sub TrackRev_DblClick()
If TrackRev.Value = True Then
    TrackRev.Value = False
End If

End Sub

Private Sub UplinkIndex_Change()
'****** put Icom rigs into memory mode*****
'       or put off CAT on Yaesu rigs
If UplinkHandle% <> 0 Then
'this is not completely accurate, but try not to park a radio that
'is being used by other band...
    If DownlinkHandle% = UplinkHandle% And DownlinkModel$ = UplinkModel$ Then
    Else
        Call ParkUplinkRadio
    End If
End If
'put accesory ports to zero
If UplinkAccPort% > 0 Then
    Call OutPort(UplinkAccPort%, 0)
End If
'Terminate uplink radio control port
'only if same handle is not in use by Downlink or Rotor
If (DownlinkHandle% <> UplinkHandle%) And (RotorHandle% <> UplinkHandle%) Then
    ClosePort (UplinkHandle%)
End If
UplinkHandle% = 0

'open the selected radio's ports and initialize public
'variables with comms settings
If Cdbl2(UplinkIndex.text) <> 0 Then
    'if there is any error opening the port -> release unlink handle
    If OpenUplinkPort <> 0 Then
        UplinkHandle% = 0
        Call frmMessage.ShowMessage("Error opening" + Chr$(13) _
            + "Downlink Radio's Port", 10)
        Exit Sub
    End If
    If UplinkHandle% <> 0 Then
        Call ActivateUplinkRadio
    End If
    'if both radios use the same acc.port we OR the values,
    'otherwise they are left independent
    If DownlinkAccPort% = UplinkAccPort% Then
        If DownlinkAccPort% <> 0 Then
            DownlinkAccPortValue% = DownlinkAccPortValue% Or UplinkAccPortValue%
            Call OutPort(DownlinkAccPort%, CInt(DownlinkAccPortValue%))
        End If
    Else
        If UplinkAccPort% <> 0 Then
            Call OutPort(UplinkAccPort%, CInt(UplinkAccPortValue%))
        End If
    End If
End If
End Sub
Private Sub UplinkIndex_Click()
    Call UplinkIndex_Change
End Sub
Private Sub UpdateRotorButton_Click()
'if rotor port is a COM, it's different from the
'radio port and it can be opened -> we open it
If frmRotor.RotorType.text <> "None" And _
    Left$(frmRotor.RotorPort.text, 3) = "COM" And _
    RotorHandle% = 0 Then
    Call OpenRotorPort
    OpenedRotorPortFlag = True
End If
Az = CInt(Cdbl2(Azimuth.text))
El = CInt(Cdbl2(Elevation.text))
Call UpdateRotor(Az, El)
If OpenedRotorPortFlag Then
    'TrakBox needs to be restored to Main Menu after angles
    'were manually sent.
    'If frmRotor.RotorType.text = "TrakBox" Then
    '    Call TBRotorSetTerminal(RotorHandle%)
    'End If
    Call ClosePort(RotorHandle%)
    RotorHandle% = 0
End If
End Sub
Private Sub UpdateRadioButton_Click()
a$ = DownlinkIndex.text
DownlinkIndex.text = LTrim(RTrim(a$))
a$ = UplinkIndex.text
UplinkIndex.text = LTrim(RTrim(a$))
Call UpdateDownlink
Call UpdateUplink
End Sub

Private Sub ddelink_Click()
    'launch the DDE setup window
    frmDdelink.Show
End Sub

'Devuelve el primer numero encontrado en un string
Function Firstnum(s As String)
    For f = 1 To Len(s)
        For i = (Len(s) - f + 1) To 1 Step -1
            If IsNumeric(Left(s, i)) Then
                Firstnum = Cdbl2(Left(s, i))
                Exit Function
            End If
        Next
    Next
End Function

'Devuelve la primer palabra encontrada en un string
Function Firststring(s As String)
    For p% = 1 To Len(s)
        If Mid(s, p%, 1) <> " " Then
            Exit For
        End If
    Next
    If p% < Len(s) Then
        For f% = p% + 1 To Len(s)
            If IsControl(Asc(Mid(s, f%, 1))) Or Asc(Mid(s, f%, 1)) = 32 Then
                f% = f% - 1
                Exit For
            End If
        Next
    Else
        f% = Len(s)
    End If
    Firststring = Mid(s, p%, f% - p% + 1)
End Function
'This function returns the Nth word from an input string
'words are assumed separated by one or more spaces.
Function PickWord(s As String, n%)
LenString% = Len(s)
f% = 1
While f% <= LenString%
    While Mid$(s, f%, 1) = " "
        f% = f% + 1
        If f% > LenString% Then Exit Function
    Wend
    n% = n% - 1
    If n% = 0 Then
        PickWord = Firststring(Mid$(s, f%))
        Exit Function
    End If
    While Mid$(s, f%, 1) <> " " And f% <= LenString%
        f% = f% + 1
    Wend
Wend
End Function
'deveulve true si es un caracter de control (8,9,10 o 13)
Function IsControl(b As Byte)
    If b = 8 Or b = 9 Or b = 10 Or b = 13 Then
        IsControl = True
    Else
        IsControl = False
    End If
End Function
Function Array2String$(a)
s$ = ""
For f% = 0 To LenB(a) - 1
    s$ = s$ + Chr$(a(f%))
Next
Array2String$ = s$
End Function


Private Sub Form_Load()
    On Error Resume Next
    
    ' Check if logging is active
    If frmRadio.CheckLog.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Program Started" + " " + Str(Time)
        
        Print #f%, "Configuration Section: Config:"
        reg = GetAllSettings("WiSP_DDE_Client", "Config")
        For j% = LBound(reg) To UBound(reg)
            Print #f%, reg(j%, 0) + ": " + reg(j%, 1)
        Next
        
        i% = 1
        While GetSetting("WiSP_DDE_Client", "Rig" + LTrim$(Str$(i%)), "Radio_model", "-") <> "-"
            Print #f%, "Configuration Section: " + "Rig" + LTrim$(Str$(i%))
            reg = GetAllSettings("WiSP_DDE_Client", "Rig" + LTrim$(Str$(i%)))
            For j% = LBound(reg) To UBound(reg)
                Print #f%, reg(j%, 0) + ": " + reg(j%, 1)
            Next
            i% = i% + 1
        Wend
            
        i% = 1
        While GetSetting("WiSP_DDE_Client", "Sat" + LTrim$(Str$(i%)), "SatName", "") <> ""
            Print #f%, "Configuration Section: " + "Sat" + LTrim$(Str$(i%))
            reg = GetAllSettings("WiSP_DDE_Client", "Sat" + LTrim$(Str$(i%)))
            For j% = LBound(reg) To UBound(reg)
                Print #f%, reg(j%, 0) + ": " + reg(j%, 1)
            Next
            i% = i% + 1
        Wend
            
        Close f%
    End If
    
    'Single pass: if the command line argument "S" is specified, then
    'the program will exit after the first pass ends.
    'the flag is set to 2, when a satellite appears and the flag is 2
    'it will be set to 1.
    'if NO SATELLITE appears and flag is 1, program will exit
    SinglePass% = 0
    If InStr(LCase$(Command()), "s") Then
        ' Check if logging is active
        If frmRadio.CheckLog.Value Then
            f% = FreeFile
            Open "Radio_Log.txt" For Append As f%
            Print #f%, "Single-Pass mode detected" + " " + Str(Time)
            Close f%
        End If
        
        SinglePass% = 2
        frmMain.Caption = "WiSPDDE (SINGLE PASS)"
    End If
    
    'Reset configuration: if the command line argument "R" is
    'present all the configuration will be erased.
    If InStr(LCase$(Command()), "r") Then
        ' Check if logging is active
        If frmRadio.CheckLog.Value Then
            f% = FreeFile
            Open "Radio_Log.txt" For Append As f%
            Print #f%, "Removing All Registry Information" + ", " + Str(Time)
            Close f%
        End If
        
        frmMain.Remove_Registry
    End If
    
    ' Defaults for first time running:
    DDE_source_default = "GSC"
    DDE_Topic_Default = "Tracking"
    DDE_Item_Default = "Tracking"
    DDE_Period_Default = "3"
    
    RotorUpdateComplete = True
    
    Rotor_com_default = "None"
    Rotor_baud_default = 4800
    Rotor_mode_default = "None"
    Rotor_step_default = "5"
    'NOTE:
    'Rotor flipping and stop position defaults are
    'false (unchecked), so the soft will not configure
    'itself for auto flipping detection and the default
    'stop pos. will be North.
    
    frmMain.Left = Cdbl2(GetSetting("WiSP_DDE_Client", "Config", "WindowPositionLeft", "0"))
    frmMain.Top = Cdbl2(GetSetting("WiSP_DDE_Client", "Config", "WindowPositionTop", "0"))
    frmMain.Height = Cdbl2(GetSetting("WiSP_DDE_Client", "Config", "WindowHeight", "5595"))
    frmMain.Width = Cdbl2(GetSetting("WiSP_DDE_Client", "Config", "WindowWidth", "3525"))
    If (frmMain.Left > Screen.Width) Or (frmMain.Left < 0) Then _
        frmMain.Left = 0
    If (frmMain.Top > Screen.Height) Or (frmMain.Top < 0) Then _
        frmMain.Top = 0
    frmMain.WindowState = Normal
        
    
    '*****Configuration retrieveing from registry:*****
    '**DDE Link:**
    'DDE source setting is retreived and if not yet
    'initialized in the Windows registry (first time run)
    'the default is employed
    'retrieve previous WiSP/Station selection
    'if first run, WiSP will be selected
    frmDdelink.DDEFormat.text = GetSetting("WiSP_DDE_Client", "Config", "Dde_format", WiSP)
    If frmDdelink.DDEFormat.text = "Nova" Then
        frmDdelink.Command3.Enabled = True
    Else
        frmDdelink.Command3.Enabled = False
    End If
    frmDdelink.SourceApplication.text = GetSetting("WiSP_DDE_Client", "Config", "Dde_source", DDE_source_default)
    'Same for Topic
    frmDdelink.Topic.text = GetSetting("WiSP_DDE_Client", "Config", "Dde_topic", DDE_Topic_Default)
    'And Item
    frmDdelink.Item.text = GetSetting("WiSP_DDE_Client", "Config", "Dde_item", DDE_Item_Default)
    'And the DDE-server purge interval
    frmDdelink.Interval.text = GetSetting("WiSP_DDE_Client", "Config", "Dde_period", DDE_Period_Default)
    
    'The following lines setup WiSPDDE as client
    'they are not applicable in the case of DDE
    'server
    If frmDdelink.DDEFormat.text = "WiSP" Or _
        frmDdelink.DDEFormat.text = "SatPC32" Or _
        frmDdelink.DDEFormat.text = "Nova" Or _
        frmDdelink.DDEFormat.text = "Station" Or _
        frmDdelink.DDEFormat.text = "Satscape" Or _
        frmDdelink.DDEFormat.text = "Winorbit" Or _
        frmDdelink.DDEFormat.text = "Orbitron" Or _
        frmDdelink.DDEFormat.text = "WXtrack" Or _
        frmDdelink.DDEFormat.text = "ARS" Then
    
        ' Check if logging is active
        If frmDdelink.CheckLog.Value Then
            f% = FreeFile
            Open "DDELink_Log.txt" For Append As f%
            Print #f%, "Configuring Program as DDE Client" + " " + Str(Time); ""
            Close f%
        End If
            
        'Set the proper value for the Timer (convert to
        'milliseconds)
        If Cdbl2(frmDdelink.Interval.text) <> 0 Then
            DDEPollTimer.Interval = Cdbl2(frmDdelink.Interval.text) * 1000
            DDEPollTimer.Enabled = True
        Else
            ' Check if logging is active
            If frmDdelink.CheckLog.Value Then
                f% = FreeFile
                Open "DDELink_Log.txt" For Append As f%
                Print #f%, "Disabling DDEPoll Timer" + " " + Str(Time); ""
                Close f%
            End If
            
            DDEPollTimer.Enabled = False
        End If
        
        'Source and Topic goes together separated by '|':
        DDELabel.LinkTopic = frmDdelink.SourceApplication.text + "|" + frmDdelink.Topic.text
        DDELabel.LinkItem = frmDdelink.Item
        'Try to open DDE link as client and with the LinkMode according
        'to DDEPollTimer value (Manual or Notify mode):
        Call DDE_Test
        'if DDE update is set to automatic (timer1 disabled
        'because refresh rate set to 0 seconds) we update
        'manually the first time as it may take long for the
        'satellite entry to be updated by server DDE application
        If DDEPollTimer.Enabled = False Then
            Call DDELabel_LinkNotify
        End If
        
        'if DDE is from NFW we have to send "TUNE ON" message:
        If frmDdelink.DDEFormat.text = "Nova" Then
            ' Check if logging is active
            If frmDdelink.CheckLog.Value Then
                f% = FreeFile
                Open "DDELink_Log.txt" For Append As f%
                Print #f%, "Sending TUNE ON DDE Command to Nova" + " " + Str(Time); ""
                Close f%
            End If
            
            DDELabel.Caption = "TUNE ON"
            DDELabel.LinkPoke
        End If
        
        Satellite.text = ""

'   SatPC32 has been modified to work as DDE Server (like WiSP)
'   So this code is no longer needed...
    'if DDE will be from SatPC32, configure WiSPDDE as DDE Server:
'    ElseIf frmDdelink.DDEFormat.text = "SatPC32" Then
'
'        DDEPollTimer.Enabled = False
'        DDELabel.LinkMode = 0
'        Satellite.text = "No DDE Link"
'        'in this case Client is SatPC, server is WiSPDDE
'        If frmDdelink.DDEFormat.text = "SatPC32" Then
'            'Set WiSPDDE as source application:
'            frmRotddesvr.LinkTopic = "RotServConv"
'            frmRotddesvr.LinkMode = 1
'            frmCatddesvr.LinkTopic = "CatServConv"
'            frmCatddesvr.LinkMode = 1
'        End If
'
    'in no DDE, disable all DDE links...
    Else
        DDEPollTimer.Enabled = False
        DDELabel.LinkMode = 0
        Satellite.text = "No DDE Link"
    End If
    
    
    '**Radio interface:**
    'Add as many entries to the "radio Index" combo as we find
    'in the registry...
    'do it twice, once for uplink and once for downlink selections
    UplinkIndex.AddItem "None"
    i% = 1
    Do
        UplinkIndex.AddItem LTrim$(Str$(i%))
        i% = i% + 1
    Loop Until GetSetting("WiSP_DDE_Client", "Rig" + LTrim$(Str$(i%)), "Radio_model", "-") = "-"
        
    DownlinkIndex.AddItem "None"
    i% = 1
    Do
        DownlinkIndex.AddItem LTrim$(Str$(i%))
        i% = i% + 1
    Loop Until GetSetting("WiSP_DDE_Client", "Rig" + LTrim$(Str$(i%)), "Radio_model", "-") = "-"
    
    '**Rotor interface:**
    'get rotor config from registry:
    Call frmRotor.Form_Load
    
    RadioControlCountdown = 3
    
    '*****SETUP TIMERS*****:
    'Set Radio time-out timer interval to .2 sec.
    RadioTimerCount = 0
    RadioTimer.Interval = 200
    RadioTimer.Enabled = False
    'RadioTimer.Enabled = True

    'allow excess time for the radio to process commands in case
    'it is busy (attending VFO knob for example)
    RadioControlDelay(1) = 500
    RadioControlDelay(2) = 500
    RadioControlDelay(3) = 500
    'enable radio loop:
    RadioControlLoopTimer.Interval = 1000
    RadioControlLoopTimer.Enabled = True
    'Slider Timer to 1 sec
    SliderTimer.Interval = 2000
    SliderTimer.Enabled = False


    'retrieve transparent tuning setup from registry:
    frmMain.RotorAuto.Value = GetSetting("WiSP_DDE_Client", "Config", "Rotor_Auto", 1)
    frmMain.SliderDownlink.Value = GetSetting("WiSP_DDE_Client", "Config", "Slider_Downlink", 1)
    frmMain.SliderUplink.Value = GetSetting("WiSP_DDE_Client", "Config", "Slider_Uplink", 0)
    frmMain.TrackRev.Value = GetSetting("WiSP_DDE_Client", "Config", "Track_Rev", True)
    frmMain.TrackDir.Value = GetSetting("WiSP_DDE_Client", "Config", "Track_Dir", False)
    
    'setup controls for normal operation:
    DownlinkRSSI.Min = 0
    DownlinkRSSI.Max = 255
    DownlinkRSSI.Enabled = False
    DownlinkRSSI.Value = 0
    
    'load decimal separator character
    frmDdelink.Decimal.text = GetSetting("WiSP_DDE_Client", "Config", "Decimal_Separator", ".")
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
' Check if logging is active
If frmRadio.CheckLog.Value Then
    f% = FreeFile
    Open "Radio_Log.txt" For Append As f%
    Print #f%, "Closing Program" + ", " + Str(Time)
    Close f%
End If


DDELabel.LinkMode = 0
DDELabel.LinkTopic = ""
frmRotddesvr.LinkMode = 0
frmRotddesvr.LinkTopic = ""
frmCatddesvr.LinkMode = 0
frmCatddesvr.LinkTopic = ""

'save transparent tuning setup to registry:
SaveSetting "WiSP_DDE_Client", "Config", "Rotor_Auto", frmMain.RotorAuto.Value
SaveSetting "WiSP_DDE_Client", "Config", "Slider_Downlink", frmMain.SliderDownlink.Value
SaveSetting "WiSP_DDE_Client", "Config", "Slider_Uplink", frmMain.SliderUplink.Value
SaveSetting "WiSP_DDE_Client", "Config", "Track_Rev", frmMain.TrackRev.Value
SaveSetting "WiSP_DDE_Client", "Config", "Track_Dir", frmMain.TrackDir.Value


'save window position:
If (frmMain.Left < Screen.Width) And (frmMain.Left > 0) _
    And (frmMain.Top < Screen.Height) And (frmMain.Top > 0) Then
    SaveSetting "WiSP_DDE_Client", "Config", "WindowPositionLeft", frmMain.Left
    SaveSetting "WiSP_DDE_Client", "Config", "WindowPositionTop", frmMain.Top
    SaveSetting "WiSP_DDE_Client", "Config", "WindowHeight", frmMain.Height
    SaveSetting "WiSP_DDE_Client", "Config", "WindowWidth", frmMain.Width
End If
'put accesory ports to zero
If DownlinkAccPort% > 0 Then
    Call OutPort(DownlinkAccPort%, 0)
End If
If UplinkAccPort% > 0 Then
    Call OutPort(UplinkAccPort%, 0)
End If

'bye-bye procedure...
'TrackBox back to Main menu:
If frmRotor.RotorType.text = "TrakBox" And RotorHandle% Then
        'wait up to 65 secs for answer to previous command
        'from TrakBox:
        'at 10 sec. show waiting message
        i% = 0
        Do
            b$ = TBRotorReadFrame(RotorHandle%)
            i% = i% + 1
            If i = 10 Then
                frmMessage.Label1.Caption = "Waiting for answer from TrakBox" + _
                    Chr$(13) + Chr$(13) + "Stop Waiting?"
                frmMessage.Show
            End If
        Loop Until Firstnum(b$) <> 0 Or i% > 65 Or _
            frmMessage.Tag = "OK"
        frmMessage.Hide
    Call WriteToPort("Q", RotorHandle%)
End If

If DownlinkHandle% <> 0 Then
    Call ParkDownlinkRadio
End If
If UplinkHandle% <> 0 Then
    Call ParkUplinkRadio
End If
'Terminate both radio control ports
ClosePort (DownlinkHandle%)
ClosePort (UplinkHandle%)
DownlinkHandle% = 0
UplinkHandle% = 0
DownlinkIndex.text = ""
UplinkIndex.text = ""
'And Rotor port too
ClosePort (RotorHandle%)
RotorHandle% = 0
DDEPollTimer.Enabled = False
End
End Sub

Private Sub Satellite_Change()
'Every time the sat. changes:
RadioControlCountdown = 10
'If sat changes *and* auto-flip is enabled then
'we store the first value of Azimuth and initialize
'the down-counter to compare this value later
If frmRotor.RotorAutoFlip.Value Then
    RotorControlCountdown = 3
    RotorAz = Cdbl2(Azimuth.text)
'if auto-flip disabled we will never flip
Else
    RotorControlCountdown = -1
    RotorFlip = False
End If

'********** PARK ANTENNAS ************************
'also use the checking to process the single-pass feature!
'also blank DDE link fields...
If InStr(Satellite.text, "NO SATELLITE") Or _
    InStr(Satellite.text, "No DDE Link") Then
    DownlinkDDEFreq.text = ""
    UplinkDDEFreq.text = ""
    DownlinkMode.text = ""
    UplinkMode.text = ""
    If frmRotor.RotorAzPark.text <> "" Or _
        frmRotor.RotorElPark.text <> "" Then
        Azimuth.text = frmRotor.RotorAzPark.text
        Elevation.text = frmRotor.RotorElPark.text
        Call UpdateRotorButton_Click
    End If
    If SinglePass% = 1 Then Call Close_Click
Else
    If SinglePass% = 2 Then SinglePass% = 1
End If

'Close rotor port only if same handle is not being used by
'Uplink or Downlink...
If (RotorHandle% <> DownlinkHandle%) And (RotorHandle% <> UplinkHandle%) Then
    'TrakBox needs to be restored to Main Menu:
    If frmRotor.RotorType.text = "TrakBox" And RotorHandle% Then
        Call TBRotorSetTerminal(RotorHandle%)
    End If
    ClosePort (RotorHandle%)
End If
RotorHandle% = 0

'park radios also...
DownlinkIndex.text = ""
UplinkIndex.text = ""

'*********Satellite selection (only for Nova for Windows):********
If frmDdelink.DDEFormat.text = "Nova" Then
    'Test all available sats for name
    i% = 1
    SatSelectedFlag% = False
    Do
        'select the sat secuentially
        'we will look at the registry entry of each sat until an apropriate one is
        'found...
        'check sat name is equal to currently tracked sat
        'and satellite is enabled
        If UCase(LTrim(GetSetting("WiSP_DDE_Client", "Sat" + LTrim$(Cdbl2(i%)), "SatName", ""))) = UCase(LTrim(Satellite.text)) _
                And GetSetting("WiSP_DDE_Client", "Sat" + LTrim$(Cdbl2(i%)), "SatSatEnabled", 0) Then
            frmSats.SatIndex.text = LTrim(Str(i%))
            SatSelectedFlag% = True
            'we need to enter aproximate values (no doppler) in downlink &
            'uplink freqs in order to make auto-rig-selection
            DownlinkDDEFreq.text = Str(Cdbl2(frmSats.SatDownlinkFreq.text))
            UplinkDDEFreq.text = Str(Cdbl2(frmSats.SatUplinkFreq.text))
            DownlinkMode.text = frmSats.SatDownlinkMode.text
            UplinkMode.text = frmSats.SatUplinkMode.text
            TrackDir.Value = frmSats.SatDirTrack.Value
            TrackRev.Value = frmSats.SatRevTrack.Value
        End If
        i% = i% + 1
    Loop Until GetSetting("WiSP_DDE_Client", "Sat" + LTrim$(Str$(i%)), "SatName", "-") = "-" Or RadioSelectedFlag%
    If Not SatSelectedFlag% Then
        frmSats.SatIndex.text = ""
    End If
End If


'*********Radio auto-selection:********
'Test all available rigs for auto-selection condition satisfaction
'first look for a suitable downlink rig
i% = 1
RadioSelectedFlag% = False
Do
    'select the rig secuentially
    'we will look at the registry entry of each radio until an apropriate one is
    'found...
    'check to satisfy *every* condition imposed to this rig
    If Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + LTrim$(Cdbl2(i%)), "Radio_enable", 0)) Then
        If Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + LTrim$(Cdbl2(i%)), "Radio_AutoSelDownlink", 0)) Then
            If (InStr(GetSetting("WiSP_DDE_Client", "Rig" + LTrim$(Cdbl2(i%)), "Radio_AutoSelSats", ""), Satellite.text) _
                And Satellite.text <> "") _
                Or (InStr(LCase(GetSetting("WiSP_DDE_Client", "Rig" + LTrim$(Cdbl2(i%)), "Radio_AutoSelSats", "")), "all") _
                And (InStr(Satellite.text, "NO SATELLITE") = 0) And Satellite.text <> "No DDE Link" _
                And Satellite.text <> "") Then
                If InStr(GetSetting("WiSP_DDE_Client", "Rig" + LTrim$(Cdbl2(i%)), "Radio_AutoSelModes", ""), DownlinkMode.text) _
                    Or InStr(LCase(GetSetting("WiSP_DDE_Client", "Rig" + LTrim$(Cdbl2(i%)), "Radio_AutoSelModes", "")), "all") Then
                    freqs$ = GetSetting("WiSP_DDE_Client", "Rig" + LTrim$(Cdbl2(i%)), "Radio_AutoSelFreqs", " ")
                    If InStr(LCase(freqs$), "all") Then
                        DownlinkIndex.text = LTrim(Str(i%))
                        RadioSelectedFlag% = True
                    Else
                        For f% = 1 To Len(freqs$)
                            If IsNumeric(Mid$(freqs$, f%, 1)) Then
                                lower$ = Firstnum(Mid$(freqs$, f%))
                                f% = f% + Len(lower$)
                                If Mid$(freqs$, f%, 1) = "-" Then
                                    f% = f% + 1
                                    upper$ = Firstnum(Mid$(freqs$, f%))
                                    f% = f% + Len(upper$)
                                    If Cdbl2(lower$) <= Cdbl2(DownlinkDDEFreq.text) And Cdbl2(DownlinkDDEFreq.text) <= Cdbl2(upper$) Then
                                        DownlinkIndex.text = LTrim(Str(i%))
                                        RadioSelectedFlag% = True
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            End If
        End If
    End If
    i% = i% + 1
Loop Until GetSetting("WiSP_DDE_Client", "Rig" + LTrim$(Str$(i%)), "Radio_model", "-") = "-" Or RadioSelectedFlag%
If Not RadioSelectedFlag% Then
    DownlinkIndex.text = ""
End If
'Test all available rigs for auto-selection condition satisfaction
'now look for a suitable uplink rig
i% = 1
RadioSelectedFlag% = False
Do
    'select the rig secuentially
    'we will look at the registry entro of each radio until an apropriate one is
    'found...
    'check to satisfy *every* condition imposed to this rig
    If Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + LTrim$(Cdbl2(i%)), "Radio_enable", 0)) Then
        If Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + LTrim$(Cdbl2(i%)), "Radio_AutoSelUplink", 0)) Then
            If (InStr(GetSetting("WiSP_DDE_Client", "Rig" + LTrim$(Cdbl2(i%)), "Radio_AutoSelSats", ""), Satellite.text) _
                And Satellite.text <> "") _
                Or (InStr(LCase(GetSetting("WiSP_DDE_Client", "Rig" + LTrim$(Cdbl2(i%)), "Radio_AutoSelSats", "")), "all") _
                And (InStr(Satellite.text, "NO SATELLITE") = 0) And Satellite.text <> "No DDE Link" _
                And Satellite.text <> "") Then
                If InStr(GetSetting("WiSP_DDE_Client", "Rig" + LTrim$(Cdbl2(i%)), "Radio_AutoSelModes", ""), UplinkMode.text) _
                    Or InStr(LCase(GetSetting("WiSP_DDE_Client", "Rig" + LTrim$(Cdbl2(i%)), "Radio_AutoSelModes", "")), "all") Then
                    freqs$ = GetSetting("WiSP_DDE_Client", "Rig" + LTrim$(Cdbl2(i%)), "Radio_AutoSelFreqs", " ")
                    If InStr(LCase(freqs$), "all") Then
                        UplinkIndex.text = LTrim(Str(i%))
                        RadioSelectedFlag% = True
                    Else
                        For f% = 1 To Len(freqs$)
                            If IsNumeric(Mid$(freqs$, f%, 1)) Then
                                lower$ = Firstnum(Mid$(freqs$, f%))
                                f% = f% + Len(lower$)
                                If Mid$(freqs$, f%, 1) = "-" Then
                                    f% = f% + 1
                                    upper$ = Firstnum(Mid$(freqs$, f%))
                                    f% = f% + Len(upper$)
                                    If Cdbl2(lower$) <= Cdbl2(UplinkDDEFreq.text) And Cdbl2(UplinkDDEFreq.text) <= Cdbl2(upper$) Then
                                        UplinkIndex.text = LTrim(Str(i%))
                                        RadioSelectedFlag% = True
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            End If
        End If
    End If
    i% = i% + 1
Loop Until GetSetting("WiSP_DDE_Client", "Rig" + LTrim$(Str$(i%)), "Radio_model", "-") = "-" Or RadioSelectedFlag%
If Not RadioSelectedFlag% Then
    UplinkIndex.text = ""
End If

'wait until no events are waiting to be handled. This may be necesary
'to ensure radio ports are finished to open before we attempt to
'open a rotor port in the same COM device...
'DoEvents

'if rotor port is a COM, it's different from the
'radio port and it can be opened -> we open it as long as
'there is valid info from DDE source
If frmRotor.RotorType.text <> "None" And _
    Left$(frmRotor.RotorPort.text, 3) = "COM" Then
    'but not if there is no satellite!
    If InStr(Satellite.text, "NO SATELLITE") Or _
        InStr(Satellite.text, "No DDE Link") Then
    Else
        Call OpenRotorPort
    End If
End If
End Sub
Private Sub radio_Click()
    'launch radio setup window
    frmRadio.Show
End Sub

Private Sub rotor_Click()
    'launch rotor setup window
    frmRotor.Show
End Sub

'Devuelve 1 si el enlace DDE esta activo
Function DDE_Test()
'in order to prevent errors when starting the DDE link,
'we provide this simple routine to check before.
    On Error GoTo Error
    DDE_Test = False
    If DDEPollTimer.Enabled = True Then
        DDELabel.LinkMode = vbLinkManual
    Else
        DDELabel.LinkMode = vbLinkNotify
    End If
    DDE_Test = True
Error:
End Function
'When DDE query-interval is non-zero this routine is executed
'periodically. When DDE query-interval is zero, this routine is
'executed on each LinkNotify event of the DDE.
'Also this is only used when WiSPDDE acts as a DDE-Client
'(when used with WiSP, Nova, Satscape etc.).
'In case WiSPDDE acts as DDE-Server,
'this routine is not used. A completely different method is used
'based on commands received from the Client (see frmCatddesvr &
'frmRotddesvr forms).
'Formerly SatPC32 operated as DDEClient, but this has changed and
'WiSPDDE doesn't need to act as server to get data from SatPC32.
Private Sub DDEPollTimer_Timer()
    Dim DopplerMHz As Double

'    On Error GoTo NoDDE
    'if DDE link is OK
    If DDE_Test Then
        'we purge the DDE server
        DDELabel.LinkRequest
        'if DDE message is no satellite
        'we do not process any further
        If InStr(DDELabel.Caption, "NO SATELLITE") <> 0 Then
            Satellite.ForeColor = &H808080
            Satellite.text = DDELabel.Caption
        Else

'*************************************************************
'***Timed sat DDE info retrieving from various DDE sources:***
'As a general rule downlink & uplink freqs and modes
'are updated first and then the tracked satellite's name.
'This is to ensure freqs & modes are correct when rig
'auto-selection is performed (every time satellite changes)
'*************************************************************
            If frmDdelink.DDEFormat.text = "WiSP" Or _
                frmDdelink.DDEFormat.text = "SatPC32" Or _
                frmDdelink.DDEFormat.text = "Satscape" Or _
                frmDdelink.DDEFormat.text = "Orbitron" Or _
                frmDdelink.DDEFormat.text = "WXtrack" Then
            '********DDE in WiSP, Satscape or WXtrack mode*******
                'check if Radio DDE is enabled before
                'updating radio fields:
                'retrieve uplink freq.
                a = InStr(DDELabel.Caption, "UP")
                If a Then
                    UplinkDDEFreq.text = Str$(0.000001 * Fix(Cdbl2(Firststring(Mid(DDELabel.Caption, a + 2)))))
                End If
                'retrieve downlink freq.
                a = InStr(DDELabel.Caption, "DN")
                If a Then
                    DownlinkDDEFreq.text = Str$(0.000001 * Fix(Cdbl2(Firststring(Mid(DDELabel.Caption, a + 2)))))
                End If
                'retrieve uplink mode
                a = InStr(DDELabel.Caption, "UM")
                If a Then
                    UplinkMode.text = Firststring(Mid(DDELabel.Caption, a + 2))
                End If
                'retrieve downlink mode
                a = InStr(DDELabel.Caption, "DM")
                If a Then
                    DownlinkMode.text = Firststring(Mid(DDELabel.Caption, a + 2))
                End If
                'check if Rotor DDE is enabled before
                'updating Rotor fields:
                'retrieve Azimuth
                a = InStr(DDELabel.Caption, "AZ")
                If a And RotorAuto.Value = 1 Then
                    Azimuth.text = Firststring(Mid(DDELabel.Caption, a + 2))
                End If
                'Elevation
                a = InStr(DDELabel.Caption, "EL")
                If a And RotorAuto.Value = 1 Then
                    Elevation.text = Firststring(Mid(DDELabel.Caption, a + 2))
                End If
                'and the sat's name
                a = InStr(DDELabel.Caption, "SN")
                If a Then
                    Satellite.ForeColor = &HFF0000
                    Satellite.text = Firststring(Mid(DDELabel.Caption, a + 2))
                End If
            End If
            
            '*****DDE in Station Program mode******
            '(no radio nor satellite info)
            If frmDdelink.DDEFormat.text = "Station" Then
                'check if Rotor DDE is enabled:
                a = InStr(DDELabel.Caption, "|")
                If a And RotorAuto.Value Then
                    Azimuth.text = Str(Firstnum(DDELabel.Caption))
                    Elevation.text = Str(Firstnum(Mid$(DDELabel.Caption, a + 1)))
                End If
                Satellite.text = "Station"
            End If
            
            '********DDE in WinOrbit mode************
            If frmDdelink.DDEFormat.text = "Winorbit" Then
                'Probamos por partes, ya que el WinOrbit lo da separado
                'updating radio fields:
                'winorbit gives a doppler sample:
                'CDbl function needs to be used instead of Val as it takes
                'care of international variation of decimal separator
                'and WinOrbit takes care of this too.
                DDELabel.LinkItem = "FreqMHz"
                DDELabel.LinkRequest
                If DDELabel.Caption <> "" Then
                    a = Cdbl2(DDELabel.Caption) * 1000000#
                End If
                DDELabel.LinkItem = "DopplerHertz"
                DDELabel.LinkRequest
                If DDELabel.Caption <> "" Then
                    DopplerRatio = (a + Cdbl2(DDELabel.Caption)) / a
                End If
                'retrieve uplink freq.
                DDELabel.LinkItem = "UplinkMHz"
                DDELabel.LinkRequest
                If DDELabel.Caption <> "" Then
                    UplinkDDEFreq.text = Str(Fix(Cdbl2(DDELabel.Caption) / DopplerRatio))
                End If
                'retrieve downlink freq.
                DDELabel.LinkItem = "DownlinkMHz"
                DDELabel.LinkRequest
                If DDELabel.Caption <> "" Then
                    DownlinkDDEFreq.text = Str(Fix(Cdbl2(DDELabel.Caption) * DopplerRatio))
                End If
                'retrieve uplink mode
                DDELabel.LinkItem = "ModeInfo"
                DDELabel.LinkRequest
                UplinkMode.text = DDELabel.Caption
                'retrieve downlink mode
                DDELabel.LinkItem = "ModeInfo"
                DDELabel.LinkRequest
                DownlinkMode.text = DDELabel.Caption
                'updating Rotor fields:
                'retrieve Azimuth
                DDELabel.LinkItem = "AzimuthDegrees"
                DDELabel.LinkRequest
                If DDELabel.Caption <> "" Then
                    If RotorAuto.Value Then Azimuth.text = Str(Cdbl2(DDELabel.Caption))
                End If
                'Elevation
                DDELabel.LinkItem = "ElevationDegrees"
                DDELabel.LinkRequest
                If DDELabel.Caption <> "" Then
                    If RotorAuto.Value Then Elevation.text = Str(Cdbl2(DDELabel.Caption))
                End If
                'and the sat's name
                DDELabel.LinkItem = "SatelliteName"
                DDELabel.LinkRequest
                Satellite.ForeColor = &HFF0000
                Satellite.text = DDELabel.Caption
            End If
            
            '*****DDE in Nova Program mode******
            If frmDdelink.DDEFormat.text = "Nova" Then
                a = InStr(DDELabel.Caption, "AH:")
                If a Then
                    NovaVisible = Firststring(Mid(DDELabel.Caption, a + 3))
                Else
                'if no valid DDE string, NFW
                'might be waiting for start command:
                    DDELabel.Caption = "TUNE ON"
                    DDELabel.LinkPoke
                End If
                If Left(NovaVisible, 1) = "Y" Then
                    ' Azimuth
                    a = InStr(DDELabel.Caption, "AZ:")
                    If a And (RotorAuto.Value = 1) Then
                        Azimuth.text = Firststring(Mid(DDELabel.Caption, a + 3))
                    End If
                    'Elevation
                    a = InStr(DDELabel.Caption, "EL:")
                    If a And (RotorAuto.Value = 1) Then
                        Elevation.text = Firststring(Mid(DDELabel.Caption, a + 3))
                    End If
                    
                    ' in the case of NFW sat name is updated
                    'before freqs. because to compute exact freqs
                    'we need to know what sat is to be tracked
                    'Modes are set only once when satellite is
                    'changed
                    SatelliteName = Firststring(DDELabel.Caption)
                    Satellite.ForeColor = &HFF0000
                    Satellite.text = SatelliteName
                    
                    'compute freqs only if a valid sat is being tracked
                    If frmSats.SatIndex.text <> "" Then
                        a = InStr(DDELabel.Caption, "RR:")
                        If a Then
                            RR = Cdbl2(Firststring(Mid(DDELabel.Caption, a + 3)))
                            DopplerMHz = -(Cdbl2(frmSats.SatDownlinkFreq.text)) * RR / 299792.458
                            DownlinkDDEFreq.text = Str(0.000001 * Fix(1000000# * (Cdbl2(frmSats.SatDownlinkFreq.text) + DopplerMHz)))
                            DopplerMHz = -(Cdbl2(frmSats.SatUplinkFreq.text)) * RR / 299792.458
                            'if dual downlink is selected, use uplink radio as second dnlink
                            If frmSats.Sat2Dnlink Then
                                UplinkDDEFreq.text = Str(0.000001 * Fix(1000000# * (Cdbl2(frmSats.SatUplinkFreq.text) + DopplerMHz)))
                            Else
                                UplinkDDEFreq.text = Str(0.000001 * Fix(1000000# * (Cdbl2(frmSats.SatUplinkFreq.text) - DopplerMHz)))
                            End If
                        End If
                    End If
                Else
                    Satellite.ForeColor = &H808080
                    Satellite.text = "NO SATELLITE"
                End If
            End If
            
            'In case satellite name changed and some port had to
            'be opened, finish pending events.
            'DoEvents
            
            If frmRotor.RotorType.text <> "" Then
                '*****ROTOR CONTROL PROCESSING*****
                'Flip detection:
                'if the down-counter reached zero, we calculate
                'azimuth variation from the first value stored
                'before in RotorAz (when we initialized the
                'downcounter)
                If RotorControlCountdown = 0 Then
                    RotorDeltaAz = Cdbl2(Azimuth.text) - RotorAz
                    RotorFlip = False
                    'if stop position is South:
                    If frmRotor.RotorSouth.Value = True Then
                        'if the sat appeared from the West and
                        'azimuth is decreasing -> we flip
                        If (180 < RotorAz) And (RotorAz < 360) And (RotorDeltaAz < 0) Then
                            RotorFlip = True
                        End If
                        'if the sat appeared from the East and
                        'az. is increasing -> we flip
                        If (0 < RotorAz) And (RotorAz < 180) And (RotorDeltaAz >= 0) Then
                            RotorFlip = True
                        End If
                    Else
                        'if stop position is North:
                        'if the sat appeared from the West and
                        'asimuth is increasing -> we flip
                        If (180 < RotorAz) And (RotorAz < 360) And (RotorDeltaAz >= 0) Then
                            RotorFlip = True
                        End If
                        'if the sat appeared from the East and
                        'az. is decreasing -> we flip
                        If (0 < RotorAz) And (RotorAz < 180) And (RotorDeltaAz <= 0) Then
                            RotorFlip = True
                        End If
                    End If
                    'in any other case we do not flip
                    'to ensure that flip check is not run again
                    'until sat changes
                    RotorControlCountdown = -1
                End If
                'if rotor down-counter is not zero yet, we keep
                'counting and do not send any command to the
                'rotor controller
                If RotorControlCountdown > 0 Then
                    RotorControlCountdown = RotorControlCountdown - 1
                Else
                '*** ACTUAL ROTOR CONTROL PROCESSING ***
                'if rotor down-counter has reached zero (or -1)
                'rotor control is processed:
                    Az = Cdbl2(frmMain.Azimuth.text)
                    El = Cdbl2(frmMain.Elevation.text)
                    If RotorFlip Then
                        Frame1.Caption = "Rotor (Flipped)"
                        El = 180 - El
                        Az = Az - 180
                        If Az < 0 Then
                            Az = Az + 360
                        End If
                    Else
                        Frame1.Caption = "Rotor"
                    End If
                    If RotorAuto.Value Then
                        If Cdbl2(frmRotor.RotorStep.text) <> 0 Then
                            Az = Cdbl2(frmRotor.RotorStep.text) * _
                                (CInt(Az / Cdbl2(frmRotor.RotorStep.text)))
                            El = Cdbl2(frmRotor.RotorStep.text) * _
                                (CInt(El / Cdbl2(frmRotor.RotorStep.text)))
                        End If
                        Call UpdateRotor(Az, El)
                    End If
                End If
            End If
            
            '****RADIO CONTROL PROCESSING*****
            'syncronize with radio control loop routine,
            'to update just after it.
            'Do
            'Loop Until RadioControlLoopTimer.Enabled = False
            'routine is running now
'            Do
'            Loop Until RadioControlLoopTimer.Enabled = True
            'routine has just ended...
            'disable the loop routine while we update radios...
'            RadioControlLoopTimer.Enabled = False
            
            
'            RadioControlLoopTimer.Enabled = True

            
        End If
    Else
NoDDE:
        Satellite.text = "No DDE Link"
    End If
End Sub
Function RadioDownlinkEnabled()
Dim flag As Boolean
    flag = True
    'if TNC freq. control for PSK downlink is selected
    If frmRadio.RadioTNCUD.Value = 1 Then
        'we check that dnlink mode is single-side-band
         If (LCase$(DownlinkMode.text) = "usb" Or LCase$(DownlinkMode.text) = "lsb") Then
            flag = False
            'and if we haven't reached yet the predefined
            'max. number of frequency corrections:
            If RadioControlCountdown > 0 Then
                'enable freq. correction and
                'decrement the down-counter
                flag = True
                RadioControlCountdown = RadioControlCountdown - 1
                'if we have reached the dncount, but the
                'sat is low (Elev. < 3deg), we still correct
                'doppler. This is because the TNC may not lock
                'and thus will not be able to perform a good
                'freq. control
            ElseIf Cdbl2(Elevation.text) < 3 Then
                flag = True
            End If
        End If
    End If
    RadioDownlinkEnabled = flag
End Function
'Get the tuning limits of current band
Function IC821RadioReadBandLimits(Address%, Handle%) As Variant
Dim Limit(1 To 2) As Long
NRep% = 0
Do
    InString = ReadFromPort(Handle%)
    'limits are initialized so that if no valid reply is
    'detected there will be no band-switch
    Limit(1) = 0
    Limit(2) = 999999999
    IC821RadioReadBandLimits = ""
    'Send bytes to CI-V. Begin with two syncs (FE) then the
    'destination address, then the origin address (E0 is
    'always the master), then the command code and
    'its data (no data in this case) and finally the end
    'code (FD)
    Call WriteToPort(Chr$(&HFE) + Chr$(&HFE) + _
        Chr$(Address%) + Chr$(&HE0) + _
        Chr$(&H2) + Chr$(&HFD), Handle%)
    Call WaitOutBuffEmpty(Handle%)
    NRep% = NRep% + 1
    'wait up to 10 times timeout-time for a frame correctly
    'addressed to bus master
    'this filters-out echoed frames.
    b = Timer
    Do
        a$ = IC821RadioReadFrame(Handle%)
    Loop Until (Cdbl2(PickWord(a$, 1)) = 11) Or (Abs(Timer - b) > RadioReplyTimeout(Handle%) * 5)
    'check frame code 11
    If Cdbl2(PickWord(a$, 1)) <> 11 Then
        OK% = 0
    Else
        Limit(1) = Cdbl2(PickWord(a$, 2))
        Limit(2) = Cdbl2(PickWord(a$, 3))
        OK% = 1
    End If
Loop Until OK% = 1 Or NRep% > 5
If OK% = 0 Then
    Call frmMessage.ShowMessage("Comms. Error during" + Chr$(13) _
        + "Icom CI-V Read Band Limits", 10)
End If
IC821RadioReadBandLimits = Limit()
End Function
'Get the frequency of current band
Function IC821RadioReadFreq(Address%, Handle%) As Double
NRep% = 0
Do
    InString = ReadFromPort(Handle%)
    'Send bytes to CI-V. Begin with two syncs (FE) then the
    'destination address, then the origin address (E0 is
    'always the master), then the command code and
    'its data (no data in this case) and finally the end
    'code (FD)
    Call WriteToPort(Chr$(&HFE) + Chr$(&HFE) + _
        Chr$(Address%) + Chr$(&HE0) + _
        Chr$(&H3) + Chr$(&HFD), Handle%)
    Call WaitOutBuffEmpty(Handle%)
    NRep% = NRep% + 1
    'wait up to 10 times timeout-time for a frame correctly
    'addressed to bus master
    'this filters-out echoed frames.
    b = Timer
    Do
        a$ = IC821RadioReadFrame(Handle%)
    Loop Until (Cdbl2(PickWord(a$, 1)) = 10) Or (Abs(Timer - b) > RadioReplyTimeout(Handle%) * 5)
    'check frame code 10
    If Cdbl2(PickWord(a$, 1)) <> 10 Then
    '    Call frmMessage.ShowMessage("Error during" + Chr$(13) _
            + "Icom CI-V Read Frequency", 10)
        IC821RadioReadFreq = 0
        OK% = 0
    Else
        IC821RadioReadFreq = Cdbl2(PickWord(a$, 2))
        OK% = 1
    End If
Loop Until OK% = 1 Or NRep% > 3
End Function
'Get the duplex offset frequency
Function IC821RadioReadOffset(Address%, Handle%) As Double
NRep% = 0
Do
    InString = ReadFromPort(Handle%)
    'Send bytes to CI-V. Begin with two syncs (FE) then the
    'destination address, then the origin address (E0 is
    'always the master), then the command code and
    'its data (no data in this case) and finally the end
    'code (FD)
    Call WriteToPort(Chr$(&HFE) + Chr$(&HFE) + _
        Chr$(Address%) + Chr$(&HE0) + _
        Chr$(&HC) + Chr$(&HFD), Handle%)
    Call WaitOutBuffEmpty(Handle%)
    NRep% = NRep% + 1
    'wait up to 10 times timeout-time for a frame correctly
    'addressed to bus master
    'this filters-out echoed frames.
    b = Timer
    Do
        a$ = IC821RadioReadFrame(Handle%)
    Loop Until (Cdbl2(PickWord(a$, 1)) = 60) Or (Abs(Timer - b) > RadioReplyTimeout(Handle%) * 5)
    'check frame code 10
    If Cdbl2(PickWord(a$, 1)) <> 60 Then
    '    Call frmMessage.ShowMessage("Error during" + Chr$(13) _
            + "Icom CI-V Read Offset Frequency", 10)
        IC821RadioReadOffset = 0
        OK% = 0
    Else
        IC821RadioReadOffset = Cdbl2(PickWord(a$, 2))
        OK% = 1
    End If
Loop Until OK% = 1 Or NRep% > 3
End Function
'Get Squelch status
Function IC821RadioReadSquelch(Address%, Handle%) As Double
NRep% = 0
Do
    InString = ReadFromPort(Handle%)
    'Send bytes to CI-V. Begin with two syncs (FE) then the
    'destination address, then the origin address (E0 is
    'always the master), then the command code and
    'its data (no data in this case) and finally the end
    'code (FD)
    Call WriteToPort(Chr$(&HFE) + Chr$(&HFE) + _
        Chr$(Address%) + Chr$(&HE0) + _
        Chr$(&H15) + Chr$(&H1) + Chr$(&HFD), Handle%)
    Call WaitOutBuffEmpty(Handle%)
    NRep% = NRep% + 1
    'wait up to 10 times timeout-time for a frame correctly
    'addressed to bus master
    'this filters-out echoed frames.
    b = Timer
    Do
        a$ = IC821RadioReadFrame(Handle%)
    Loop Until (Cdbl2(PickWord(a$, 1)) = 10) Or (Abs(Timer - b) > RadioReplyTimeout(Handle%) * 5)
    'check frame code 10
    If Cdbl2(PickWord(a$, 1)) <> 10 Then
    '    Call frmMessage.ShowMessage("Error during" + Chr$(13) _
            + "Icom CI-V Read Frequency", 10)
        IC821RadioReadSquelch = 0
        OK% = 0
    Else
        IC821RadioReadSquelch = Cdbl2(PickWord(a$, 2))
        OK% = 1
    End If
Loop Until OK% = 1 Or NRep% > 5
End Function
Function IC706RadioReadRSSI(Address%, Handle%) As Double
NRep% = 0
Do
    InString = ReadFromPort(Handle%)
    'Send bytes to CI-V. Begin with two syncs (FE) then the
    'destination address, then the origin address (E0 is
    'always the master), then the command code and
    'its data (no data in this case) and finally the end
    'code (FD)
    Call WriteToPort(Chr$(&HFE) + Chr$(&HFE) + _
        Chr$(Address%) + Chr$(&HE0) + _
        Chr$(&H15) + Chr$(&H2) + Chr$(&HFD), Handle%)
    Call WaitOutBuffEmpty(Handle%)
    NRep% = NRep% + 1
    'wait up to 10 times timeout-time for a frame correctly
    'addressed to bus master
    'this filters-out echoed frames.
    b = Timer
    Do
        a$ = IC821RadioReadFrame(Handle%)
    Loop Until (Cdbl2(PickWord(a$, 1)) = 10) Or (Abs(Timer - b) > RadioReplyTimeout(Handle%) * 5)
    'check frame code 10
    If Cdbl2(PickWord(a$, 1)) <> 40 Then
    '    Call frmMessage.ShowMessage("Error during" + Chr$(13) _
            + "Icom CI-V Read Frequency", 10)
        IC706RadioReadRSSI = 0
        OK% = 0
    Else
        IC706RadioReadRSSI = Cdbl2(PickWord(a$, 2))
        OK% = 1
    End If
Loop Until OK% = 1 Or NRep% > 5
End Function
'CI-V frame readback routine
'receives chars from radio port until 0xFD is detected (end of frame sync)
'then parses the complete frame for valid data.
'The output of this function is a string beginning with the decoded
'frame type code followed by the data available separated by spaces.
'the following type codes can be returned: (freqs. in Hz)
'code "00:" - frame not understood pass timeout period.
'code "01:" - acknowlegde frame (OK message) was received)
'code "02:" - radio reported an error (NG message).
'code "03:" - frame was not directed to bus master (PC)
'code "10:" - frequency information (of currently selected band)
'code "11:" - band limits frequencies (lower freq first then upper)
'code "20:" - operating mode infor (of currently sel. band)
'code "30:" - PTT (TX) status
'code "40:" - RSSI
'code "50:" - IF bandwidth.
'code "60:" - offset freq.
'code "70:" - band limits
Function IC821RadioReadFrame(Handle%) As String
'This will hold the received frame
Dim InBuff(100) As Integer
'initialize buffer pointer..
InBuffPtr% = 0
'Frame-Ending signal:
FrameEndChar% = &HFD
'this will indicate the position where the 0xFD byte is found.
'which indicated end of frame.
'while it is -1 it means it haven't been found yet. after it
'is found we need to be sure that the reqd. number of bytes
'is received
fend% = -1
f% = 0
'set endint time...
a = Timer
Do
    'Wait until we receive bytes
    'or a time-out occurs
    Select Case Handle%
    Case 1
        Do
            
        Loop Until (MSComm1.InBufferCount >= 1) Or (Abs(Timer - a) > RadioReplyTimeout(Handle%))
        b = Abs(Timer - a)
        c = MSComm1.InBufferCount
        InString = MSComm1.Input
    Case 2
        Do
            
        Loop Until (MSComm2.InBufferCount >= 1) Or Abs(Timer - a) > RadioReplyTimeout(Handle%)
        InString = MSComm2.Input
    Case 3
        Do
            
        Loop Until (MSComm3.InBufferCount >= 1) Or Abs(Timer - a) > RadioReplyTimeout(Handle%)
        InString = MSComm3.Input
    End Select
    'We pass the received bytes to the buffer
    For i% = 0 To LenB(InString) - 1
        InBuff(InBuffPtr%) = InString(i%)
        'we will examine the buffer searching for 0xFD:
        If InString(i%) = &HFD Then fend% = InBuffPtr%
        InBuffPtr% = InBuffPtr% + 1
    Next
Loop While fend% = -1 And Abs(Timer - a) < RadioReplyTimeout(Handle%)
'if loop ended due to timeout:
If fend% = -1 Then
    s$ = "00:"
    IC821RadioReadFrame = s$
    Exit Function
Else
    'otherwise examine frame:
    'check that the frame has at least 6 bytes
    If fend% < 5 Then
        s$ = "00:"
        IC821RadioReadFrame = s$
        Exit Function
    End If
    'look for start of frame and
    'check integrity at the same time:
    For fstart% = fend% - 4 To 1 Step -1
        If InBuff(fstart%) = &HFE Then
            fstart% = fstart% - 1
            If InBuff(fstart%) = &HFE Then
                'frame is complete (starts with FE,FE and ends with FD)
                'now check it IS addressed to bus master (E0):
                If InBuff(fstart% + 2) = &HE0 Then
                    Exit For
                Else
                    'it's not addressed to bus master...
                    s$ = "03:"
                    IC821RadioReadFrame = s$
                    Exit Function
                End If
            Else
                s$ = "00:"
                IC821RadioReadFrame = s$
                Exit Function
            End If
        End If
    Next
    If InBuff(fstart%) <> &HFE Then
        s$ = "00:"
        IC821RadioReadFrame = s$
        Exit Function
    End If
    'so we have a valid frame...
    'check frame is directed to controller (bus master)
    If InBuff(fstart% + 2) <> &HE0 Then
        s$ = "03:"
        IC821RadioReadFrame = s$
        Exit Function
    End If
    'check for OK message:
    If InBuff(fend% - 1) = &HFB Then
        s$ = "01:"
        IC821RadioReadFrame = s$
        Exit Function
    End If
    'check for NG message:
    If InBuff(fend% - 1) = &HFA Then
        s$ = "02:"
        IC821RadioReadFrame = s$
        Exit Function
    End If
    'check for Band Limits answer:
    If InBuff(fstart% + 4) = &H2 Then
        s$ = "11: "
        edge& = 0
        For f% = fstart% + 9 To fstart% + 5 Step -1
            edge& = edge& * 10
            edge& = edge& + (InBuff(f%) And &HF0) / 16
            edge& = edge& * 10
            edge& = edge& + (InBuff(f%) And &HF)
        Next
        s$ = s$ + Str$(edge&) + " "
        edge& = 0
        For f% = fstart% + 15 To fstart% + 11 Step -1
            edge& = edge& * 10
            edge& = edge& + (InBuff(f%) And &HF0) / 16
            edge& = edge& * 10
            edge& = edge& + (InBuff(f%) And &HF)
        Next
        s$ = s$ + Str$(edge&)
        IC821RadioReadFrame = s$
        Exit Function
    End If
    'check for freq. answer:
    If InBuff(fstart% + 4) = &H3 Then
        s$ = "10: "
        edge& = 0
        For f% = fstart% + 9 To fstart% + 5 Step -1
            edge& = edge& * 10
            edge& = edge& + (InBuff(f%) And &HF0) / 16
            edge& = edge& * 10
            edge& = edge& + (InBuff(f%) And &HF)
        Next
        s$ = s$ + Str$(edge&)
        IC821RadioReadFrame = s$
        Exit Function
    End If
    'check for mode answer:
    If InBuff(fstart% + 4) = &H4 Then
        s$ = "20: "
        Select Case InBuff(fstart% + 5)
            Case 0
            s$ = s$ + "LSB"
            Case 1
            s$ = s$ + "USB"
            Case 2
            s$ = s$ + "AM"
            Case 3
            s$ = s$ + "CW"
            Case 5
            'IC-R7000 variation:
            If InBuff(fstart% + 6) = 0 Then
                s$ = s$ + "SSB"
            Else
                s$ = s$ + "FM"
            End If
            Case 6
            s$ = s$ + "FM-W"
        End Select
    End If
    'check for offset freq. answer:
    If InBuff(fstart% + 4) = &HC Then
        s$ = "60: "
        edge& = 0
        For f% = fstart% + 7 To fstart% + 5 Step -1
            edge& = edge& * 10
            edge& = edge& + (InBuff(f%) And &HF0) / 16
            edge& = edge& * 10
            edge& = edge& + (InBuff(f%) And &HF)
        Next
        'offset freq comes in hundreds
        edge& = edge& * 100
        s$ = s$ + Str$(edge&)
        IC821RadioReadFrame = s$
        Exit Function
    End If
    'check for RSSI answer:
    If InBuff(fstart% + 4) = &H15 And InBuff(fstart% + 5) = &H2 Then
        s$ = "40: "
        edge& = 0
        For f% = fstart% + 6 To fstart% + 7
            edge& = edge& * 10
            edge& = edge& + (InBuff(f%) And &HF0) / 16
            edge& = edge& * 10
            edge& = edge& + (InBuff(f%) And &HF)
        Next
        s$ = s$ + Str$(edge&)
        IC821RadioReadFrame = s$
        Exit Function
    End If
End If
End Function
'Exchange Main<->Sub bands on Icom rigs
Sub IC821RadioMS(Address%, Bidir%, Handle%)
    InString = ReadFromPort(Handle%)
    'Send bytes to CI-V. Begin with two syncs (FE) then the
    'destination address, then the origin address (E0 is
    'always the master), then the command code and
    'its data and finally the end code (FD)
    Call WriteToPort(Chr$(&HFE) + Chr$(&HFE) + _
        Chr$(Address%) + Chr$(&HE0) + Chr$(&H7) + _
        Chr$(&HB0) + Chr$(&HFD), Handle%)
    Call WaitOutBuffEmpty(Handle%)
    If Bidir% Then
        'wait up to 10 times timeout-time for a frame correctly
        'addressed to bus master
        'this filters-out echoed frames.
        b = Timer
        Do
            a$ = IC821RadioReadFrame(Handle%)
        Loop Until (Cdbl2(PickWord(a$, 1)) = 1) Or (Abs(Timer - b) > RadioReplyTimeout(Handle%) * 5)
        If Cdbl2(PickWord(a$, 1)) <> 1 Then
            Call frmMessage.ShowMessage("Comms. Error during" + Chr$(13) _
                + "Icom CI-V Main<->Sub Command", 10)
        End If
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub

Sub IC821RadioMain(Address%, Bidir%, Handle%)
    InString = ReadFromPort(Handle%)
    'Send bytes to CI-V. Begin with two syncs (FE) then the
    'destination address, then the origin address (E0 is
    'always the master), then the command code and
    'its data and finally the end code (FD)
     Call WriteToPort(Chr$(&HFE) + Chr$(&HFE) + _
        Chr$(Address%) + Chr$(&HE0) + Chr$(&H7) + _
        Chr$(&HD0) + Chr$(&HFD), Handle%)
    
    Call WaitOutBuffEmpty(Handle%)
   If Bidir% Then
        'wait up to 5 times timeout-time for a frame correctly
        'addressed to bus master
        'this filters-out echoed frames.
        b = Timer
        Do
            a$ = IC821RadioReadFrame(Handle%)
        Loop Until (Cdbl2(PickWord(a$, 1)) = 1) Or (Abs(Timer - b) > RadioReplyTimeout(Handle%) * 5)
        If Cdbl2(PickWord(a$, 1)) <> 1 Then
            Call frmMessage.ShowMessage("Comms. Error during" + Chr$(13) _
                + "Icom CI-V Set MainBand", 10)
        End If
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub

Sub IC821RadioSub(Address%, Bidir%, Handle%)
    InString = ReadFromPort(Handle%)
    'Send bytes to CI-V. Begin with two syncs (FE) then the
    'destination address, then the origin address (E0 is
    'always the master), then the command code and
    'its data and finally the end code (FD)
    a$ = Chr$(&HFE) + Chr$(&HFE) + _
        Chr$(Address%) + Chr$(&HE0) + Chr$(&H7) + _
        Chr$(&HD1) + Chr$(&HFD)
    Call WriteToPort(a$, Handle%)
    Call WaitOutBuffEmpty(Handle%)
    If Bidir% Then
        'wait up to 5 times timeout-time for a frame correctly
        'addressed to bus master
        'this filters-out echoed frames.
        b = Timer
        Do
            a$ = IC821RadioReadFrame(Handle%)
            X$ = X$ + a$
        Loop Until (Cdbl2(PickWord(a$, 1)) = 1) Or (Abs(Timer - b) > RadioReplyTimeout(Handle%) * 5)
        If Cdbl2(PickWord(a$, 1)) <> 1 Then
            Call frmMessage.ShowMessage("Comms. Error during" + Chr$(13) _
                + "Icom CI-V Set SubBand", 10)
        End If
    Else
        a = Timer
        Do
            
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub

Function IC821RadioSetFreq(Freq, Address%, Bidir%, Handle%) As Integer
freq2 = Freq
NRep% = 0
Do
    a$ = ""
    Freq = freq2
    If Freq < 0 Then Freq = 0
    InString = ReadFromPort(Handle%)
    'Send bytes to CI-V. Begin with two syncs (FE) then the
    'destination address, then the origin address (E0 is
    'always the master), then the command code and
    'its data and finally the end code (FD)
    Call WriteToPort(Chr$(&HFE) + Chr$(&HFE) + _
        Chr$(Address%) + Chr$(&HE0) + Chr$(&H5), Handle%)
    a$ = a$ + Chr$(&HFE) + Chr$(&HFE) + Chr$(Address%) + Chr$(&HE0) + Chr$(&H5)
    For f% = 1 To 5
        n% = Fix(Freq - 10 * Fix(Freq / 10)) ' G6LVB 17 Jan 2001 Changed Freq& to Freq and altered to make work for 2.048GHz overflow bug
        If n% >= 10 Then
            n% = 0
        End If
        Freq = Fix(Freq / 10)
        n% = n% + 16 * Fix(Freq - 10 * Fix(Freq / 10))
        Freq = Fix(Freq / 10)
        Call WriteToPort(Chr$(n%), Handle%)
        a$ = a$ + Chr$(n%)
    Next
    Call WriteToPort(Chr$(&HFD), Handle%)
    a$ = a$ + Chr$(&HFD)
    ' Check if logging is active
    If frmRadio.CheckLogCom.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Sending SetFreq. at " + Str(freq2) + "Hz Command: " + StrToHex(a$) + " to " + DownlinkModel$ + " downlink radio on stream" + Str(DownlinkHandle%) + ", " + Str(Time)
        Close f%
    End If
    Call WaitOutBuffEmpty(Handle%)
    NRep% = NRep% + 1
    If Bidir% Then
        'wait up to 3 times timeout-time for a frame correctly
        'addressed to bus master
        'this filters-out echoed frames.
        b = Timer
        Do
            a$ = IC821RadioReadFrame(Handle%)
            c$ = c$ + a$
        Loop Until (Cdbl2(PickWord(a$, 1)) = 1) Or (Abs(Timer - b) > RadioReplyTimeout(Handle%))
        If Cdbl2(PickWord(a$, 1)) <> 1 Then
            OK% = 0
        Else
            OK% = 1
        End If
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
        OK% = 1
    End If
Loop Until OK% = 1 Or NRep% > 3
If OK% = 0 Then
    Call frmMessage.ShowMessage("Comms. Error during" + Chr$(13) _
        + "Icom CI-V Set Frequency", 5)
End If
'a = IC821RadioReadFreq(DownlinkCIVAddress%, DownlinkHandle%)
'If (a <> freq2) And (a <> 0) Then
'    a = a
'End If
IC821RadioSetFreq = OK%
End Function

Sub IC821RadioCancelSplit(Address%, Bidir%, Handle%)
    InString = ReadFromPort(Handle%)
    'Send bytes to CI-V. Begin with two syncs (FE) then the
    'destination address, then the origin address (E0 is
    'always the master), then the command code and
    'its data and finally the end code (FD)
    Call WriteToPort(Chr$(&HFE) + Chr$(&HFE) + _
        Chr$(Address%) + Chr$(&HE0) + Chr$(&HF) + _
        Chr$(&H0) + Chr$(&HFD), Handle%)
    Call WaitOutBuffEmpty(Handle%)
    If Bidir% Then
        'wait up to 5 times timeout-time for a frame correctly
        'addressed to bus master
        'this filters-out echoed frames.
        b = Timer
        Do
            a$ = IC821RadioReadFrame(Handle%)
            resp = Cdbl2(PickWord(a$, 1))
        Loop Until (resp = 1) Or (Abs(Timer - b) > RadioReplyTimeout(Handle%) * 5)
        If Cdbl2(PickWord(a$, 1)) <> 1 Then
            Call frmMessage.ShowMessage("Comms. Error during" + Chr$(13) _
                + "Icom CI-V Cancel Split Command", 10)
        End If
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
Sub IC821RadioSetSplit(Address%, Bidir%, Handle%)
    InString = ReadFromPort(Handle%)
    'Send bytes to CI-V. Begin with two syncs (FE) then the
    'destination address, then the origin address (E0 is
    'always the master), then the command code and
    'its data and finally the end code (FD)
    Call WriteToPort(Chr$(&HFE) + Chr$(&HFE) + _
        Chr$(Address%) + Chr$(&HE0) + Chr$(&HF) + _
        Chr$(&H1) + Chr$(&HFD), Handle%)
    Call WaitOutBuffEmpty(Handle%)
    If Bidir% Then
        'wait up to 5 times timeout-time for a frame correctly
        'addressed to bus master
        'this filters-out echoed frames.
        b = Timer
        Do
            a$ = IC821RadioReadFrame(Handle%)
            resp = Cdbl2(PickWord(a$, 1))
        Loop Until (resp = 1) Or (Abs(Timer - b) > RadioReplyTimeout(Handle%) * 5)
        If Cdbl2(PickWord(a$, 1)) = 0 Then
            Call frmMessage.ShowMessage("Comms. Error during" + Chr$(13) _
                + "Icom CI-V Set Split Command", 10)
        End If
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub

Sub IC821RadioCancelDuplex(Address%, Bidir%, Handle%)
    InString = ReadFromPort(Handle%)
    'Send bytes to CI-V. Begin with two syncs (FE) then the
    'destination address, then the origin address (E0 is
    'always the master), then the command code and
    'its data and finally the end code (FD)
    Call WriteToPort(Chr$(&HFE) + Chr$(&HFE) + _
        Chr$(Address%) + Chr$(&HE0) + Chr$(&HF) + _
        Chr$(&H10) + Chr$(&HFD), Handle%)
    Call WaitOutBuffEmpty(Handle%)
    If Bidir% Then
        'wait up to 5 times timeout-time for a frame correctly
        'addressed to bus master
        'this filters-out echoed frames.
        b = Timer
        Do
            a$ = IC821RadioReadFrame(Handle%)
            resp = Cdbl2(PickWord(a$, 1))
        Loop Until (resp = 1) Or (Abs(Timer - b) > RadioReplyTimeout(Handle%) * 5)
        If Cdbl2(PickWord(a$, 1)) <> 1 Then
            Call frmMessage.ShowMessage("Comms. Error during" + Chr$(13) _
                + "Icom CI-V Cancel Duplex Command", 10)
        End If
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
Sub IC821RadioSetDuplexPlus(Address%, Bidir%, Handle%)
    InString = ReadFromPort(Handle%)
    'Send bytes to CI-V. Begin with two syncs (FE) then the
    'destination address, then the origin address (E0 is
    'always the master), then the command code and
    'its data and finally the end code (FD)
    Call WriteToPort(Chr$(&HFE) + Chr$(&HFE) + _
        Chr$(Address%) + Chr$(&HE0) + Chr$(&HF) + _
        Chr$(&H12) + Chr$(&HFD), Handle%)
    Call WaitOutBuffEmpty(Handle%)
    If Bidir% Then
        'wait up to 5 times timeout-time for a frame correctly
        'addressed to bus master
        'this filters-out echoed frames.
        b = Timer
        Do
            a$ = IC821RadioReadFrame(Handle%)
            resp = Cdbl2(PickWord(a$, 1))
        Loop Until (resp = 1) Or (Abs(Timer - b) > RadioReplyTimeout(Handle%) * 5)
        If Cdbl2(PickWord(a$, 1)) <> 1 Then
            Call frmMessage.ShowMessage("Comms. Error during" + Chr$(13) _
                + "Icom CI-V Cancel Duplex Command", 10)
        End If
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
Sub IC821RadioSetDuplexMinus(Address%, Bidir%, Handle%)
    InString = ReadFromPort(Handle%)
    'Send bytes to CI-V. Begin with two syncs (FE) then the
    'destination address, then the origin address (E0 is
    'always the master), then the command code and
    'its data and finally the end code (FD)
    Call WriteToPort(Chr$(&HFE) + Chr$(&HFE) + _
        Chr$(Address%) + Chr$(&HE0) + Chr$(&HF) + _
        Chr$(&H11) + Chr$(&HFD), Handle%)
    Call WaitOutBuffEmpty(Handle%)
    If Bidir% Then
        'wait up to 5 times timeout-time for a frame correctly
        'addressed to bus master
        'this filters-out echoed frames.
        b = Timer
        Do
            a$ = IC821RadioReadFrame(Handle%)
            resp = Cdbl2(PickWord(a$, 1))
        Loop Until (resp = 1) Or (Abs(Timer - b) > RadioReplyTimeout(Handle%) * 5)
        If Cdbl2(PickWord(a$, 1)) <> 1 Then
            Call frmMessage.ShowMessage("Comms. Error during" + Chr$(13) _
                + "Icom CI-V Cancel Duplex Command", 10)
        End If
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
'Sets oper. mode for Icom radios
Sub IC821RadioSetMode(mode$, Address%, Bidir%, Handle%)
    Select Case LCase$(mode$)
        Case Is = "fm-n", "fm", "fm-w"
            m% = 5
        Case Is = "lsb"
            m% = 0
        Case Is = "usb"
            m% = 1
        Case Is = "cw", "cw-n"
            m% = 3
        Case Else
            m% = -1
    End Select
    If m% >= 0 Then
        InString = ReadFromPort(Handle%)
        'Send bytes to CI-V. Begin with two syncs (FE) then the
        'destination address, then the origin address (E0 is
        'always the master), then the command code and
        'its data and finally the end code (FD)
        Call WriteToPort(Chr$(&HFE) + Chr$(&HFE) + _
            Chr$(Address%) + Chr$(&HE0) + _
            Chr$(&H6), Handle%)
        Call WriteToPort(Chr$(m%), Handle%)
        Call WriteToPort(Chr$(&HFD), Handle%)
        Call WaitOutBuffEmpty(Handle%)
        If Bidir% Then
        'wait up to 10 times timeout-time for a frame correctly
        'addressed to bus master
        'this filters-out echoed frames.
        b = Timer
        Do
            a$ = IC821RadioReadFrame(Handle%)
        Loop Until (Cdbl2(PickWord(a$, 1)) = 1) Or (Abs(Timer - b) > RadioReplyTimeout(Handle%) * 5)
            If Cdbl2(PickWord(a$, 1)) <> 1 Then
                Call frmMessage.ShowMessage("Comms. Error during" + Chr$(13) _
                    + "Icom CI-V Set Mode", 10)
            End If
        Else
            a = Timer
            Do
            Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
        End If
    End If
End Sub
Function IC821RadioSetOffset(Freq, Address%, Bidir%, Handle%) As Integer
'They might pass a negative offset freq
Freq = Abs(Freq)
'Check offset is <100MHz
If (Freq >= 100000000) Then
    'If limit is exceeded return with 0 (error)
    IC821RadioSetOffset = 0
    Return
End If
freq2 = Freq
NRep% = 0
Do
    a$ = ""
    Freq = freq2
    If Freq < 100 Then Freq = 0
    'Offset Resolution is 100Hz
    Freq = Freq / 100
    InString = ReadFromPort(Handle%)
    'Send bytes to CI-V. Begin with two syncs (FE) then the
    'destination address, then the origin address (E0 is
    'always the master), then the command code and
    'its data and finally the end code (FD)
    Call WriteToPort(Chr$(&HFE) + Chr$(&HFE) + _
        Chr$(Address%) + Chr$(&HE0) + Chr$(&HD), Handle%)
    a$ = a$ + Chr$(&HFE) + Chr$(&HFE) + Chr$(Address%) + Chr$(&HE0) + Chr$(&HD)
    'Offset freq has only 3 BCD bytes
    For f% = 1 To 3
        n% = Fix(Freq - 10 * Fix(Freq / 10)) ' G6LVB 17 Jan 2001 Changed Freq& to Freq and altered to make work for 2.048GHz overflow bug
        If n% >= 10 Then
            n% = 0
        End If
        Freq = Fix(Freq / 10)
        n% = n% + 16 * Fix(Freq - 10 * Fix(Freq / 10))
        Freq = Fix(Freq / 10)
        Call WriteToPort(Chr$(n%), Handle%)
        a$ = a$ + Chr$(n%)
    Next
    Call WriteToPort(Chr$(&HFD), Handle%)
    a$ = a$ + Chr$(&HFD)
    ' Check if logging is active
    If frmRadio.CheckLogCom.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Sending SetFreq. at " + Str(freq2) + "Hz Command: " + StrToHex(a$) + " to " + DownlinkModel$ + " downlink radio on stream" + Str(DownlinkHandle%) + ", " + Str(Time)
        Close f%
    End If
    Call WaitOutBuffEmpty(Handle%)
    NRep% = NRep% + 1
    If Bidir% Then
        'wait up to 3 times timeout-time for a frame correctly
        'addressed to bus master
        'this filters-out echoed frames.
        b = Timer
        Do
            a$ = IC821RadioReadFrame(Handle%)
            c$ = c$ + a$
        Loop Until (Cdbl2(PickWord(a$, 1)) = 1) Or (Abs(Timer - b) > RadioReplyTimeout(Handle%))
        If Cdbl2(PickWord(a$, 1)) <> 1 Then
            OK% = 0
        Else
            OK% = 1
        End If
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
        OK% = 1
    End If
Loop Until OK% = 1 Or NRep% > 1
'When rig is on TX it can't update offset, we don't report this as error
'If OK% = 0 Then
'    Call frmMessage.ShowMessage("Comms. Error during" + Chr$(13) _
'        + "Icom CI-V Set Frequency", 5)
'End If
IC821RadioSetOffset = OK%
End Function
'Sets operating mode for IC-R7000 receiver:
Sub ICR7000RadioSetMode(mode$, Address%, Bidir%, Handle%)
    Select Case LCase$(mode$)
        Case Is = "fm-n", "fm"
            m% = &H502
        Case Is = "fm-w"
            m% = &H501
        Case Is = "lsb"
            m% = &H500
        Case Is = "usb"
            m% = &H500
        Case Is = "cw", "cw-n"
            m% = &H500
        Case Else
            m% = -1
    End Select
    If m% >= 0 Then
        InString = ReadFromPort(Handle%)
        'Send bytes to CI-V. Begin with two syncs (FE) then the
        'destination address, then the origin address (E0 is
        'always the master), then the command code and
        'its data and finally the end code (FD)
        Call WriteToPort(Chr$(&HFE) + Chr$(&HFE) + _
            Chr$(Address%) + Chr$(&HE0) + _
            Chr$(&H6), Handle%)
        ' some mode codes for the R7000 are two-byte so:
        If m% > &HFF Then
            Call WriteToPort(Chr$((m% And &HFF00) / &H100), Handle%)
            m% = m% And &HFF
        End If
        Call WriteToPort(Chr$(m%), Handle%)
        Call WriteToPort(Chr$(&HFD), Handle%)
        Call WaitOutBuffEmpty(Handle%)
        If Bidir% Then
        'wait up to 10 times timeout-time for a frame correctly
        'addressed to bus master
        'this filters-out echoed frames.
        b = Timer
        Do
            a$ = IC821RadioReadFrame(Handle%)
        Loop Until (Cdbl2(PickWord(a$, 1)) = 1) Or (Abs(Timer - b) > RadioReplyTimeout(Handle%) * 5)
            If Cdbl2(PickWord(a$, 1)) <> 1 Then
                Call frmMessage.ShowMessage("Comms. Error during" + Chr$(13) _
                    + "Icom CI-V Set Mode", 10)
            End If
        Else
            a = Timer
            Do
            Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
        End If
    End If
End Sub
'Sets operating mode for IC-R8500 receiver:
Sub ICR8500RadioSetMode(mode$, Address%, Bidir%, Handle%)
    Select Case LCase$(mode$)
        Case Is = "fm-n"
            m% = &H502
        Case Is = "fm"
            m% = &H501
        Case Is = "fm-w"
            m% = &H601
        Case Is = "lsb"
            m% = &H1
        Case Is = "usb"
            m% = &H101
        Case Is = "cw"
            m% = &H500
        Case Is = "cw-n"
            m% = &H500
        Case Is = "am"
            m% = &H202
        Case Else
            m% = -1
    End Select
    If m% >= 0 Then
        InString = ReadFromPort(Handle%)
        'Send bytes to CI-V. Begin with two syncs (FE) then the
        'destination address, then the origin address (E0 is
        'always the master), then the command code and
        'its data and finally the end code (FD)
        Call WriteToPort(Chr$(&HFE) + Chr$(&HFE) + _
            Chr$(Address%) + Chr$(&HE0) + _
            Chr$(&H6), Handle%)
        'send high byte of the mode code
        Call WriteToPort(Chr$((m% And &HFF00) / &H100), Handle%)
        m% = m% And &HFF
        'and lower byte..
        Call WriteToPort(Chr$(m%), Handle%)
        Call WriteToPort(Chr$(&HFD), Handle%)
        Call WaitOutBuffEmpty(Handle%)
        If Bidir% Then
        'wait up to 10 times timeout-time for a frame correctly
        'addressed to bus master
        'this filters-out echoed frames.
        b = Timer
        Do
            a$ = IC821RadioReadFrame(Handle%)
        Loop Until (Cdbl2(PickWord(a$, 1)) = 1) Or (Abs(Timer - b) > RadioReplyTimeout(Handle%) * 5)
            If Cdbl2(PickWord(a$, 1)) <> 1 Then
                Call frmMessage.ShowMessage("Comms. Error during" + Chr$(13) _
                    + "Icom CI-V Set Mode", 10)
            End If
        Else
            a = Timer
            Do
            Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
        End If
    End If
End Sub
'Turns on IC-R8500 receiver:
Sub ICR8500RadioOn(Address%, Bidir%, Handle%)
    InString = ReadFromPort(Handle%)
    'Send bytes to CI-V. Begin with two syncs (FE) then the
    'destination address, then the origin address (E0 is
    'always the master), then the command code and
    'its data and finally the end code (FD)
    Call WriteToPort(Chr$(&HFE) + Chr$(&HFE) + _
        Chr$(Address%) + Chr$(&HE0) + _
        Chr$(&H18) + Chr$(&H1) + Chr$(&HFD), Handle%)
    Call WaitOutBuffEmpty(Handle%)
    If Bidir% Then
        'wait up to 10 times timeout-time for a frame correctly
        'addressed to bus master
        'this filters-out echoed frames.
        b = Timer
        Do
            a$ = IC821RadioReadFrame(Handle%)
        Loop Until (Cdbl2(PickWord(a$, 1)) = 1) Or (Abs(Timer - b) > RadioReplyTimeout(Handle%) * 5)
        If Cdbl2(PickWord(a$, 1)) <> 1 Then
            Call frmMessage.ShowMessage("Comms. Error during" + Chr$(13) _
                + "Icom CI-V Radio On", 10)
        End If
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
'Turns off IC-R8500 receiver:
Sub ICR8500RadioOff(Address%, Bidir%, Handle%)
    InString = ReadFromPort(Handle%)
    'Send bytes to CI-V. Begin with two syncs (FE) then the
    'destination address, then the origin address (E0 is
    'always the master), then the command code and
    'its data and finally the end code (FD)
    Call WriteToPort(Chr$(&HFE) + Chr$(&HFE) + _
        Chr$(Address%) + Chr$(&HE0) + _
        Chr$(&H18) + Chr$(&H0) + Chr$(&HFD), Handle%)
    Call WaitOutBuffEmpty(Handle%)
    If Bidir% Then
        'wait up to 10 times timeout-time for a frame correctly
        'addressed to bus master
        'this filters-out echoed frames.
        b = Timer
        Do
            a$ = IC821RadioReadFrame(Handle%)
        Loop Until (Cdbl2(PickWord(a$, 1)) = 1) Or (Abs(Timer - b) > RadioReplyTimeout(Handle%) * 5)
        If Cdbl2(PickWord(a$, 1)) <> 1 Then
            Call frmMessage.ShowMessage("Comms. Error during" + Chr$(13) _
                + "Icom CI-V Radio Off", 10)
        End If
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
'Set Icom radios to VFO mode
Sub IC821RadioVFO(Address%, Bidir%, Handle%)
    InString = ReadFromPort(Handle%)
    'Send bytes to CI-V. Begin with two syncs (FE) then the
    'destination address, then the origin address (E0 is
    'always the master), then the command code and
    'its data and finally the end code (FD)
     Call WriteToPort(Chr$(&HFE) + Chr$(&HFE) + _
        Chr$(Address%) + Chr$(&HE0) + Chr$(&H7) + _
        Chr$(&HFD), Handle%)
    Call WaitOutBuffEmpty(Handle%)
    If Bidir% Then
        'wait up to 10 times timeout-time for a frame correctly
        'addressed to bus master
        'this filters-out echoed frames.
        b = Timer
        Do
            a$ = IC821RadioReadFrame(Handle%)
            X$ = X$ + a$
            resp = Cdbl2(PickWord(a$, 1))
        Loop Until (resp = 1) Or (Abs(Timer - b) > RadioReplyTimeout(Handle%) * 5)
        If Cdbl2(PickWord(a$, 1)) <> 1 Then
            Call frmMessage.ShowMessage("Comms. Error during" + Chr$(13) _
                + "Icom CI-V Set VFO", 10)
        End If
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
'Set Icom radios to VFO-A mode
'returns the frame code for which it exits (00: not understood, 01: OK, 02: NG, 03: not
'addressed to master etc.)
Function IC706RadioSetVFOA(Address%, Bidir%, Handle%) As Integer
    InString = ReadFromPort(Handle%)
    'Send bytes to CI-V. Begin with two syncs (FE) then the
    'destination address, then the origin address (E0 is
    'always the master), then the command code and
    'its data and finally the end code (FD)
     Call WriteToPort(Chr$(&HFE) + Chr$(&HFE) + _
        Chr$(Address%) + Chr$(&HE0) + Chr$(&H7) + Chr$(&H0) + _
        Chr$(&HFD), Handle%)
    Call WaitOutBuffEmpty(Handle%)
    If Bidir% Then
        'wait up to 10 times timeout-time for a frame correctly
        'addressed to bus master
        'this filters-out echoed frames.
        b = Timer
        Do
            a$ = IC821RadioReadFrame(Handle%)
            X$ = X$ + a$
            resp = Cdbl2(PickWord(a$, 1))
        Loop Until (resp = 1) Or (Abs(Timer - b) > RadioReplyTimeout(Handle%) * 5)
        If resp = 2 Then
            a = a
        End If
        If Cdbl2(PickWord(a$, 1)) <> 1 Then
            Call frmMessage.ShowMessage("Comms. Error during" + Chr$(13) _
                + "Icom CI-V Set VFO-A", 10)
        End If
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
    IC706RadioSetVFOA = resp
End Function
'Set Icom radios to VFO-B mode
'returns the frame code for which it exits (00: not understood, 01: OK, 02: NG, 03: not
'addressed to master etc.)
Function IC706RadioSetVFOB(Address%, Bidir%, Handle%) As Integer
    InString = ReadFromPort(Handle%)
    'Send bytes to CI-V. Begin with two syncs (FE) then the
    'destination address, then the origin address (E0 is
    'always the master), then the command code and
    'its data and finally the end code (FD)
     Call WriteToPort(Chr$(&HFE) + Chr$(&HFE) + _
        Chr$(Address%) + Chr$(&HE0) + Chr$(&H7) + Chr$(&H1) + _
        Chr$(&HFD), Handle%)
    Call WaitOutBuffEmpty(Handle%)
    If Bidir% Then
        'wait up to 10 times timeout-time for a frame correctly
        'addressed to bus master
        'this filters-out echoed frames.
        b = Timer
        Do
            a$ = IC821RadioReadFrame(Handle%)
            X$ = X$ + a$
            resp = Cdbl2(PickWord(a$, 1))
        Loop Until (resp = 1) Or (Abs(Timer - b) > RadioReplyTimeout(Handle%) * 5)
        If Cdbl2(PickWord(a$, 1)) <> 1 Then
            Call frmMessage.ShowMessage("Comms. Error during" + Chr$(13) _
                + "Icom CI-V Set VFO-B", 10)
        End If
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
    IC706RadioSetVFOB = resp
End Function
'Set Icom Radios to Memory mode
Sub IC821RadioMem(Address%, Bidir%, Handle%)
    InString = ReadFromPort(Handle%)
    'Send bytes to CI-V. Begin with two syncs (FE) then the
    'destination address, then the origin address (E0 is
    'always the master), then the command code and
    'its data and finally the end code (FD)
     Call WriteToPort(Chr$(&HFE) + Chr$(&HFE) + _
        Chr$(Address%) + Chr$(&HE0) + Chr$(&H8) + _
        Chr$(&HFD), Handle%)
    Call WaitOutBuffEmpty(Handle%)
    If Bidir% Then
        'wait up to 5 times timeout-time for a frame correctly
        'addressed to bus master
        'this filters-out echoed frames.
        b = Timer
        Do
            a$ = IC821RadioReadFrame(Handle%)
            X$ = X$ + a$
        Loop Until (Cdbl2(PickWord(a$, 1)) = 1) Or (Abs(Timer - b) > RadioReplyTimeout(Handle%) * 5)
        If Cdbl2(PickWord(a$, 1)) <> 1 Then
            Call frmMessage.ShowMessage("Comms. Error during" + Chr$(13) _
                + "Icom CI-V Set Mem", 10)
        End If
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub

Sub FT847RadioCATOn(Handle%)
    InString = ReadFromPort(Handle%)
    'Send 5 bytes to Yaesu.
    Call WriteToPort(Chr$(&H0) + Chr$(&H0) + _
        Chr$(&H0) + Chr$(&H0) + Chr$(&H0), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub

Sub FT847RadioCATOff(Handle%)
    InString = ReadFromPort(Handle%)
    'Send 5 bytes to Yaesu.
    Call WriteToPort(Chr$(&H0) + Chr$(&H0) + _
        Chr$(&H0) + Chr$(&H0) + Chr$(&H80), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub

Sub FT847RadioSatOn(Handle%)
    InString = ReadFromPort(Handle%)
    'Send 5 bytes to Yaesu.
    Call WriteToPort(Chr$(&H0) + Chr$(&H0) + _
        Chr$(&H0) + Chr$(&H0) + Chr$(&H4E), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub

Sub FT847RadioSatOff(Handle%)
    InString = ReadFromPort(Handle%)
    'Send 5 bytes to Yaesu.
    Call WriteToPort(Chr$(&H0) + Chr$(&H0) + _
        Chr$(&H0) + Chr$(&H0) + Chr$(&H8E), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
Sub FT847RadioShiftPlus(Handle%)
    InString = ReadFromPort(Handle%)
    'Send 5 bytes to Yaesu.
    Call WriteToPort(Chr$(&H49) + Chr$(&H0) + _
        Chr$(&H0) + Chr$(&H0) + Chr$(&H9), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
Sub FT847RadioShiftMinus(Handle%)
    InString = ReadFromPort(Handle%)
    'Send 5 bytes to Yaesu.
    Call WriteToPort(Chr$(&H9) + Chr$(&H0) + _
        Chr$(&H0) + Chr$(&H0) + Chr$(&H9), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
Sub FT847RadioShiftOff(Handle%)
    InString = ReadFromPort(Handle%)
    'Send 5 bytes to Yaesu.
    Call WriteToPort(Chr$(&H89) + Chr$(&H0) + _
        Chr$(&H0) + Chr$(&H0) + Chr$(&H9), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
Sub FT847RadioSetShiftFreq(Freq&, Handle%)
    InString = ReadFromPort(Handle%)
    'Send 4 bytes to Yaesu
    For f% = 1 To 4
        n% = Fix(Freq& / 100000000)
        Freq& = Freq& Mod 100000000
    
        n2% = Fix(Freq& / 10000000)
        Freq& = Freq& Mod 10000000
    
        Freq& = Freq& * 100
    
        Call WriteToPort(Chr$(16 * n% + n2%), Handle%)
    Next
    'and the 5th.:
    Call WriteToPort(Chr$(&HF9), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
Sub FT847RadioCTCSSRXOff(Handle%)
    InString = ReadFromPort(Handle%)
    'Send 5 bytes to Yaesu.
    Call WriteToPort(Chr$(&H8A) + Chr$(&H0) + _
        Chr$(&H0) + Chr$(&H0) + Chr$(&H1A), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
Sub FT847RadioCTCSSTXOff(Handle%)
    InString = ReadFromPort(Handle%)
    'Send 5 bytes to Yaesu.
    Call WriteToPort(Chr$(&H8A) + Chr$(&H0) + _
        Chr$(&H0) + Chr$(&H0) + Chr$(&H2A), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
'Add PA2EON
Sub FT847RadioCTCSSTXOn(Handle%)
    InString = ReadFromPort(Handle%)
    'Send 5 bytes to Yaesu.
    Call WriteToPort(Chr$(&H4A) + Chr$(&H0) + _
        Chr$(&H0) + Chr$(&H0) + Chr$(&H2A), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub


Sub FT847RadioSetRXMode(mode$, Handle%)
    Select Case LCase$(mode$)
        Case Is = "fm-n"
            m% = &H88
        Case Is = "fm", "fm-w"
            m% = &H8
        Case Is = "lsb"
            m% = &H0
        Case Is = "usb"
            m% = &H1
        Case Is = "cw"
            m% = &H2
        Case Is = "cw-n"
            m% = &H82
        Case Else
            m% = -1
    End Select
    If m% >= 0 Then
        InString = ReadFromPort(Handle%)
        'Send 5 bytes to Yaesu.
         Call WriteToPort(Chr$(m%) + Chr$(&H0) + _
            Chr$(&H0) + Chr$(&H0) + Chr$(&H17), Handle%)
        InString = ReadFromPort(Handle%)
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub

Sub FT847RadioSetTXMode(mode$, Handle%)
    Select Case LCase$(mode$)
        Case Is = "fm-n"
            m% = &H88
        Case Is = "fm", "fm-w"
            m% = &H8
        Case Is = "lsb"
            m% = &H0
        Case Is = "usb"
            m% = &H1
        Case Is = "cw"
            m% = &H2
        Case Is = "cw-n"
            m% = &H82
        Case Else
            m% = -1
    End Select
    If m% >= 0 Then
        InString = ReadFromPort(Handle%)
        'Send 5 bytes to Yaesu.
         Call WriteToPort(Chr$(m%) + Chr$(&H0) + _
            Chr$(&H0) + Chr$(&H0) + Chr$(&H27), Handle%)
        InString = ReadFromPort(Handle%)
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
Sub FT847RadioSetMainMode(mode$, Handle%)
    Select Case LCase$(mode$)
        Case Is = "fm-n"
            m% = &H88
        Case Is = "fm", "fm-w"
            m% = &H8
        Case Is = "lsb"
            m% = &H0
        Case Is = "usb"
            m% = &H1
        Case Is = "cw"
            m% = &H2
        Case Is = "cw-n"
            m% = &H82
        Case Else
            m% = -1
    End Select
    If m% >= 0 Then
        InString = ReadFromPort(Handle%)
        'Send 5 bytes to Yaesu.
         Call WriteToPort(Chr$(m%) + Chr$(&H0) + _
            Chr$(&H0) + Chr$(&H0) + Chr$(&H7), Handle%)
        InString = ReadFromPort(Handle%)
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub

Sub FT847RadioSetRXFreq(Freq&, Handle%)
    InString = ReadFromPort(Handle%)
    'Send 4 bytes to Yaesu
    For f% = 1 To 4
        n% = Fix(Freq& / 100000000)
        Freq& = Freq& Mod 100000000
    
        n2% = Fix(Freq& / 10000000)
        Freq& = Freq& Mod 10000000
    
        Freq& = Freq& * 100
    
        Call WriteToPort(Chr$(16 * n% + n2%), Handle%)
    Next
    'and the 5th.:
    Call WriteToPort(Chr$(&H11), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub

Sub FT847RadioSetTXFreq(Freq&, Handle%)
    InString = ReadFromPort(Handle%)
    'Send 4 bytes to Yaesu
    For f% = 1 To 4
        n% = Fix(Freq& / 100000000)
        Freq& = Freq& Mod 100000000
    
        n2% = Fix(Freq& / 10000000)
        Freq& = Freq& Mod 10000000
    
        Freq& = Freq& * 100
    
        Call WriteToPort(Chr$(16 * n% + n2%), Handle%)
    Next
    'and the 5th.:
    Call WriteToPort(Chr$(&H21), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
Sub FT847RadioSetMainFreq(Freq&, Handle%)
    InString = ReadFromPort(Handle%)
    'Send 4 bytes to Yaesu
    For f% = 1 To 4
        n% = Fix(Freq& / 100000000)
        Freq& = Freq& Mod 100000000
    
        n2% = Fix(Freq& / 10000000)
        Freq& = Freq& Mod 10000000
    
        Freq& = Freq& * 100
    
        Call WriteToPort(Chr$(16 * n% + n2%), Handle%)
    Next
    'and the 5th.:
    Call WriteToPort(Chr$(&H1), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
'FT-847 frame readback routine
'waits up to 200mS for 5 bytes to come from the radio.
'Valid frames are 5 bytes long only.
'**To read RSSI of FT847, VR5000 routines must be used.**
'Parses the complete frame for valid data.
'The output of this function is a string beginning with the decoded
'frame type code followed by the data available separated by spaces.
'the following type codes can be returned: (freqs. in Hz)
'code "00:" - frame not understood pass timeout period.
'code "1020:" - frequency and mode information (can be of main
'               VFO, sat RX VFO or sat TX VFO depending on query
'               command sent previously)
Function FT847RadioReadFrame(Handle%) As String
'This will hold the received frame
Dim InBuff(10) As Integer
'initialize buffer pointer..
InBuffPtr% = 0
FrameSize% = 5
'set start time...
a = Timer
Do
    'Wait until we receive bytes
    'or a time-out occurs
    Select Case Handle%
    Case 1
        Do
            
        Loop Until (MSComm1.InBufferCount >= 1) Or Abs(Timer - a) > RadioReplyTimeout(Handle%)
        InString = MSComm1.Input
    Case 2
        Do
            
        Loop Until (MSComm2.InBufferCount >= 1) Or Abs(Timer - a) > RadioReplyTimeout(Handle%)
        InString = MSComm2.Input
    Case 3
        Do
            
        Loop Until (MSComm3.InBufferCount >= 1) Or Abs(Timer - a) > RadioReplyTimeout(Handle%)
        InString = MSComm3.Input
    End Select
    'We pass the received bytes to the buffer
    For i% = 0 To LenB(InString) - 1
        InBuff(InBuffPtr%) = InString(i%)
        InBuffPtr% = InBuffPtr% + 1
    Next
    'if a byte was received we reset timer to wait for another
    '200mS
    If Abs(Timer - a) < RadioReplyTimeout(Handle%) Then
        a = Timer
    End If
Loop While (Abs(Timer - a) < RadioReplyTimeout(Handle%)) And (InBuffPtr% <> FrameSize%)
'if answer is not 5 bytes :
If InBuffPtr% <> 5 Then
    s$ = "00:"
    FT847RadioReadFrame = s$
    Exit Function
Else
    'otherwise examine frame:
    'we are sure that the frame has exactly 5 bytes so
    'perform a basic integrity check of frame:
    'modes between 0x09 and 0x81 are invalid
    If (InBuff(4) > &H8) And (InBuff(4) < &H82) Then
        s$ = "00:"
        FT847RadioReadFrame = s$
        Exit Function
    End If
    'if it seems correct...
    'read frequency:
    s$ = "1020: "
    edge& = 0
    For f% = 0 To 3
        edge& = edge& * 10
        edge& = edge& + (InBuff(f%) And &HF0) / 16
        edge& = edge& * 10
        edge& = edge& + (InBuff(f%) And &HF)
    Next
    edge& = edge& * 10
    s$ = s$ + Str$(edge&) + " "
    'read mode:
    Select Case InBuff(4)
        Case 0
        s$ = s$ + "LSB"
        Case 1
        s$ = s$ + "USB"
        Case 2
        s$ = s$ + "CW"
        Case 3
        s$ = s$ + "CW-R"
        Case 4
        s$ = s$ + "AM"
        Case 8
        s$ = s$ + "FM"
        Case &H82
        s$ = s$ + "CW-N"
        Case &H83
        s$ = s$ + "CW-NR"
        Case &H84
        s$ = s$ + "AM-N"
        Case &H88
        s$ = s$ + "FM-N"
    End Select
    FT847RadioReadFrame = s$
End If
End Function
'FT-847: Get the frequency of RX VFO
Function FT847RadioReadRXFreq(Handle%) As Double
InString = ReadFromPort(Handle%)
'Send sat RX VFO query command:
Call WriteToPort(Chr$(&H0) + Chr$(&H0) + _
    Chr$(&H0) + Chr$(&H0) + _
    Chr$(&H13), Handle%)
a$ = FT847RadioReadFrame(Handle%)
'check frame code 1020
If Cdbl2(PickWord(a$, 1)) <> 1020 Then
'    Call frmMessage.ShowMessage("Error during" + Chr$(13) _
        + "Yaesu Read Frequency", 10)
    FT847RadioReadRXFreq = 0
Else
    FT847RadioReadRXFreq = Cdbl2(PickWord(a$, 2))
End If
End Function
'FT-847: Get the frequency of TX VFO
Function FT847RadioReadTXFreq(Handle%) As Double
InString = ReadFromPort(Handle%)
'Send sat RX VFO query command:
Call WriteToPort(Chr$(&H0) + Chr$(&H0) + _
    Chr$(&H0) + Chr$(&H0) + _
    Chr$(&H23), Handle%)
a$ = FT847RadioReadFrame(Handle%)
'check frame code 1020
If Cdbl2(PickWord(a$, 1)) <> 1020 Then
'    Call frmMessage.ShowMessage("Error during" + Chr$(13) _
        + "Yaesu Read Frequency", 10)
    FT847RadioReadTXFreq = 0
Else
    FT847RadioReadTXFreq = Cdbl2(PickWord(a$, 2))
End If
End Function
'FT-847: Get the frequency of Main VFO
Function FT847RadioReadMainFreq(Handle%) As Double
InString = ReadFromPort(Handle%)
'Send sat RX VFO query command:
Call WriteToPort(Chr$(&H0) + Chr$(&H0) + _
    Chr$(&H0) + Chr$(&H0) + _
    Chr$(&H3), Handle%)
a$ = FT847RadioReadFrame(Handle%)
'check frame code 1020
If Cdbl2(PickWord(a$, 1)) <> 1020 Then
'    Call frmMessage.ShowMessage("Error during" + Chr$(13) _
        + "Yaesu Read Frequency", 10)
    FT847RadioReadMainFreq = 0
Else
    FT847RadioReadMainFreq = Cdbl2(PickWord(a$, 2))
End If
End Function

'FT736.... procedures are for FT-736 CAT. Only those that
'differ from FT-847's version are rewrote.
Sub FT736RadioSatOn(Handle%)
    InString = ReadFromPort(Handle%)
    'Send 5 bytes to Yaesu.
    Call WriteToPort(Chr$(&H0) + Chr$(&H0) + _
        Chr$(&H0) + Chr$(&H0) + Chr$(&HE), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
'Yaesu 736 RX freq setting
Sub FT736RadioSetRXFreq(Freq&, Handle%)
    InString = ReadFromPort(Handle%)
    'Send 4 bytes to Yaesu
    For f% = 1 To 4
        n% = Fix(Freq& / 100000000)
        Freq& = Freq& Mod 100000000
        
        n2% = Fix(Freq& / 10000000)
        Freq& = Freq& Mod 10000000
    
        Freq& = Freq& * 100
    
        Call WriteToPort(Chr$(16 * n% + n2%), Handle%)
    Next
    'and the 5th.:
    Call WriteToPort(Chr$(&H1E), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
'Yaesu 736 TX freq setting.
Sub FT736RadioSetTXFreq(Freq&, Handle%)
    InString = ReadFromPort(Handle%)
    'Send 4 bytes to Yaesu
    For f% = 1 To 4
        n% = Fix(Freq& / 100000000)
        Freq& = Freq& Mod 100000000
    
        n2% = Fix(Freq& / 10000000)
        Freq& = Freq& Mod 10000000
    
        Freq& = Freq& * 100
    
        Call WriteToPort(Chr$(16 * n% + n2%), Handle%)
    Next
    'and the 5th.:
    Call WriteToPort(Chr$(&H2E), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
Sub FT100RadioSetVFOA(Handle%) ' G6LVB 16 Dec 2000 FT-100
    InString = ReadFromPort(Handle%)
    'Send 5 bytes to Yaesu.
    a$ = Chr$(&H0) + Chr$(&H0) + _
        Chr$(&H0) + Chr$(&H0) + Chr$(&H5)

    ' Check if command logging is active
    If frmRadio.CheckLogCom.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Setting FT-100 VFO A, sending:" + StrToHex(a$)
        Close f%
    End If
    
    Call WriteToPort(a$, Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
Sub FT100RadioSetVFOB(Handle%)
    InString = ReadFromPort(Handle%)
    'Send 5 bytes to Yaesu.
    a$ = Chr$(&H0) + Chr$(&H0) + _
        Chr$(&H0) + Chr$(&H1) + Chr$(&H5)

    ' Check if command logging is active
    If frmRadio.CheckLogCom.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Setting FT-100 VFO B, sending:" + StrToHex(a$)
        Close f%
    End If
    
    Call WriteToPort(a$, Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub


Sub FT100RadioSetCTCSSOff(Handle%) ' G6LVB 16 Dec 2000 FT-100
    InString = ReadFromPort(Handle%)
    'Send 5 bytes to Yaesu.
    a$ = Chr$(&H0) + Chr$(&H0) + _
        Chr$(&H0) + Chr$(&H0) + Chr$(&H92)
    
    ' Check if command logging is active
    If frmRadio.CheckLogCom.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Setting FT-100 CTCSS Off, sending:" + StrToHex(a$)
        Close f%
    End If
    
    Call WriteToPort(a$, Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub

Sub FT100RadioSetSplitOff(Handle%) ' G6LVB 16 Dec 2000 FT-100
    InString = ReadFromPort(Handle%)
    'Send 5 bytes to Yaesu.
    a$ = Chr$(&H0) + Chr$(&H0) + _
        Chr$(&H0) + Chr$(&H0) + Chr$(&H1)

    ' Check if command logging is active
    If frmRadio.CheckLogCom.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Setting FT-100 Rpt Off, sending:" + StrToHex(a$)
        Close f%
    End If

    Call WriteToPort(a$, Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
Sub FT100RadioSetSplitOn(Handle%)
    InString = ReadFromPort(Handle%)
    'Send 5 bytes to Yaesu.
    a$ = Chr$(&H0) + Chr$(&H0) + _
        Chr$(&H0) + Chr$(&H1) + Chr$(&H1)

    ' Check if command logging is active
    If frmRadio.CheckLogCom.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Setting FT-100 Split Off, sending:" + StrToHex(a$)
        Close f%
    End If
    
    Call WriteToPort(a$, Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub

Sub FT100RadioSetRptOff(Handle%) ' G6LVB 16 Dec 2000 FT-100
    InString = ReadFromPort(Handle%)
    'Send 5 bytes to Yaesu.
    a$ = Chr$(&H0) + Chr$(&H0) + _
        Chr$(&H0) + Chr$(&H0) + Chr$(&H84)

    ' Check if command logging is active
    If frmRadio.CheckLogCom.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Setting FT-100 Rpt Off, sending:" + StrToHex(a$)
        Close f%
    End If
    
    Call WriteToPort(a$, Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
'FT-100 Status-Update frame readback routine
'waits up to 200mS for 16 bytes to come from the radio.
'Valid frames are 16 bytes long only.
'**To read RSSI of FT847, VR5000 routines must be used.**
'Parses the complete frame for valid data.
'The output of this function is a string beginning with the decoded
'frame type code followed by the data available separated by spaces.
'the following type codes can be returned: (freqs. in Hz)
'code "00:" - frame not understood pass timeout period.
'code "1020:" - frequency and mode information
'
Function FT100RadioReadStatusFrame(Handle%) As String
'This will hold the received frame
Dim InBuff(20) As Integer
'initialize buffer pointer..
InBuffPtr% = 0
'Frame size validation
FrameSize% = 16
'set start time...
a = Timer
'RadioControlDelay(Handle%) will be fixed at 200mS as this time is part of
'Yaesu protocol specification
Do
    'Wait until we receive bytes
    'or a time-out occurs
    Select Case Handle%
    Case 1
        Do
            
        Loop Until (MSComm1.InBufferCount >= 1) Or Abs(Timer - a) > RadioReplyTimeout(Handle%)
        InString = MSComm1.Input
    Case 2
        Do
            
        Loop Until (MSComm2.InBufferCount >= 1) Or Abs(Timer - a) > RadioReplyTimeout(Handle%)
        InString = MSComm2.Input
    Case 3
        Do
            
        Loop Until (MSComm3.InBufferCount >= 1) Or Abs(Timer - a) > RadioReplyTimeout(Handle%)
        InString = MSComm3.Input
    End Select
    'We pass the received bytes to the buffer
    For i% = 0 To LenB(InString) - 1
        InBuff(InBuffPtr%) = InString(i%)
        InBuffPtr% = InBuffPtr% + 1
    Next
    'if a byte was received we reset timer to wait for another
    '200mS
    If Abs(Timer - a) < RadioReplyTimeout(Handle%) Then
        a = Timer
    End If
Loop While (Abs(Timer - a) < RadioReplyTimeout(Handle%)) And (InBuffPtr% <> FrameSize%)

' Check if command logging is active
If frmRadio.CheckLogCom.Value Then
    f% = FreeFile
    Open "Radio_Log.txt" For Append As f%
    Print #f%, "Reading FT-100 status frame, received:" + ArrayToHex(InBuff, InBuffPtr% - 1)
    Close f%
End If

'if answer is not 16 bytes :
If InBuffPtr% <> 16 Then
    s$ = "00:"
    FT100RadioReadStatusFrame = s$
    Exit Function
Else
    'otherwise examine frame:
    'we are sure that the frame has exactly 16 bytes so
    'if it seems correct...
    'read frequency:
    s$ = "1020: "
    edge& = 0
    For f% = 1 To 4
        edge& = edge& * 256
        edge& = edge& + InBuff(f%)
    Next
    edge& = edge& * 1.25
    s$ = s$ + Str$(edge&) + " "
    'read mode:
    'mode info is in 4 LSbits of 6th sent byte.
    Select Case (InBuff(5) And 15)
        Case 0
        s$ = s$ + "LSB"
        Case 1
        s$ = s$ + "USB"
        Case 2
        s$ = s$ + "CW"
        Case 3
        s$ = s$ + "CW-R"
        Case 4
        s$ = s$ + "AM"
        Case 5
        s$ = s$ + "DIG"
        Case 6
        s$ = s$ + "FM"
        Case 7
        s$ = s$ + "FM-W"
    End Select
    FT100RadioReadStatusFrame = s$
End If
End Function
'FT-100: Get the frequency of active VFO
Function FT100RadioReadFreq(Handle%) As Double
InString = ReadFromPort(Handle%)
'Send Status Update query command:
a$ = Chr$(&H0) + Chr$(&H0) + _
    Chr$(&H0) + Chr$(&H0) + _
    Chr$(&H10)

' Check if command logging is active
If frmRadio.CheckLogCom.Value Then
    f% = FreeFile
    Open "Radio_Log.txt" For Append As f%
    Print #f%, "Reading FT-100 freq., sending:" + StrToHex(a$)
    Close f%
End If

Call WriteToPort(a$, Handle%)
a$ = FT100RadioReadStatusFrame(Handle%)
'check frame code 1020
If Cdbl2(PickWord(a$, 1)) <> 1020 Then
'    Call frmMessage.ShowMessage("Error during" + Chr$(13) _
        + "Yaesu Read Frequency", 10)
    FT100RadioReadFreq = 0
Else
    FT100RadioReadFreq = Cdbl2(PickWord(a$, 2))
End If
End Function
Sub FT817RadioSetCTCSSOff(Handle%) ' G6LVB 16 Dec 2000 FT-817
    InString = ReadFromPort(Handle%)
    'Send 5 bytes to Yaesu.
    a$ = Chr$(&H8A) + Chr$(&H0) + _
        Chr$(&H0) + Chr$(&H0) + Chr$(&HA)
    
    ' Check if command logging is active
    If frmRadio.CheckLogCom.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Setting Off FT-817 CTCSS, sending:" + StrToHex(a$)
        Close f%
    End If
        
    Call WriteToPort(a$, Handle%)
    
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
Sub FT817RadioToggleVFO(Handle%)
    InString = ReadFromPort(Handle%)
    'Send 5 bytes to Yaesu.
    a$ = Chr$(&H0) + Chr$(&H0) + _
        Chr$(&H0) + Chr$(&H0) + Chr$(&H81)
    
    ' Check if command logging is active
    If frmRadio.CheckLogCom.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Toggling FT-817 VFO, sending:" + StrToHex(a$)
        Close f%
    End If
    
    Call WriteToPort(a$, Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
Sub FT817RadioWakeUp(Handle%)
    InString = ReadFromPort(Handle%)
    'Send 5 bytes to Yaesu.
    a$ = Chr$(&HFF) + Chr$(&HFF) + _
        Chr$(&HFF) + Chr$(&HFF) + Chr$(&HFF)
    
    ' Check if command logging is active
    If frmRadio.CheckLogCom.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Waking-Up FT-817, sending:" + StrToHex(a$)
        Close f%
    End If
    
    Call WriteToPort(a$, Handle%)
    
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    
    'Send 5 bytes to Yaesu.
    a$ = Chr$(&H0) + Chr$(&H0) + _
        Chr$(&H0) + Chr$(&H0) + Chr$(&HF)
    
    ' Check if command logging is active
    If frmRadio.CheckLogCom.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Waking-Up FT-817, sending:" + StrToHex(a$)
        Close f%
    End If
    
    Call WriteToPort(a$, Handle%)
    
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub

Sub FT817RadioSetSplitOff(Handle%) ' G6LVB 16 Dec 2000 FT-817
    InString = ReadFromPort(Handle%)
    'Send 5 bytes to Yaesu.
    a$ = Chr$(&H0) + Chr$(&H0) + _
        Chr$(&H0) + Chr$(&H0) + Chr$(&H82)
    
    ' Check if command logging is active
    If frmRadio.CheckLogCom.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Setting FT-817 Split Off, sending:" + StrToHex(a$)
        Close f%
    End If
    
    Call WriteToPort(a$, Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
Sub FT817RadioSetSplitOn(Handle%)
    InString = ReadFromPort(Handle%)
    'Send 5 bytes to Yaesu.
    a$ = Chr$(&H0) + Chr$(&H0) + _
        Chr$(&H0) + Chr$(&H0) + Chr$(&H2)
    
    ' Check if command logging is active
    If frmRadio.CheckLogCom.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Setting FT-817 Split On, sending:" + StrToHex(a$)
        Close f%
    End If
    
    Call WriteToPort(a$, Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub

Sub FT817RadioSetRptOff(Handle%) ' G6LVB 16 Dec 2000 FT-817
    InString = ReadFromPort(Handle%)
    'Send 5 bytes to Yaesu.
    a$ = Chr$(&H89) + Chr$(&H0) + _
        Chr$(&H0) + Chr$(&H0) + Chr$(&H9)
    
    ' Check if command logging is active
    If frmRadio.CheckLogCom.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Setting FT-817 Rpt Off, sending:" + StrToHex(a$)
        Close f%
    End If
    
    Call WriteToPort(a$, Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
Sub FT100RadioSetMode(mode$, Handle%) ' G6LVB 16 Dec 2000 FT-100
    Select Case LCase$(mode$)
        Case Is = "fm-n"
            m% = &H6
        Case Is = "fm", "fm-w"
            m% = &H7
        Case Is = "lsb"
            m% = &H0
        Case Is = "usb"
            m% = &H1
        Case Is = "cw"
            m% = &H2
        Case Is = "cw-n"
            m% = &H3
        Case Else
            m% = -1
    End Select
    If m% >= 0 Then
        InString = ReadFromPort(Handle%)
        'Send 5 bytes to Yaesu.
         a$ = Chr$(&H0) + _
            Chr$(&H0) + Chr$(&H0) + Chr$(m%) + Chr$(&HC)
    
        ' Check if command logging is active
        If frmRadio.CheckLogCom.Value Then
            f% = FreeFile
            Open "Radio_Log.txt" For Append As f%
            Print #f%, "Setting FT-100 Mode to " + mode$ + ", sending:" + StrToHex(a$)
            Close f%
        End If
    
         Call WriteToPort(a$, Handle%)
        InString = ReadFromPort(Handle%)
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub

Sub FT817RadioSetMode(mode$, Handle%) 'G6LVB 16 Dec 2000 FT-817
    Select Case LCase$(mode$)
        Case Is = "fm-n"
            m% = &H8
        Case Is = "fm", "fm-w"
            m% = &H8
        Case Is = "lsb"
            m% = &H0
        Case Is = "usb"
            m% = &H1
        Case Is = "cw"
            m% = &H2
        Case Is = "cw-n"
            m% = &H3
        Case Else
            m% = -1
    End Select
    If m% >= 0 Then
        InString = ReadFromPort(Handle%)
        'Send 5 bytes to Yaesu.
         a$ = Chr$(m%) + Chr$(&H0) + _
            Chr$(&H0) + Chr$(&H0) + Chr$(&H7)
        
        ' Check if command logging is active
        If frmRadio.CheckLogCom.Value Then
            f% = FreeFile
            Open "Radio_Log.txt" For Append As f%
            Print #f%, "Setting FT-817 Mode to " + mode$ + ", sending:" + StrToHex(a$)
            Close f%
        End If
            
         Call WriteToPort(a$, Handle%)
        
        InString = ReadFromPort(Handle%)
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
Sub FT897RadioSetMode(mode$, Handle%)
    Select Case LCase$(mode$)
        Case Is = "fm-n"
            m% = &H88
        Case Is = "fm", "fm-w"
            m% = &H8
        Case Is = "lsb"
            m% = &H0
        Case Is = "usb"
            m% = &H1
        Case Is = "cw"
            m% = &H2
        Case Is = "cw-n"
            m% = &H3
        Case Else
            m% = -1
    End Select
    If m% >= 0 Then
        InString = ReadFromPort(Handle%)
        'Send 5 bytes to Yaesu.
         Call WriteToPort(Chr$(m%) + Chr$(&H0) + _
            Chr$(&H0) + Chr$(&H0) + Chr$(&H7), Handle%)
        InString = ReadFromPort(Handle%)
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
'FT-817: Get the frequency of active VFO
Function FT817RadioReadFreq(Handle%) As Double
InString = ReadFromPort(Handle%)
'Send freq query command:
a$ = (Chr$(&H0) + Chr$(&H0) + _
    Chr$(&H0) + Chr$(&H0) + _
    Chr$(&H3))

' Check if command logging is active
If frmRadio.CheckLogCom.Value Then
    f% = FreeFile
    Open "Radio_Log.txt" For Append As f%
    Print #f%, "Reading FT-817 freq., sending:" + StrToHex(a$)
    Close f%
End If

Call WriteToPort(a$, Handle%)

a$ = FT817RadioReadStatusFrame(Handle%)
'check frame code 1020
If Cdbl2(PickWord(a$, 1)) <> 1020 Then
'    Call frmMessage.ShowMessage("Error during" + Chr$(13) _
        + "Yaesu Read Frequency", 10)
    FT817RadioReadFreq = 0
Else
    FT817RadioReadFreq = Cdbl2(PickWord(a$, 2))
End If
End Function
'FT-817: Get the PTT Status
Function FT817RadioReadPTT(Handle%) As String
InString = ReadFromPort(Handle%)
'Send TX status query command:
a$ = (Chr$(&H0) + Chr$(&H0) + _
    Chr$(&H0) + Chr$(&H0) + _
    Chr$(&HF7))

' Check if command logging is active
If frmRadio.CheckLogCom.Value Then
    f% = FreeFile
    Open "Radio_Log.txt" For Append As f%
    Print #f%, "Reading FT-817 PTT status, sending:" + StrToHex(a$)
    Close f%
End If

Call WriteToPort(a$, Handle%)

a$ = FT817RadioReadTXStatusFrame(Handle%)

If Cdbl2(PickWord(a$, 1)) = 30 Then
    FT817RadioReadPTT = PickWord(a$, 2)
ElseIf Cdbl2(PickWord(a$, 1)) = 2 Then
    FT817RadioReadPTT = "OFF"
Else
'    Call frmMessage.ShowMessage("Error during" + Chr$(13) _
        + "Yaesu Read Frequency", 10)
    FT817RadioReadPTT = ""
End If
End Function
Sub FT100RadioSetFreq(Freq&, Handle%) ' G6LVB 16 Dec 2000 FT-100
    InString = ReadFromPort(Handle%)
    'Send 4 bytes to Yaesu
    'G6LVB 16 Dec 2000 FT-100 does bytes the other way around to other Yaesus
    Freq& = Freq& / 10

    For f% = 1 To 4
        n2% = Fix(Freq& Mod 10)
        Freq& = Fix(Freq& / 10)
        n% = Fix(Freq& Mod 10)
        Freq& = Fix(Freq& / 10)
        
        a$ = a$ + Chr$(16 * n% + n2%)
    Next
    'and the 5th.:
    a$ = a$ + Chr$(&HA)
    
    ' Check if command logging is active
    If frmRadio.CheckLogCom.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Setting FT-100 freq., sending:" + StrToHex(a$)
        Close f%
    End If
    
    Call WriteToPort(a$, Handle%)
    InString = ReadFromPort(Handle%)
    
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub

Sub FT817RadioSetFreq(Freq&, Handle%) ' G6LVB 16 Dec 2000 FT-817
    
    b$ = Str(Freq&)
    InString = ReadFromPort(Handle%)
    'Send 4 bytes to Yaesu
    For f% = 1 To 4
        n% = Fix(Freq& / 100000000)
        Freq& = Freq& Mod 100000000
    
        n2% = Fix(Freq& / 10000000)
        Freq& = Freq& Mod 10000000
    
        Freq& = Freq& * 100
    
        a$ = a$ + Chr$(16 * n% + n2%)
    Next
    'and the 5th.:
    a$ = a$ + Chr$(&H1)
    
    ' Check if command logging is active
    If frmRadio.CheckLogCom.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Setting FT-817 freq. to" + b$ + ", sending:" + StrToHex(a$)
        Close f%
    End If
    
    Call WriteToPort(a$, Handle%)
    InString = ReadFromPort(Handle%)
    
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
'FT-817 Status-Update frame readback routine
'waits up to 200mS for 5 bytes to come from the radio.
'Valid frames are 5 bytes long only.
'Parses the complete frame for valid data.
'The output of this function is a string beginning with the decoded
'frame type code followed by the data available separated by spaces.
'the following type codes can be returned: (freqs. in Hz)
'code "00:" - frame not understood pass timeout period.
'code "1020:" - frequency and mode information
'
Function FT817RadioReadStatusFrame(Handle%) As String
'This will hold the received frame
Dim InBuff(20) As Integer
'initialize buffer pointer..
InBuffPtr% = 0
'Frame size validation
FrameSize% = 5
'set start time...
a = Timer
'RadioControlDelay(Handle%) will be fixed at 200mS as this time is part of
'Yaesu protocol specification
Do
    'Wait until we receive bytes
    'or a time-out occurs
    Select Case Handle%
    Case 1
        Do
            
        Loop Until (MSComm1.InBufferCount >= 1) Or Abs(Timer - a) > RadioReplyTimeout(Handle%)
        InString = MSComm1.Input
    Case 2
        Do
            
        Loop Until (MSComm2.InBufferCount >= 1) Or Abs(Timer - a) > RadioReplyTimeout(Handle%)
        InString = MSComm2.Input
    Case 3
        Do
            
        Loop Until (MSComm3.InBufferCount >= 1) Or Abs(Timer - a) > RadioReplyTimeout(Handle%)
        InString = MSComm3.Input
    End Select
    'We pass the received bytes to the buffer
    For i% = 0 To LenB(InString) - 1
        InBuff(InBuffPtr%) = InString(i%)
        InBuffPtr% = InBuffPtr% + 1
    Next
    'if a byte was received we reset timer to wait for another
    '200mS
    If Abs(Timer - a) < RadioReplyTimeout(Handle%) Then
        a = Timer
    End If
Loop While (Abs(Timer - a) < RadioReplyTimeout(Handle%)) And (InBuffPtr% <> FrameSize%)

' Check if command logging is active
If frmRadio.CheckLogCom.Value Then
    f% = FreeFile
    Open "Radio_Log.txt" For Append As f%
    Print #f%, "Reading FT-817 status frame, received:" + ArrayToHex(InBuff, InBuffPtr% - 1)
    Close f%
End If

'if answer is not 5 bytes :
If InBuffPtr% <> 5 Then
    s$ = "00:"
    FT817RadioReadStatusFrame = s$
    Exit Function
Else
    'otherwise examine frame:
    'we are sure that the frame has exactly 5 bytes so
    'if it seems correct...
    'read frequency:
    s$ = "1020: "
    edge& = 0
    For f% = 0 To 3
        edge& = edge& * 10
        edge& = edge& + (InBuff(f%) And &HF0) / 16
        edge& = edge& * 10
        edge& = edge& + (InBuff(f%) And &HF)
    Next
    edge& = edge& * 10
    s$ = s$ + Str$(edge&) + " "
    'read mode:
    Select Case InBuff(4)
        Case 0
        s$ = s$ + "LSB"
        Case 1
        s$ = s$ + "USB"
        Case 2
        s$ = s$ + "CW"
        Case 3
        s$ = s$ + "CW-R"
        Case 4
        s$ = s$ + "AM"
        Case 6
        s$ = s$ + "FM-W"
        Case 8
        s$ = s$ + "FM"
        Case &HA
        s$ = s$ + "DIG"
        Case &HC
        s$ = s$ + "PKT"
    End Select
    FT817RadioReadStatusFrame = s$
End If
End Function
'FT-817 TX-Status-Update frame readback routine
'waits up to 200mS for 1 bytes to come from the radio.
'Valid frames are 1 bytes long only.
'Parses the complete frame for valid data.
'The output of this function is a string beginning with the decoded
'frame type code followed by the data available separated by spaces.
'the following type codes can be returned: (freqs. in Hz)
'code "00:" - frame not understood pass timeout period.
'code "02:" - frame understood but not valid
'code "30:" - PTT Status
'
Function FT817RadioReadTXStatusFrame(Handle%) As String
'This will hold the received frame
Dim InBuff(20) As Integer
'initialize buffer pointer..
InBuffPtr% = 0
'Frame size validation
FrameSize% = 1
'set start time...
a = Timer
'RadioControlDelay(Handle%) will be fixed at 200mS as this time is part of
'Yaesu protocol specification
Do
    'Wait until we receive bytes
    'or a time-out occurs
    Select Case Handle%
    Case 1
        Do
            
        Loop Until (MSComm1.InBufferCount >= 1) Or Abs(Timer - a) > RadioReplyTimeout(Handle%)
        InString = MSComm1.Input
    Case 2
        Do
            
        Loop Until (MSComm2.InBufferCount >= 1) Or Abs(Timer - a) > RadioReplyTimeout(Handle%)
        InString = MSComm2.Input
    Case 3
        Do
            
        Loop Until (MSComm3.InBufferCount >= 1) Or Abs(Timer - a) > RadioReplyTimeout(Handle%)
        InString = MSComm3.Input
    End Select
    'We pass the received bytes to the buffer
    For i% = 0 To LenB(InString) - 1
        InBuff(InBuffPtr%) = InString(i%)
        InBuffPtr% = InBuffPtr% + 1
    Next
    'if a byte was received we reset timer to wait for another
    '200mS
    If Abs(Timer - a) < RadioReplyTimeout(Handle%) Then
        a = Timer
    End If
Loop While (Abs(Timer - a) < RadioReplyTimeout(Handle%)) And (InBuffPtr% <> FrameSize%)

' Check if command logging is active
If frmRadio.CheckLogCom.Value Then
    f% = FreeFile
    Open "Radio_Log.txt" For Append As f%
    Print #f%, "Reading FT-817 TX status frame, received:" + ArrayToHex(InBuff, InBuffPtr% - 1)
    Close f%
End If

'if answer is not the correct size :
If InBuffPtr% <> FrameSize% Then
    s$ = "00:"
    FT817RadioReadTXStatusFrame = s$
'    frmMain.Caption = "PTT Unknown"
    Exit Function
Else
    'otherwise examine frame:
    'we are sure that the frame has exactly the correct size so
    'if it seems correct...
    'read frequency:
    If InBuff(0) = &HFF Then
        'request not valid (most probably radio is on RX)
        s$ = "02:"
    ElseIf InBuff(0) And &H80 Then
        'if bit 7 is active, radio PTT is not Keyed
        s$ = "30: "
        s$ = s$ + "OFF "
    Else
        s$ = "30: "
        s$ = s$ + "ON "
    End If
    FT817RadioReadTXStatusFrame = s$
End If
End Function
'EB2CTA: Sets operating mode for AR-3000A receiver
Sub AR3000ARadioSetMode(mode$, Handle%)
    Select Case LCase$(mode$)
        Case Is = "fm-w"
            mAR$ = "W"
            m% = 0
        Case Is = "fm-n", "fm"
            mAR$ = "N"
            m% = 1
        Case Is = "am"
            mAR$ = "A"
            m% = 2
        Case Is = "usb"
            mAR$ = "U"
            m% = 3
        Case Is = "lsb"
            mAR$ = "L"
            m% = 4
        Case Is = "cw", "cw-n"
            mAR$ = "C"
            m% = 5
       Case Else
            mAR$ = ""
            m% = -1
    End Select
    If m% >= 0 Then
        InString = ReadFromPort(Handle%)
        Call WriteToPort(mAR$, Handle%)
        Call WriteToPort(Chr$(13), Handle%)
        InString = ReadFromPort(Handle%)
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
Sub AR3000ARadioSetFreq(Freq, Handle%) ' EB2CTA 5 Ago 2005 Changed Freq
    InString = ReadFromPort(Handle%)
    'format the frequency for the AR-3000A
    freqtext = Format(Cdbl2(Freq), "000000000") ' EB2CTA 5 Ago 2005 Changed Freq de 9 digitos
    freqtext1$ = Left(freqtext, 4) + "." + Right(freqtext, 5)
    Call WriteToPort(freqtext1$, Handle%) ' Envia Freq a RX
    Call WriteToPort(Chr$(13), Handle%) ' Envia <CR> a RX
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub

Sub AR3000ARadioOff(Handle%)
    'instring = ReadFromPort(Handle%)
    'Send VA to set VFO A of the AR-8000
    'the return, CR
    'Call WriteToPort("T" + Chr$(13), Handle%)
    'Send EX to exit form remote control from AR-8000
    'the return, CR
'    Call WriteToPort("EX" + Chr$(13), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub

Sub AR3000ARadioOn(Handle%)
    InString = ReadFromPort(Handle%)
    'Send the minimal step to AR-3000A
    'Send ATT off to AR-3000A
    'the return, CR
    Step& = 5
    freqtext = Format(Cdbl2(Step&), "00000")
    freqtext1$ = Left(freqtext, 3) + "." + Right(freqtext, 2) + "S"
    Call WriteToPort(freqtext1$, Handle%) ' Envia Step a RX
    Call WriteToPort(Chr$(13), Handle%) ' Envia <CR> a RX
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    Call WriteToPort("T", Handle%) ' Envia ATT off a RX
    Call WriteToPort(Chr$(13), Handle%) ' Envia <CR> a RX
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    Call WriteToPort("G", Handle%) ' Envia MUTE-OUT auto con Squelch a RX
    Call WriteToPort(Chr$(13), Handle%) ' Envia <CR> a RX
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    
End Sub

'Sets operating mode for AR-8000 receiver:
Sub AR8000RadioSetMode(mode$, Handle%)
    Select Case LCase$(mode$)
        Case Is = "fm-w"
            m% = 0
        Case Is = "fm-n", "fm"
            m% = 1
        Case Is = "am"
            m% = 2
        Case Is = "usb"
            m% = 3
        Case Is = "lsb"
            m% = 4
        Case Is = "cw", "cw-n"
            m% = 5
       Case Else
            m% = -1
    End Select
    If m% >= 0 Then
        'To pass form numbre to ASCII
        m% = m% + 48
        InString = ReadFromPort(Handle%)
        'Send bytes to serial port. Begin with MD (Mode) then the
        'data mode
        Call WriteToPort("MD" + Chr$(m%) + Chr$(13), Handle%)
        InString = ReadFromPort(Handle%)
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
'Sets operating mode for AR-5000 receiver:
Sub AR5000RadioSetMode(mode$, Handle%)
    Select Case LCase$(mode$)
        Case Is = "fm-w", "fm", "fm-n"
            m% = 0
        Case Is = "am"
            m% = 1
        Case Is = "lsb"
            m% = 2
        Case Is = "usb"
            m% = 3
        Case Is = "cw"
            m% = 4
       Case Else
            m% = -1
    End Select
    If m% >= 0 Then
        'To pass form numbre to ASCII
        m% = m% + 48
        InString = ReadFromPort(Handle%)
        'Send bytes to serial port. Begin with MD (Mode) then the
        'data mode
        Call WriteToPort("MD" + Chr$(m%) + Chr$(13), Handle%)
        
        InString = ReadFromPort(Handle%)
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
'Sets bandwidth for AR-5000 receiver:
Sub AR5000RadioSetBW(FilterKHz, Handle%)
    Select Case FilterKHz
    Case Is < 1
        m% = 0
    Case Is < 4.5
        m% = 1
    Case Is < 10
        m% = 2
    Case Is < 25
        m% = 3
    Case Is < 50
        m% = 4
    Case Is < 150
        m% = 5
    Case Else
        m% = 6
    End Select
    If m% >= 0 Then
        'To pass form numbre to ASCII
        m% = m% + 48
        InString = ReadFromPort(Handle%)
        'Send bytes to serial port. Begin with MD (Mode) then the
        'data mode
        Call WriteToPort("BW" + Chr$(m%) + Chr$(13), Handle%)
        InString = ReadFromPort(Handle%)
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub

Sub AR8000RadioSetFreq(Freq, Handle%) ' G6LVB 30 Dec 2000 Changed Freq& to Freq for freqs>2G
    InString = ReadFromPort(Handle%)
    'Send RF to set the frecuency, and then the data and
    'the return=13
    Call WriteToPort("RF", Handle%)
    'format the frequency for the AR-8000
    freqtext = Format(Cdbl2(Freq), "0000000000") ' G6LVB 30 Dec 2000 Changed Freq& to Freq for freqs>2G
    For f% = 1 To 10
       n% = Mid(freqtext, f%, 1)
       n% = n% + 48
       Call WriteToPort(Chr$(n%), Handle%)
    Next
    Call WriteToPort(Chr$(13), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub


Sub AR8000RadioOff(Handle%)
    InString = ReadFromPort(Handle%)
    'Send VA to set VFO A of the AR-8000
    'the return, CR
    Call WriteToPort("VA" + Chr$(13), Handle%)
    'Send EX to exit form remote control from AR-8000
    'the return, CR
    Call WriteToPort("EX" + Chr$(13), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub

Sub AR8000RadioOn(Handle%)
    InString = ReadFromPort(Handle%)
    'Send ST to enter the minimal step to AR-8000
    'the return, CR
    Call WriteToPort("ST", Handle%)
    Step& = 50
    freqtext = Format(Cdbl2(Step&), "000000")
    For f% = 1 To 6
       n% = Mid(freqtext, f%, 1)
       n% = n% + 48
       Call WriteToPort(Chr$(n%), Handle%)
    Next
    Call WriteToPort(Chr$(13), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
'VR-5000 receiver CAT On:
Sub VR5000RadioCATOn(Handle%)
InString = ReadFromPort(Handle%)
Call WriteToPort(Chr$(0), Handle%)
Call WriteToPort(Chr$(0), Handle%)
Call WriteToPort(Chr$(0), Handle%)
Call WriteToPort(Chr$(0), Handle%)
Call WriteToPort(Chr$(0), Handle%)
a = Timer
Do
Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)

End Sub
'VR-5000 receiver CAT Off:
Sub VR5000RadioCATOff(Handle%)
InString = ReadFromPort(Handle%)
Call WriteToPort(Chr$(0), Handle%)
Call WriteToPort(Chr$(0), Handle%)
Call WriteToPort(Chr$(0), Handle%)
Call WriteToPort(Chr$(0), Handle%)
Call WriteToPort(Chr$(&H80), Handle%)
a = Timer
Do
Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)

End Sub

'VR-5000 receiver Main frequency set:
Sub VR5000RadioSetMainFreq(Freq, Handle%)
InString = ReadFromPort(Handle%)
Call WriteToPort(Chr$(Fix((Freq / 10) / 2 ^ 24) And &HFF), Handle%)
Call WriteToPort(Chr$(Fix((Freq / 10) / 2 ^ 16) And &HFF), Handle%)
Call WriteToPort(Chr$(Fix((Freq / 10) / 2 ^ 8) And &HFF), Handle%)
Call WriteToPort(Chr$(Fix(Freq / 10) And &HFF), Handle%)
Call WriteToPort(Chr$(1), Handle%)
a = Timer
Do
Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)

End Sub

'VR-5000 receiver Sub frequency set:
Sub VR5000RadioSetSubFreq(Freq, Handle%)
InString = ReadFromPort(Handle%)
Call WriteToPort(Chr$(Fix((Freq / 10) / 2 ^ 24) And &HFF), Handle%)
Call WriteToPort(Chr$(Fix((Freq / 10) / 2 ^ 16) And &HFF), Handle%)
Call WriteToPort(Chr$(Fix((Freq / 10) / 2 ^ 8) And &HFF), Handle%)
Call WriteToPort(Chr$(Fix(Freq / 10) And &HFF), Handle%)
Call WriteToPort(Chr$(&H31), Handle%)
a = Timer
Do
Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)

End Sub

'VR-5000 set mode:
Sub VR5000RadioSetMainMode(m$, Handle%)

InString = ReadFromPort(Handle%)

Select Case LCase$(m$)
Case "lsb"
    a = &H0
Case "usb"
    a = &H1
Case "cw"
    a = &H2
Case "am"
    a = &H4
Case "am-w"
    a = &H44
Case "fm-w"
    a = &H48
Case "am-n"
    a = &H84
Case "fm-n"
    a = &H88
End Select
Call WriteToPort(Chr$(a), Handle%)
Call WriteToPort(Chr$(&H21), Handle%)
Call WriteToPort(Chr$(0), Handle%)
Call WriteToPort(Chr$(0), Handle%)
Call WriteToPort(Chr$(&H7), Handle%)
a = Timer
Do
Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)

End Sub

'VR-5000 set Sub mode:
Sub VR5000RadioSetSubMode(m$, Handle%)

InString = ReadFromPort(Handle%)

Select Case LCase$(m$)
Case "lsb"
    a = &H0
Case "usb"
    a = &H1
Case "cw"
    a = &H2
Case "am"
    a = &H4
Case "am-w"
    a = &H44
Case "fm-w"
    a = &H48
Case "am-n"
    a = &H84
Case "fm-n"
    a = &H88
End Select
Call WriteToPort(Chr$(a), Handle%)
Call WriteToPort(Chr$(&H21), Handle%)
Call WriteToPort(Chr$(0), Handle%)
Call WriteToPort(Chr$(0), Handle%)
Call WriteToPort(Chr$(&H37), Handle%)
a = Timer
Do
Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
'VR-5000 frame readback routine
'waits up to 200mS for a char to come from the radio.
'Valid frames are 1 bytes long only.
'The output of this function is a string beginning with the decoded
'frame type code followed by the data available separated by spaces.
'the following type codes can be returned: (freqs. in Hz)
'code "00:" - frame not understood pass timeout period.
'code "40:" - RSSI
Function VR5000RadioReadFrame(Handle%) As String
'This will hold the received frame
Dim InBuff(10) As Integer
'initialize buffer pointer..
InBuffPtr% = 0
FrameSize% = 1
'set start time...
a = Timer
'RadioControlDelay(Handle%) will be fixed at 200mS as this time is part of
'Yaesu protocol specification
Do
    'Wait until we receive bytes
    'or a time-out occurs
    Select Case Handle%
    Case 1
        Do
            
        Loop Until (MSComm1.InBufferCount >= 1) Or Abs(Timer - a) > RadioReplyTimeout(Handle%)
        InString = MSComm1.Input
    Case 2
        Do
            
        Loop Until (MSComm2.InBufferCount >= 1) Or Abs(Timer - a) > RadioReplyTimeout(Handle%)
        InString = MSComm2.Input
    Case 3
        Do
            
        Loop Until (MSComm3.InBufferCount >= 1) Or Abs(Timer - a) > RadioReplyTimeout(Handle%)
        InString = MSComm3.Input
    End Select
    'We pass the received bytes to the buffer
    For i% = 0 To LenB(InString) - 1
        InBuff(InBuffPtr%) = InString(i%)
        InBuffPtr% = InBuffPtr% + 1
    Next
    'if a byte was received we reset timer to wait for another
    '200mS
    If Abs(Timer - a) < RadioReplyTimeout(Handle%) Then
        a = Timer
    End If
'if a byte was received we know answer is complete:
Loop While (Abs(Timer - a) < RadioReplyTimeout(Handle%)) And (InBuffPtr% <> FrameSize%)
'if answer is greater than 1 or is nil:
If InBuffPtr% = 0 Or InBuffPtr% > 1 Then
    s$ = "00:"
    VR5000RadioReadFrame = s$
    Exit Function
Else
    'otherwise examine frame:
    'we know buffer size is 1byte
    s$ = "40: "
    'take the 5 least sign. bits and scale to 8 bits:
    a = CInt((InBuff(0) And &H1F) * 8.226)
    s$ = s$ + Str$(a)
    VR5000RadioReadFrame = s$
    Exit Function
End If
End Function

'VR-5000 read s-meter:
Function VR5000RadioReadRSSI(Handle%) As Integer
InString = ReadFromPort(Handle%)
Call WriteToPort(Chr$(0), Handle%)
Call WriteToPort(Chr$(0), Handle%)
Call WriteToPort(Chr$(0), Handle%)
Call WriteToPort(Chr$(0), Handle%)
Call WriteToPort(Chr$(&HE7), Handle%)
'wait upto 200mS for receiver answer:
a$ = VR5000RadioReadFrame(Handle%)
'check frame code 40 for RSSI
If Cdbl2(PickWord(a$, 1)) <> 40 Then
'    Call frmMessage.ShowMessage("Error during" + Chr$(13) _
        + "Yaesu Read RSSI", 10)
    VR5000RadioReadRSSI = 0
Else
    VR5000RadioReadRSSI = CInt(PickWord(a$, 2))
End If
End Function

Sub ICPCRRadioSet(Freq&, mode$, FilterKHz, Handle%)
    InString = ReadFromPort(Handle%)
    'begin PCR freq. set with KO
    Call WriteToPort("K0", Handle%)
    'PCR need 10 digits for the frequency
    freqtext$ = Format(Cdbl2(Freq&), "0000000000")
    Call WriteToPort(freqtext$, Handle%)
    Select Case LCase(mode$)
    Case "lsb"
        m$ = "00"
    Case "usb"
        m$ = "01"
    Case "am"
        m$ = "02"
    Case "cw", "cw-n"
        m$ = "03"
    Case "fm-n", "fm"
        m$ = "05"
    Case "fm-w"
        m$ = "06"
    End Select
    'send mode...
    Call WriteToPort(m$, Handle%)
    Select Case FilterKHz
    Case Is < 4.5
        f$ = "00"
    Case Is < 10
        f$ = "01"
    Case Is < 25
        f$ = "02"
    Case Is < 100
        f$ = "03"
    Case Else
        f$ = "04"
    End Select
    'and filter...
    Call WriteToPort(f$ + "00", Handle%)
    'end with CR+LF and thats all.
    Call WriteToPort(Chr$(13) + Chr$(10), Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    InString = ReadFromPort(Handle%)
End Sub
Sub ICPCRRadioOn(Handle%)
    InString = ReadFromPort(Handle%)
    'PCR ON command:
    Call WriteToPort("H101", Handle%)
    'end comman with CR+LF
    Call WriteToPort(Chr$(13) + Chr$(10), Handle%)
    InString = ReadFromPort(Handle%)
    'Auto-update off command:
    Call WriteToPort("G300", Handle%)
    'end command with CR+LF
    Call WriteToPort(Chr$(13) + Chr$(10), Handle%)
    InString = ReadFromPort(Handle%)
    'Center IF shift  command:
    Call WriteToPort("J4380", Handle%)
    'end comman with CR+LF
    Call WriteToPort(Chr$(13) + Chr$(10), Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    InString = ReadFromPort(Handle%)
End Sub
Sub ICPCRRadioOff(Handle%)
    InString = ReadFromPort(Handle%)
    'PCR ON command:
    Call WriteToPort("H100", Handle%)
    'end comman with CR+LF
    Call WriteToPort(Chr$(13) + Chr$(10), Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    InString = ReadFromPort(Handle%)
End Sub
Function ICPCRRadioReadRSSI(Handle%) As Integer
    InString = ReadFromPort(Handle%)
    'PCR RSSI query command:
    Call WriteToPort("I1?", Handle%)
    'end command with CR+LF
    Call WriteToPort(Chr$(13) + Chr$(10), Handle%)
    a = Timer
    Do
        
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    s$ = Array2String$(ReadFromPort(Handle%))
    If Len(s$) > 0 Then
        If IsControl(Asc(Mid$(s$, 1, 1))) Then
            s$ = Mid(s$, 2)
        End If
    End If
    p% = InStr(s$, "I1")
    If p% Then
        ICPCRRadioReadRSSI = Cdbl2("&H" + LTrim(Mid(s$, p% + 2, 2)))
    Else
        ICPCRRadioReadRSSI = 0
    End If
End Function

Sub ICPCRRadioSetVol(Volume%, Handle%)
    InString = ReadFromPort(Handle%)
    'PCR ON command:
    Call WriteToPort("J40", Handle%)
    v$ = Hex(Volume)
    v$ = String(2 - Len(v$), "0") + v$
    Call WriteToPort(v$, Handle%)
    'end with CR+LF and thats all.
    Call WriteToPort(Chr$(13) + Chr$(10), Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    InString = ReadFromPort(Handle%)
End Sub
'Sets operating mode for TM-D700 txcvr:
Sub TMD700RadioSetMode(mode$, Handle%)
    Select Case LCase$(mode$)
        Case Is = "fm-w", "fm", "fm-n"
            m% = 0
        Case Is = "am"
            m% = 1
       Case Else
            m% = -1
    End Select
    If m% >= 0 Then
        'To pass form numbre to ASCII
        m% = m% + 48
        InString = ReadFromPort(Handle%)
        'Send bytes to serial port. Begin with MD (Mode) then the
        'data mode
        Call WriteToPort("MD " + Chr$(m%) + Chr$(13), Handle%)
        
        InString = ReadFromPort(Handle%)
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
'Sets full-duplex mode for TM-D700 txcvr:
Sub TMD700RadioSetDual(Handle%)
    InString = ReadFromPort(Handle%)
    'Send bytes to serial port.
    Call WriteToPort("DTB 3" + Chr$(13), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
'Sets band A for TM-D700 txcvr:
Sub TMD700RadioSetA(Handle%)
    InString = ReadFromPort(Handle%)
    'Send bytes to serial port.
    Call WriteToPort("BC 0,1" + Chr$(13), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
'Sets band B for TM-D700 txcvr:
Sub TMD700RadioSetB(Handle%)
    InString = ReadFromPort(Handle%)
    'Send bytes to serial port.
    Call WriteToPort("BC 1,0" + Chr$(13), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
'Cancells auto repeater offset for TM-D700 txcvr:
Sub TMD700RadioCancelSplit(Handle%)
    InString = ReadFromPort(Handle%)
    'Send bytes to serial port.
    Call WriteToPort("ARO 0" + Chr$(13), Handle%)
    Call WriteToPort("SFT 0" + Chr$(13), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
'Cancells Tone for TM-D700 txcvr:
Sub TMD700RadioCancelTone(Handle%)
    InString = ReadFromPort(Handle%)
    'Send bytes to serial port.
    Call WriteToPort("CT 0" + Chr$(13), Handle%)
    Call WriteToPort("TO 0" + Chr$(13), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
'Sets BandA-TX, BandB-RX for TM-D700 txcvr:
Sub TMD700RadioSetATX_BRX(Handle%)
    InString = ReadFromPort(Handle%)
    'Send bytes to serial port.
    Call WriteToPort("DTB 2" + Chr$(13), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
'Sets BandB-TX, BandA-RX for TM-D700 txcvr:
Sub TMD700RadioSetBTX_ARX(Handle%)
    InString = ReadFromPort(Handle%)
    'Send bytes to serial port.
    Call WriteToPort("DTB 3" + Chr$(13), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
'Sets downlink operating freq&mode for TM-D700 txcvr:
'*****This is not allright/properly tested******
Sub TMD700RadioSetRX(Freq&, mode$, Handle%)
    InString = ReadFromPort(Handle%)
    Call WriteToPort("BUF 0, ", Handle%)
    'PCR need 10 digits for the frequency
    freqtext$ = Format(Cdbl2(Freq&), "00000000000")
    Call WriteToPort(freqtext$ + ", ", Handle%)
    Call WriteToPort("0, 0, 0, 0, 0, 0, 1, 0, 1, 0, ", Handle%)
    Select Case LCase$(mode$)
        Case Is = "am"
            m% = 1
        Case Else
            m% = 0
    End Select
    'To pass form numbre to ASCII
    m% = m% + 48
    InString = ReadFromPort(Handle%)
    'Send bytes to serial port. Begin with MD (Mode) then the
    'data mode
    Call WriteToPort(Chr$(m%) + Chr$(13), Handle%)
    
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
'Sets uplink operating freq&mode for TM-D700 txcvr:
'******** This is not allright/properly tested *******
Sub TMD700RadioSetTX(Freq&, mode$, Handle%)
    InString = ReadFromPort(Handle%)
    Call WriteToPort("BUF 1, ", Handle%)
    'PCR need 10 digits for the frequency
    freqtext$ = Format(Cdbl2(Freq&), "00000000000")
    Call WriteToPort(freqtext$ + ", ", Handle%)
    Call WriteToPort("0, 0, 0, 0, 0, 0, 1, 0, 1, 0, ", Handle%)
    Select Case LCase$(mode$)
        Case Is = "am"
            m% = 1
        Case Else
            m% = 0
    End Select
    'To pass form numbre to ASCII
    m% = m% + 48
    InString = ReadFromPort(Handle%)
    'Send bytes to serial port. Begin with MD (Mode) then the
    'data mode
    Call WriteToPort(Chr$(m%) + Chr$(13), Handle%)
    
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
'Sets operating freq of selected band for TM-D700 txcvr:
'freq is rounded to the nearest 5KHz
Sub TMD700RadioSetFreq(Freq&, Handle%)
    InString = ReadFromPort(Handle%)
    Call WriteToPort("FQ ", Handle%)
    'need 10 digits for the frequency
    freqtext$ = Format(5000 * CLng(Freq& / 5000), "00000000000")
    'Tuning step fixed at 5KHz:
    Call WriteToPort(freqtext$ + ",0" + Chr$(13), Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
'Cancells split for TS-790 txcvr:
Sub TS790RadioCancelSplit(Handle%, Bidir%)
    InString = ReadFromPort(Handle%)
    'Send bytes to serial port.
    Call WriteToPort("SP0;", Handle%)
    InString = ReadFromPort(Handle%)
    If Bidir% Then
        Call TS790RadioWaitAck(Handle%)
        InString = ReadFromPort(Handle%)
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
'Cancells scan for TS-790 txcvr:
Sub TS790RadioCancelScan(Handle%, Bidir%)
    InString = ReadFromPort(Handle%)
    'Send bytes to serial port.
    Call WriteToPort("SC0;", Handle%)
    InString = ReadFromPort(Handle%)
    If Bidir% Then
        Call TS790RadioWaitAck(Handle%)
        InString = ReadFromPort(Handle%)
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
'Sets operating mode for TS-790 txcvr:
Sub TS790RadioSetMode(mode$, Handle%, Bidir%)
    Select Case LCase$(mode$)
        Case Is = "fm-w", "fm", "fm-n"
            m% = 4
        Case Is = "lsb", "am"
            m% = 1
        Case Is = "usb"
            m% = 2
        Case Is = "cw"
            m% = 3
        Case Is = "cw-n"
            m% = 7
       Case Else
            m% = -1
    End Select
    If m% >= 0 Then
        'To pass form numbre to ASCII
        m% = m% + 48
        InString = ReadFromPort(Handle%)
        'Send bytes to serial port. Begin with MD (Mode) then the
        'data mode
        Call WriteToPort("MD" + Chr$(m%) + ";", Handle%)
        
        InString = ReadFromPort(Handle%)
        If Bidir% Then
            Call TS790RadioWaitAck(Handle%)
            InString = ReadFromPort(Handle%)
        Else
            a = Timer
            Do
            Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
        End If
    End If
End Sub
'Sets main band for TS-790 txcvr:
Sub TS790RadioSetMain(Handle%, Bidir%)
    InString = ReadFromPort(Handle%)
    'Send bytes to serial port.
    Call WriteToPort("DC0;", Handle%)
    InString = ReadFromPort(Handle%)
    If Bidir% Then
        Call TS790RadioWaitAck(Handle%)
        InString = ReadFromPort(Handle%)
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
'Sets sub band for TS-790 txcvr:
Sub TS790RadioSetSub(Handle%, Bidir%)
    InString = ReadFromPort(Handle%)
    'Send bytes to serial port.
    Call WriteToPort("DC1;", Handle%)
    InString = ReadFromPort(Handle%)
    If Bidir% Then
        Call TS790RadioWaitAck(Handle%)
        InString = ReadFromPort(Handle%)
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
'Sets VFO A frequency for TS-790 txcvr:
Sub TS790RadioSetVFOA(Freq#, Handle%, Bidir%)
    InString = ReadFromPort(Handle%)
    '790 need 10 digits for the frequency
    freqtext$ = Format(Freq#, "00000000000")
    Call WriteToPort("FA" + freqtext$ + ";", Handle%)
    InString = ReadFromPort(Handle%)
    If Bidir% Then
        Call TS790RadioWaitAck(Handle%)
        InString = ReadFromPort(Handle%)
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
'Sets VFO B frequency for TS-790 txcvr:
Sub TS790RadioSetVFOB(Freq#, Handle%, Bidir%)
    InString = ReadFromPort(Handle%)
    '790 need 10 digits for the frequency
    freqtext$ = Format(Freq#, "00000000000")
    Call WriteToPort("FB" + freqtext$ + ";", Handle%)
    InString = ReadFromPort(Handle%)
    If Bidir% Then
        Call TS790RadioWaitAck(Handle%)
        InString = ReadFromPort(Handle%)
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
'Reads VFO A frequency for TS-790 txcvr:
Function TS790RadioReadVFOA(Handle%) As Double
InString = ReadFromPort(Handle%)
Call WriteToPort("FA" + ";", Handle%)
a$ = TS790RadioReadFrame(Handle%)
'check frame code 10
If Cdbl2(PickWord(a$, 1)) <> 10 Then
'    Call frmMessage.ShowMessage("Error during" + Chr$(13) _
        + "Kenwood Read Frequency", 10)
    TS790RadioReadVFOA = 0
Else
    TS790RadioReadVFOA = Cdbl2(PickWord(a$, 2))
End If
End Function
'Reads VFO B frequency for TS-790 txcvr:
Function TS790RadioReadVFOB(Handle%) As Double
InString = ReadFromPort(Handle%)
Call WriteToPort("FB" + ";", Handle%)
a$ = TS790RadioReadFrame(Handle%)
'check frame code 10
If Cdbl2(PickWord(a$, 1)) <> 11 Then
'    Call frmMessage.ShowMessage("Error during" + Chr$(13) _
        + "Kenwood Read Frequency", 10)
    TS790RadioReadVFOB = 0
Else
    TS790RadioReadVFOB = Cdbl2(PickWord(a$, 2))
End If
End Function
'Reads Sub band signal strength for TS-790 txcvr, returns 0-255
Function TS790RadioReadSubRSSI(Handle%) As Integer
InString = ReadFromPort(Handle%)
Call WriteToPort("SM;", Handle%)
a$ = TS790RadioReadFrame(Handle%)
'check frame code 40
If Cdbl2(PickWord(a$, 1)) <> 40 Then
'    Call frmMessage.ShowMessage("Error during" + Chr$(13) _
        + "Kenwood Read RSSI", 10)
    TS790RadioReadSubRSSI = 0
Else
    TS790RadioReadSubRSSI = 17 * Cdbl2(PickWord(a$, 2))
End If
End Function
'Selects VFO A for TS-790 txcvr:
Sub TS790RadioSelectVFOA(Handle%, Bidir%)
    InString = ReadFromPort(Handle%)
    Call WriteToPort("FN0;", Handle%)
    InString = ReadFromPort(Handle%)
    If Bidir% Then
        Call TS790RadioWaitAck(Handle%)
        InString = ReadFromPort(Handle%)
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
'Selects VFO B for TS-790 txcvr:
Sub TS790RadioSelectVFOB(Handle%, Bidir%)
    InString = ReadFromPort(Handle%)
    Call WriteToPort("FN1;", Handle%)
    InString = ReadFromPort(Handle%)
    If Bidir% Then
        Call TS790RadioWaitAck(Handle%)
        InString = ReadFromPort(Handle%)
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
'Selects MEMORY for TS-790 txcvr:
Sub TS790RadioSelectMEM(Handle%, Bidir%)
    InString = ReadFromPort(Handle%)
    Call WriteToPort("FN2;", Handle%)
    InString = ReadFromPort(Handle%)
    If Bidir% Then
        Call TS790RadioWaitAck(Handle%)
        InString = ReadFromPort(Handle%)
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
'Sends four semicolons to TS-790 txcvr:
Sub TS790RadioSendSync(Handle%)
    InString = ReadFromPort(Handle%)
    Call WriteToPort(";;;;", Handle%)
    InString = ReadFromPort(Handle%)
    a = Timer
    Do
    Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
End Sub
'Store all chars received until a ";" and return them as a string
'including the ";" up to 14 chars
'if nothing is received within 250mS there is a time-out exit
Function TS790RadioWaitAck(Handle%)
'This will hold the received frame
Dim InBuff%(41)
'initialize buffer pointer. A new buffer is needed to hold
'received bytes to perform a search for the header. This
'pointer is needed so as to put bytes consecutively one after
'another without overwriting previous bytes.
InBuffPtr% = 0
'this will indicate the position where the header is found.
'while it is -1 it means it haven't been found yet. after it
'is found we need to be sure that the reqd. number of bytes
'is received
found% = -1
ackstring$ = ""
a = Timer
Do
    'Wait until we receive bytes
    'or a time-out occurs
    Select Case Handle%
    Case 1
        Do
            
        Loop Until MSComm1.InBufferCount >= 1 Or Abs(Timer - a) > RadioReplyTimeout(Handle%)
        InString = MSComm1.Input
    Case 2
        Do
            
        Loop Until MSComm2.InBufferCount >= 1 Or Abs(Timer - a) > RadioReplyTimeout(Handle%)
        InString = MSComm2.Input
    Case 3
        Do
            
        Loop Until MSComm3.InBufferCount >= 1 Or Abs(Timer - a) > RadioReplyTimeout(Handle%)
        InString = MSComm3.Input
    End Select
    'We pass the received bytes to the buffer
    f% = 0
    For InBuffPtr% = InBuffPtr% To InBuffPtr% + LenB(InString) - 1
        If InBuffPtr% < 40 Then
            InBuff%(InBuffPtr%) = InString(f%)
        End If
        f% = f% + 1
    Next
    'we will examine the buffer searching for ";":
    For f% = 0 To InBuffPtr%
        If Chr$(InBuff(f%)) = ";" Then
            found% = f%
            Exit For
        End If
    Next
    'Let rig answers be longer
    If Abs(Timer - a) < RadioReplyTimeout(Handle%) Then
        a = Timer
    End If
Loop While found% < 0 And Abs(Timer - a) < RadioReplyTimeout(Handle%)
If Abs(Timer - a) > RadioControlDelay(Handle%) Then
Else
    'pass only 13 last chars to output string:
    For f% = found% - 13 To found%
        If f% >= 0 Then ackstring$ = ackstring$ + Chr$(InBuff(f%))
    Next
End If
TS790RadioWaitAck = ackstring$
End Function
'Cancells split for TS-2000 txcvr:
Sub TS2000RadioCancelSplit(Handle%, Bidir%) ' 2 March 2000 G6LVB TS-2000
    InString = ReadFromPort(Handle%)
    'Send bytes to serial port.
    Call WriteToPort("FT1;", Handle%)
    InString = ReadFromPort(Handle%)
    If Bidir% Then
        Call TS790RadioWaitAck(Handle%)
        InString = ReadFromPort(Handle%)
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
Sub TS2000RadioSatOn(Handle%, Bidir%) ' 2 March 2000 G6LVB TS-2000
    InString = ReadFromPort(Handle%)
    'Send bytes to serial port.
    Call WriteToPort("SA1010000;", Handle%) ' 13 Feb 2002 G6LVB was "SA1010110;"
    InString = ReadFromPort(Handle%)
    If Bidir% Then
        Call TS790RadioWaitAck(Handle%)
        InString = ReadFromPort(Handle%)
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
Sub TS2000RadioSatOff(Handle%, Bidir%) ' 2 March 2000 G6LVB TS-2000
    InString = ReadFromPort(Handle%)
    'Send bytes to serial port.
    Call WriteToPort("SA0000000;", Handle%) ' 13 Feb 2002 G6LVB was "SA0000110;"
    InString = ReadFromPort(Handle%)
    If Bidir% Then
        Call TS790RadioWaitAck(Handle%)
        InString = ReadFromPort(Handle%)
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
'Sets main band for TS-2000 txcvr:
Sub TS2000RadioSetMain(Handle%, Bidir%) ' 2 March 2000 G6LVB TS-2000
    InString = ReadFromPort(Handle%)
    'Send bytes to serial port.
    Call WriteToPort("SA1010000;", Handle%) ' 13 Feb 2002 G6LVB was "SA1010110;"
    InString = ReadFromPort(Handle%)
    If Bidir% Then
        Call TS790RadioWaitAck(Handle%)
        InString = ReadFromPort(Handle%)
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
'Sets sub band for TS-2000 txcvr:
Sub TS2000RadioSetSub(Handle%, Bidir%) ' 2 March 2000 G6LVB TS-2000
    InString = ReadFromPort(Handle%)
    'Send bytes to serial port.
    Call WriteToPort("SA1011000;", Handle%) ' 13 Feb 2002 G6LVB was "SA1011110;"
    InString = ReadFromPort(Handle%)
    If Bidir% Then
        Call TS790RadioWaitAck(Handle%)
        InString = ReadFromPort(Handle%)
    Else
        a = Timer
        Do
        Loop Until Abs(Timer - a) > RadioControlDelay(Handle%)
    End If
End Sub
'Reads Sub band signal strength for TS-2000 txcvr, returns 0-255
Function TS2000RadioReadSubRSSI(Handle%) As Integer
InString = ReadFromPort(Handle%)
Call WriteToPort("SM1;", Handle%)
a$ = TS790RadioReadFrame(Handle%)
'check frame code 42
If Cdbl2(PickWord(a$, 1)) <> 42 Then
'    Call frmMessage.ShowMessage("Error during" + Chr$(13) _
        + "Kenwood Read RSSI", 10)
    TS2000RadioReadSubRSSI = 0
Else
    TS2000RadioReadSubRSSI = 17 * Cdbl2(PickWord(a$, 2))
End If
End Function

'Reads Main band signal strength for TS-2000 txcvr, returns 0-255
Function TS2000RadioReadMainRSSI(Handle%) As Integer
InString = ReadFromPort(Handle%)
Call WriteToPort("SM0;", Handle%)
a$ = TS790RadioReadFrame(Handle%)
'check frame code 41
If Cdbl2(PickWord(a$, 1)) <> 41 Then
'    Call frmMessage.ShowMessage("Error during" + Chr$(13) _
        + "Kenwood Read RSSI", 10)
    TS2000RadioReadMainRSSI = 0
Else
    TS2000RadioReadMainRSSI = 8.5 * Cdbl2(PickWord(a$, 2))
End If
End Function

Sub FRG9600RadioSetFreq(Freq&, Handle%)
    Dim ts1 As Integer
    Dim ts2 As Integer
    Dim ts3 As Integer
    Dim ts4 As Integer
    Dim ts5 As Integer
    Dim sstr As String
    
    ss$ = Format(Cdbl2(Freq&), "000000000")
    ts1 = 10
    ts2 = ((Asc(Mid(ss$, 1, 1)) - Asc("0")) * 16) + (Asc(Mid(ss$, 2, 1)) - Asc("0"))
    ts3 = ((Asc(Mid(ss$, 3, 1)) - Asc("0")) * 16) + (Asc(Mid(ss$, 4, 1)) - Asc("0"))
    ts4 = ((Asc(Mid(ss$, 5, 1)) - Asc("0")) * 16) + (Asc(Mid(ss$, 6, 1)) - Asc("0"))
    ts5 = ((Asc(Mid(ss$, 7, 1)) - Asc("0")) * 16) + (Asc(Mid(ss$, 8, 1)) - Asc("0"))
    
    sstr = Chr(ts1) + Chr(ts2) + Chr(ts3) + Chr(ts4) + Chr(ts5)
    
'    MSComm1.Output = sstr
    
    StToSend$ = sstr
    Call WriteToPort(StToSend$, Handle%)
End Sub
Sub FRG9600RadioSetMode(mode$, Handle%)
    CurMode$ = CurDownlinkMode
    
    ' Only set the mode if it is different than already set
    If mode$ <> CurMode$ Then
        Select Case LCase$(mode$)
            Case Is = "lsb"
                m$ = Chr$(&H10)
            Case Is = "usb", "cw"
                m$ = Chr$(&H11)
            Case Is = "am", "am-n"
                m$ = Chr$(&H14)
            Case Is = "am-w"
                m$ = Chr$(&H15)
            Case Is = "fm", "fm-n"
                m$ = Chr$(&H16)
            Case Is = "fm-w"
                m$ = Chr$(&H17)
           Case Else
                m$ = Chr$(&H1)
        End Select
        If m$ <> Chr$(&H1) Then
            t$ = Chr$(&H0)
            Call WriteToPort(m$ + t$ + t$ + t$ + t$, Handle%)
        End If
        
        ' if its different also update the current mode flag
        CurDownlinkMode = mode$
    End If
        
End Sub
Sub FRG9600RadioOn()
    CurDownlinkMode = ""
End Sub
Sub TS711RadioSetStart(Handle%, Birdir%)
    CurDownlinkMode = ""
    CurUplinkMode = ""
    DownlinkRSSI.Enabled = False
    DownlinkRSSI.Value = 127
    ' if I have a bidirectional interface wait for a proper response
    ' otherwise just try to guess
    If Birdir% Then
        Call WriteToPort("FA;", Handle%)
        X% = 0
        Do
            a$ = TS711WaitAck(Handle%)
            If Len(a$) > 5 Then
                Exit Do
            End If
            X% = X% + 1
            If X% > 10 Then
                Exit Do
            End If
        Loop
        'If x% > 10 Then
        '    MsgBox "Can't Receive Data from radio. Please check connections and restart program", vbOKOnly, "Communication Error"
        '    End
        'End If
    Else
        Call WriteToPort("FN0;FN1;FN0;", Handle%)
    End If
    ' use VFO A/B split RX/TX operaion for in band operation
    ' i.e. talking to the space station
    '
    ' This is so I can adjust the uplink and downlink doppler independently
    DownlinkPort% = Cdbl2(Right(GetSetting("WiSP_DDE_Client", "Rig" + DownlinkIndex.text, "Radio_port"), 1))
    UplinkPort% = Cdbl2(Right(GetSetting("WiSP_DDE_Client", "Rig" + UplinkIndex.text, "Radio_port"), 1))
    If (DownlinkPort% = UplinkPort%) Then
        Call WriteToPort("SP1;", Handle%)
    Else
        Call WriteToPort("SP0;", Handle%)
    End If
End Sub
Sub TS711RadioSetMode(mode$, Handle%, Direction$, Bidir%)
    If (Direction$ = "Down") Then
        CurMode$ = CurDownlinkMode
    Else
        CurMode$ = CurUplinkMode
    End If
    
    ' Only set the mode if it is different than already set
    If mode$ <> CurMode$ Then
        ' if I have a bidirectional interface wait for a proper response
        ' otherwise just try to guess
        If Birdir% Then
            Call WriteToPort("FA;", Handle%)
            X% = 0
            Do
                a$ = TS711WaitAck(Handle%)
                If Len(a$) > 5 Then
                    Exit Do
                End If
                X% = X% + 1
                If X% > 10 Then
                    Exit Do
                End If
            Loop
            'If x% > 10 Then
            '    MsgBox "Can't Receive Data from radio. Please check connections and restart program", vbOKOnly, "Communication Error"
            '    End
            'End If
        Else
            Call WriteToPort("FN0;FN1;FN0;", Handle%)
        End If
        
        Select Case LCase$(mode$)
            Case Is = "lsb"
                m$ = "1"
            Case Is = "usb"
                m$ = "2"
            Case Is = "cw"
                m$ = "3"
            Case Is = "fm", "fm-n"
                m$ = "4"
           Case Else
                m$ = "9"
        End Select
        DownlinkPort% = Cdbl2(Right(GetSetting("WiSP_DDE_Client", "Rig" + DownlinkIndex.text, "Radio_port"), 1))
        UplinkPort% = Cdbl2(Right(GetSetting("WiSP_DDE_Client", "Rig" + UplinkIndex.text, "Radio_port"), 1))
        If m$ <> "9" Then
            If (DownlinkPort% = UplinkPort%) And (Direction$ <> "Down") Then
                ' write to port if valid
                ' use VFO B for TX inband
                Call WriteToPort("FN1;MD" + m$ + ";FN0;", Handle%)
            Else
                ' write to port if valid
                ' use VFO A for TX & RX cross band and RX inband
                Call WriteToPort("FN0;MD" + m$ + ";", Handle%)
            End If
        End If
        
        If (DownlinkPort% = UplinkPort%) Then
            Call WriteToPort("SP1;", Handle%)
        Else
            Call WriteToPort("SP0;", Handle%)
        End If
        
        ' if its different also update the current mode flag
        If (Direction$ = "Down") Then
            CurDownlinkMode = mode$
        Else
            CurUplinkMode = mode$
        End If
    
    End If
        
End Sub

Sub TS711RadioSetVFOA(Freq&, Handle%)
    freqtext$ = Format(Cdbl2(Freq&), "00000000000")
    Call WriteToPort("FA" + freqtext$ + ";", Handle%)
End Sub
Sub TS711RadioSetVFOB(Freq&, Handle%)
    freqtext$ = Format(Cdbl2(Freq&), "00000000000")
    Call WriteToPort("FB" + freqtext$ + ";", Handle%)
End Sub
Private Function TS711RadioGetVFOA(Handle%) As Double
    Dim StFreq As String
    Dim DbFreq As Double
    
    Call WriteToPort("FA;", Handle%)
    StFreq = TS711WaitAck(Handle%)
    DbFreq = Cdbl2(Left(Right(StFreq, 12), 11))
    DbFreq = DbFreq / 1000000
    TS711RadioGetVFOA = DbFreq
End Function
Private Function TS711RadioGetVFOB(Handle%) As Double
    Dim StFreq As String
    Dim DbFreq As Double
    
    Call WriteToPort("FB;", Handle%)
    StFreq = TS711WaitAck(Handle%)
    DbFreq = Cdbl2(Left(Right(StFreq, 12), 11))
    DbFreq = DbFreq / 1000000
    TS711RadioGetVFOB = DbFreq
End Function
Private Function TS711WaitAck(Handle%) As String
    a = Timer
    Do
            
    Loop Until Abs(Timer - a) > RadioReplyTimeout(Handle%)
    Select Case Handle%
        Case 1
            InSt$ = MSComm1.Input
        Case 2
            InSt$ = MSComm2.Input
        Case 3
            InSt$ = MSComm3.Input
    End Select
    InLength% = Len(InSt$)
'    MsgBox InSt$ & " : Handle - " & CStr(Handle%)
'    End
    RetString$ = ""
    If InLength% > 0 Then
        If Mid$(InSt$, InLength%, 1) = ";" Then
            RetString$ = ";"
            For i% = InLength% - 1 To 1 Step -1
                If i% <= 0 Or Mid$(InSt$, i%, 1) = ";" Then
                    Exit For
                End If
                RetString$ = Mid$(InSt$, i%, 1) + RetString$
            Next
        End If
    End If
'    MsgBox RetString$
    TS711WaitAck = RetString$
End Function

'TS-790/711/811 Readback routine
'receives chars from radio port until ";" char is detected
'then parses the complete frame received for valid answers.
'The output of this function is a string beginning with the decoded
'frame type code followed by the data available.
'the following type codes can be returned:
'code "00:" - frame not understood pass timeout period.
'code "01:" - acknowlegde frame.
'code "02:" - radio reported a command error.
'code "03:" - radio reported a communications error.
'code "10:" - VFO A (Downlink) frequency information
'code "11:" - VFO B (Uplink) freq. info
'code "40:" - RSSI
'code "41:" - Main (or downlink) RSSI (from TS-2000)
'code "42:" - Sub (or uplink) RSSI (from TS-2000)
Function TS790RadioReadFrame(Handle%) As String
'This will hold the received frame
InBuff$ = ""
'initialize buffer pointer..
InBuffPtr% = 0
'this will indicate if the ";" is found.
f% = 0
'set timeout time...
a = Timer
Do
    'Wait until we receive bytes
    'or a time-out occurs
    Select Case Handle%
    Case 1
        Do
            
        Loop Until MSComm1.InBufferCount >= 1 Or Abs(Timer - a) > RadioReplyTimeout(Handle%)
        InString = MSComm1.Input
    Case 2
        Do
            
        Loop Until MSComm2.InBufferCount >= 1 Or Abs(Timer - a) > RadioReplyTimeout(Handle%)
        InString = MSComm2.Input
    Case 3
        Do
            
        Loop Until MSComm3.InBufferCount >= 1 Or Abs(Timer - a) > RadioReplyTimeout(Handle%)
        InString = MSComm3.Input
    End Select
    'We pass the received bytes to the buffer
    For i% = 0 To LenB(InString) - 1
        InBuff$ = InBuff$ + Chr$(InString(i%))
    Next
    'we will examine the buffer searching for end of frame:
    f% = InStr(InBuff$, ";")
Loop While f% = 0 And Abs(Timer - a) < RadioReplyTimeout(Handle%)
'if loop ended due to timeout:
If f% = 0 Then
    s$ = "00:"
    TS790RadioReadFrame = s$
    Exit Function
Else
    'otherwise examine frame:
    'command error:
    a = InStr(LCase$(InBuff$), "?")
    If a <> 0 Then
        s$ = "02:"
        TS790RadioReadFrame = s$
        Exit Function
    End If
    'overflow error:
    a = InStr(LCase$(InBuff$), "O;")
    If a <> 0 Then
        s$ = "02:"
        TS790RadioReadFrame = s$
        Exit Function
    End If
    'comms error;
    a = InStr(LCase$(InBuff$), "E;")
    If a <> 0 Then
        s$ = "03:"
        TS790RadioReadFrame = s$
        Exit Function
    End If
    'VFO A info:
    a = InStr(InBuff$, "FA")
    If a <> 0 Then
        'check answer string is complete or just an ack:
        If Len(InBuff$) = 3 Then
            s$ = "01:"
            TS790RadioReadFrame = s$
            Exit Function
        ElseIf Len(InBuff$) = 14 Then
            R# = Cdbl2(Mid$(InBuff$, a + 2, 11))
            'we have all the info...
            s$ = "10: " + Str(R#)
            TS790RadioReadFrame = s$
            Exit Function
        End If
    End If
    'VFO B info:
    a = InStr(InBuff$, "FB")
    If a <> 0 Then
        'check answer string is complete or just an ack:
        If Len(InBuff$) = 3 Then
            s$ = "01:"
            TS790RadioReadFrame = s$
            Exit Function
        ElseIf Len(InBuff$) = 14 Then
            R# = Cdbl2(Mid$(InBuff$, a + 2, 11))
            'we have all the info...
            s$ = "11: " + Str(R#)
            TS790RadioReadFrame = s$
            Exit Function
        End If
    End If
    'S-meter info from TS-2000:
    a = InStr(InBuff$, "SM0")
    If a <> 0 Then
        'check answer string is complete (8 chars):
        'f% marks end of answer frame (location of ";")
        If f% - a = 7 Then
            b = Cdbl2(Mid$(InBuff$, a + 3, 4))
            'we have all the info...
            s$ = "41: " + Str(b)
            TS790RadioReadFrame = s$
            Exit Function
        Else
            'if there are not 8 chars, answer not understood
            s$ = "00:"
            TS790RadioReadFrame = s$
            Exit Function
        End If
    End If
    'S-meter info from TS-2000:
    a = InStr(InBuff$, "SM1")
    If a <> 0 Then
        'check answer string is complete (8 chars):
        'f% marks end of answer frame (location of ";")
        If f% - a = 7 Then
            b = Cdbl2(Mid$(InBuff$, a + 3, 4))
            'we have all the info...
            s$ = "42: " + Str(b)
            TS790RadioReadFrame = s$
            Exit Function
        Else
            'if there are not 8 chars, answer not understood
            s$ = "00:"
            TS790RadioReadFrame = s$
            Exit Function
        End If
    End If
    'S-meter info from TS-790:
    a = InStr(InBuff$, "SM")
    If a <> 0 Then
        'check answer string is complete (7 chars):
        'f% marks end of answer frame (location of ";")
        If f% - a = 6 Then
            b = Cdbl2(Mid$(InBuff$, a + 2, 4))
            'we have all the info...
            s$ = "40: " + Str(b)
            TS790RadioReadFrame = s$
            Exit Function
        Else
            'if there are not 7 chars, answer not understood
            s$ = "00:"
            TS790RadioReadFrame = s$
            Exit Function
        End If
    End If
    'if it is non of the above types of frame -> time out error
    s$ = "00:"
    TS790RadioReadFrame = s$
End If

End Function

Sub ParkDownlinkRadio()

' Check if logging is active
If frmRadio.CheckLog.Value Then
    f% = FreeFile
    Open "Radio_Log.txt" For Append As f%
    Print #f%, "Parking IC-821 downlink radio on stream" + Str(DownlinkHandle%) + ", " + Str(Time)
    Close f%
End If
        
'deactivate s-meter bargraph:
DownlinkRSSI.Enabled = False
DownlinkRSSI.Value = 0
Select Case DownlinkModel$
Case "IC-821", "IC-970"
    Call IC821RadioSub(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
    Call IC821RadioMem(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
    Call IC821RadioMain(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
    Call IC821RadioVFO(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)

Case "IC-910"
    Call IC821RadioSub(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
    Call IC821RadioVFO(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
    Call IC821RadioMem(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)

Case "IC-275", "IC-475", "IC-746", "IC-706"
    Call IC821RadioVFO(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
    Call IC821RadioMem(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)

Case "IC-R8500"
    'Call ICR8500RadioOff(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)

Case "FT-817", "FT-897"
    Call FT817RadioSetSplitOff(DownlinkHandle%)

Case "FT-847"
'    Call FT847RadioSatOff(DownlinkHandle%)
    Call FT847RadioCATOff(DownlinkHandle%)

Case "FT-736"
'    Call FT847RadioSatOff(DownlinkHandle%)
    Call FT847RadioCATOff(DownlinkHandle%)

Case "AR-8000", "AR-5000"
    Call AR8000RadioOff(DownlinkHandle%)

Case "PCR-1000"
    Call ICPCRRadioOff(DownlinkHandle%)

Case "TM-D700", "TH-D7"
Case "TS-790"
'    Call TS790RadioSetSub(DownlinkHandle%)
'    Call TS790RadioSelectMEM(DownlinkHandle%)

Case "TS-2000"
'    Call TS2000RadioSatOff(DownlinkHandle%, DownlinkBidir%) ' 13 Feb 2002 G6LVB Revert back to previous non-satellite setting
End Select

End Sub
Sub ParkUplinkRadio()

' Check if logging is active
If frmRadio.CheckLog.Value Then
    f% = FreeFile
    Open "Radio_Log.txt" For Append As f%
    Print #f%, "Parking IC-821 uplink radio on stream" + Str(UplinkHandle%) + ", " + Str(Time)
    Close f%
End If
    
Select Case UplinkModel$
Case "IC-821", "IC-970"
    Call IC821RadioSub(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
    Call IC821RadioMain(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
    Call IC821RadioVFO(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
    Call IC821RadioMem(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)

Case "IC-910"
    Call IC821RadioMain(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
    Call IC821RadioSub(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
    Call IC821RadioVFO(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
    Call IC821RadioMem(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)

Case "IC-275", "IC-475", "IC-746", "IC-706"
    Call IC821RadioVFO(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
    Call IC821RadioMem(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)

Case "FT-817", "FT-897"
    Call FT817RadioSetSplitOff(UplinkHandle%)

Case "FT-847", "FT-736"
'    Call FT847RadioSatOff(UplinkHandle%)
    Call FT847RadioCATOff(UplinkHandle%)

Case "TM-D700", "TH-D7"
Case "TS-790"
'    Call TS790RadioSetMain(UplinkHandle%)
'    Call TS790RadioSelectMEM(UplinkHandle%)
End Select

End Sub
Sub ActivateDownlinkRadio()

' Check if logging is active
If frmRadio.CheckLog.Value Then
    f% = FreeFile
    Open "Radio_Log.txt" For Append As f%
    Print #f%, "Activating IC-821 downlink radio on stream" + Str(DownlinkHandle%) + ", " + Str(Time)
    Close f%
End If
    
Select Case DownlinkModel$
Case "IC-821", "IC-970"
'   Icom radios can be controlled with an unidirectional interface
'   in which data can only flow PC -> Radio.
'   There is a configuration option to support this kind of interface
    DownlinkRSSI.Enabled = False
    
    If DownlinkSplit% = 1 Then
        'if downlink radio is set to split-mode
        'We will use duplex mode instead
        Call IC821RadioVFO(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
    Else
        Call IC821RadioCancelDuplex(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
        Call IC821RadioSub(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
    End If

    If DownlinkBidir% Then
        If Cdbl2(DownlinkDDEFreq.text) = 0 Then
            'if no DDE freq. available for radio -> read freq. from radio
            DownlinkFreq.text = Str$(DownlinkLO# + IC821RadioReadFreq(DownlinkCIVAddress%, DownlinkHandle%) / 1000000#)
        Else
            DownlinkFreq.text = DownlinkDDEFreq.text
        End If
        DownlinkCorrection = 0
        'get the freq.limits of the band under
        'focus. If freq. is out of limits
        'we will exchange MAIN<->SUB
        a = IC821RadioReadBandLimits(DownlinkCIVAddress%, DownlinkHandle%)
        If (a(1) + 1000000# * DownlinkLO#) > 1000000# * Cdbl2(DownlinkFreq.text) Or _
            (a(2) + 1000000# * DownlinkLO#) < 1000000# * Cdbl2(DownlinkFreq.text) Then
            'This will exchange MAIN<->SUB bands
            Call IC821RadioMS(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
        End If
        Call IC821RadioSetFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
    Else
        Call IC821RadioSetFreq(1000000# * (Cdbl2(DownlinkDDEFreq.text) - DownlinkLO#), DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
    End If
    Call IC821RadioSetMode(DownlinkMode.text, DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)

Case "IC-910"
'  ***IC910 maps main/sub bands straigh dislike 821 which inverts them in SAT mode.***
'   Icom radios can be controlled with an unidirectional interface
'   in which data can only flow PC -> Radio.
'   There is a configuration option to support this kind of interface
    DownlinkRSSI.Enabled = False
    
    If DownlinkSplit% = 1 Then
        'if downlink radio is set to split-mode
        'We will use duplex mode instead
        Call IC821RadioVFO(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
    Else
        Call IC821RadioCancelDuplex(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
        Call IC821RadioMain(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
    End If

    If DownlinkBidir% Then
        If Cdbl2(DownlinkDDEFreq.text) = 0 Then
            'if no DDE freq. available for radio -> read freq. from radio
            DownlinkFreq.text = Str$(DownlinkLO# + IC821RadioReadFreq(DownlinkCIVAddress%, DownlinkHandle%) / 1000000#)
        Else
            DownlinkFreq.text = DownlinkDDEFreq.text
        End If
        DownlinkCorrection = 0
        'get the freq.limits of the band under
        'focus. If freq. is out of limits
        'we will exchange MAIN<->SUB
        a = IC821RadioReadBandLimits(DownlinkCIVAddress%, DownlinkHandle%)
        If (a(1) + 1000000# * DownlinkLO#) > 1000000# * Cdbl2(DownlinkFreq.text) Or _
            (a(2) + 1000000# * DownlinkLO#) < 1000000# * Cdbl2(DownlinkFreq.text) Then
            'This will exchange MAIN<->SUB bands
            Call IC821RadioMS(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
        End If
        Call IC821RadioSetFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
    Else
        Call IC821RadioSetFreq(1000000# * (Cdbl2(DownlinkDDEFreq.text) - DownlinkLO#), DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
    End If
    Call IC821RadioSetMode(DownlinkMode.text, DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)

Case "IC-275", "IC-475", "IC-746", "IC-706"
'   Icom radios can be controlled in with an unidirectional interface
'   in which data can only flow PC -> Radio.
'   There is a configuration option to support this kind of interface
    DownlinkRSSI.Enabled = False
    
    If DownlinkSplit% = 1 Then
        'if downlink radio is set to split-mode
        'We will use duplex mode instead
        Call IC821RadioVFO(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
    Else
        Call IC821RadioCancelDuplex(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
    End If
    
    If DownlinkBidir% Then
       If Cdbl2(DownlinkDDEFreq.text) = 0 Then
           'if no DDE freq. available for radio -> read freq. from radio
           DownlinkFreq.text = Str$(DownlinkLO# + IC821RadioReadFreq(DownlinkCIVAddress%, DownlinkHandle%) / 1000000#)
       Else
            DownlinkFreq.text = DownlinkDDEFreq.text
       End If
        DownlinkCorrection = 0
       'get the freq.limits of the band under
       'focus. If reqd. freq. is out of limits
       'we will not update the rig
       a = IC821RadioReadBandLimits(DownlinkCIVAddress%, DownlinkHandle%)
       If (a(1) + 1000000# * DownlinkLO#) > 1000000# * Cdbl2(DownlinkFreq.text) Or _
           (a(2) + 1000000# * DownlinkLO#) < 1000000# * Cdbl2(DownlinkFreq.text) Then
       Else
           Call IC821RadioSetFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
       End If
    Else
       Call IC821RadioSetFreq(1000000# * (Cdbl2(DownlinkDDEFreq.text) - DownlinkLO#), DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
    End If
    Call IC821RadioSetMode(DownlinkMode.text, DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)

Case "IC-R8500"
    DownlinkRSSI.Enabled = False
    Call ICR8500RadioOn(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
    If DownlinkBidir% Then
       If Cdbl2(DownlinkDDEFreq.text) = 0 Then
           'if no DDE freq. available for radio -> read freq. from radio
           DownlinkFreq.text = Str$(DownlinkLO# + IC821RadioReadFreq(DownlinkCIVAddress%, DownlinkHandle%) / 1000000#)
       Else
            DownlinkFreq.text = DownlinkDDEFreq.text
       End If
       DownlinkCorrection = (Cdbl2(DownlinkFreq.text) - Cdbl2(DownlinkDDEFreq.text)) * 1000000#
       'get the freq.limits of the band under
       'focus. If reqd. freq. is out of limits
       'we will not update the rig
       a = IC821RadioReadBandLimits(DownlinkCIVAddress%, DownlinkHandle%)
       If (a(1) + 1000000# * DownlinkLO#) > 1000000# * Cdbl2(DownlinkFreq.text) Or _
           (a(2) + 1000000# * DownlinkLO#) < 1000000# * Cdbl2(DownlinkFreq.text) Then
       Else
           Call IC821RadioSetFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
           Call ICR8500RadioSetMode(DownlinkMode.text, DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
       End If
    Else
       Call IC821RadioSetFreq(1000000# * (Cdbl2(DownlinkDDEFreq.text) - DownlinkLO#), DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
       Call ICR8500RadioSetMode(DownlinkMode.text, DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
    End If

Case "IC-R7000"
    DownlinkRSSI.Enabled = False
    Call ICR8500RadioOn(DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
    If DownlinkBidir% Then
       If Cdbl2(DownlinkDDEFreq.text) = 0 Then
           'if no DDE freq. available for radio -> read freq. from radio
           DownlinkFreq.text = Str$(DownlinkLO# + IC821RadioReadFreq(DownlinkCIVAddress%, DownlinkHandle%) / 1000000#)
       Else
            DownlinkFreq.text = DownlinkDDEFreq.text
       End If
       DownlinkCorrection = (Cdbl2(DownlinkFreq.text) - Cdbl2(DownlinkDDEFreq.text)) * 1000000#
       'get the freq.limits of the band under
       'focus. If reqd. freq. is out of limits
       'we will not update the rig
       a = IC821RadioReadBandLimits(DownlinkCIVAddress%, DownlinkHandle%)
       If (a(1) + 1000000# * DownlinkLO#) > 1000000# * Cdbl2(DownlinkFreq.text) Or _
           (a(2) + 1000000# * DownlinkLO#) < 1000000# * Cdbl2(DownlinkFreq.text) Then
       Else
           Call IC821RadioSetFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
           Call ICR7000RadioSetMode(DownlinkMode.text, DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
       End If
    Else
       Call IC821RadioSetFreq(1000000# * (Cdbl2(DownlinkDDEFreq.text) - DownlinkLO#), DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
       Call ICR7000RadioSetMode(DownlinkMode.text, DownlinkCIVAddress%, DownlinkBidir%, DownlinkHandle%)
    End If
Case "FT-847"
    DownlinkRSSI.Enabled = False
    Call FT847RadioCATOn(DownlinkHandle%)
    Call FT847RadioCTCSSRXOff(DownlinkHandle%)
    
    If Cdbl2(DownlinkDDEFreq.text) = 0 Then
        'if no DDE freq. available for radio -> read freq. from radio
        DownlinkFreq.text = Str$(DownlinkLO# + FT847RadioReadRXFreq(DownlinkHandle%) / 1000000#)
    Else
        DownlinkFreq.text = DownlinkDDEFreq.text
    End If
    DownlinkCorrection = (Cdbl2(DownlinkFreq.text) - Cdbl2(DownlinkDDEFreq.text)) * 1000000#
    
    If DownlinkSplit% = 1 Then
        'if downlink radio is set to split-mode
        'We will use Repeater Offset for uplink freq.
        Call FT847RadioSatOff(DownlinkHandle%)
        'Call FT847RadioShiftPlus(DownlinkHandle%)
        Call FT847RadioSetMainFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%)
        Call FT847RadioSetMainMode(DownlinkMode.text, DownlinkHandle%)
    Else
        Call FT847RadioSatOn(DownlinkHandle%)
        Call FT847RadioShiftOff(DownlinkHandle%)
        Call FT847RadioSetRXFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%)
        Call FT847RadioSetRXMode(DownlinkMode.text, DownlinkHandle%)
    End If
    

Case "FT-817"
    Call FT817RadioWakeUp(DownlinkHandle%)
    
    If DownlinkSplit% = 1 Then
        Call FT817RadioSetSplitOn(DownlinkHandle%)
    Else
        Call FT817RadioSetSplitOff(DownlinkHandle%)
    End If
    If Cdbl2(DownlinkDDEFreq.text) <> 0 Then
        DownlinkFreq.text = DownlinkDDEFreq.text
    Else
        'if no DDE freq. available for radio -> read freq. from radio
        DownlinkFreq.text = Str$(DownlinkLO# + FT817RadioReadFreq(DownlinkHandle%) / 1000000#)
    End If
    DownlinkCorrection = (Cdbl2(DownlinkFreq.text) - Cdbl2(DownlinkDDEFreq.text)) * 1000000#
    
    Call FT817RadioSetFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%)
    Call FT817RadioSetMode(DownlinkMode.text, DownlinkHandle%)
    
Case "FT-897"
    If DownlinkSplit% = 1 Then
        Call FT817RadioSetSplitOn(DownlinkHandle%)
    Else
        Call FT817RadioSetSplitOff(DownlinkHandle%)
    End If
    If Cdbl2(DownlinkDDEFreq.text) <> 0 Then
        DownlinkFreq.text = DownlinkDDEFreq.text
    Else
        'if no DDE freq. available for radio -> read freq. from radio
        DownlinkFreq.text = Str$(DownlinkLO# + FT817RadioReadFreq(DownlinkHandle%) / 1000000#)
    End If
    DownlinkCorrection = (Cdbl2(DownlinkFreq.text) - Cdbl2(DownlinkDDEFreq.text)) * 1000000#
    
    Call FT817RadioSetFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%)
    Call FT897RadioSetMode(DownlinkMode.text, DownlinkHandle%)
    
Case "FT-736"
    DownlinkRSSI.Enabled = False
    Call FT847RadioCATOn(DownlinkHandle%)
'    Call FT736RadioSatOn(DownlinkHandle%)
    If DownlinkDDEFreq.text <> "" Then
        DownlinkFreq.text = DownlinkDDEFreq.text
    End If
    DownlinkCorrection = (Cdbl2(DownlinkFreq.text) - Cdbl2(DownlinkDDEFreq.text)) * 1000000#
    Call FT736RadioSetRXFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%)
    Call FT847RadioSetRXMode(DownlinkMode.text, DownlinkHandle%)

Case "FT-100"
    If DownlinkSplit% = 1 Then
        'if downlink radio is set to split-mode
        Call FT100RadioSetSplitOn(DownlinkHandle%)
        'We will use VFO-A of FT100 as downlink band:
        Call FT100RadioSetVFOA(DownlinkHandle%)
    Else
        Call FT100RadioSetSplitOff(DownlinkHandle%)
    End If
    DownlinkRSSI.Enabled = False
    If Cdbl2(DownlinkDDEFreq.text) = 0 Then
        'if no DDE freq. available for radio -> read freq. from radio
        DownlinkFreq.text = Str$(DownlinkLO# + FT100RadioReadFreq(DownlinkHandle%) / 1000000#)
    Else
        DownlinkFreq.text = DownlinkDDEFreq.text
    End If
    DownlinkCorrection = (Cdbl2(DownlinkFreq.text) - Cdbl2(DownlinkDDEFreq.text)) * 1000000#
    Call FT100RadioSetFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%)
    Call FT100RadioSetMode(DownlinkMode.text, DownlinkHandle%)

Case "TM-D700", "TH-D7"
    DownlinkRSSI.Enabled = False
    DownlinkFreq.text = DownlinkDDEFreq.text
    DownlinkCorrection = (Cdbl2(DownlinkFreq.text) - Cdbl2(DownlinkDDEFreq.text)) * 1000000#
    'only Band B can handle UHF frequencies
    If Cdbl2(DownlinkFreq.text) >= 300 Then
        Call TMD700RadioSetB(DownlinkHandle%)
        Call TMD700RadioSetATX_BRX(DownlinkHandle%)
    Else
        Call TMD700RadioSetA(DownlinkHandle%)
        Call TMD700RadioSetBTX_ARX(DownlinkHandle%)
    End If
    Call TMD700RadioCancelTone(DownlinkHandle%)
    Call TMD700RadioCancelSplit(DownlinkHandle%)

Case "TS-790"
    DownlinkRSSI.Enabled = True
    Call TS790RadioSetSub(DownlinkHandle%, DownlinkBidir%)
    'Call TS790RadioCancelScan(DownlinkHandle%, DownlinkBidir%)
    'Call TS790RadioCancelSplit(DownlinkHandle%, DownlinkBidir%)
    If Cdbl2(DownlinkDDEFreq.text) = 0 Then
        'if no DDE freq. available for radio -> read freq. from radio
        DownlinkFreq.text = Str$(DownlinkLO# + TS790RadioReadVFOA(DownlinkHandle%) / 1000000#)
    Else
        DownlinkFreq.text = DownlinkDDEFreq.text
    End If
    DownlinkCorrection = (Cdbl2(DownlinkFreq.text) - Cdbl2(DownlinkDDEFreq.text)) * 1000000#
    Call TS790RadioSetVFOA(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%, DownlinkBidir%)
    Call TS790RadioSetMode(DownlinkMode.text, DownlinkHandle%, DownlinkBidir%)

Case "TS-2000"
    DownlinkRSSI.Enabled = True
'    Call TS2000RadioSatOn(DownlinkHandle%, DownlinkBidir%)
    Call TS2000RadioCancelSplit(DownlinkHandle%, DownlinkBidir%)
    Call TS2000RadioSetMain(DownlinkHandle%, DownlinkBidir%) ' G6LVB 11 Feb 2002 Replaced TS790 call
    'check if downlink freq was initialized:
    If Cdbl2(DownlinkDDEFreq.text) = 0 Then
        'if no DDE freq. available for radio -> read freq. from radio
        DownlinkFreq.text = Str$(DownlinkLO# + TS790RadioReadVFOA(DownlinkHandle%) / 1000000#)
    Else
        DownlinkFreq.text = DownlinkDDEFreq.text
    End If
    
    ' G6LVB 3 Feb 2002 We need to be smart here: TS-2000 will not let you set downlink
    ' to same band as uplink, so we need to move the uplink away somewhere safe...
    ' This assumes that the downlink is set prior to the uplink. This needs to be done to
    ' TS-790 code too... this is the result of a nasty boolean truth table, so apologies!
    TS2000TxBandInit$ = BandDesignator(TS790RadioReadVFOB(DownlinkHandle%))
    TS2000RxBandInit$ = BandDesignator(TS790RadioReadVFOA(DownlinkHandle%))
    TS2000RxBandRequest$ = BandDesignator(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#))
    If (TS2000RxBandRequest$ = TS2000TxBandInit$) Then
        If (TS2000RxBandInit$ = "H" And TS2000TxBandInit$ = "V") _
            Or (TS2000RxBandInit$ = "V" And TS2000TxBandInit$ = "H") Then
            Call TS790RadioSetVFOB(436000000#, DownlinkHandle%, DownlinkBidir%)
        Else
            If (TS2000RxBandInit$ = "H" And TS2000TxBandInit$ = "U") _
            Or (TS2000RxBandInit$ = "H" And TS2000TxBandInit$ = "L") _
            Or (TS2000RxBandInit$ = "U" And TS2000TxBandInit$ = "H") _
            Or (TS2000RxBandInit$ = "L" And TS2000TxBandInit$ = "H") Then
                Call TS790RadioSetVFOB(145900000#, DownlinkHandle%, DownlinkBidir%)
            Else
                Call TS790RadioSetVFOB(29500000#, DownlinkHandle%, DownlinkBidir%)
            End If
        End If
    End If
    
    DownlinkCorrection = (Cdbl2(DownlinkFreq.text) - Cdbl2(DownlinkDDEFreq.text)) * 1000000#
    Call TS790RadioSetVFOA(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%, DownlinkBidir%)
    Call TS790RadioSetMode(DownlinkMode.text, DownlinkHandle%, DownlinkBidir%)

Case "AR-8000"
    DownlinkRSSI.Enabled = False
    Call AR8000RadioOn(DownlinkHandle%)
    If DownlinkDDEFreq.text <> "" Then
        DownlinkFreq.text = DownlinkDDEFreq.text
    End If
    DownlinkCorrection = (Cdbl2(DownlinkFreq.text) - Cdbl2(DownlinkDDEFreq.text)) * 1000000#
    Call AR8000RadioSetFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%)
    Call AR8000RadioSetMode(DownlinkMode.text, DownlinkHandle%)

Case "AR-3000A" 'EB2CTA
    DownlinkRSSI.Enabled = True
    Call AR3000ARadioOn(DownlinkHandle%)
    If DownlinkDDEFreq.text <> "" Then
        DownlinkFreq.text = DownlinkDDEFreq.text
    End If
    DownlinkCorrection = (Cdbl2(DownlinkFreq.text) - Cdbl2(DownlinkDDEFreq.text)) * 1000000#
    Call AR3000ARadioSetFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%)
    Call AR3000ARadioSetMode(DownlinkMode.text, DownlinkHandle%)
    
    InString = ReadFromPort(Handle%)
    
Case "AR-5000"
    DownlinkRSSI.Enabled = False
    Call AR8000RadioOn(DownlinkHandle%)
    If DownlinkDDEFreq.text <> "" Then
        DownlinkFreq.text = DownlinkDDEFreq.text
    End If
    DownlinkCorrection = (Cdbl2(DownlinkFreq.text) - Cdbl2(DownlinkDDEFreq.text)) * 1000000#
    Call AR8000RadioSetFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%)
    Call AR5000RadioSetMode(DownlinkMode.text, DownlinkHandle%)
    Call AR5000RadioSetBW(DownlinkFilter, DownlinkHandle%)

Case "PCR-1000"
    DownlinkRSSI.Enabled = True
    Call ICPCRRadioOn(DownlinkHandle%)
    If DownlinkDDEFreq.text <> "" Then
        DownlinkFreq.text = DownlinkDDEFreq.text
    End If
    DownlinkCorrection = (Cdbl2(DownlinkFreq.text) - Cdbl2(DownlinkDDEFreq.text)) * 1000000#
    Call ICPCRRadioSet(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkMode.text, DownlinkFilter, DownlinkHandle%)
    Call ICPCRRadioSetVol(DownlinkVolume%, DownlinkHandle%)
'    DownlinkRSSI.Value = CInt(ICPCRRadioReadRSSI(DownlinkHandle%))

Case "TS-711", "TS-811"
    Call TS711RadioSetStart(DownlinkHandle%, DownlinkBidir%)
    If DownlinkDDEFreq.text <> "" Then
        DownlinkFreq.text = DownlinkDDEFreq.text
    End If
    DownlinkCorrection = (Cdbl2(DownlinkFreq.text) - Cdbl2(DownlinkDDEFreq.text)) * 1000000#
    Call TS711RadioSetMode(DownlinkMode.text, DownlinkHandle%, "Down", DownlinkBidir%)
    Call TS711RadioSetVFOA(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%)

Case "FRG-9600"
    Call FRG9600RadioOn
    If DownlinkDDEFreq.text <> "" Then
        DownlinkFreq.text = DownlinkDDEFreq.text
    End If
    DownlinkCorrection = (Cdbl2(DownlinkFreq.text) - Cdbl2(DownlinkDDEFreq.text)) * 1000000#
    Call FRG9600RadioSetMode(DownlinkMode.text, DownlinkHandle%)

Case "VR-5000"
    DownlinkRSSI.Enabled = True
    Call VR5000RadioCATOn(DownlinkHandle%)
    'as we cannot read VR5000 freq, we copy DDEFreq into it:
    If DownlinkDDEFreq.text <> "" Then
        DownlinkFreq.text = DownlinkDDEFreq.text
    End If
    DownlinkCorrection = (Cdbl2(DownlinkFreq.text) - Cdbl2(DownlinkDDEFreq.text)) * 1000000#
    Call VR5000RadioSetMainFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%)
    Call VR5000RadioSetMainMode(DownlinkMode.text, DownlinkHandle%)

Case "TrakBox"
    If DownlinkDDEFreq.text <> "" Then
        DownlinkFreq.text = DownlinkDDEFreq.text
    End If
    DownlinkCorrection = (Cdbl2(DownlinkFreq.text) - Cdbl2(DownlinkDDEFreq.text)) * 1000000#
    Call TBRadioSetRXFreq(1000000# * (Cdbl2(DownlinkFreq.text) - DownlinkLO#), DownlinkHandle%)
    Call TBRadioSetRXMode(DownlinkMode.text, DownlinkHandle%)
End Select

End Sub
Sub ActivateUplinkRadio()

' Check if logging is active
If frmRadio.CheckLog.Value Then
    f% = FreeFile
    Open "Radio_Log.txt" For Append As f%
    Print #f%, "Activating IC-821 uplink radio on stream" + Str(UplinkHandle%) + ", " + Str(Time)
    Close f%
End If
    
Select Case UplinkModel$
Case "IC-821", "IC-970"
    If UplinkSplit% = 1 Then
        'if downlink radio is set to split-mode
        'We will duplex mode
        Call IC821RadioVFO(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
    Else
        Call IC821RadioCancelDuplex(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
        Call IC821RadioMain(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
    End If

    If UplinkBidir% Then
        If Cdbl2(UplinkDDEFreq.text) = 0 Then
            'if no DDE freq. available for radio -> read freq. from radio
            If UplinkSplit% = 1 Then
                'to get uplink freq in duplex se need to read offset and RX freq
                d# = IC821RadioReadOffset(UplinkCIVAddress%, UplinkHandle%)
                'we don't know what sign the offset has, assume positive
            Else
                d# = 0
            End If
            
            d# = d# + IC821RadioReadFreq(UplinkCIVAddress%, UplinkHandle%)
            If d# <> 0 Then
                d# = d# + 1000000# * UplinkLO#
            End If
            
            UplinkFreq.text = Str$(d# / 1000000#)
        Else
            UplinkFreq.text = UplinkDDEFreq.text
        End If
        UplinkCorrection = 0
        'get the freq.limits of the band under
        'focus. If freq. is out of limits we'll try
        'swapping bands
        a = IC821RadioReadBandLimits(UplinkCIVAddress%, UplinkHandle%)
        If (a(1) + 1000000# * UplinkLO#) > (1000000# * Cdbl2(UplinkFreq.text)) Or _
            (a(2) + 1000000# * UplinkLO#) < (1000000# * Cdbl2(UplinkFreq.text)) Then
            'This will exchange MAIN<->SUB bands
            Call IC821RadioMS(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
        End If
        Call IC821RadioSetMode(UplinkMode.text, UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
        If UplinkSplit% = 1 Then
            'if downlink radio is set to split-mode
            'we'll use duplex mode
            'need to know if TX is greater or lower than RX
            If ((Cdbl2(UplinkFreq.text) - UplinkLO#) > (Cdbl2(DownlinkFreq.text) - DownlinkLO#)) Then
                Call IC821RadioSetDuplexPlus(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
                UplinkDuplexPlus% = 1
            Else
                Call IC821RadioSetDuplexMinus(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
                UplinkDuplexPlus% = 0
            End If
            Call IC821RadioSetOffset(1000000# * ((Cdbl2(UplinkFreq.text) - UplinkLO#) - (Cdbl2(DownlinkFreq.text) - DownlinkLO#)) _
                , UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
        Else
            Call IC821RadioSetFreq(1000000# * (Cdbl2(UplinkFreq.text) - UplinkLO#), UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
        End If
    Else
        Call IC821RadioSetMode(UplinkMode.text, UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
        If UplinkSplit% = 1 Then
            'if downlink radio is set to split-mode
            'We will duplex mode
            If ((Cdbl2(UplinkDDEFreq.text) - UplinkLO#) > (Cdbl2(DownlinkDDEFreq.text) - DownlinkLO#)) Then
                Call IC821RadioSetDuplexPlus(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
            Else
                Call IC821RadioSetDuplexMinus(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
            End If
            Call IC821RadioSetOffset(1000000# * ((Cdbl2(UplinkDDEFreq.text) - UplinkLO#) - (Cdbl2(DownlinkDDEFreq.text) - DownlinkLO#)) _
                , UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
        Else
            Call IC821RadioSetFreq(1000000# * (Cdbl2(UplinkDDEFreq.text) - UplinkLO#), UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
        End If
    End If
    
    If UplinkSplit% = 1 Then
    Else
        Call IC821RadioSub(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
    End If

Case "IC-910"
'  ***IC910 maps main/sub bands straigh dislike 821 which inverts them in SAT mode.***
    
    If UplinkSplit% = 1 Then
        'if downlink radio is set to split-mode
        'We will duplex mode
        Call IC821RadioVFO(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
    Else
        Call IC821RadioCancelDuplex(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
        Call IC821RadioSub(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
    End If

    If UplinkBidir% Then
        If Cdbl2(UplinkDDEFreq.text) = 0 Then
            'if no DDE freq. available for radio -> read freq. from radio
            If UplinkSplit% = 1 Then
                'to get uplink freq in duplex se need to read offset and RX freq
                d# = IC821RadioReadOffset(UplinkCIVAddress%, UplinkHandle%)
                'we don't know what sign the offset has, assume positive
            Else
                d# = 0
            End If
            
            d# = d# + IC821RadioReadFreq(UplinkCIVAddress%, UplinkHandle%)
            If d# <> 0 Then
                d# = d# + 1000000# * UplinkLO#
            End If
            
            UplinkFreq.text = Str$(d# / 1000000#)
        Else
            UplinkFreq.text = UplinkDDEFreq.text
        End If
        UplinkCorrection = 0
        'get the freq.limits of the band under
        'focus. If freq. is out of limits we'll try
        'swapping bands
        a = IC821RadioReadBandLimits(UplinkCIVAddress%, UplinkHandle%)
        If (a(1) + 1000000# * UplinkLO#) > (1000000# * Cdbl2(UplinkFreq.text)) Or _
            (a(2) + 1000000# * UplinkLO#) < (1000000# * Cdbl2(UplinkFreq.text)) Then
            'This will exchange MAIN<->SUB bands
            Call IC821RadioMS(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
        End If
        Call IC821RadioSetFreq(1000000# * (Cdbl2(UplinkFreq.text) - UplinkLO#), UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
    Else
        Call IC821RadioSetFreq(1000000# * (Cdbl2(UplinkDDEFreq.text) - UplinkLO#), UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
    End If
    Call IC821RadioSetMode(UplinkMode.text, UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
    
    If UplinkSplit% = 1 Then
    Else
        Call IC821RadioMain(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
    End If

Case "IC-275", "IC-475", "IC-746", "IC-706"
    If UplinkSplit% = 1 Then
        'if downlink radio is set to split-mode
        'We will duplex mode
        Call IC821RadioVFO(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
    Else
        Call IC821RadioCancelDuplex(UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
    End If

     If UplinkBidir% Then
        If Cdbl2(UplinkDDEFreq.text) = 0 Then
            'if no DDE freq. available for radio -> read freq. from radio
            If UplinkSplit% = 1 Then
                'to get uplink freq in duplex se need to read offset and RX freq
                d# = IC821RadioReadOffset(UplinkCIVAddress%, UplinkHandle%)
                'we don't know what sign the offset has, assume positive
            Else
                d# = 0
            End If
            
            d# = d# + IC821RadioReadFreq(UplinkCIVAddress%, UplinkHandle%)
            If d# <> 0 Then
                d# = d# + 1000000# * UplinkLO#
            End If
            
            UplinkFreq.text = Str$(d# / 1000000#)
        Else
            UplinkFreq.text = UplinkDDEFreq.text
        End If
        UplinkCorrection = (Cdbl2(UplinkFreq.text) - Cdbl2(UplinkDDEFreq.text)) * 1000000#
        'get the freq.limits of the band under
        'focus. If reqd. freq. is out of limits
        'we will not update the rig
        a = IC821RadioReadBandLimits(UplinkCIVAddress%, UplinkHandle%)
        If (a(1) + 1000000# * UplinkLO#) > 1000000# * Cdbl2(UplinkFreq.text) Or _
            (a(2) + 1000000# * UplinkLO#) < 1000000# * Cdbl2(UplinkFreq.text) Then
        Else
            Call IC821RadioSetFreq(1000000# * (Cdbl2(UplinkFreq.text) - UplinkLO#), UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
            Call IC821RadioSetMode(UplinkMode.text, UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
        End If
    Else
        Call IC821RadioSetFreq(1000000# * (Cdbl2(UplinkDDEFreq.text) - UplinkLO#), UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
        Call IC821RadioSetMode(UplinkMode.text, UplinkCIVAddress%, UplinkBidir%, UplinkHandle%)
    End If

Case "FT-847"
    Call FT847RadioCATOn(UplinkHandle%)
    Call FT847RadioCTCSSTXOff(UplinkHandle%)
    
    If Cdbl2(UplinkDDEFreq.text) = 0 Then
        'if no DDE freq. available for radio -> read freq. from radio
        UplinkFreq.text = Str$(UplinkLO# + FT847RadioReadTXFreq(UplinkHandle%) / 1000000#)
    Else
        UplinkFreq.text = UplinkDDEFreq.text
    End If
    UplinkCorrection = (Cdbl2(UplinkFreq.text) - Cdbl2(UplinkDDEFreq.text)) * 1000000#
    
    If UplinkSplit% = 1 Then
        'if uplink radio is set to split-mode
        'We will use Repeater Offset to obtain correct TX freq.
        'Need to calculate difference between tx and rx freqs.
        a = 1000000# * ((Cdbl2(UplinkFreq.text) - UplinkLO#) - (Cdbl2(DownlinkFreq.text) - DownlinkLO#))
        If a > 0 Then
            Call FT847RadioShiftPlus(UplinkHandle%)
        Else
            Call FT847RadioShiftMinus(UplinkHandle%)
        End If
        Call FT847RadioSetShiftFreq(CLng(Abs(a)), UplinkHandle%)
    Else
        Call FT847RadioShiftOff(UplinkHandle%)
        Call FT847RadioSetTXFreq(1000000# * (Cdbl2(UplinkFreq.text) - UplinkLO#), UplinkHandle%)
        Call FT847RadioSetTXMode(UplinkMode.text, UplinkHandle%)
    End If

Case "FT-736"
    Call FT847RadioCATOn(UplinkHandle%)
'    Call FT736RadioSatOn(UplinkHandle%)
    UplinkFreq.text = UplinkDDEFreq.text
    UplinkCorrection = (Cdbl2(UplinkFreq.text) - Cdbl2(UplinkDDEFreq.text)) * 1000000#
    Call FT736RadioSetTXFreq(1000000# * (Cdbl2(UplinkFreq.text) - UplinkLO#), UplinkHandle%)
    Call FT847RadioSetTXMode(UplinkMode.text, UplinkHandle%)

Case "FT-100"
    If UplinkSplit% = 1 Then
        'if uplink radio is set to split-mode
        Call FT100RadioSetSplitOn(UplinkHandle%)
        'VFO-B will be used as uplink band:
        Call FT100RadioSetVFOB(UplinkHandle%)
    Else
        Call FT100RadioSetSplitOff(UplinkHandle%)
    End If
    If Cdbl2(UplinkDDEFreq.text) = 0 Then
        'if no DDE freq. available -> read freq. from radio
        UplinkFreq.text = Str$(UplinkLO# + FT100RadioReadFreq(UplinkHandle%) / 1000000#)
    Else
        UplinkFreq.text = UplinkDDEFreq.text
    End If
    UplinkCorrection = (Cdbl2(UplinkFreq.text) - Cdbl2(UplinkDDEFreq.text)) * 1000000#
    Call FT100RadioSetFreq(1000000# * (Cdbl2(UplinkFreq.text) - UplinkLO#), UplinkHandle%)
    Call FT100RadioSetMode(UplinkMode.text, UplinkHandle%)

Case "FT-817"
    Call FT817RadioWakeUp(UplinkHandle%)
        
    If UplinkSplit% = 1 Then
        Call FT817RadioSetSplitOn(UplinkHandle%)
    Else
        Call FT817RadioSetSplitOff(UplinkHandle%)
    End If
    
    If UplinkSplit% = 1 Then
        Call FT817RadioToggleVFO(UplinkHandle%)
    End If
    
    If Cdbl2(UplinkDDEFreq.text) <> 0 Then
        UplinkFreq.text = UplinkDDEFreq.text
    Else
        'if no DDE freq. available -> read freq. from radio
        UplinkFreq.text = Str$(UplinkLO# + FT817RadioReadFreq(UplinkHandle%) / 1000000#)
    End If
    UplinkCorrection = (Cdbl2(UplinkFreq.text) - Cdbl2(UplinkDDEFreq.text)) * 1000000#
    
    Call FT817RadioSetFreq(1000000# * (Cdbl2(UplinkFreq.text) - UplinkLO#), UplinkHandle%)
    Call FT817RadioSetMode(UplinkMode.text, UplinkHandle%)

    If UplinkSplit% = 1 Then
        Call FT817RadioToggleVFO(UplinkHandle%)
    End If
    
Case "FT-897"
    If UplinkSplit% = 1 Then
        Call FT817RadioSetSplitOn(UplinkHandle%)
    Else
        Call FT817RadioSetSplitOff(UplinkHandle%)
    End If
    
    If UplinkSplit% = 1 Then
        Call FT817RadioToggleVFO(UplinkHandle%)
    End If
    
    If Cdbl2(UplinkDDEFreq.text) <> 0 Then
        UplinkFreq.text = UplinkDDEFreq.text
    Else
        'if no DDE freq. available -> read freq. from radio
        UplinkFreq.text = Str$(UplinkLO# + FT817RadioReadFreq(UplinkHandle%) / 1000000#)
    End If
    UplinkCorrection = (Cdbl2(UplinkFreq.text) - Cdbl2(UplinkDDEFreq.text)) * 1000000#
    
    Call FT817RadioSetFreq(1000000# * (Cdbl2(UplinkFreq.text) - UplinkLO#), UplinkHandle%)
    Call FT897RadioSetMode(UplinkMode.text, UplinkHandle%)

    If UplinkSplit% = 1 Then
        Call FT817RadioToggleVFO(UplinkHandle%)
    End If
    
Case "TM-D700", "TH-D7"
    UplinkFreq.text = UplinkDDEFreq.text
    UplinkCorrection = (Cdbl2(UplinkFreq.text) - Cdbl2(UplinkDDEFreq.text)) * 1000000#
    If Cdbl2(UplinkFreq.text) > 300 Then
        Call TMD700RadioSetB(UplinkHandle%)
        Call TMD700RadioSetBTX_ARX(UplinkHandle%)
    Else
        Call TMD700RadioSetA(UplinkHandle%)
        Call TMD700RadioSetATX_BRX(UplinkHandle%)
    End If
    Call TMD700RadioCancelTone(UplinkHandle%)
    Call TMD700RadioCancelSplit(UplinkHandle%)
    If Cdbl2(UplinkFreq.text) >= 300 Then
        Call TMD700RadioSetB(UplinkHandle%)
    Else
        Call TMD700RadioSetA(UplinkHandle%)
    End If
    Call TMD700RadioSetFreq(1000000# * (Cdbl2(UplinkFreq.text) - UplinkLO#), UplinkHandle%)

Case "TS-790"
    Call TS790RadioSetMain(UplinkHandle%, UplinkBidir%)
    'Call TS790RadioCancelScan(UplinkHandle%, UplinkBidir%)
    'Call TS790RadioCancelSplit(UplinkHandle%, UplinkBidir%)
    If Cdbl2(UplinkDDEFreq.text) = 0 Then
        'if no DDE freq. available for radio -> read freq. from radio
        UplinkFreq.text = Str$(UplinkLO# + TS790RadioReadVFOA(UplinkHandle%) / 1000000#)
    Else
        UplinkFreq.text = UplinkDDEFreq.text
    End If
    UplinkCorrection = (Cdbl2(UplinkFreq.text) - Cdbl2(UplinkDDEFreq.text)) * 1000000#
    Call TS790RadioSetVFOA(1000000# * (Cdbl2(UplinkFreq.text) - UplinkLO#), UplinkHandle%, UplinkBidir%)
    Call TS790RadioSetMode(UplinkMode.text, UplinkHandle%, UplinkBidir%)

Case "TS-2000"
'    Call TS2000RadioSatOn(UplinkHandle%, UplinkBidir%)
    Call TS2000RadioCancelSplit(UplinkHandle%, UplinkBidir%)
    Call TS2000RadioSetSub(UplinkHandle%, UplinkBidir%)
    'check if uplink freq was initialized:
    If Cdbl2(UplinkDDEFreq.text) = 0 Then
        'if no DDE freq. available for radio -> read freq. from radio
        UplinkFreq.text = Str$(UplinkLO# + TS790RadioReadVFOA(UplinkHandle%) / 1000000#)
    Else
        UplinkFreq.text = UplinkDDEFreq.text
    End If
    UplinkCorrection = (Cdbl2(UplinkFreq.text) - Cdbl2(UplinkDDEFreq.text)) * 1000000#
    Call TS790RadioSetVFOB(1000000# * (Cdbl2(UplinkFreq.text) - UplinkLO#), UplinkHandle%, UplinkBidir%)
    Call TS790RadioSetMode(UplinkMode.text, UplinkHandle%, UplinkBidir%)

Case "TS-711", "TS-811"
    Call TS711RadioSetStart(UplinkHandle%, UplinkBidir%)
    If UplinkDDEFreq.text <> "" Then
        UplinkFreq.text = UplinkDDEFreq.text
    End If
    UplinkCorrection = (Cdbl2(UplinkFreq.text) - Cdbl2(UplinkDDEFreq.text)) * 1000000#
    Call TS711RadioSetMode(UplinkMode.text, UplinkHandle%, "Up", UplinkBidir%)

Case "TrakBox"
    If UplinkDDEFreq.text <> "" Then
        UplinkFreq.text = UplinkDDEFreq.text
    End If
    UplinkCorrection = (Cdbl2(UplinkFreq.text) - Cdbl2(UplinkDDEFreq.text)) * 1000000#
    Call TBRadioSetTXFreq(1000000# * (Cdbl2(UplinkFreq.text) - UplinkLO#), UplinkHandle%)
    Call TBRadioSetTXMode(UplinkMode.text, UplinkHandle%)
End Select

End Sub
Function ReadFromPort(Handle%)
Select Case Handle%
Case 1
    ReadFromPort = MSComm1.Input
Case 2
    ReadFromPort = MSComm2.Input
Case 3
    ReadFromPort = MSComm3.Input
End Select
End Function

Sub WriteToPort(s$, Handle%)
Select Case Handle%
Case 1
    MSComm1.Output = s$
Case 2
    MSComm2.Output = s$
Case 3
    MSComm3.Output = s$
End Select
End Sub

Sub WaitOutBuffEmpty(Handle%)
Select Case Handle%
Case 1
    a = Timer
    Do
    Loop Until (Abs(Timer - a) > RadioControlDelay(Handle%) Or _
        MSComm1.OutBufferCount = 0)
Case 2
    a = Timer
    Do
    Loop Until (Abs(Timer - a) > RadioControlDelay(Handle%) Or _
        MSComm2.OutBufferCount = 0)
Case 3
    a = Timer
    Do
    Loop Until (Abs(Timer - a) > RadioControlDelay(Handle%) Or _
        MSComm3.OutBufferCount = 0)
End Select

End Sub
Function OpenDownlinkPort() As Integer
    ' Check if logging is active
    If frmRadio.CheckLog.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Opening downlink port, " + Str(Time)
        Close f%
    End If
    
    'place general settings for this radio into more
    'convenient variables
    DownlinkPortName$ = GetSetting("WiSP_DDE_Client", "Rig" + DownlinkIndex.text, "Radio_port")
    DownlinkPort% = Cdbl2(Right(DownlinkPortName$, Len(DownlinkPortName$) - 3))
    DownlinkModel$ = GetSetting("WiSP_DDE_Client", "Rig" + DownlinkIndex.text, "Radio_model", "None")
    DownlinkBaud& = Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + DownlinkIndex.text, "Radio_baud"))
    DownlinkCIVAddress% = Cdbl2("&h" + GetSetting("WiSP_DDE_Client", "Rig" + DownlinkIndex.text, "Radio_address"))
    DownlinkBidir% = Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + DownlinkIndex.text, "Radio_bidir"))
    DownlinkTNCUD% = Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + DownlinkIndex.text, "Radio_tncupdn"))
    DownlinkLO# = Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + DownlinkIndex.text, "Radio_DownlinkLO", ""))
    DownlinkSplit% = Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + DownlinkIndex.text, "Radio_Split", 0))
'    RadioControlLoopTimer.Interval = Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + DownlinkIndex.text, "Radio_Delay", 1000))
    If SliderDownlink.Value = 1 Then
        RadioControlLoopTimer.Interval = Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + DownlinkIndex.text, "Radio_delay"))
    End If
'    If Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + DownlinkIndex.text, "Radio_Antenna", "1")) = 1 Then
'        RotorAuto.Value = 1
'    End If
    DownlinkVolume% = Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + DownlinkIndex.text, "Radio_Volume"))
    Select Case UCase(DownlinkMode.text)
    Case "USB", "LSB"
        DownlinkFilter = Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + DownlinkIndex.text, "Radio_ssbfilter"))
    Case "CW"
        DownlinkFilter = Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + DownlinkIndex.text, "Radio_cwfilter"))
    Case "CW-N"
        DownlinkFilter = Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + DownlinkIndex.text, "Radio_cwnfilter"))
    Case "FM-N"
        DownlinkFilter = Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + DownlinkIndex.text, "Radio_fmnfilter"))
    Case "FM"
        DownlinkFilter = Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + DownlinkIndex.text, "Radio_fmfilter"))
    Case "FM-W"
        DownlinkFilter = Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + DownlinkIndex.text, "Radio_fmwfilter"))
    End Select
    If GetSetting("WiSP_DDE_Client", "Rig" + DownlinkIndex.text, "Radio_AutoSelAccPortEnable", 0) Then
        DownlinkAccPort% = Cdbl2("&H" + GetSetting("WiSP_DDE_Client", "Rig" + DownlinkIndex.text, "Radio_AutoSelAccPortPort"))
        DownlinkAccPortValue% = Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + DownlinkIndex.text, "Radio_AutoSelAccPortValue"))
    Else
        DownlinkAccPort% = 0
    End If
    'if uplink is already using the same port, we won't open it again
    'and we will keep previous settings (baudrate etc.)
    If UplinkHandle% <> 0 And UplinkPort% = DownlinkPort% Then
        ' Check if logging is active
        If frmRadio.CheckLog.Value Then
            f% = FreeFile
            Open "Radio_Log.txt" For Append As f%
            Print #f%, "Downlink port is same as uplink port, " + Str(Time)
            Close f%
        End If
        
        DownlinkHandle% = UplinkHandle%
    
    ElseIf RotorHandle% <> 0 And RotorPort% = DownlinkPort% Then
        ' Check if logging is active
        If frmRadio.CheckLog.Value Then
            f% = FreeFile
            Open "Radio_Log.txt" For Append As f%
            Print #f%, "Downlink port is same as rotor port, " + Str(Time)
            Close f%
        End If
        
        DownlinkHandle% = RotorHandle%
    Else
        'if none of the already open ports is suitable we'll have to open another
        'at least one of the three must be available!:
        If OpenPort(DownlinkPort%, 1) Then
            On Error GoTo OpenDownlinkPortError
            DownlinkHandle% = 1
            MSComm1.Settings = Str(DownlinkBaud&) + ",N,8,1" ' Changed NE1H from ",N,8,2"
            'Tell the control to read entire buffer when Input
            'is used.
            MSComm1.RTSEnable = True
            MSComm1.InputLen = 0
            MSComm1.PortOpen = True
            MSComm1.InputMode = 1 'binary mode
        ElseIf OpenPort(DownlinkPort%, 2) Then
            On Error GoTo OpenDownlinkPortError
            DownlinkHandle% = 2
            MSComm2.Settings = Str(DownlinkBaud&) + ",N,8,1" ' Changed NE1H from ",N,8,2"
            'Tell the control to read entire buffer when Input
            'is used.
            MSComm2.RTSEnable = True
            MSComm2.InputLen = 0
            MSComm2.PortOpen = True
            MSComm2.InputMode = 1 'binary mode
        ElseIf OpenPort(DownlinkPort%, 3) Then
            On Error GoTo OpenDownlinkPortError
            DownlinkHandle% = 3
            MSComm3.Settings = Str(DownlinkBaud&) + ",N,8,1" ' Changed NE1H from ",N,8,2"
            'Tell the control to read entire buffer when Input
            'is used.
            MSComm3.RTSEnable = True
            MSComm3.InputLen = 0
            MSComm3.PortOpen = True
            MSComm3.InputMode = 1 'binary mode
        Else
            GoTo OpenDownlinkPortError
        End If
    End If
    
    'each device can have its own command delay and reply time-out
    If DownlinkHandle% <> 0 Then
        RadioControlDelay(DownlinkHandle%) = 0.001 * Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + DownlinkIndex.text, "Radio_Control_Delay", 0.2))
        RadioReplyTimeout(DownlinkHandle%) = 0.001 * Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + DownlinkIndex.text, "Radio_Reply_Timeout", 1#))
    End If
    
    ' Check if logging is active
    If frmRadio.CheckLog.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Downlink port opened succesfully, " + Str(Time)
        Close f%
    End If
    
    OpenDownlinkPort = 0

Exit Function

'Error handler:
OpenDownlinkPortError:
    ' Check if logging is active
    If frmRadio.CheckLog.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Error opening downlink port, " + Str(Time)
        Close f%
    End If
    
    OpenDownlinkPort = Err.Number
End Function

Function OpenUplinkPort() As Integer
    ' Check if logging is active
    If frmRadio.CheckLog.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Opening uplink port, " + Str(Time)
        Close f%
    End If
    
    'place general settings for this radio into more
    'convenient variables
    UplinkPortName$ = GetSetting("WiSP_DDE_Client", "Rig" + UplinkIndex.text, "Radio_port")
    UplinkPort% = Cdbl2(Right(UplinkPortName$, Len(UplinkPortName$) - 3))
    UplinkModel$ = GetSetting("WiSP_DDE_Client", "Rig" + UplinkIndex.text, "Radio_model", "None")
    UplinkBaud& = Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + UplinkIndex.text, "Radio_baud"))
    UplinkCIVAddress% = Cdbl2("&h" + GetSetting("WiSP_DDE_Client", "Rig" + UplinkIndex.text, "Radio_address"))
    UplinkBidir% = Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + UplinkIndex.text, "Radio_bidir"))
    UplinkTNCUD% = Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + UplinkIndex.text, "Radio_tncupdn"))
    UplinkLO# = Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + UplinkIndex.text, "Radio_UplinkLO"))
    UplinkSplit% = Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + UplinkIndex.text, "Radio_Split", 0))
' Add PA2EON
    UplinkCTCSS% = Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + UplinkIndex.text, "Radio_CTCSS", 0))
'    RadioControlLoopTimer.Interval = Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + UplinkIndex.text, "Radio_Delay", 250))
    If SliderUplink.Value = 1 Then
        RadioControlLoopTimer.Interval = Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + UplinkIndex.text, "Radio_delay"))
    End If
'    If Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + UplinkIndex.text, "Radio_Antenna", "1")) = 1 Then
'        RotorAuto.Value = 1
'    End If
    If GetSetting("WiSP_DDE_Client", "Rig" + UplinkIndex.text, "Radio_AutoSelAccPortEnable", 0) Then
        UplinkAccPort% = Cdbl2("&H" + GetSetting("WiSP_DDE_Client", "Rig" + UplinkIndex.text, "Radio_AutoSelAccPort"))
        UplinkAccPortValue% = Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + UplinkIndex.text, "Radio_AutoSelAccPortValue"))
    Else
        UplinkAccPort% = 0
    End If
    'if downlink is already using the port, we won't open it again
    'and we will keep previous settings (baudrate etc.)
    If DownlinkHandle% <> 0 And UplinkPort% = DownlinkPort% Then
        ' Check if logging is active
        If frmRadio.CheckLog.Value Then
            f% = FreeFile
            Open "Radio_Log.txt" For Append As f%
            Print #f%, "Uplink port is same as downlink port, " + Str(Time)
            Close f%
        End If
        
        UplinkHandle% = DownlinkHandle%
    
    ElseIf RotorHandle% <> 0 And UplinkPort% = RotorPort% Then
        ' Check if logging is active
        If frmRadio.CheckLog.Value Then
            f% = FreeFile
            Open "Radio_Log.txt" For Append As f%
            Print #f%, "Uplink port is same as rotor port, " + Str(Time)
            Close f%
        End If
        
        UplinkHandle% = RotorHandle%
    Else
        If OpenPort(UplinkPort%, 1) Then
            On Error GoTo OpenUplinkPortError
            UplinkHandle% = 1
            MSComm1.Settings = Str(UplinkBaud&) + ",N,8,1" ' Changed NE1H from ",N,8,2"
            'Tell the control to read entire buffer when Input
            'is used.
            MSComm1.RTSEnable = True
            MSComm1.InputLen = 0
            'Open the port.
            MSComm1.PortOpen = True
            MSComm1.InputMode = 1 'binary mode
        ElseIf OpenPort(UplinkPort%, 2) Then
            On Error GoTo OpenUplinkPortError
            UplinkHandle% = 2
            MSComm2.Settings = Str(UplinkBaud&) + ",N,8,1" ' Changed NE1H from ",N,8,2"
            'Tell the control to read entire buffer when Input
            'is used.
            MSComm2.RTSEnable = True
            MSComm2.InputLen = 0
            'Open the port.
            MSComm2.PortOpen = True
            MSComm2.InputMode = 1 'binary mode
        ElseIf OpenPort(UplinkPort%, 3) Then
            On Error GoTo OpenUplinkPortError
            UplinkHandle% = 3
            MSComm3.Settings = Str(UplinkBaud&) + ",N,8,1" ' Changed NE1H from ",N,8,2"
            'Tell the control to read entire buffer when Input
            'is used.
            MSComm3.RTSEnable = True
            MSComm3.InputLen = 0
            'Open the port.
            MSComm3.PortOpen = True
            MSComm3.InputMode = 1 'binary mode
        Else
            GoTo OpenUplinkPortError
        End If
    End If
    'each device can have its own delay
    If UplinkHandle% <> 0 Then
        RadioControlDelay(UplinkHandle%) = 0.001 * Cdbl2(GetSetting("WiSP_DDE_Client", "Rig" + UplinkIndex.text, "Radio_Control_Delay", 0.2))
    End If
    
    ' Check if logging is active
    If frmRadio.CheckLog.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Uplink port opened succesfuly, " + Str(Time)
        Close f%
    End If
    
    OpenUplinkPort = 0
    
Exit Function
'Error handler:
OpenUplinkPortError:
    ' Check if logging is active
    If frmRadio.CheckLog.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Error opening uplink port, " + Str(Time)
        Close f%
    End If
    
    OpenUplinkPort = Err.Number
End Function

Sub OpenRotorPort()
RotorPort% = Cdbl2(Right(frmRotor.RotorPort.text, Len(frmRotor.RotorPort.text) - 3))
If RotorPort% = DownlinkPort% And DownlinkHandle% Then
    RotorHandle% = DownlinkHandle%
ElseIf RotorPort% = UplinkPort% And UplinkHandle% Then
    RotorHandle% = UplinkHandle%
ElseIf OpenPort(RotorPort%, 1) Then
    RotorHandle% = 1
    MSComm1.Settings = frmRotor.RotorBaud.text + ",N,8,1" ' Changed NE1H from ",N,8,2"
    'Tell the control to read entire buffer when Input
    'is used.
    MSComm1.InputLen = 0
    'Open the port.
    MSComm1.PortOpen = True
    MSComm1.InputMode = 1 'binary mode
ElseIf OpenPort(RotorPort%, 2) Then
    RotorHandle% = 2
    MSComm2.Settings = frmRotor.RotorBaud.text + ",N,8,1" ' Changed NE1H from ",N,8,2"
    'Tell the control to read entire buffer when Input
    'is used.
    MSComm2.InputLen = 0
    'Open the port.
    MSComm2.PortOpen = True
    MSComm2.InputMode = 1 'binary mode
ElseIf OpenPort(RotorPort%, 3) Then
    RotorHandle% = 3
    MSComm3.Settings = frmRotor.RotorBaud.text + ",N,8,1" ' Changed NE1H from ",N,8,2"
    'Tell the control to read entire buffer when Input
    'is used.
    MSComm3.InputLen = 0
    'Open the port.
    MSComm3.PortOpen = True
    MSComm3.InputMode = 1 'binary mode
End If
RotorPaceDelaySecs = Cdbl2(frmRotor.RotorPaceDelay.text)
If RotorPaceDelaySecs > 0.1 Then RotorPaceDelaySecs = 0.1
RotorTimeOut = Cdbl2(frmRotor.RotorTimeOutDelay.text)
If RotorTimeOut = 0 Then RotorTimeOut = 0.2
'TrakBox needs to be put in Host mode prior to send data:
If frmRotor.RotorType.text = "TrakBox" And RotorHandle% Then
    Call TBRotorSetTerminal(RotorHandle%)
    Call TBRotorSetHost(RotorHandle%)
End If

End Sub

Function ClosePort(Handle%) As Integer
    On Error GoTo Error:
    
    ' Check if logging is active
    If frmRadio.CheckLog.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Closing stream" + Str(Handle%) + ", " + Str(Time)
        Close f%
    End If
        
    Select Case Handle%
    Case 1
        MSComm1.PortOpen = False
    
    Case 2
        MSComm2.PortOpen = False
    
    Case 3
        MSComm3.PortOpen = False
    
    End Select
        

Exit Function
'error handler:
Error:
    ' Check if logging is active
    If frmRadio.CheckLog.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Error closing stream" + Str(Handle%) + ", " + Str(Time)
        Close f%
    End If
        
    ClosePort = Err.Number
End Function
Function OpenPort(Port%, Handle%)
    On Error GoTo Error:
    
    ' Check if logging is active
    If frmRadio.CheckLog.Value Then
        f% = FreeFile
        Open "Radio_Log.txt" For Append As f%
        Print #f%, "Opening stream" + Str(Handle%) + " thru COM" + LTrim(Str(Port%)) + ", " + Str(Time)
        Close f%
    End If
        
    OpenPort = False
    Select Case Handle%
    Case 1
        MSComm1.CommPort = Port%
        MSComm1.PortOpen = True
        MSComm1.PortOpen = False
        OpenPort = True
    
    Case 2
        MSComm2.CommPort = Port%
        MSComm2.PortOpen = True
        MSComm2.PortOpen = False
        OpenPort = True
    
    Case 3
        MSComm3.CommPort = Port%
        MSComm3.PortOpen = True
        MSComm3.PortOpen = False
        OpenPort = True
    
    End Select

Exit Function

Error:
' Check if logging is active
If frmRadio.CheckLog.Value Then
    f% = FreeFile
    Open "Radio_Log.txt" For Append As f%
    Print #f%, "Error opening stream" + Str(Handle%) + ", " + Str(Time)
    Close f%
End If
        
End Function

Private Sub RadioTimer_Timer()
    If RadioTimerCount < 1000 Then
        RadioTimerCount = RadioTimerCount + 1
    Else
        RadioTimerCount = 0
    End If
End Sub
Function BandDesignator(Freq#) As String
b$ = ""
Select Case Freq#
Case 3000000# To 30000000#
    b$ = "H"    'HF
Case 30000000# To 60000000#
    b$ = "Vl"   'VHF Low
Case 60000000# To 390000000#
    b$ = "V"    'VHF
Case 390000000# To 600000000#
    b$ = "U"    'UHF
Case 600000000# To 1550000000#
    b$ = "L"
Case 1550000000# To 3900000000#
    b$ = "S"
Case 3900000000# To 6200000000#
    b$ = "C"
Case 6200000000# To 10900000000#
    b$ = "X"
Case 10900000000# To 18000000000#
    b$ = "Ku"
Case 18000000000# To 26500000000#
    b$ = "K"
Case 26500000000# To 40000000000#
    b$ = "Ka"
End Select
BandDesignator = b$
End Function

Sub Remove_Registry()
    DeleteSetting "WiSP_DDE_Client"
End Sub

Private Sub Form_Terminate()
    Call Close_Click
End Sub

Private Sub verayuda_Click()
    Call Shell("write leeme.txt", 1)
End Sub

Private Sub viewhelp_Click()
    Call Shell("write readme.txt", 1)
End Sub
Sub ErrorHandler(text$, fatal%)
If fatal% Then
    a$ = "Fatal "
End If
a$ = a$ + "Error: " + text$
frmMessage.Label1.Caption = a$
frmMessage.Show
'wait until user aknowledges:
Do
    DoEvents
Loop While frmMessage.Visible
If fatal% Then Form_Unload (0)
End Sub

'we used old Val() function thoroughly in the program but it was not locale-aware
'changing to Cdbl() caused problems, Cdbl2() tries to solve this.
Function Cdbl2(a) As Double

i = InStrRev(a, frmDdelink.Decimal.text)
If ((i <> 0)) Then
    Mid(a, i, 1) = "."
End If
Cdbl2 = Val(a)

End Function
'Converts a string of chars into a string of text representing the original chars in
'hexadecimal
Function StrToHex(s$) As String
a$ = ""
For f% = 1 To Len(s$)
    a$ = a$ + Hex(Asc(Mid(s$, f%, 1))) + "h "
Next
StrToHex = a$
End Function

Function ArrayToHex(arr, fin) As String
a$ = ""
If fin > UBound(arr) Then
    fin = UBound(arr)
End If
For f% = LBound(arr) To fin
    a$ = a$ + Hex(arr(f%)) + "h "
Next
ArrayToHex = a$
End Function
