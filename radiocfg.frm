VERSION 5.00
Begin VB.Form frmRadio 
   Caption         =   "Radio Settings"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4740
   Icon            =   "radiocfg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox RadioCTCSS 
      Caption         =   "CTCSS On/Off"
      Height          =   495
      Left            =   1680
      TabIndex        =   55
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox RadioReplyTime 
      Height          =   247
      Left            =   117
      TabIndex        =   53
      ToolTipText     =   "Waiting period for reply from radio (milliseconds)."
      Top             =   4680
      Width           =   1183
   End
   Begin VB.CheckBox CheckLogCom 
      Caption         =   "Log Commands"
      Height          =   247
      Left            =   3000
      TabIndex        =   52
      ToolTipText     =   "Log Radio Commands to a file named ""Radio_Log.txt""."
      Top             =   1440
      Width           =   1410
   End
   Begin VB.CheckBox CheckLog 
      Caption         =   "Log Events"
      Height          =   247
      Left            =   3000
      TabIndex        =   51
      ToolTipText     =   "Log Radio control events to a file named ""Radio_Log.txt""."
      Top             =   1080
      Width           =   1300
   End
   Begin VB.TextBox RadioControlDelay 
      Height          =   285
      Left            =   117
      TabIndex        =   49
      ToolTipText     =   "Delay time after commands sent to radio (milliseconds)."
      Top             =   4095
      Width           =   1183
   End
   Begin VB.CheckBox RadioSplit 
      Caption         =   "Split Mode"
      Height          =   481
      Left            =   1680
      TabIndex        =   48
      ToolTipText     =   "For half-duplex rigs and same-band up/dnlinks. VFO-A is downlink, VFO-B is uplink."
      Top             =   1320
      Width           =   1183
   End
   Begin VB.TextBox RadioDelay 
      Height          =   285
      Left            =   120
      TabIndex        =   44
      ToolTipText     =   "Interval between succesive updates to the radio (milliseconds)."
      Top             =   5280
      Width           =   1183
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frequency Converters:"
      Height          =   1066
      Left            =   1920
      TabIndex        =   39
      Top             =   2760
      Width           =   2587
      Begin VB.TextBox RadioDownlinkLOFreq 
         Height          =   285
         Left            =   1080
         TabIndex        =   41
         ToolTipText     =   "Sets the Receive Local Oscillator Frequency (leave blank if no converter used)"
         Top             =   255
         Width           =   960
      End
      Begin VB.TextBox RadioUplinkLOFreq 
         Height          =   285
         Left            =   1080
         TabIndex        =   40
         ToolTipText     =   "Sets the Transmit Local Oscillator Frequency (leave blank if no converter used)"
         Top             =   600
         Width           =   960
      End
      Begin VB.Label Label5 
         Caption         =   "MHz"
         Height          =   255
         Left            =   2100
         TabIndex        =   47
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "MHz"
         Height          =   255
         Left            =   2100
         TabIndex        =   46
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Dnconv. LO:"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "Upconv. LO:"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.CommandButton RadioDelete 
      Caption         =   "Delete least Radio"
      Height          =   495
      Left            =   3600
      TabIndex        =   38
      ToolTipText     =   "WARNING!, deletes the highest index radio."
      Top             =   6720
      Width           =   975
   End
   Begin VB.CheckBox RadioAntenna 
      Caption         =   "Antenna Steerable"
      Height          =   480
      Left            =   3000
      TabIndex        =   37
      ToolTipText     =   "Uncheck if this radio has an omnidirectional antenna attached."
      Top             =   120
      Width           =   1650
   End
   Begin VB.TextBox RadioVolume 
      Height          =   285
      Left            =   120
      TabIndex        =   35
      ToolTipText     =   "Volume setting (for PCR-1000 only)"
      Top             =   3510
      Width           =   1183
   End
   Begin VB.CheckBox RadioEnable 
      Caption         =   "Enable Radio"
      Height          =   481
      Left            =   1680
      TabIndex        =   31
      ToolTipText     =   "Radio enabling, necesary to access every other field."
      Top             =   120
      Width           =   1305
   End
   Begin VB.CommandButton RadioAutoSelConfig 
      Caption         =   "Auto Selection"
      Height          =   481
      Left            =   120
      TabIndex        =   30
      ToolTipText     =   "Shows Automatic Selection paremeters window."
      Top             =   6720
      Width           =   1066
   End
   Begin VB.CommandButton RadioFinish 
      Caption         =   "Close"
      Height          =   495
      Left            =   2520
      TabIndex        =   29
      ToolTipText     =   "Close this window."
      Top             =   6720
      Width           =   858
   End
   Begin VB.CommandButton RadioSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   1440
      TabIndex        =   28
      ToolTipText     =   "Save settings to Windows Registry."
      Top             =   6720
      Width           =   858
   End
   Begin VB.TextBox RadioAddress 
      Height          =   285
      Left            =   120
      TabIndex        =   27
      ToolTipText     =   "For Icom rigs, the address the rig is configured to respond to."
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ComboBox RadioModel 
      Height          =   315
      Left            =   120
      TabIndex        =   26
      ToolTipText     =   "Radio model selection."
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ComboBox RadioIndex 
      Height          =   315
      ItemData        =   "radiocfg.frx":030A
      Left            =   120
      List            =   "radiocfg.frx":030C
      Sorted          =   -1  'True
      TabIndex        =   25
      ToolTipText     =   "Picks one radio from the list."
      Top             =   360
      Width           =   1215
   End
   Begin VB.CheckBox RadioBidir 
      Caption         =   "Bidirectional Interface"
      Height          =   481
      Left            =   1680
      TabIndex        =   24
      ToolTipText     =   "Uncheck if radio interface is only PC->Rig."
      Top             =   600
      Width           =   1305
   End
   Begin VB.Frame RadioFilters 
      Caption         =   "Filter Asignments:"
      Height          =   2587
      Left            =   2040
      TabIndex        =   6
      ToolTipText     =   "Only for some receivers, you can customize each mode's filter bandwidth."
      Top             =   3960
      Width           =   2353
      Begin VB.TextBox RadioSSBFilter 
         Height          =   285
         Left            =   720
         TabIndex        =   32
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox RadioFMWFilter 
         Height          =   285
         Left            =   720
         TabIndex        =   11
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox RadioFMNFilter 
         Height          =   285
         Left            =   720
         TabIndex        =   10
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox RadioFMFilter 
         Height          =   285
         Left            =   720
         TabIndex        =   9
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox RadioCWNFilter 
         Height          =   285
         Left            =   720
         TabIndex        =   8
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox RadioCWFilter 
         Height          =   285
         Left            =   720
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "KHz"
         Height          =   255
         Left            =   1680
         TabIndex        =   23
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label18 
         Caption         =   "KHz"
         Height          =   255
         Left            =   1680
         TabIndex        =   22
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "KHz"
         Height          =   255
         Left            =   1680
         TabIndex        =   21
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "KHz"
         Height          =   255
         Left            =   1680
         TabIndex        =   20
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "KHz"
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "KHz"
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "FM:"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "CW-N:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "FM-N:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "CW:"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "FM-W:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "SSB:"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.CheckBox RadioTNCUD 
      Caption         =   "TNC UP/DN SSB Dwnlink"
      Height          =   598
      Left            =   3000
      TabIndex        =   4
      ToolTipText     =   "For PSK downlinks, check this if your TNC can control this Rigs frequency with Up/Down keys."
      Top             =   480
      Width           =   1650
   End
   Begin VB.ComboBox RadioBaud 
      Height          =   273
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Baud rate selection."
      Top             =   2880
      Width           =   1215
   End
   Begin VB.ComboBox RadioPort 
      Height          =   273
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Communications port selection."
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label RadioReplyTimeoutLabel 
      Caption         =   "Reply timeout (mSec)"
      Height          =   247
      Left            =   117
      TabIndex        =   54
      Top             =   4446
      Width           =   1534
   End
   Begin VB.Label RadioControlDelayLabel 
      Caption         =   "Command delay (mSec)"
      Height          =   247
      Left            =   117
      TabIndex        =   50
      Top             =   3861
      Width           =   1534
   End
   Begin VB.Label RadioDelayText 
      Caption         =   "Loop delay (mSec.)"
      Height          =   255
      Left            =   120
      TabIndex        =   45
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label RadioVolumeText 
      Caption         =   "Volume:"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label RadioAddressText 
      Caption         =   "Address (hex):"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label RadioModelText 
      Caption         =   "Model:"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Radio Index:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label RadioBaudText 
      Caption         =   "Baud Rate:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label RadioPortText 
      Caption         =   "Port:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "frmRadio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub RadioAutoSelConfig_Click()
    frmAutoSel.Show
End Sub
Private Sub RadioDelete_Click()
    'remove last RigN folder from registry:
    'Add as many entries to the "radio Index" combo as we find
    'in the registry...
    i% = 0
    Do
        i% = i% + 1
    Loop Until GetSetting("WiSP_DDE_Client", "Rig" + LTrim$(Str$(i%)), "Radio_model", "-") = "-"
    If i% > 1 Then
        DeleteSetting "WiSP_DDE_Client", "Rig" + LTrim(Str(i% - 1))
    End If
    'update the index combo as we may have more radios now!
    'Add as many entries to the "radio Index" combo as we find
    'in the registry...
    a% = RadioIndex.ListIndex
    RadioIndex.Clear
    i% = 0
    Do
        i% = i% + 1
        RadioIndex.AddItem LTrim$(Str$(i%))
    Loop Until GetSetting("WiSP_DDE_Client", "Rig" + LTrim$(Str$(i%)), "Radio_model", "-") = "-"
    If a% >= RadioIndex.ListCount Then
        RadioIndex.ListIndex = RadioIndex.ListCount - 1
    Else
        RadioIndex.ListIndex = a%
    End If
End Sub

Private Sub RadioEnable_Click()
If RadioIndex.text <> "" Then
    Select Case RadioModel.text
    Case "IC-746", "IC-275", "IC-475", "IC-706"
        RadioAddress.Enabled = True
        RadioAddressText.Enabled = True
        RadioBidir.Enabled = True
        RadioFilters.Enabled = False
        RadioSSBFilter.Enabled = False
        RadioCWFilter.Enabled = False
        RadioCWNFilter.Enabled = False
        RadioFMFilter.Enabled = False
        RadioFMNFilter.Enabled = False
        RadioFMWFilter.Enabled = False
        RadioVolume.Enabled = False
        RadioVolumeText.Enabled = False
        RadioSplit.Enabled = True
        CheckLogCom.Enabled = False
    Case "IC-821", "IC-910", "IC-970", "IC-R7000", "IC-R8500"
        RadioAddress.Enabled = True
        RadioAddressText.Enabled = True
        RadioBidir.Enabled = True
        RadioFilters.Enabled = False
        RadioSSBFilter.Enabled = False
        RadioCWFilter.Enabled = False
        RadioCWNFilter.Enabled = False
        RadioFMFilter.Enabled = False
        RadioFMNFilter.Enabled = False
        RadioFMWFilter.Enabled = False
        RadioVolume.Enabled = False
        RadioVolumeText.Enabled = False
        RadioSplit.Enabled = True
        CheckLogCom.Enabled = True
   Case "FT-100", "FT-817", "FT-897"
        RadioAddress.Enabled = False
        RadioAddressText.Enabled = False
        RadioBidir.Enabled = False
        RadioBidir.Value = 1
        RadioFilters.Enabled = False
        RadioSSBFilter.Enabled = False
        RadioCWFilter.Enabled = False
        RadioCWNFilter.Enabled = False
        RadioFMFilter.Enabled = False
        RadioFMNFilter.Enabled = False
        RadioFMWFilter.Enabled = False
        RadioVolume.Enabled = False
        RadioVolumeText.Enabled = False
        RadioSplit.Enabled = True
        CheckLogCom.Enabled = True
   Case "FT-736", "AR-8000", "TM-D700", "TH-D7", "TS-2000"  ' G6LVB 16 Dec 2000 Added FT-100 & FT-817
        RadioAddress.Enabled = False
        RadioAddressText.Enabled = False
        RadioBidir.Enabled = False
        RadioBidir.Value = 1
        RadioFilters.Enabled = False
        RadioSSBFilter.Enabled = False
        RadioCWFilter.Enabled = False
        RadioCWNFilter.Enabled = False
        RadioFMFilter.Enabled = False
        RadioFMNFilter.Enabled = False
        RadioFMWFilter.Enabled = False
        RadioVolume.Enabled = False
        RadioVolumeText.Enabled = False
        RadioSplit.Enabled = False
        RadioSplit.Value = 0
        CheckLogCom.Enabled = False
   Case "FT-847"
        RadioAddress.Enabled = False
        RadioAddressText.Enabled = False
        RadioBidir.Enabled = False
        RadioBidir.Value = 1
        RadioFilters.Enabled = False
        RadioSSBFilter.Enabled = False
        RadioCWFilter.Enabled = False
        RadioCWNFilter.Enabled = False
        RadioFMFilter.Enabled = False
        RadioFMNFilter.Enabled = False
        RadioFMWFilter.Enabled = False
        RadioVolume.Enabled = False
        RadioVolumeText.Enabled = False
        RadioSplit.Enabled = True
        RadioSplit.Value = 0
        'Add PA2EON
        RadioCTCSS.Enabled = True
        RadioCTCSS.Value = 0
        CheckLogCom.Enabled = False
    Case "TS-790"
        RadioAddress.Enabled = False
        RadioAddressText.Enabled = False
        RadioBidir.Enabled = True
        RadioFilters.Enabled = False
        RadioSSBFilter.Enabled = False
        RadioCWFilter.Enabled = False
        RadioCWNFilter.Enabled = False
        RadioFMFilter.Enabled = False
        RadioFMNFilter.Enabled = False
        RadioFMWFilter.Enabled = False
        RadioVolume.Enabled = False
        RadioVolumeText.Enabled = False
        RadioSplit.Value = 0
        RadioSplit.Enabled = False
        CheckLogCom.Enabled = False
   Case "AR-3000A"  ' EB2CTA 5 Ago 2005 Added AR-3000A
        RadioAddress.Enabled = False
        RadioAddressText.Enabled = False
        RadioBidir.Enabled = False
        RadioBidir.Value = 1
        RadioFilters.Enabled = False
        RadioSSBFilter.Enabled = False
        RadioCWFilter.Enabled = False
        RadioCWNFilter.Enabled = False
        RadioFMFilter.Enabled = False
        RadioFMNFilter.Enabled = False
        RadioFMWFilter.Enabled = False
        RadioVolume.Enabled = True  ' S-Meter on
        RadioVolumeText.Enabled = True  ' S-meter on
        RadioSplit.Enabled = False
        RadioSplit.Value = 0
        CheckLogCom.Enabled = False
   Case "AR-5000", "VR-5000"
        RadioAddress.Enabled = False
        RadioAddressText.Enabled = False
        RadioBidir.Value = 0
        RadioBidir.Enabled = False
        RadioFilters.Enabled = True
        RadioSSBFilter.Enabled = True
        RadioCWFilter.Enabled = True
        RadioCWNFilter.Enabled = True
        RadioFMFilter.Enabled = True
        RadioFMNFilter.Enabled = True
        RadioFMWFilter.Enabled = True
        RadioVolume.Enabled = False
        RadioVolumeText.Enabled = False
        RadioSplit.Value = 0
        RadioSplit.Enabled = False
        CheckLogCom.Enabled = False
    Case "PCR-1000"
        RadioAddress.Enabled = False
        RadioAddressText.Enabled = False
        RadioBidir.Enabled = False
        RadioBidir.Value = 1
        RadioFilters.Enabled = True
        RadioSSBFilter.Enabled = True
        RadioCWFilter.Enabled = True
        RadioCWNFilter.Enabled = True
        RadioFMFilter.Enabled = True
        RadioFMNFilter.Enabled = True
        RadioFMWFilter.Enabled = True
        RadioVolume.Enabled = True
        RadioVolumeText.Enabled = True
        RadioSplit.Value = 0
        RadioSplit.Enabled = False
        RadioBaud.text = "9600"
        CheckLogCom.Enabled = False
    Case "TS-711", "TS-811"  'VK2JXI
        RadioAddress.Enabled = False
        RadioAddressText.Enabled = False
        RadioBidir.Enabled = True
        RadioFilters.Enabled = False
        RadioSSBFilter.Enabled = False
        RadioCWFilter.Enabled = False
        RadioCWNFilter.Enabled = False
        RadioFMFilter.Enabled = False
        RadioFMNFilter.Enabled = False
        RadioFMWFilter.Enabled = False
        RadioVolume.Enabled = False
        RadioVolumeText.Enabled = False
        RadioSplit.Value = 0
        RadioSplit.Enabled = False
        CheckLogCom.Enabled = False
    Case "FRG-9600"
        RadioAddress.Enabled = False
        RadioAddressText.Enabled = False
        RadioBidir.Value = 0
        RadioBidir.Enabled = False
        RadioFilters.Enabled = False
        RadioSSBFilter.Enabled = False
        RadioCWFilter.Enabled = False
        RadioCWNFilter.Enabled = False
        RadioFMFilter.Enabled = False
        RadioFMNFilter.Enabled = False
        RadioFMWFilter.Enabled = False
        RadioVolume.Enabled = False
        RadioVolumeText.Enabled = False
        RadioSplit.Value = 0
        RadioSplit.Enabled = False
        RadioBaud.text = "4800"
        CheckLogCom.Enabled = False
    Case "TrakBox"
        RadioAddress.Enabled = False
        RadioAddressText.Enabled = False
        RadioBidir.Value = 1
        RadioBidir.Enabled = False
        RadioFilters.Enabled = False
        RadioSSBFilter.Enabled = False
        RadioCWFilter.Enabled = False
        RadioCWNFilter.Enabled = False
        RadioFMFilter.Enabled = False
        RadioFMNFilter.Enabled = False
        RadioFMWFilter.Enabled = False
        RadioVolume.Enabled = False
        RadioVolumeText.Enabled = False
        RadioSplit.Value = 0
        RadioSplit.Enabled = False
        CheckLogCom.Enabled = False

    End Select
Else
    RadioEnable.Value = 0
End If
  
If RadioEnable.Value Then
    RadioModel.Enabled = True
    RadioModelText.Enabled = True
    RadioPort.Enabled = True
    RadioPortText.Enabled = True
    RadioBaud.Enabled = True
    RadioBaudText.Enabled = True
    RadioTNCUD.Enabled = True
    RadioAntenna.Enabled = True
    RadioAutoSelConfig.Enabled = True
    RadioDownlinkLOFreq.Enabled = True
    RadioUplinkLOFreq.Enabled = True
    RadioControlDelay.Enabled = True
    RadioControlDelayLabel.Enabled = True
    RadioDelay.Enabled = True
    RadioDelayText.Enabled = True
Else
    RadioDelay.Enabled = False
    RadioDelayText.Enabled = False
    RadioAutoSelConfig.Enabled = False
    RadioModel.Enabled = False
    RadioModelText.Enabled = False
    RadioPort.Enabled = False
    RadioPortText.Enabled = False
    RadioBaud.Enabled = False
    RadioBaudText.Enabled = False
    RadioTNCUD.Enabled = False
    RadioAntenna.Enabled = False
    RadioAddress.Enabled = False
    RadioAddressText.Enabled = False
    RadioBidir.Enabled = False
    RadioFilters.Enabled = False
    RadioSSBFilter.Enabled = False
    RadioCWFilter.Enabled = False
    RadioCWNFilter.Enabled = False
    RadioFMFilter.Enabled = False
    RadioFMNFilter.Enabled = False
    RadioFMWFilter.Enabled = False
    RadioVolume.Enabled = False
    RadioVolumeText.Enabled = False
    RadioDownlinkLOFreq.Enabled = False
    RadioUplinkLOFreq.Enabled = False
    RadioSplit.Enabled = False
    'add PA2EON
    RadioCTCSS.Enabled = False
    RadioControlDelay.Enabled = False
    RadioControlDelayLabel.Enabled = False
End If
End Sub

Private Sub RadioFinish_Click()

    frmRadio.Hide

End Sub


Private Sub Form_Load()
    'Add as many entries to the "radio Index" combo as we find
    'in the registry...
    RadioIndex.Clear
    i% = 0
    Do
        i% = i% + 1
        RadioIndex.AddItem LTrim$(Str$(i%))
    Loop Until GetSetting("WiSP_DDE_Client", "Rig" + LTrim$(Str$(i%)), "Radio_model", "-") = "-"
    RadioPort.Clear
    RadioPort.AddItem "None"    ' Add each item to list.
    RadioPort.AddItem "COM1"
    RadioPort.AddItem "COM2"
    RadioPort.AddItem "COM3"
    RadioPort.AddItem "COM4"
    RadioPort.AddItem "COM5"
    RadioPort.AddItem "COM6"
    RadioPort.AddItem "COM7"
    RadioPort.AddItem "COM8"
    RadioPort.AddItem "COM9"
    RadioPort.AddItem "COM10"
    RadioPort.AddItem "COM11"
    RadioPort.AddItem "COM12"
    RadioPort.AddItem "COM13"
    RadioPort.AddItem "COM14"
    RadioPort.AddItem "COM15"
    RadioPort.AddItem "COM16"
    RadioPort.AddItem "COM17"
    RadioPort.AddItem "COM18"
    RadioPort.AddItem "COM19"
    RadioPort.AddItem "COM20"
    RadioModel.Clear
    RadioModel.AddItem "None"
    RadioModel.AddItem "IC-821"
    RadioModel.AddItem "IC-910"
    RadioModel.AddItem "IC-970"
    RadioModel.AddItem "IC-275"
    RadioModel.AddItem "IC-475"
    RadioModel.AddItem "IC-746"
    RadioModel.AddItem "IC-706"
    RadioModel.AddItem "IC-R7000"
    RadioModel.AddItem "IC-R8500"
    RadioModel.AddItem "PCR-1000"
    RadioModel.AddItem "FT-847"
    RadioModel.AddItem "FT-736"
    RadioModel.AddItem "FT-100" ' G6LVB added 16 Dec 2000
    RadioModel.AddItem "FT-817" ' G6LVB added 16 Dec 2000
    RadioModel.AddItem "FT-897"
    RadioModel.AddItem "FRG-9600"   'VK2KXI
    RadioModel.AddItem "VR-5000"
    RadioModel.AddItem "AR-3000A"   ' 5 Ago 2005 EB2CTA AR-3000A
    RadioModel.AddItem "AR-5000"
    RadioModel.AddItem "AR-8000"
    RadioModel.AddItem "TH-D7"
    RadioModel.AddItem "TM-D700"
    RadioModel.AddItem "TS-790"
    RadioModel.AddItem "TS-2000" ' 2 March 2000 G6LVB TS-2000
    RadioModel.AddItem "TS-711" ' 14 May 2001 VK2JXI TS711/TS811
    RadioModel.AddItem "TS-811" ' 14 May 2001 VK2JXI TS711/TS811
    RadioModel.AddItem "TrakBox"

    RadioBaud.Clear
    RadioBaud.AddItem "1200"    ' Add each item to list.
    RadioBaud.AddItem "2400"
    RadioBaud.AddItem "4800"
    RadioBaud.AddItem "9600"
    RadioBaud.AddItem "19200"
    RadioBaud.AddItem "38400"
    RadioBaud.AddItem "57600"
    'This sub will hide, show, enable & disable each structure
    'as needed for the particular radio selected
    Call RadioEnable_Click
    
    'Recall radio control events logging option
    CheckLog.Value = GetSetting("WiSP_DDE_Client", "Config", "Radio_log", 0)
    
    'Recall radio commands logging option
    CheckLogCom.Value = GetSetting("WiSP_DDE_Client", "Config", "Radio_log_commands", 0)
    
    'Initially disable command logging option (later when a radio is selected will
    'eventually be enabled.
    CheckLogCom.Enabled = False
    
End Sub

Private Sub RadioIndex_Change()
    'update settings for selected radio
    'Retrieve configuration for selected radio:
    RadioEnable.Value = GetSetting("WiSP_DDE_Client", "Rig" + LTrim$(RadioIndex.text), "Radio_enable", 0)
    RadioModel.text = GetSetting("WiSP_DDE_Client", "Rig" + LTrim$(RadioIndex.text), "Radio_model", "None")
    
    'Enable & disable fields as needed:
    Call RadioEnable_Click
    
    RadioBaud.text = GetSetting("WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_baud", "9600")
    RadioPort.text = GetSetting("WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_port", "None")
    RadioTNCUD.Value = GetSetting("WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_tncupdn", 0)
    RadioAntenna.Value = GetSetting("WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_Antenna", 1)
    RadioDelay.text = GetSetting("WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_delay", "500")
    RadioControlDelay.text = GetSetting("WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_Control_delay", "200")
    RadioReplyTime.text = GetSetting("WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_Reply_Timeout", "1000")
    
    RadioDownlinkLOFreq.text = GetSetting("WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_DownlinkLO", "")
    RadioUplinkLOFreq.text = GetSetting("WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_UplinkLO", "")
    
    RadioSplit.Value = GetSetting("WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_Split", 0)
    'Add PA2EON
    RadioCTCSS.Value = GetSetting("WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_CTCSS", 0)
    
    'not all radios need all info to be stored...
    RadioAddress.text = GetSetting("WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_address", "")
    RadioBidir.Value = GetSetting("WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_bidir", 0)
    RadioSSBFilter.text = GetSetting("WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_ssbfilter", "")
    RadioCWFilter.text = GetSetting("WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_cwfilter", "")
    RadioCWNFilter.text = GetSetting("WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_cwnfilter", "")
    RadioFMFilter.text = GetSetting("WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_fmfilter", "")
    RadioFMNFilter.text = GetSetting("WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_fmnfilter", "")
    RadioFMWFilter.text = GetSetting("WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_fmwfilter", "")
    RadioVolume.text = GetSetting("WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_Volume", "")
    
    'update also auto-selection parameters for current rig:
    frmAutoSel.Refresh
End Sub
'make clicking the same as typing:
Private Sub RadioIndex_Click()
    Call RadioIndex_Change
End Sub
'enable/disable structures when radio model changes:
Private Sub RadioModel_Change()
    Call RadioEnable_Click
End Sub

Private Sub RadioModel_Click()
    Call RadioEnable_Click
End Sub

Private Sub RadioSave_Click()
    'force PCR1000 speed to 9600baud
    If RadioModel.text = "PCR-1000" Then
        RadioBaud.text = "9600"
    End If
    'force FRG9600 speed to 4800baud
    If RadioModel.text = "FRG-9600" Then
        RadioBaud.text = "4800"
    End If
    If RadioIndex.text <> "" Then
        'Save settings to windows registry...
        SaveSetting "WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_enable", RadioEnable.Value
        SaveSetting "WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_model", RadioModel.text
        SaveSetting "WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_port", RadioPort.text
        SaveSetting "WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_baud", RadioBaud.text
        SaveSetting "WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_tncupdn", RadioTNCUD.Value
        SaveSetting "WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_Antenna", RadioAntenna.Value
        SaveSetting "WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_address", RadioAddress.text
        SaveSetting "WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_bidir", RadioBidir.Value
        SaveSetting "WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_ssbfilter", RadioSSBFilter.text
        SaveSetting "WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_cwfilter", RadioCWFilter.text
        SaveSetting "WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_cwnfilter", RadioCWNFilter.text
        SaveSetting "WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_fmfilter", RadioFMFilter.text
        SaveSetting "WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_fmnfilter", RadioFMNFilter.text
        SaveSetting "WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_fmwfilter", RadioFMWFilter.text
        SaveSetting "WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_delay", RadioDelay.text
        SaveSetting "WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_DownlinkLO", RadioDownlinkLOFreq.text
        SaveSetting "WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_UplinkLO", RadioUplinkLOFreq.text
        SaveSetting "WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_Split", RadioSplit.Value
        'Add PA2EON
        SaveSetting "WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_CTCSS", RadioCTCSS.Value
        SaveSetting "WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_Control_delay", RadioControlDelay.text
        SaveSetting "WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_Reply_Timeout", RadioReplyTime.text
        
        If RadioVolume.Enabled Then SaveSetting "WiSP_DDE_Client", "Rig" + RadioIndex.text, "Radio_Volume", RadioVolume.text
        'update the index combo as we may have more radios now!
    
        '**Radio interface:**
        'Add as many entries to the "radio Index" combo as we find
        'in the registry...
        'do it twice, once for uplink and once for downlink selections
        frmMain.UplinkIndex.Clear
        frmMain.UplinkIndex.AddItem "None"
        i% = 1
        Do
            frmMain.UplinkIndex.AddItem LTrim$(Str$(i%))
            i% = i% + 1
        Loop Until GetSetting("WiSP_DDE_Client", "Rig" + LTrim$(Str$(i%)), "Radio_model", "-") = "-"
        
        frmMain.DownlinkIndex.Clear
        frmMain.DownlinkIndex.AddItem "None"
        i% = 1
        Do
            frmMain.DownlinkIndex.AddItem LTrim$(Str$(i%))
            i% = i% + 1
        Loop Until GetSetting("WiSP_DDE_Client", "Rig" + LTrim$(Str$(i%)), "Radio_model", "-") = "-"
        
    End If
    
    'Store radio events logging option
    SaveSetting "WiSP_DDE_Client", "Config", "Radio_log", CheckLog.Value

    'Store radio commands logging option
    SaveSetting "WiSP_DDE_Client", "Config", "Radio_log_commands", CheckLogCom.Value

End Sub
