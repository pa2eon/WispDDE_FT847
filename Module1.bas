Attribute VB_Name = "Module1"

'Direct port access port (uses DriverLINX drivers):
'*  DriverLINX Port I/O Driver Interface
'*  <cp> Copyright 1996 Scientific Software Tools, Inc.<nl>
'*  All Rights Reserved.<nl>
'*  DriverLINX is a registered trademark of Scientific Software Tools, Inc.
Public Declare Function DlPortReadPortUchar Lib "dlportio.dll" (ByVal Port As Long) As Byte
Public Declare Sub DlPortWritePortUchar Lib "dlportio.dll" (ByVal Port As Long, ByVal Value As Byte)

'Decimal and separator according to locale settings (initialized at startup)
Public Dec_Sep$
'Punctuation character NOT representing decimal separator (if decimal is "." this is ",")
Public Not_Dec_Sep$

'***OLD*** Direct port command functions
'Declare Function InPort Lib "InOut.dll" (ByVal Address As Integer) As Integer
'Declare Function OutPort Lib "InOut.dll" (ByVal Address As Integer, ByVal Data As Integer) As Integer

'RadioControlCountdown is used in the PSK downlink to send
'only a few control frames to the rig at the beggining of
'the pass, this variable counts these frames.
'
'RotorControlCountdown is used in the flip detection to
'evaluate the change in Az. from the first DDE burst to
'the N-th burst. RotorControlCountdown holds this N.
Public RadioControlCountdown, RotorControlCountdown As Integer

'This will store the first Azimuth value when the sat appears
'(or is attended by WiSP)
Public RotorAz As Single
'After some seconds, the first Azimuth value is compared with
'current value and the variation is stored in RotorDeltaAz
'this is for the auto flip-detection
Public RotorDeltaAz As Single
Public RotorFlip As Boolean

Public RotorPort%
Public RotorHandle%
Public RotorPaceDelaySecs As Single
Public RotorTimeOut As Single
Public RotorUpdateComplete As Boolean

'accesory counter to limit the waiting time for the rig to
'reply to a command (acknowledge etc.)
Public RadioTimerCount As Integer
Public RadioTimeOut As Double
'Each device (rig) has its own control delay
Public RadioControlDelay(3) As Double
'Each device (rig) has its own reply timeout
Public RadioReplyTimeout(3) As Double

'flag to detect the command-line argument "S" to instruct
'wispdde to close itself at the end of the first pass
Public SinglePass%

'variables that store selected rigs comm settings
Public DownlinkHandle%
Public DownlinkPort%
Public DownlinkModel$
Public DownlinkBaud&
Public DownlinkCIVAddress%
Public DownlinkBidir%
Public DownlinkTNCUD%
Public DownlinkAccPort%
Public DownlinkAccPortValue%
Public DownlinkFilter
Public DownlinkVolume%
Public DownlinkLO#
Public DownlinkSplit%
Public DownlinkTXOn%

Public UplinkHandle%
Public UplinkPort%
Public UplinkModel$
Public UplinkBaud&
Public UplinkCIVAddress%
Public UplinkBidir%
Public UplinkTNCUD%
Public UplinkAccPort%
Public UplinkAccPortValue%
Public UplinkLO#
Public UplinkSplit%
'Add PA2EON
Public UplinkCTCSS%
Public UplinkDuplexPlus%


'variables used to provide "transparent tuning" of radios
Public DownlinkCorrection As Double
Public UplinkCorrection As Double
Public SliderCorrection As Double
Public ButtonCorrection As Double
'this will hold last (good) freq read from radio
Public RD#
Public RU#
'flags that indicate updates in queue
Public uu%
Public ud%

Public SliderUplinkMemory%
Public SliderDownlinkMemory%



'getting rid of silly names:
Function InPort(Port As Integer) As Byte
InPort = DlPortReadPortUchar(CLng(Port))
End Function
Sub OutPort(Port As Integer, Value As Byte)
Call DlPortWritePortUchar(CLng(Port), Value)
End Sub


