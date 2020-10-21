VERSION 5.00
Begin VB.Form frmRotor 
   Caption         =   "Rotor Settings"
   ClientHeight    =   5421
   ClientLeft      =   65
   ClientTop       =   351
   ClientWidth     =   3341
   Icon            =   "rotorcfg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5421
   ScaleWidth      =   3341
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox CheckLog 
      Caption         =   "Log Events"
      Enabled         =   0   'False
      Height          =   247
      Left            =   1755
      TabIndex        =   28
      Top             =   2925
      Width           =   1417
   End
   Begin VB.TextBox RotorTimeOutDelay 
      Height          =   285
      Left            =   1800
      TabIndex        =   26
      Top             =   2460
      Width           =   735
   End
   Begin VB.TextBox RotorPaceDelay 
      Height          =   285
      Left            =   120
      TabIndex        =   24
      ToolTipText     =   "Character-rate pace down delay time in secs. for TrakBox."
      Top             =   2460
      Width           =   735
   End
   Begin VB.CheckBox Rotor450Deg 
      Caption         =   "Az. 450deg."
      Height          =   247
      Left            =   1755
      TabIndex        =   23
      ToolTipText     =   "Check if interface is capable of driving 450deg rotors"
      Top             =   1872
      Width           =   1534
   End
   Begin VB.Frame Frame2 
      Caption         =   "Park"
      Height          =   1455
      Left            =   1800
      TabIndex        =   18
      ToolTipText     =   "Parking position, will be effective when no satellite is tracked."
      Top             =   3276
      Width           =   1335
      Begin VB.TextBox RotorElPark 
         Height          =   285
         Left            =   240
         TabIndex        =   22
         ToolTipText     =   "Elevation angle to set after satellite goes, may be left blank."
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox RotorAzPark 
         Height          =   285
         Left            =   240
         TabIndex        =   20
         ToolTipText     =   "Azimuth angle to set after satellite goes, may be left blank."
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Elevation:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Azimuth:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Offset"
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Offset angles to add or substract."
      Top             =   3276
      Width           =   1335
      Begin VB.TextBox RotorElOffset 
         Height          =   285
         Left            =   240
         TabIndex        =   17
         ToolTipText     =   "Elevation angle added to that shown."
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox RotorAzOffset 
         Height          =   285
         Left            =   240
         TabIndex        =   15
         ToolTipText     =   "Azimuth angle added to that shown."
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Elevation:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Azimuth:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton Finish 
      Caption         =   "Close"
      Height          =   495
      Left            =   1920
      TabIndex        =   12
      ToolTipText     =   "Close this window."
      Top             =   4797
      Width           =   975
   End
   Begin VB.CheckBox RotorBidir 
      Caption         =   "Bidirectional Interface"
      Height          =   364
      Left            =   1755
      TabIndex        =   11
      ToolTipText     =   "Uncheck if interface is only PC->Control Box."
      Top             =   819
      Width           =   1534
   End
   Begin VB.CheckBox RotorSouth 
      Caption         =   "South stop"
      Height          =   247
      Left            =   1755
      TabIndex        =   10
      ToolTipText     =   "Substract 180deg from azimuth angle, except for GS-232 interface."
      Top             =   1521
      Width           =   1534
   End
   Begin VB.CheckBox RotorAutoFlip 
      Caption         =   "Auto flip detect"
      Height          =   364
      Left            =   1755
      TabIndex        =   9
      ToolTipText     =   "Check if DDE Server is unable of flipping, uncheck for WiSP."
      Top             =   1170
      Width           =   1534
   End
   Begin VB.TextBox RotorStep 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Internally round angles to nearest integer multiple of this value."
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Save 
      Caption         =   "Save"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      ToolTipText     =   "Save settings to registry."
      Top             =   4797
      Width           =   975
   End
   Begin VB.ComboBox RotorBaud 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Baudrate selection, only for interfaces that use serial ports."
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ComboBox RotorPort 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      ToolTipText     =   "Rotor interface port selection."
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox RotorType 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Rotor interface type selection."
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label RotorTimeOutDelayLabel 
      Caption         =   "Time Out (Secs.):"
      Height          =   195
      Left            =   1800
      TabIndex        =   27
      Top             =   2220
      Width           =   1455
   End
   Begin VB.Label RotorPaceDelayLabel 
      Caption         =   "Pace Delay (Secs.):"
      Height          =   195
      Left            =   120
      TabIndex        =   25
      Top             =   2220
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Step (deg.):"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Baud Rate:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Port:"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Interface Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmRotor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Finish_Click()
    frmRotor.Hide
End Sub

Sub Form_Load()
    RotorPort.Clear
    RotorPort.AddItem "None"    ' Add each item to list.
    RotorPort.AddItem "COM1"
    RotorPort.AddItem "COM2"
    RotorPort.AddItem "COM3"
    RotorPort.AddItem "COM4"
    RotorPort.AddItem "LPT1 (378)"
    RotorPort.AddItem "LPT2 (278)"
    RotorPort.AddItem "LPT3 (3BC)"
    
    RotorType.Clear
    RotorType.AddItem "None"
    RotorType.AddItem "GS-232"
    RotorType.AddItem "FODTrack"
    RotorType.AddItem "CI-V"
    RotorType.AddItem "IF-100"
    RotorType.AddItem "TrakBox"
    RotorType.AddItem "EASYCOMM-I"
    RotorType.AddItem "RC2800PX"
    
    RotorBaud.Clear
    RotorBaud.AddItem "1200"    ' Add each item to list.
    RotorBaud.AddItem "2400"
    RotorBaud.AddItem "4800"
    RotorBaud.AddItem "9600"
    RotorBaud.AddItem "19200"
    RotorBaud.AddItem "38400"
    RotorBaud.AddItem "57600"
    
    'get rotor config from registry
    'get port (maybe COM or LPT)
    RotorPort.text = GetSetting("WiSP_DDE_Client", "Config", "Rotor_com", Rotor_com_default)
    'get baudrate (useless for LPT)
    RotorBaud.text = GetSetting("WiSP_DDE_Client", "Config", "Rotor_baud", Rotor_baud_default)
    'get controller type
    RotorType.text = GetSetting("WiSP_DDE_Client", "Config", "Rotor_mode", Rotor_mode_default)
    'get step
    RotorStep.text = GetSetting("WiSP_DDE_Client", "Config", "Rotor_step", Rotor_step_default)
    'get auto-flip enabling
    RotorAutoFlip.Value = GetSetting("WiSP_DDE_Client", "Config", "Rotor_flip", 0)
    'get Bidirectional Interface enabling
    RotorBidir.Value = GetSetting("WiSP_DDE_Client", "Config", "Rotor_bidir", 0)
    'get stop position selection (True=South)
    RotorSouth.Value = GetSetting("WiSP_DDE_Client", "Config", "Rotor_stop", 0)
    'get Azimuth offset:
    RotorAzOffset.text = GetSetting("WiSP_DDE_Client", "Config", "Rotor_AzOf")
    'get Elevation offset:
    RotorElOffset.text = GetSetting("WiSP_DDE_Client", "Config", "Rotor_ElOf")
    'get Azimuth park position:
    RotorAzPark.text = GetSetting("WiSP_DDE_Client", "Config", "Rotor_AzPark")
    'get Elevation park position:
    RotorElPark.text = GetSetting("WiSP_DDE_Client", "Config", "Rotor_ElPark")
    'get 450 degrees Azimuth flag
    Rotor450Deg.Value = GetSetting("WiSP_DDE_Client", "Config", "Rotor_450Deg", 0)
    'get pace down delay time:
    RotorPaceDelay.text = GetSetting("WiSP_DDE_Client", "Config", "Rotor_PaceDelay")
    'get pace time-out period:
    RotorTimeOutDelay.text = GetSetting("WiSP_DDE_Client", "Config", "Rotor_TimeOutDelay")
    
    CheckLog.Enabled = False
    
End Sub

Private Sub RotorType_Change()
Select Case RotorType.text
Case "TrakBox"
    RotorBidir.Value = 1
    RotorBidir.Enabled = False
    RotorAutoFlip.Value = 0
    RotorAutoFlip.Enabled = False
    RotorPaceDelay.Enabled = True
    RotorPaceDelayLabel.Enabled = True

Case "CI-V", "FODTrack", "IF-100", "GS-232", "RC2800PX"
    RotorBidir.Enabled = True
    RotorAutoFlip.Enabled = True
    RotorPaceDelay.Enabled = False
    RotorPaceDelay.text = ""
    RotorPaceDelayLabel.Enabled = False

Case "EASYCOMM-I"
    RotorBidir.Enabled = True
    RotorAutoFlip.Enabled = True
    RotorPaceDelay.Enabled = True
    RotorPaceDelayLabel.Enabled = True

End Select
End Sub

Private Sub RotorType_Click()
Call RotorType_Change
End Sub

Private Sub Save_Click()
    SaveSetting "WiSP_DDE_Client", "Config", "Rotor_mode", RotorType.text
    SaveSetting "WiSP_DDE_Client", "Config", "Rotor_com", RotorPort.text
    SaveSetting "WiSP_DDE_Client", "Config", "Rotor_baud", RotorBaud.text
    SaveSetting "WiSP_DDE_Client", "Config", "Rotor_step", RotorStep.text
    SaveSetting "WiSP_DDE_Client", "Config", "Rotor_flip", RotorAutoFlip.Value
    SaveSetting "WiSP_DDE_Client", "Config", "Rotor_stop", RotorSouth.Value
    SaveSetting "WiSP_DDE_Client", "Config", "Rotor_bidir", RotorBidir.Value
    SaveSetting "WiSP_DDE_Client", "Config", "Rotor_AzOf", RotorAzOffset.text
    SaveSetting "WiSP_DDE_Client", "Config", "Rotor_ElOf", RotorElOffset.text
    SaveSetting "WiSP_DDE_Client", "Config", "Rotor_AzPark", RotorAzPark.text
    SaveSetting "WiSP_DDE_Client", "Config", "Rotor_ElPark", RotorElPark.text
    SaveSetting "WiSP_DDE_Client", "Config", "Rotor_450Deg", Rotor450Deg.Value
    SaveSetting "WiSP_DDE_Client", "Config", "Rotor_PaceDelay", RotorPaceDelay.text
    SaveSetting "WiSP_DDE_Client", "Config", "Rotor_TimeOutDelay", RotorTimeOutDelay.text
End Sub
