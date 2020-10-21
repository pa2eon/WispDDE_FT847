VERSION 5.00
Begin VB.Form frmAutoSel 
   Caption         =   "Radio Auto-Selection Configuration"
   ClientHeight    =   3640
   ClientLeft      =   65
   ClientTop       =   351
   ClientWidth     =   5564
   Icon            =   "autosel.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3640
   ScaleWidth      =   5564
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton AutoSelReset 
      Caption         =   "Reset to Defaults"
      Height          =   495
      Left            =   3600
      TabIndex        =   15
      ToolTipText     =   "WARNING!, resets all fields to sample values."
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton AutoSelFinish 
      Caption         =   "Close"
      Height          =   495
      Left            =   2280
      TabIndex        =   12
      ToolTipText     =   "Close this window."
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton AutoSelSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   960
      TabIndex        =   11
      ToolTipText     =   "Save settings to Windows Registry."
      Top             =   3000
      Width           =   975
   End
   Begin VB.CheckBox AutoSelAccPortEnable 
      Caption         =   "Enable Acc. port"
      Height          =   195
      Left            =   3720
      TabIndex        =   10
      ToolTipText     =   "Check to enable sending data to accesory ports."
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame AutoSelAccPort 
      Caption         =   "Accesory port config:"
      Height          =   2295
      Left            =   3600
      TabIndex        =   1
      ToolTipText     =   "Use an accesory port to control relays and such."
      Top             =   480
      Width           =   1815
      Begin VB.TextBox AutoSelAccPortValue 
         Height          =   285
         Left            =   360
         TabIndex        =   18
         ToolTipText     =   "Data to send to accesory port when this radio gets selected, if same acc.port is used for up&downlink both values will be ORed."
         Top             =   1200
         Width           =   975
      End
      Begin VB.ComboBox AutoSelAccPortPort 
         Height          =   315
         Left            =   360
         TabIndex        =   17
         ToolTipText     =   "Accesory port address selection, can use predetermined values or also type-in non-standard ports."
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton AutoSelAccPortTest 
         Caption         =   "Test"
         Height          =   375
         Left            =   360
         TabIndex        =   16
         ToolTipText     =   "Toggle acc.port data between entered value and 0 (resting value)."
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Port value (dec):"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Port address(hex):"
         Height          =   252
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   1332
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Conditions to select this Radio:"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.TextBox AutoSelFreqs 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Type the frequency range(s) this radio is capable of tuning, '144-148 420-450' for example."
         Top             =   2160
         Width           =   3015
      End
      Begin VB.TextBox AutoSelModes 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Type those modes you want to receive with this radio or 'ALL'."
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox AutoSelSatellites 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Type those satellites you want to receive with this radio or 'ALL'."
         Top             =   960
         Width           =   3015
      End
      Begin VB.CheckBox AutoSelDownlink 
         Caption         =   "Downlink"
         Height          =   195
         Left            =   1800
         TabIndex        =   4
         ToolTipText     =   "Check to enable this radio as downlink channel."
         Top             =   360
         Width           =   975
      End
      Begin VB.CheckBox AutoSelUplink 
         Caption         =   "Uplink"
         Height          =   195
         Left            =   600
         TabIndex        =   3
         ToolTipText     =   "Check to enable this radio as uplink channel."
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Freq. ranges in MHz (or 'ALL'):"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "Modes (or 'ALL'):"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Satellites (or 'ALL'):"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmAutoSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AutoSelAccPortEnable_Click()
    If AutoSelAccPortEnable.Value Then
        AutoSelAccPort.Enabled = True
        AutoSelAccPortPort.Enabled = True
        AutoSelAccPortValue.Enabled = True
    Else
        AutoSelAccPort.Enabled = False
        AutoSelAccPortPort.Enabled = False
        AutoSelAccPortValue.Enabled = False
    End If
End Sub

Private Sub AutoSelAccPortTest_Click()
If AutoSelAccPortPort.text <> "" And _
    AutoSelAccPortValue.text <> "" Then
    If AutoSelAccPortTest.Caption = "Test" Then
        Call OutPort(frmMain.Cdbl2("&H" + frmMain.Firststring(AutoSelAccPortPort.text)), frmMain.Cdbl2(AutoSelAccPortValue.text))
        AutoSelAccPortTest.Caption = "TESTING"
        AutoSelAccPortPort.Enabled = False
        AutoSelAccPortValue.Enabled = False
    Else
        Call OutPort(frmMain.Cdbl2("&H" + frmMain.Firststring(AutoSelAccPortPort.text)), 0)
        AutoSelAccPortTest.Caption = "Test"
        AutoSelAccPortPort.Enabled = True
        AutoSelAccPortValue.Enabled = True
    End If
End If
End Sub

Private Sub AutoSelFinish_Click()
    frmAutoSel.Hide
End Sub

Private Sub AutoSelReset_Click()
    AutoSelSatellites.text = "UO-22 KO-23 KO-25 AO-16 UO-36"
    AutoSelModes.text = "USB LSB CW CW-N FM FM-N FM-W"
    AutoSelFreqs.text = "28-30 144-148 420-450"
    AutoSelAccPortEnable.Value = False
    AutoSelAccPortPort.text = "278"
    AutoSelAccPortValue.text = "16"
End Sub

Sub AutoSelSave_Click()
    'Save settings to windows registry...
    SaveSetting "WiSP_DDE_Client", "Rig" + frmRadio.RadioIndex.text, "Radio_AutoSelUplink", AutoSelUplink.Value
    SaveSetting "WiSP_DDE_Client", "Rig" + frmRadio.RadioIndex.text, "Radio_AutoSelDownlink", AutoSelDownlink.Value
    SaveSetting "WiSP_DDE_Client", "Rig" + frmRadio.RadioIndex.text, "Radio_AutoSelSats", AutoSelSatellites.text
    SaveSetting "WiSP_DDE_Client", "Rig" + frmRadio.RadioIndex.text, "Radio_AutoSelModes", AutoSelModes.text
    SaveSetting "WiSP_DDE_Client", "Rig" + frmRadio.RadioIndex.text, "Radio_AutoSelFreqs", AutoSelFreqs.text
    SaveSetting "WiSP_DDE_Client", "Rig" + frmRadio.RadioIndex.text, "Radio_AutoSelAccPortEnable", AutoSelAccPortEnable.Value
    SaveSetting "WiSP_DDE_Client", "Rig" + frmRadio.RadioIndex.text, "Radio_AutoSelAccPortPort", AutoSelAccPortPort.text
    SaveSetting "WiSP_DDE_Client", "Rig" + frmRadio.RadioIndex.text, "Radio_AutoSelAccPortValue", AutoSelAccPortValue.text
End Sub

Sub Form_Activate()
    Call Form_Load
End Sub

Private Sub Form_Load()
    frmAutoSel.Caption = "Radio number " + frmRadio.RadioIndex.text + " Auto-Selection Configuration"
    frmAutoSel.Frame1.Caption = "Conditions to select Radio number " + frmRadio.RadioIndex.text
    AutoSelAccPortPort.Clear
    AutoSelAccPortPort.AddItem "378 (LPT1)"
    AutoSelAccPortPort.AddItem "278"
    AutoSelAccPortPort.AddItem "3BC"
    AutoSelUplink.Value = GetSetting("WiSP_DDE_Client", "Rig" + frmRadio.RadioIndex.text, "Radio_AutoSelUplink", 0)
    AutoSelDownlink.Value = GetSetting("WiSP_DDE_Client", "Rig" + frmRadio.RadioIndex.text, "Radio_AutoSelDownlink", 0)
    AutoSelSatellites.text = GetSetting("WiSP_DDE_Client", "Rig" + frmRadio.RadioIndex.text, "Radio_AutoSelSats", "")
    AutoSelModes.text = GetSetting("WiSP_DDE_Client", "Rig" + frmRadio.RadioIndex.text, "Radio_AutoSelModes", "")
    AutoSelFreqs.text = GetSetting("WiSP_DDE_Client", "Rig" + frmRadio.RadioIndex.text, "Radio_AutoSelFreqs", "")
    AutoSelAccPortEnable.Value = GetSetting("WiSP_DDE_Client", "Rig" + frmRadio.RadioIndex.text, "Radio_AutoSelAccPortEnable", 0)
    AutoSelAccPortPort.text = GetSetting("WiSP_DDE_Client", "Rig" + frmRadio.RadioIndex.text, "Radio_AutoSelAccPortPort", "")
    AutoSelAccPortValue.text = GetSetting("WiSP_DDE_Client", "Rig" + frmRadio.RadioIndex.text, "Radio_AutoSelAccPortValue", "")
    Call AutoSelAccPortEnable_Click
End Sub
