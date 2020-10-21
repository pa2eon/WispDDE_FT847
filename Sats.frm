VERSION 5.00
Begin VB.Form frmSats 
   Caption         =   "Satellite Data ( for NFW)"
   ClientHeight    =   3718
   ClientLeft      =   65
   ClientTop       =   351
   ClientWidth     =   4147
   LinkTopic       =   "Form1"
   ScaleHeight     =   3718
   ScaleWidth      =   4147
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Sat2Dnlink 
      Caption         =   "Uplink channel is second dnlink."
      Height          =   364
      Left            =   2340
      TabIndex        =   19
      ToolTipText     =   "Check this to have both channes doppler corrected for downlink."
      Top             =   1287
      Width           =   1534
   End
   Begin VB.CheckBox SatDirTrack 
      Caption         =   "Direct Track"
      Height          =   255
      Left            =   2340
      TabIndex        =   17
      Top             =   936
      Width           =   1183
   End
   Begin VB.CheckBox SatRevTrack 
      Caption         =   "Reverse Track"
      Height          =   255
      Left            =   2340
      TabIndex        =   16
      Top             =   585
      Width           =   1183
   End
   Begin VB.CheckBox SatSatEnabled 
      Caption         =   "Enable this Satellite"
      Height          =   375
      Left            =   2340
      TabIndex        =   15
      Top             =   182
      Width           =   1155
   End
   Begin VB.TextBox SatName 
      Height          =   285
      Left            =   240
      TabIndex        =   13
      ToolTipText     =   "Satellite name as in Nova's database."
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton SatDeleteLast 
      Caption         =   "Delete least Satellite"
      Height          =   495
      Left            =   2760
      TabIndex        =   12
      ToolTipText     =   "WARNING!. Deletes the highest index satellte."
      Top             =   3055
      Width           =   975
   End
   Begin VB.CommandButton SatClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   1560
      TabIndex        =   11
      ToolTipText     =   "Close this window."
      Top             =   3055
      Width           =   975
   End
   Begin VB.CommandButton SatSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   360
      TabIndex        =   10
      ToolTipText     =   "Save settings to Windows registry."
      Top             =   3055
      Width           =   975
   End
   Begin VB.TextBox SatDownlinkMode 
      Height          =   285
      Left            =   2280
      TabIndex        =   7
      ToolTipText     =   "RX downlink modulation mode."
      Top             =   2574
      Width           =   1095
   End
   Begin VB.TextBox SatUplinkMode 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "TX uplink modulation mode."
      Top             =   2574
      Width           =   1095
   End
   Begin VB.TextBox SatDownlinkFreq 
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      ToolTipText     =   "Downlink frequency of the satellite."
      Top             =   1911
      Width           =   1575
   End
   Begin VB.TextBox SatUplinkFreq 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Uplink frequency of the satellite."
      Top             =   1911
      Width           =   1575
   End
   Begin VB.ComboBox SatIndex 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Index number of the satelite to edit, or pick the last number to add a new satellite."
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Direct Track"
      Height          =   260
      Left            =   2574
      TabIndex        =   18
      Top             =   1027
      Width           =   1092
   End
   Begin VB.Label Label8 
      Caption         =   "Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Downlink mode:"
      Height          =   260
      Left            =   2275
      TabIndex        =   9
      Top             =   2340
      Width           =   1573
   End
   Begin VB.Label Label4 
      Caption         =   "Uplink mode:"
      Height          =   260
      Left            =   234
      TabIndex        =   8
      Top             =   2340
      Width           =   1092
   End
   Begin VB.Label Label3 
      Caption         =   "Downlink freq.:"
      Height          =   260
      Left            =   2275
      TabIndex        =   5
      Top             =   1677
      Width           =   1573
   End
   Begin VB.Label Label2 
      Caption         =   "Satellite Index:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Uplink freq.:"
      Height          =   260
      Left            =   234
      TabIndex        =   2
      Top             =   1677
      Width           =   1573
   End
End
Attribute VB_Name = "frmSats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    'Add as many entries to the "sat Index" combo as we find
    'in the registry...
    SatIndex.Clear
    i% = 0
    Do
        i% = i% + 1
        SatIndex.AddItem LTrim$(Str$(i%))
    Loop Until GetSetting("WiSP_DDE_Client", "Sat" + LTrim$(Str$(i%)), "SatName", "-") = "-"
    
    SatName.text = ""
    SatUplinkFreq.text = ""
    SatDownlinkFreq.text = ""
    SatUplinkMode.text = ""
    SatDownlinkMode.text = ""
    SatDirTrack.Value = 0
    SatRevTrack.Value = 0
    
    Call SatSatEnabled_Click
    
End Sub


Private Sub SatClose_Click()
    frmSats.Hide
End Sub

Private Sub SatDeleteLast_Click()
    'remove last SatN folder from registry:
    'Add as many entries to the "sat Index" combo as we find
    'in the registry...
    SatIndex.Clear
    i% = 0
    Do
        i% = i% + 1
        SatIndex.AddItem LTrim$(Str$(i%))
    Loop Until (GetSetting("WiSP_DDE_Client", "Sat" + LTrim$(Str$(i%)), "SatName", "-") = "-")
    
    If i% > 1 Then
        DeleteSetting "WiSP_DDE_Client", "Sat" + LTrim(Str(i% - 1))
    End If
    
    'update the index combo as we may have more sats now!
    'Add as many entries to the "sat Index" combo as we find
    'in the registry...
    a% = SatIndex.ListIndex
    SatIndex.Clear
    i% = 0
    Do
        i% = i% + 1
        SatIndex.AddItem LTrim$(Str$(i%))
    Loop Until (GetSetting("WiSP_DDE_Client", "Sat" + LTrim$(Str$(i%)), "SatName", "-") = "-")
    
    SatIndex.ListIndex = a%
    
    
    If a% >= SatIndex.ListCount Then
        SatIndex.ListIndex = SatIndex.ListCount - 1
    Else
        SatIndex.ListIndex = a%
    End If
    
    Call SatIndex_Change

End Sub

Private Sub SatDirTrack_Click()
If SatDirTrack.Value = 1 Then
    SatRevTrack.Value = 0
End If

End Sub

Private Sub SatIndex_Change()
    'update settings for selected satellite
    'Retrieve configuration for selected sat:
    SatName.text = GetSetting("WiSP_DDE_Client", "Sat" + LTrim$(SatIndex.text), "SatName", "")
    SatUplinkFreq.text = GetSetting("WiSP_DDE_Client", "Sat" + LTrim$(SatIndex.text), "SatUplinkFreq", "")
    SatDownlinkFreq.text = GetSetting("WiSP_DDE_Client", "Sat" + LTrim$(SatIndex.text), "SatDownlinkFreq", "")
    SatUplinkMode.text = GetSetting("WiSP_DDE_Client", "Sat" + LTrim$(SatIndex.text), "SatUplinkMode", "")
    SatDownlinkMode.text = GetSetting("WiSP_DDE_Client", "Sat" + LTrim$(SatIndex.text), "SatDownlinkMode", "")
    SatDirTrack.Value = GetSetting("WiSP_DDE_Client", "Sat" + LTrim$(SatIndex.text), "SatDirTrack", 0)
    SatRevTrack.Value = GetSetting("WiSP_DDE_Client", "Sat" + LTrim$(SatIndex.text), "SatRevTrack", 0)
    SatSatEnabled.Value = GetSetting("WiSP_DDE_Client", "Sat" + LTrim$(SatIndex.text), "SatSatEnabled", 0)
    Sat2Dnlink.Value = GetSetting("WiSP_DDE_Client", "Sat" + LTrim$(SatIndex.text), "Sat2Dnlink", 0)
End Sub

Private Sub SatIndex_Click()
    Call SatIndex_Change
End Sub

Private Sub SatRevTrack_Click()
If SatRevTrack.Value = 1 Then
    SatDirTrack.Value = 0
End If
End Sub

Private Sub SatSatEnabled_Click()
If SatIndex.text = "" Then
    SatSatEnabled.Value = 0
End If
If SatSatEnabled.Value Then
    SatName.Enabled = True
    SatUplinkFreq.Enabled = True
    SatDownlinkFreq.Enabled = True
    SatUplinkMode.Enabled = True
    SatDownlinkMode.Enabled = True
    SatDirTrack.Enabled = True
    SatRevTrack.Enabled = True
    Sat2Dnlink.Enabled = True
    
    Label1.Enabled = True
    Label2.Enabled = True
    Label3.Enabled = True
    Label4.Enabled = True
    Label5.Enabled = True
    Label7.Enabled = True
    Label8.Enabled = True
Else
    SatName.Enabled = False
    SatUplinkFreq.Enabled = False
    SatDownlinkFreq.Enabled = False
    SatUplinkMode.Enabled = False
    SatDownlinkMode.Enabled = False
    SatDirTrack.Enabled = False
    SatRevTrack.Enabled = False
    Sat2Dnlink.Enabled = False

    Label1.Enabled = False
    Label2.Enabled = False
    Label3.Enabled = False
    Label4.Enabled = False
    Label5.Enabled = False
    Label7.Enabled = False
    Label8.Enabled = False
End If


End Sub

Private Sub SatSave_Click()
    If SatIndex.text <> "" Then
        'Save settings to windows registry...
        SaveSetting "WiSP_DDE_Client", "Sat" + SatIndex.text, "SatName", SatName.text
        SaveSetting "WiSP_DDE_Client", "Sat" + SatIndex.text, "SatDownlinkFreq", SatDownlinkFreq.text
        SaveSetting "WiSP_DDE_Client", "Sat" + SatIndex.text, "SatDownlinkMOde", SatDownlinkMode.text
        SaveSetting "WiSP_DDE_Client", "Sat" + SatIndex.text, "SatUplinkFreq", SatUplinkFreq.text
        SaveSetting "WiSP_DDE_Client", "Sat" + SatIndex.text, "SatUplinkMode", SatUplinkMode.text
        SaveSetting "WiSP_DDE_Client", "Sat" + SatIndex.text, "SatSatEnabled", SatSatEnabled.Value
        SaveSetting "WiSP_DDE_Client", "Sat" + SatIndex.text, "SatDirTrack", SatDirTrack.Value
        SaveSetting "WiSP_DDE_Client", "Sat" + SatIndex.text, "SatRevTrack", SatRevTrack.Value
        SaveSetting "WiSP_DDE_Client", "Sat" + SatIndex.text, "Sat2Dnlink", Sat2Dnlink.Value

        'update the index combo as we may have more sats now!
        'Add as many entries to the "sat Index" combo as we find
        'in the registry...
        a% = SatIndex.ListIndex
        SatIndex.Clear
        i% = 0
        Do
            i% = i% + 1
            SatIndex.AddItem LTrim$(Str$(i%))
        Loop Until (GetSetting("WiSP_DDE_Client", "Sat" + LTrim$(Str$(i%)), "SatName", "-") = "-")
        SatIndex.ListIndex = a%
    End If

End Sub
