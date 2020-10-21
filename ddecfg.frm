VERSION 5.00
Begin VB.Form frmDdelink 
   Caption         =   "DDE Settings"
   ClientHeight    =   3458
   ClientLeft      =   65
   ClientTop       =   351
   ClientWidth     =   3419
   Icon            =   "ddecfg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3458
   ScaleWidth      =   3419
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox CheckLog 
      Caption         =   "Log Events"
      Enabled         =   0   'False
      Height          =   247
      Left            =   1872
      TabIndex        =   15
      Top             =   2340
      Width           =   1300
   End
   Begin VB.TextBox Decimal 
      Height          =   247
      Left            =   117
      TabIndex        =   13
      Text            =   "."
      Top             =   2457
      Width           =   247
   End
   Begin VB.ComboBox DDEFormat 
      Height          =   315
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "Select the tracking program to provide frequency calculations."
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Satellite Data"
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      ToolTipText     =   "When using Nova, a satellites frequencies database is kept."
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      ToolTipText     =   "Close this window."
      Top             =   2808
      Width           =   975
   End
   Begin VB.TextBox Interval 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      ToolTipText     =   $"ddecfg.frx":030A
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox Item 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "DDE link item for frequency and mode info of satellite tracking application."
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Topic 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      ToolTipText     =   "DDE link topic for frequency and mode info of satellite tracking application."
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox SourceApplication 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "DDE name of satellite tracking application."
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   495
      Left            =   468
      TabIndex        =   2
      ToolTipText     =   "Save settings to Windows Registry."
      Top             =   2808
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Decimal separator:"
      Height          =   247
      Left            =   117
      TabIndex        =   14
      Top             =   2223
      Width           =   1300
   End
   Begin VB.Label DDE 
      Caption         =   "Receive DDE from:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Query Interval (sec.):"
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Link Item:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Link Topic:"
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Source Application:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "frmDdelink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    SaveSetting "WiSP_DDE_Client", "Config", "Dde_source", SourceApplication.text
    SaveSetting "WiSP_DDE_Client", "Config", "Dde_topic", Topic.text
    SaveSetting "WiSP_DDE_Client", "Config", "Dde_item", Item.text
    SaveSetting "WiSP_DDE_Client", "Config", "Dde_period", Interval.text
    SaveSetting "WiSP_DDE_Client", "Config", "Dde_format", DDEFormat.text
    'make sure decimal separator is OK
    If (Decimal.text <> "." And Decimal.text <> ",") Then
        frmDdelink.Decimal.text = "."
    End If
    SaveSetting "WiSP_DDE_Client", "Config", "Decimal_Separator", Decimal.text
    'if WiSPDDE is to be client, we update settings,
    If DDEFormat.text <> "SatPC32" Then
        'Set the proper value for the Timer (convert to
        'milliseconds)
        If frmMain.Cdbl2(Interval.text) <> 0 Then
            frmMain.DDEPollTimer.Interval = frmMain.Cdbl2(Interval.text) * 1000
            frmMain.DDEPollTimer.Enabled = True
        Else
            frmMain.DDEPollTimer.Enabled = False
        End If
        'Source and Topic goes together separated by '|':
        frmMain.DDELabel.LinkTopic = SourceApplication.text + "|" + Topic.text
        frmMain.DDELabel.LinkItem = Item.text
        frmMain.DDE_Test
    End If
End Sub

Private Sub Command2_Click()
    frmDdelink.Hide
End Sub

Private Sub Command3_Click()
    frmSats.Show
End Sub

Private Sub DDEFormat_Change()
    Select Case DDEFormat.text
    Case Is = "WiSP"
        SourceApplication.text = "GSC"
        Topic.text = "Tracking"
        Item.text = "Tracking"
        Interval.text = "3"
        
        Command3.Enabled = False
    Case Is = "Station"
        SourceApplication.text = "Station"
        Topic.text = "Tracking"
        Item.text = "General"
        Interval.text = "3"
        
        Command3.Enabled = False

    Case Is = "Winorbit"
        SourceApplication.text = "WinOrbit"
        Topic.text = "TrackingInfo"
        Item.text = "SatelliteName"
        Interval.text = "3"
    
        Command3.Enabled = False

    Case Is = "Nova"
        SourceApplication.text = "NFW32"
        Topic.text = "NFW_DATA"
        Item.text = "NFW_SERVER"
        Interval.text = "3"
    
        Command3.Enabled = True

    Case Is = "SatPC32"
        SourceApplication.text = "SatPC32"
        Topic.text = "SatPcDdeConv"
        Item.text = "SatPcDdeItem"
        Interval.text = "3"
        
        Command3.Enabled = False
        
    Case Is = "Satscape"
        SourceApplication.text = "Satscape"
        Topic.text = "Tracking"
        Item.text = "Tracking"
        Interval.text = "3"
    
        Command3.Enabled = False

    Case Is = "WXtrack"
        SourceApplication.text = "WXtrack"
        Topic.text = "Tracking"
        Item.text = "Tracking"
        Interval.text = "3"
    
        Command3.Enabled = False

    Case Is = "Orbitron"
        SourceApplication.text = "Orbitron"
        Topic.text = "Tracking"
        Item.text = "TrackingData"
        Interval.text = "1"
    
        Command3.Enabled = False

    End Select
    

End Sub

Private Sub DDEFormat_Click()
    Call DDEFormat_Change
End Sub

Private Sub Form_Load()
    
    DDEFormat.AddItem "None"
    DDEFormat.AddItem "Orbitron"
    DDEFormat.AddItem "WiSP"
    DDEFormat.AddItem "Station"
    DDEFormat.AddItem "Winorbit"
    DDEFormat.AddItem "Nova"
    DDEFormat.AddItem "SatPC32"
    DDEFormat.AddItem "Satscape"
    DDEFormat.AddItem "WXtrack"

    CheckLog.Enabled = False
    CheckLog.Value = 0
    
End Sub
