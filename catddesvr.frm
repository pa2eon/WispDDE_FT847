VERSION 5.00
Begin VB.Form frmCatddesvr 
   Caption         =   "Cat DDE Server"
   ClientHeight    =   3172
   ClientLeft      =   65
   ClientTop       =   364
   ClientWidth     =   4680
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   ScaleHeight     =   3172
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmCatddesvr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
'SatPC32 no longer acts as Client, it acts as Server so standard
'routines (client) work ok for this program now.

'A command from DDE Client (ie SatPC32) is received
'Check if selected DDE mode is SatPC32
If frmDdelink.DDEFormat.text = "SatPC32" Then
    If CmdStr = "ClientOn" Then
        'frmMain.Satellite = "Tracked by " + frmDdelink.DDEFormat.text
    End If
    If CmdStr = "Close" Then
        frmMain.Satellite = "No DDE Link"
    End If
    'Set-Mode command:
    If InStr(CmdStr, "SM") Then
        a = InStr(CmdStr, "SM") + 2
        frmMain.DownlinkMode.text = frmMain.Firststring( _
            Mid$(CmdStr, a))
        'seek for space that separates up-link mode from dn-link mode:
        a = InStr(CmdStr, " ")
        frmMain.UplinkMode.text = frmMain.Firststring( _
            Mid$(CmdStr, a + 1))
    End If
    'Set-Frequency command:
    If InStr(CmdStr, "SF") Then
        a = InStr(CmdStr, "SF") + 2
        frmMain.DownlinkDDEFreq.text = Str$(0.001 * frmMain.Cdbl2(frmMain.Firststring( _
            Mid$(CmdStr, a))))
        'seek for space that separates up-link from dn-link:
        a = InStr(CmdStr, " ")
        frmMain.UplinkDDEFreq.text = Str$(0.001 * frmMain.Cdbl2(frmMain.Firststring( _
            Mid$(CmdStr, a))))
            
        'change sat name only after receiving some freq...
        If frmMain.Satellite = "No DDE Link" Then
            frmMain.Satellite = "Tracked by " + frmDdelink.DDEFormat.text
            'changing sat name will trigger auto radio selection, so
            'wait untill done...
            'DoEvents
        End If
            
        '****RADIO CONTROL PORCESSING*****
        Call frmMain.UpdateDownlink
        Call frmMain.UpdateUplink
        
    End If
    Label1.Caption = CmdStr
    'frmCatddesvr.Show
End If

End Sub

Private Sub Form_LinkOpen(Cancel As Integer)
    'frmCatddesvr.Show
    

End Sub
