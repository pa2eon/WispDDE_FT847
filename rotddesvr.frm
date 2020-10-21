VERSION 5.00
Begin VB.Form frmRotddesvr 
   Caption         =   "Rotor DDE Server"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmRotddesvr"
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
        frmMain.Satellite = "Tracked by " + frmDdelink.DDEFormat.text
    End If
    If CmdStr = "Close" Then
        frmMain.Satellite = "No DDE Link"
    End If
    'if command begins in number: angles are comming...
    If IsNumeric(Left$(CmdStr, 1)) Then
        frmMain.Azimuth.text = Str$(frmMain.Cdbl2(frmMain.Firststring(CmdStr)))
        a = InStr(CmdStr, " ")
        frmMain.Elevation.text = Str$(frmMain.Cdbl2(frmMain.Firststring(Mid$(CmdStr, a))))
        
        'rotor control is processed:
        Az = frmMain.Cdbl2(frmMain.Azimuth.text)
        El = frmMain.Cdbl2(frmMain.Elevation.text)
        If frmMain.RotorAuto.Value Then
            Az = frmMain.Cdbl2(frmRotor.RotorStep.text) * _
                (CInt(Az / frmMain.Cdbl2(frmRotor.RotorStep.text)))
            El = frmMain.Cdbl2(frmRotor.RotorStep.text) * _
                (CInt(El / frmMain.Cdbl2(frmRotor.RotorStep.text)))
            Call frmMain.UpdateRotor(Az, El)
        End If
    End If
    
    
    Label1.Caption = CmdStr
    'frmRotddesvr.Show

End If

End Sub

Private Sub Form_LinkOpen(Cancel As Integer)
    'frmRotddesvr.Show

End Sub
