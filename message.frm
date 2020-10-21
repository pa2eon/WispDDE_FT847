VERSION 5.00
Begin VB.Form frmMessage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Message"
   ClientHeight    =   1605
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3780
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer MessageTimer 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   0
      Top             =   1140
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   855
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "message.frx":0000
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CancelButton_Click()
    frmMessage.Tag = "Cancel"
    frmMessage.Hide
End Sub

Private Sub Form_Load()
    frmMessage.Tag = ""
End Sub

Private Sub MessageTimer_Timer()
Hide
MessageTimer.Enabled = False
End Sub

Private Sub OKButton_Click()
    frmMessage.Tag = "OK"
    frmMessage.Hide
End Sub

Sub ShowMessage(m$, timeout As Double)
    If MessageTimer.Enabled = True Then Exit Sub
    Label1.Caption = m$
    Tag = ""
    Show
    Beep
    MessageTimer.Enabled = True
End Sub
