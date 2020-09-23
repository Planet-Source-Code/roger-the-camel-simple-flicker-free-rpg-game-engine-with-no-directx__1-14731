VERSION 5.00
Begin VB.Form frmIP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect"
   ClientHeight    =   1515
   ClientLeft      =   5730
   ClientTop       =   4860
   ClientWidth     =   3915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   3915
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "IP Address:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   810
   End
End
Attribute VB_Name = "frmIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConnect_Click()
On Error GoTo Errorhandler

GameForm.Winsock1.Close
GameForm.Winsock1.RemoteHost = txtIP.Text
GameForm.Winsock1.RemotePort = 1002
GameForm.Winsock1.Connect
GameForm.mnuDis.Enabled = True
GameForm.mnuLis.Enabled = False
GameForm.mnuConn.Enabled = False
GameForm.StatusBar1.Panels.Item(1).Text = "Connected to " & txtIP.Text
receiveMap = True
Unload Me
Exit Sub
Errorhandler:
MsgBox "Couldn't connect to " & txtIP.Text, vbExclamation, "Induced Evolution"
End Sub

