VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form GameForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Induced Evolution"
   ClientHeight    =   6270
   ClientLeft      =   645
   ClientTop       =   2400
   ClientWidth     =   6285
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   6285
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   195
      Left            =   3420
      TabIndex        =   8
      Top             =   6060
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   5625
      Width           =   5220
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   7
      Top             =   6000
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5900
            Text            =   "Unconnected"
            TextSave        =   "Unconnected"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   6195
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4800
      Width           =   6255
      Begin VB.Label Label3 
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   480
         Width           =   6255
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   240
         Width           =   6255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   6255
      End
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   5640
      Width           =   855
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5160
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4680
      Top             =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5640
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   27
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "game form.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "game form.frx":0C54
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "game form.frx":18A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "game form.frx":24FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "game form.frx":3150
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "game form.frx":3DA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "game form.frx":49F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "game form.frx":564C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "game form.frx":62A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "game form.frx":6EF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "game form.frx":7B48
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "game form.frx":879C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "game form.frx":93F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "game form.frx":A044
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "game form.frx":AC98
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "game form.frx":B8EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "game form.frx":C540
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "game form.frx":D194
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "game form.frx":DDE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "game form.frx":EA3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "game form.frx":F690
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "game form.frx":102E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "game form.frx":10F38
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "game form.frx":11B8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "game form.frx":127E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "game form.frx":13434
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "game form.frx":14088
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox sq 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   0
      Picture         =   "game form.frx":14CDC
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   480
   End
   Begin VB.Menu mnuCon 
      Caption         =   "&Connection"
      Begin VB.Menu mnuConn 
         Caption         =   "&Connect"
      End
      Begin VB.Menu mnuLis 
         Caption         =   "&Listen"
      End
      Begin VB.Menu mnuDis 
         Caption         =   "&Disconnect"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "GameForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xMan As Long
Dim yMan As Long
Dim i As Integer
Dim ScrX As Integer
Dim ScrY As Integer
Dim Sending As Boolean
Dim downmap As String

Private Sub cmdSend_Click()
sendText Text2.Text
Label3.Caption = Label2.Caption
Label2.Caption = Label1.Caption
Label1.Caption = Text2.Text
Label3.ForeColor = Label2.ForeColor
Label2.ForeColor = Label1.ForeColor
Label1.ForeColor = &HFF0000
Text2.Text = ""
Text2.SetFocus
End Sub

Private Sub Form_Load()
Let screenw = 13
Let screenh = 10
Let startpoint = 0
sq(0).AutoSize = True
sq(0).Picture = ImageList1.ListImages.Item(Grass).Picture
receiveMap = False
For i = 1 To screenh
For j = 1 To screenw
If i = 1 And j = 1 Then Let j = 2
Let a = (((i - 1) * screenw) + j - 1)
Load sq(a)
sq(a).AutoSize = True
sq(a).Picture = ImageList1.ListImages.Item(Grass).Picture
Let sq(a).Left = (j - 1) * sq(a).Width
Let sq(a).Top = (i - 1) * sq(a).Height
Let sq(a).Visible = True
Next j
Next i

'On Error GoTo dialog

LoadMap App.Path & "\map1.txt"
ScrX = startscrX(startpoint)
ScrY = startscrY(startpoint)

level1 startscrX(startpoint), startscrY(startpoint)
xMan = startpointX(startpoint)
yMan = startpointY(startpoint)
AddImage xMan, yMan, Front
Exit Sub

dialog:
MsgBox "Map not found", vbExclamation, "Induced Evolution"
End Sub

Private Sub mnuConn_Click()
frmIP.Show
End Sub

Private Sub mnuDis_Click()
d = MsgBox("Are you sure you want to disconnect from " & Winsock1.RemoteHost, vbOKCancel + vbQuestion, "Induced Evolution")
If d = 2 Then Exit Sub
Winsock1.Close
StatusBar1.Panels.Item(1).Text = "Unconnected"
mnuConn.Enabled = True
mnuLis.Enabled = True
mnuDis.Enabled = False
End Sub

Private Sub mnuLis_Click()
Winsock1.LocalPort = 1002
Winsock1.Listen
StatusBar1.Panels.Item(1).Text = "Listening on..." & Winsock1.LocalIP
End Sub


Private Sub sq_GotFocus(Index As Integer)
Text2.SetFocus
End Sub

Private Sub sq_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyEscape Then End
'    If KeyCode = vbKeyLeft Then Key = "A"
'    If KeyCode = vbKeyRight Then Key = "D"
'    If KeyCode = vbKeyUp Then Key = "W"
'    If KeyCode = vbKeyDown Then Key = "S"
'    If KeyCode = vbKeyP Then MoveClouds = Not MoveClouds
End Sub

Private Sub sq_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'If Button = 1 Then
'    sq(Index).Cls
'    ImageList1.ListImages.Item(Tree).Draw sq(Index).hDC, 0, 0, 1
'    sq(Index).Tag = "solid"
'End If
'If Button = 2 Then
'    sq(Index).Cls
'    ImageList1.ListImages.Item(Hole).Draw sq(Index).hDC, 0, 0, 1
'    sq(Index).Tag = ""
'End If
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyLeft Then Key = "A"
    If KeyCode = vbKeyRight Then Key = "D"
    If KeyCode = vbKeyUp Then Key = "W"
    If KeyCode = vbKeyDown Then Key = "S"
    If KeyCode = vbKeyP Then MoveClouds = Not MoveClouds
End Sub
Private Sub Timer1_Timer()
If Key2 = "S" Then
    sq(IndexFind(xMan - 1, yMan)).Cls
    sq(IndexFind(xMan, yMan)).Cls

    AddImage xMan - 1, yMan, Grass
    AddImage xMan, yMan, Front
    Key2 = ""
ElseIf Key2 = "W" Then
    sq(IndexFind(xMan + 1, yMan)).Cls
    sq(IndexFind(xMan, yMan)).Cls

    AddImage xMan + 1, yMan, Grass
    AddImage xMan, yMan, Back
    Key2 = ""
ElseIf Key2 = "A" Then
    sq(IndexFind(xMan, yMan + 1)).Cls
    sq(IndexFind(xMan, yMan)).Cls

    AddImage xMan, yMan + 1, Grass
    AddImage xMan, yMan, Leftt
    Key2 = ""
ElseIf Key2 = "D" Then
    sq(IndexFind(xMan, yMan - 1)).Cls
    sq(IndexFind(xMan, yMan)).Cls

    AddImage xMan, yMan - 1, Grass
    AddImage xMan, yMan, Rightt
    Key2 = ""
ElseIf Key = "S" Then
    If xMan = 10 Then
        ScrX = ScrX + 1
        level1 ScrX, ScrY
        xMan = 1
        AddImage xMan, yMan, Front
        Key = ""
        GoTo 10
    End If
    DetectObj xMan + 1, yMan
    If Key = "q" Then GoTo 10
    xMan = xMan + 1
    sq(IndexFind(xMan - 1, yMan)).Cls
    sq(IndexFind(xMan, yMan)).Cls
    
    AddImage xMan - 1, yMan, FrontD
    AddImage xMan, yMan, FrontU
    Key = ""
    Key2 = "S"
ElseIf Key = "W" Then
If xMan = 1 Then
        ScrX = ScrX - 1
        level1 ScrX, ScrY
        xMan = 10
        AddImage xMan, yMan, Back
        Key = ""
        GoTo 10
    End If
    DetectObj xMan - 1, yMan
    If Key = "q" Then GoTo 10
    xMan = xMan - 1
    sq(IndexFind(xMan + 1, yMan)).Cls
    sq(IndexFind(xMan, yMan)).Cls

    AddImage xMan + 1, yMan, BackD
    AddImage xMan, yMan, BackU
    Key = ""
    Key2 = "W"
ElseIf Key = "A" Then
If yMan = 1 Then
        ScrY = ScrY - 1
        level1 ScrX, ScrY
        yMan = 13
        AddImage xMan, yMan, Leftt
        Key = ""
        GoTo 10
    End If
    DetectObj xMan, yMan - 1
    If Key = "q" Then GoTo 10
    yMan = yMan - 1
    sq(IndexFind(xMan, yMan + 1)).Cls
    sq(IndexFind(xMan, yMan)).Cls

    AddImage xMan, yMan + 1, LeftD
    AddImage xMan, yMan, LeftU
    Key = ""
    Key2 = "A"
ElseIf Key = "D" Then
If yMan = 13 Then
        ScrY = ScrY + 1
        level1 ScrX, ScrY
        yMan = 1
        AddImage xMan, yMan, Rightt
        Key = ""
        GoTo 10
    End If
    DetectObj xMan, yMan + 1
    If Key = "q" Then GoTo 10
    yMan = yMan + 1
    sq(IndexFind(xMan, yMan - 1)).Cls
    sq(IndexFind(xMan, yMan)).Cls

    AddImage xMan, yMan - 1, RightD
    AddImage xMan, yMan, RightU
    Key = ""
    Key2 = "D"
End If
10
End Sub

Private Sub Winsock1_Close()
MsgBox "Disconnected", vbInformation, "Induced Evolution"
Winsock1.Close
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close
Winsock1.Accept requestID
StatusBar1.Panels.Item(1).Text = "Connected to " & Winsock1.RemoteHost
sendMap App.Path & "\map1.txt"
End Sub

Public Sub sendText(TextMsg As String)
    'Make sure there is a connection
If Winsock1.State <> sckConnected Then Exit Sub
    'Send the string
Winsock1.SendData "TEXT!" & TextMsg
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim strdata As String
Winsock1.GetData strdata
DoEvents
If Left(strdata, 4) = "MAP!" Then
    downmap = ""
    Sending = True
    strdata = Right(strdata, Len(strdata) - 4)
    If Right(strdata, 4) = "!END" Then
        Sending = False
        strdata = Left(strdata, Len(strdata) - 4)
        downmap = downmap & strdata
        makefile downmap
        Exit Sub
    End If
    downmap = downmap & strdata
    Exit Sub
End If

If Sending = True Then
    If Right(strdata, 4) = "!END" Then
        Sending = False
        strdata = Left(strdata, Len(strdata) - 4)
        downmap = downmap & strdata
        makefile downmap
        Exit Sub
    End If
    downmap = downmap & strdata
    Exit Sub
End If

If Left(strdata, 12) = "@31*7%43(&45" Then
    strdata = Right(strdata, Len(strdata) - 12)
    If Mid(strdata, 13, 1) = 1 Then
    End If
    Exit Sub
End If

If Left(strdata, 5) = "TEXT!" Then
strdata = Right(strdata, Len(strdata) - 5)

Label3.Caption = Label2.Caption
Label2.Caption = Label1.Caption
Label1.Caption = strdata
Label3.ForeColor = Label2.ForeColor
Label2.ForeColor = Label1.ForeColor
Label1.ForeColor = &HFF&
End If
End Sub

Public Sub sendControl(Control As Integer)
    'Make sure there is a connection
If Winsock1.State <> sckConnected Then Exit Sub
    'Send the string
Winsock1.SendData "@31*7%43(&45" & Control
End Sub
Public Sub sendMap(FilePath As String)
    'Make sure there is a connection
If Winsock1.State <> sckConnected Then Exit Sub
Filedata = ""
Open FilePath For Input As #1
   DoEvents
   Filedata = Input(LOF(1), 1)
   Close #1
   MsgBox Filedata
   Winsock1.SendData "MAP!" & Filedata & "!END"
   DoEvents
End Sub

Private Sub Winsock1_SendComplete()
ProgressBar1.Value = 0
End Sub

Private Sub Winsock1_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
ProgressBar1.Max = bytesSent + bytesRemaining
ProgressBar1.Value = bytesSent
End Sub

Public Sub makefile(MapStr As String)
Open App.Path & "\temp.txt" For Output As #2
Print #2, downmap
Close #2
startpoint = 1
LoadMap App.Path & "\temp.txt"
ScrX = startscrX(startpoint)
ScrY = startscrY(startpoint)

level1 startscrX(startpoint), startscrY(startpoint)
xMan = startpointX(startpoint)
yMan = startpointY(startpoint)
AddImage xMan, yMan, Front

End Sub
