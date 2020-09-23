Attribute VB_Name = "Module1"
'map numbers, each number in txt file = the const name
Global screenw As Long
Global screenh As Long
Public Const Grass As Long = 1
Public Const Tree As Long = 2
Public Const Hole As Long = 3
Public Const Front As Long = 4
Public Const Back As Long = 5
Public Const Leftt As Long = 6
Public Const Rightt As Long = 7
Public Const FrontD As Long = 8
Public Const BackD As Long = 10
Public Const LeftD As Long = 12
Public Const RightD As Long = 14
Public Const FrontU As Long = 9
Public Const BackU As Long = 11
Public Const LeftU As Long = 13
Public Const RightU As Long = 15
Public Const Mist As Long = 16
Public Const Pistol As Long = 17
Public Const Shotgun As Long = 18
Public Const Mines As Long = 19
Public Const LeftCorner As Long = 20
Public Const Wall As Long = 21
Public Const RightCorner As Long = 22
Public Const LeftUpWall As Long = 23
Public Const RightUpWall As Long = 24
Public Const LeftBackCorner As Long = 25
Public Const BackWall As Long = 26
Public Const RightBackCorner As Long = 27
'for remebering what key was pressed
Public Key As String
Public Key2 As String
'for reading the map txt
Public MapRead As String
Public MapWriteR As Long
Public MapWriteC As Long
Dim MapLength As Long
Dim Loopy As Long
Dim Loopy2 As Long
'for storing the map
Global data(400, 400) As String
Global dRow As Integer
Global dCol As Integer
'constants for controls
Public Const DirUp As Integer = 1
Public Const DirDown As Integer = 2
Public Const DirLeft As Integer = 3
Public Const DirRight As Integer = 4
'opponents position
Global enemyX As Integer
Global enemyY As Integer
Global receiveMap As Boolean
'startpoints
Global startscrX(10) As Integer
Global startscrY(10) As Integer
Global startpointX(10) As Integer
Global startpointY(10) As Integer
Global stp As Integer


Sub LoadMap(FilePath As String)
Open FilePath For Input As #1
MapWriteR = 0
MapWriteC = 0
Do
    On Error GoTo Done
    Line Input #1, MapRead
    MapLength = Len(MapRead)
    For Loopy = 1 To MapLength
        data(MapWriteR, MapWriteC) = Mid(MapRead, Loopy, 1)
        If data(MapWriteR, MapWriteC) = "*" Then
            startscrX(stp) = 1
            MapWriteRB = MapWriteR
            Do While MapWriteRB >= 10
            MapWriteRB = MapWriteRB - 10
            startscrX(stp) = startscrX(stp) + 1
            Loop
            
            startpointX(stp) = MapWriteRB + 1
            
            startscrY(stp) = 1
            MapWriteCB = MapWriteC
            Do While MapWriteCB >= 13
            MapWriteCB = MapWriteCB - 13
            startscrY(stp) = startscrY(stp) + 1
            Loop
            
            startpointY(stp) = MapWriteCB + 1
            stp = stp + 1
        End If
        MapWriteC = MapWriteC + 1
    Next
    MapWriteC = 0
    MapWriteR = MapWriteR + 1
Loop
Done:
Close #1
End Sub

Sub level1(XMap As Integer, YMap As Integer)
For Loopy = 0 To 129
    GameForm.sq(Loopy).Tag = ""
    GameForm.sq(Loopy).Cls
Next Loopy
For Loopy = 0 To 9
    For Loopy2 = 0 To 12
        dRow = Loopy + ((XMap * 10) - 10)
        dCol = Loopy2 + ((YMap * 13) - 13)
        If data(dRow, dCol) = "" Or data(dRow, dCol) = " " Then
            AddImage Loopy + 1, Loopy2 + 1, Mist, "solid"
        ElseIf data(dRow, dCol) = "0" Then
            AddImage Loopy + 1, Loopy2 + 1, Grass
        ElseIf data(dRow, dCol) = "1" Then
            AddImage Loopy + 1, Loopy2 + 1, Tree, "solid"
        ElseIf data(dRow, dCol) = "2" Then
            AddImage Loopy + 1, Loopy2 + 1, Hole
        ElseIf data(dRow, dCol) = "3" Then
            AddImage Loopy + 1, Loopy2 + 1, LeftCorner, "solid"
        ElseIf data(dRow, dCol) = "4" Then
            AddImage Loopy + 1, Loopy2 + 1, Wall, "solid"
        ElseIf data(dRow, dCol) = "5" Then
            AddImage Loopy + 1, Loopy2 + 1, RightCorner, "solid"
        ElseIf data(dRow, dCol) = "6" Then
            AddImage Loopy + 1, Loopy2 + 1, LeftUpWall, "solid"
        ElseIf data(dRow, dCol) = "7" Then
            AddImage Loopy + 1, Loopy2 + 1, RightUpWall, "solid"
        ElseIf data(dRow, dCol) = "8" Then
            AddImage Loopy + 1, Loopy2 + 1, LeftBackCorner, "solid"
        ElseIf data(dRow, dCol) = "9" Then
            AddImage Loopy + 1, Loopy2 + 1, BackWall, "solid"
        ElseIf data(dRow, dCol) = "a" Then
            AddImage Loopy + 1, Loopy2 + 1, RightBackCorner, "solid"
        ElseIf data(dRow, dCol) = "b" Then
            AddImage Loopy + 1, Loopy2 + 1, Pistol, "pistol"
        ElseIf data(dRow, dCol) = "c" Then
            AddImage Loopy + 1, Loopy2 + 1, Shotgun, "shotgun"
        ElseIf data(dRow, dCol) = "d" Then
            AddImage Loopy + 1, Loopy2 + 1, Mines, "mines"
        ElseIf data(dRow, dCol) = "*" Then
            AddImage Loopy + 1, Loopy2 + 1, Grass, "startpoint"
        End If
    Next Loopy2
Next Loopy
End Sub
Public Sub AddImage(row As Long, col As Long, Ind As Long, Optional Property As String)
GameForm.ImageList1.ListImages.Item(Ind).Draw GameForm.sq((col - 1) + screenw * (row - 1)).hDC, 0, 0, 1
If IsMissing(Property) = False Then
    If Property = "solid" Then GameForm.sq((col - 1) + screenw * (row - 1)).Tag = "solid"
    If Property = "Fall In" Then GameForm.sq((col - 1) + screenw * (row - 1)).Tag = "Fall In"
End If
End Sub
Public Sub DetectObj(row As Long, col As Long)
If GameForm.sq((col - 1) + screenw * (row - 1)).Tag = "solid" Then Key = "q"

End Sub
Public Function IndexFind(row As Long, col As Long) As Long
IndexFind = (col - 1) + screenw * (row - 1)
End Function
