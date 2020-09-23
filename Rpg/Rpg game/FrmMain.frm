VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7200
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicChar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3360
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      Top             =   3360
      Width           =   480
   End
   Begin VB.PictureBox PicMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7200
      Left            =   0
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   0
      Width           =   7200
      Begin VB.PictureBox PicCharset 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1950
         Left            =   1200
         Picture         =   "FrmMain.frx":0000
         ScaleHeight     =   130
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   71
         TabIndex        =   3
         Top             =   2040
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Timer TmrKeys 
         Enabled         =   0   'False
         Interval        =   120
         Left            =   120
         Top             =   120
      End
   End
   Begin VB.PictureBox PicTileset 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   18720
      Left            =   10320
      ScaleHeight     =   1248
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   7680
      Visible         =   0   'False
      Width           =   3840
   End
   Begin MSComDlg.CommonDialog Com 
      Left            =   6360
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu Mnu_LoadLevel 
      Caption         =   "Load level"
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Api's
Private Declare Function TransparentBlt Lib "msimg32" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'Private types
Private Type TilePos
    X As Integer
    Y As Integer
    IsAnObject As Boolean
End Type

Private Type Map
    TileCoordinates() As TilePos
End Type

Private Enum Side
    Left
    Right
    Up
    Down
End Enum

'Private Variables
Private Layer(1 To 4) As Map
Private MapWidth As Integer
Private MapHeight As Integer
Private CurrentTilesetFile As String
Private Counter As Integer
Private CurrentFrame As Integer
Private CurrentTile As TilePos

'Load a level from text
Private Sub LoadLevel(Filename As String)
Dim File As Integer
Dim sFile As String
Dim sTemp(0 To 4) As String
Dim sBuf As String
Dim X As Integer
Dim Y As Integer
Dim Counter As Long
Dim LayerIndex As Integer
Dim i As Integer

DoEvents

File = FreeFile

Open Filename For Input As File
sFile = Input(LOF(File), 1)
sFile = Replace(sFile, vbCrLf, "")
sTemp(0) = Mid(Split(sFile, "[Layer]")(0), 1)
sTemp(1) = Mid(Split(sFile, "[Layer]")(1), 1)
sTemp(2) = Mid(Split(sFile, "[Layer]")(2), 1)
sTemp(3) = Mid(Split(sFile, "[Layer]")(3), 1)
sTemp(4) = Mid(Split(sFile, "[Layer]")(4), 1)

Close File


MapWidth = Split(sTemp(0), "^")(0)
MapHeight = Split(sTemp(0), "^")(1)
CurrentTilesetFile = Split(sTemp(0), "^")(2)

PicTileset.Picture = LoadPicture(App.Path & "\Tilesets\" & CurrentTilesetFile)

PicMap.Width = 32 * MapWidth
PicMap.Height = 32 * MapHeight

For i = 1 To 4
    ReDim Layer(i).TileCoordinates(1 To MapWidth, 1 To MapHeight)
    
    For Y = 1 To MapHeight
        For X = 1 To MapWidth
            Layer(i).TileCoordinates(X, Y).X = -1
            Layer(i).TileCoordinates(X, Y).Y = -1
        Next X
    Next Y
Next i

For LayerIndex = 1 To 4
    Counter = 0
    For Y = 1 To MapHeight
        For X = 1 To MapWidth
            sBuf = Split(sTemp(LayerIndex), "]")(Counter)
            Layer(LayerIndex).TileCoordinates(X, Y).X = Split(sBuf, "*")(0)
            Layer(LayerIndex).TileCoordinates(X, Y).Y = Split(sBuf, "*")(1)
            Layer(LayerIndex).TileCoordinates(X, Y).IsAnObject = Split(sBuf, "*")(2)
            Counter = Counter + 1
        Next X
    Next Y
Next LayerIndex

DrawLayer (1), False
DrawLayer (2), False
DrawLayer (3), False
DrawLayer (4), False
End Sub

'Draw layer
Private Sub DrawLayer(LayerIndex As Integer, Optional CleanUp As Boolean = True)
Dim X As Integer, Y As Integer
Dim CurrentSelectedTile As TilePos
Dim CurrentPaintTile As TilePos

If CleanUp = True Then PicMap.Cls

For Y = 1 To MapHeight
    For X = 1 To MapWidth
        CurrentPaintTile.X = Layer(LayerIndex).TileCoordinates(X, Y).X
        CurrentPaintTile.Y = Layer(LayerIndex).TileCoordinates(X, Y).Y
        CurrentSelectedTile.X = X
        CurrentSelectedTile.Y = Y
        PaintOneTile CurrentPaintTile, CurrentSelectedTile
    Next X
Next Y

PicMap.Refresh
End Sub

'Paint one tile at the moment
Private Sub PaintOneTile(PaintTile As TilePos, Tile As TilePos)
DoEvents
TransparentBlt PicMap.hDC, Tile.X * 32 - 32, Tile.Y * 32 - 32, 32, 32, PicTileset.hDC, PaintTile.X * 32, PaintTile.Y * 32, 32, 32, RGB(84, 138, 150)
End Sub

'Load level
Private Sub Mnu_LoadLevel_Click()
With Com
    .InitDir = App.Path
    .Filename = ""
    .DialogTitle = "Load map"
    .DefaultExt = "*.map"
    .Filter = "Map Files|*.MAP"
    .ShowOpen
    If .Filename = "" Then Exit Sub
    LoadLevel (.Filename)
End With

TmrKeys.Enabled = True
End Sub


Private Sub DrawCharFrame(Side As Side)
Dim CurrentPaintTile As TilePos
Dim i As Integer

Select Case Side
    Case Left
        For i = 1 To 4
            CurrentPaintTile.X = Layer(i).TileCoordinates(CurrentTile.X - 1, CurrentTile.Y).X
            CurrentPaintTile.Y = Layer(i).TileCoordinates(CurrentTile.X - 1, CurrentTile.Y).Y
            TransparentBlt PicChar.hDC, 0, 0, 32, 32, PicTileset.hDC, CurrentPaintTile.X * 32, CurrentPaintTile.Y * 32, 32, 32, RGB(84, 138, 150)
        Next i
    Case Right
        For i = 1 To 4
            CurrentPaintTile.X = Layer(i).TileCoordinates(CurrentTile.X + 1, CurrentTile.Y).X
            CurrentPaintTile.Y = Layer(i).TileCoordinates(CurrentTile.X + 1, CurrentTile.Y).Y
            TransparentBlt PicChar.hDC, 0, 0, 32, 32, PicTileset.hDC, CurrentPaintTile.X * 32, CurrentPaintTile.Y * 32, 32, 32, RGB(84, 138, 150)
        Next i
    Case Up
        For i = 1 To 4
            CurrentPaintTile.X = Layer(i).TileCoordinates(CurrentTile.X, CurrentTile.Y - 1).X
            CurrentPaintTile.Y = Layer(i).TileCoordinates(CurrentTile.X, CurrentTile.Y - 1).Y
            TransparentBlt PicChar.hDC, 0, 0, 32, 32, PicTileset.hDC, CurrentPaintTile.X * 32, CurrentPaintTile.Y * 32, 32, 32, RGB(84, 138, 150)
        Next i
    Case Down
        For i = 1 To 4
            CurrentPaintTile.X = Layer(i).TileCoordinates(CurrentTile.X, CurrentTile.Y + 1).X
            CurrentPaintTile.Y = Layer(i).TileCoordinates(CurrentTile.X, CurrentTile.Y + 1).Y
            TransparentBlt PicChar.hDC, 0, 0, 32, 32, PicTileset.hDC, CurrentPaintTile.X * 32, CurrentPaintTile.Y * 32, 32, 32, RGB(84, 138, 150)
        Next i
End Select
End Sub

Private Sub TmrKeys_Timer()
Dim Buf As String

Buf = "0,1,0,2"

CurrentTile.X = Split((PicChar.Left - PicMap.Left) / 32, ",")(0) + 1
CurrentTile.Y = Split((PicChar.Top - PicMap.Top) / 32, ",")(0) + 1

If Not GetAsyncKeyState(vbKeyUp) = 0 Then  'up
    If Collision(CurrentTile.X, CurrentTile.Y - 1) = True Then Exit Sub
    
    DrawCharFrame Up
    
    TransparentBlt PicChar.hDC, 0, 0, PicCharset.Width / 3, PicCharset.Height / 4, PicCharset.hDC, PicCharset.Width / 3 * CurrentFrame, 0, PicCharset.Width / 3, PicCharset.Height / 4, RGB(255, 255, 255)
    If Not PicChar.Top = 224 Then
        PicChar.Top = PicChar.Top - 32
        Exit Sub
    End If
    If PicMap.Top < 0 Then
        PicMap.Top = PicMap.Top + 32
    Else
        PicChar.Top = PicChar.Top - 32
    End If
    Counter = Counter + 1
    If Counter = 4 Then Counter = 0
    CurrentFrame = Split(Buf, ",")(Counter)
End If

If Not GetAsyncKeyState(vbKeyDown) = 0 Then  'down
    If Collision(CurrentTile.X, CurrentTile.Y + 1) = True Then Exit Sub
    
    DrawCharFrame Down
    
    TransparentBlt PicChar.hDC, 0, 0, PicCharset.Width / 3, PicCharset.Height / 4, PicCharset.hDC, PicCharset.Width / 3 * CurrentFrame, PicCharset.Height / 4 * 2, PicCharset.Width / 3, PicCharset.Height / 4, RGB(255, 255, 255)
    If Not PicChar.Top = 224 Then
        PicChar.Top = PicChar.Top + 32
        Exit Sub
    End If
    If Not PicMap.Top <= 0 - (PicMap.Height - 480) Then
        PicMap.Top = PicMap.Top - 32
    Else
        PicChar.Top = PicChar.Top + 32
    End If
    Counter = Counter + 1
    If Counter = 4 Then Counter = 0
    CurrentFrame = Split(Buf, ",")(Counter)
End If

If Not GetAsyncKeyState(vbKeyLeft) = 0 Then  'left
    If Collision(CurrentTile.X - 1, CurrentTile.Y) = True Then Exit Sub
    
    DrawCharFrame Left
    
    TransparentBlt PicChar.hDC, 0, 0, PicCharset.Width / 3, PicCharset.Height / 4, PicCharset.hDC, PicCharset.Width / 3 * CurrentFrame, PicCharset.Height / 4 * 3, PicCharset.Width / 3, PicCharset.Height / 4, RGB(255, 255, 255)
    If Not PicChar.Left = 224 Then
        PicChar.Left = PicChar.Left - 32
        Exit Sub
    End If
    If PicMap.Left < 0 Then
        PicMap.Left = PicMap.Left + 32
    Else
        PicChar.Left = PicChar.Left - 32
    End If
    Counter = Counter + 1
    If Counter = 4 Then Counter = 0
    CurrentFrame = Split(Buf, ",")(Counter)
End If

If Not GetAsyncKeyState(vbKeyRight) = 0 Then  'right
    If Collision(CurrentTile.X + 1, CurrentTile.Y) = True Then Exit Sub
    
    DrawCharFrame Right
    
    TransparentBlt PicChar.hDC, 0, 0, PicCharset.Width / 3, PicCharset.Height / 4, PicCharset.hDC, PicCharset.Width / 3 * CurrentFrame, PicCharset.Height / 4 * 1, PicCharset.Width / 3, PicCharset.Height / 4, RGB(255, 255, 255)
    If Not PicChar.Left = 224 Then
        PicChar.Left = PicChar.Left + 32
        Exit Sub
    End If
    If Not PicMap.Left <= 0 - (PicMap.Width - 480) Then
        PicMap.Left = PicMap.Left - 32
    Else
        PicChar.Left = PicChar.Left + 32
    End If
    Counter = Counter + 1
    If Counter = 4 Then Counter = 0
    CurrentFrame = Split(Buf, ",")(Counter)
End If
End Sub

Private Function Collision(X As Integer, Y As Integer) As Boolean
On Error GoTo FndErr

Dim i As Integer
Dim Buffer As Integer

For i = 1 To 4
    If Layer(i).TileCoordinates(X, Y).IsAnObject = True Then
        Buffer = 1
    End If
Next i

If Buffer = 1 Then Collision = True
Exit Function

FndErr:
Collision = True
End Function
