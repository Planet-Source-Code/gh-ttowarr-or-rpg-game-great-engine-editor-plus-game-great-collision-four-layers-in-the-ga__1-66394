VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10200
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   13935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   680
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   929
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar ToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImgLstToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   22
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Opentileset"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SavePic"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "x1"
            ImageIndex      =   8
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "x2"
            ImageIndex      =   9
            Style           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "x3"
            ImageIndex      =   10
            Style           =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "x4"
            ImageIndex      =   11
            Style           =   1
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   12
            Style           =   1
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   6
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   7
            Style           =   1
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   12
            Style           =   1
         EndProperty
      EndProperty
      Begin MSComctlLib.Slider SldSize 
         Height          =   315
         Left            =   4080
         TabIndex        =   4
         Top             =   30
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Min             =   2
         Max             =   20
         SelStart        =   2
         TickStyle       =   3
         TickFrequency   =   10
         Value           =   2
      End
   End
   Begin VB.PictureBox PicTileset 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   18720
      Left            =   0
      ScaleHeight     =   1248
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   1
      Top             =   360
      Width           =   3840
      Begin VB.Shape ShpBig 
         BorderColor     =   &H00FF0000&
         BorderStyle     =   3  'Dot
         Height          =   480
         Left            =   0
         Top             =   0
         Width           =   480
      End
      Begin VB.Shape ShpTile 
         BorderColor     =   &H000000FF&
         Height          =   480
         Left            =   0
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.HScrollBar ScrHMap 
      Height          =   255
      Left            =   4080
      Max             =   0
      TabIndex        =   6
      Top             =   9960
      Width           =   9600
   End
   Begin VB.VScrollBar ScrVMap 
      Height          =   9615
      Left            =   13680
      Max             =   0
      TabIndex        =   5
      Top             =   360
      Width           =   255
   End
   Begin VB.VScrollBar ScrTileset 
      Enabled         =   0   'False
      Height          =   9615
      Left            =   3840
      Max             =   0
      TabIndex        =   2
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox Pic 
      BorderStyle     =   0  'None
      Height          =   8055
      Index           =   1
      Left            =   -240
      ScaleHeight     =   8055
      ScaleWidth      =   4335
      TabIndex        =   8
      Top             =   2160
      Width           =   4335
   End
   Begin VB.PictureBox Pic 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   0
      Left            =   13680
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   7
      Top             =   9960
      Width           =   255
   End
   Begin VB.PictureBox PicMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9600
      Left            =   4080
      ScaleHeight     =   640
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   3
      Top             =   360
      Width           =   9600
      Begin VB.Shape ShpPlace 
         BorderColor     =   &H00FF0000&
         BorderStyle     =   3  'Dot
         Height          =   480
         Left            =   0
         Top             =   0
         Width           =   480
      End
   End
   Begin MSComDlg.CommonDialog Com 
      Left            =   600
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgLstToolbar 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":0000
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":0112
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":0224
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":0336
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":0448
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":055A
            Key             =   "Arc"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":066C
            Key             =   "Rectangle"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":077E
            Key             =   "x1"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":0AD0
            Key             =   "x2"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":0E22
            Key             =   "x3"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1174
            Key             =   "x4"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":14C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1818
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1BAA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu Mnu_File 
      Caption         =   "&File"
      Begin VB.Menu Mnu_LoadTileset 
         Caption         =   "&Load tileset"
      End
      Begin VB.Menu Mnu_Open 
         Caption         =   "&Load map"
         Shortcut        =   ^O
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_Save 
         Caption         =   "&Save map"
         Shortcut        =   ^S
      End
      Begin VB.Menu Mnu_SavePicture 
         Caption         =   "&Save Picture"
      End
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

'Private types
Private Type TilePos
    X As Integer
    Y As Integer
    IsAnObject As Boolean
End Type

Private Type Map
    TileCoordinates() As TilePos
End Type

'Private Variables
Private MouseDownTile As TilePos
Private MouseUpTile As TilePos
Private Layer(1 To 4) As Map
Private MapWidth As Integer
Private MapHeight As Integer
Private CurrentLayer As Integer
Private CurrentTilesetFile As String

'Start the program
Private Sub Form_Load()
Dim i As Integer, Y As Integer, X As Integer

MapWidth = InputBox("Please fill in the width of the map you want to make", "Map width")
MapHeight = InputBox("Please fill in the height of the map you want to make", "Map height")

PicMap.Width = 32 * MapWidth
PicMap.Height = 32 * MapHeight

ScrHMap.Max = MapWidth - 20
ScrVMap.Max = MapHeight - 20

If ScrHMap.Max <= 0 Then ScrHMap.Enabled = False Else ScrHMap.Enabled = True
If ScrVMap.Max <= 0 Then ScrVMap.Enabled = False Else ScrVMap.Enabled = True

For i = 1 To 4
    ReDim Layer(i).TileCoordinates(1 To MapWidth, 1 To MapHeight)
    
    For Y = 1 To MapHeight
        For X = 1 To MapWidth
            Layer(i).TileCoordinates(X, Y).X = -1
            Layer(i).TileCoordinates(X, Y).Y = -1
        Next X
    Next Y
Next i

CurrentLayer = 1
End Sub

'Save a picture of the map
Private Sub SaveMapPicture()
With Com
    .InitDir = App.Path
    .DialogTitle = "Save picture of the map"
    .DefaultExt = "*.bmp"
    .Filter = "Bitmap Files|*.BMP"
    .Filename = ""
    .ShowSave
    If .Filename = "" Then Exit Sub
    DrawLayer (1), False
    DrawLayer (2), False
    DrawLayer (3), False
    DrawLayer (4), False
    SavePicture PicMap.Image, .Filename
End With
End Sub

'Menu button load tileset
Private Sub Mnu_LoadTileset_Click()
With Com
    .InitDir = App.Path
    .Filename = ""
    .DialogTitle = "Load tileset"
    .DefaultExt = "*.bmp"
    .Filter = "Tileset Files|*.BMP"
    .ShowOpen
    If .Filename = "" Then Exit Sub
    PicTileset.Picture = LoadPicture(App.Path & "\Tilesets\" & .FileTitle)
    CurrentTilesetFile = .FileTitle
End With

ScrTileset.Max = (PicTileset.Height / 32) - 20
If ScrTileset.Max <= 0 Then ScrTileset.Enabled = False Else ScrTileset.Enabled = True
End Sub

'Load file
Private Sub Mnu_Open_Click()
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
End Sub

'save level
Private Sub Mnu_Save_Click()
With Com
    .InitDir = App.Path
    .DialogTitle = "Save map"
    .DefaultExt = "*.map"
    .Filename = ""
    .Filter = "Map Files|*.MAP"
    .ShowSave
    If .Filename = "" Then Exit Sub
    SaveLevel (.Filename)
End With
End Sub

'Menu button save map picture
Private Sub Mnu_SavePicture_Click()
SaveMapPicture
End Sub

'Map click
Private Sub PicMap_Click()
On Error Resume Next

Dim X As Integer, Y As Integer
Dim SelectedTile As TilePos
Dim CurrentTile As TilePos

If ToolBar.Buttons(13).Value = tbrPressed Then
    If ShpBig.Width = 32 And ShpBig.Height = 32 Then
            CurrentTile.X = ShpBig.Left / 32
            CurrentTile.Y = ShpBig.Top / 32

            For Y = 0 To SldSize.Value - 1
                For X = 0 To SldSize.Value - 1
                    SelectedTile.X = ShpPlace.Left / 32 + X + 1
                    SelectedTile.Y = ShpPlace.Top / 32 + Y + 1
                    PaintOneTile CurrentTile, SelectedTile
                Next X
            Next Y
    End If
Else
    If ShpBig.Width = 32 And ShpBig.Height = 32 Then
        SelectedTile.X = ShpPlace.Left / 32 + 1
        SelectedTile.Y = ShpPlace.Top / 32 + 1
        CurrentTile.X = ShpBig.Left / 32
        CurrentTile.Y = ShpBig.Top / 32
        PaintOneTile CurrentTile, SelectedTile
    Else
        For Y = 0 To ShpBig.Height / 32 - 1
            For X = 0 To ShpBig.Width / 32 - 1
                CurrentTile.X = ShpBig.Left / 32 + X
                CurrentTile.Y = ShpBig.Top / 32 + Y

                SelectedTile.X = ShpPlace.Left / 32 + X + 1
                SelectedTile.Y = ShpPlace.Top / 32 + Y + 1
                PaintOneTile CurrentTile, SelectedTile
            Next X
        Next Y
    End If
End If
PicMap.Refresh
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
        PaintOneTile CurrentPaintTile, CurrentSelectedTile, False
    Next X
Next Y

PicMap.Refresh
End Sub

'Paint one tile at the moment
Private Sub PaintOneTile(PaintTile As TilePos, Tile As TilePos, Optional Save As Boolean = True)
If Save = True Then
    Layer(CurrentLayer).TileCoordinates(Tile.X, Tile.Y).X = PaintTile.X
    Layer(CurrentLayer).TileCoordinates(Tile.X, Tile.Y).Y = PaintTile.Y
    If ToolBar.Buttons(22).Value = tbrPressed Then
        Layer(CurrentLayer).TileCoordinates(Tile.X, Tile.Y).IsAnObject = True
    Else
        Layer(CurrentLayer).TileCoordinates(Tile.X, Tile.Y).IsAnObject = False
    End If
End If

TransparentBlt PicMap.HDC, Tile.X * 32 - 32, Tile.Y * 32 - 32, 32, 32, PicTileset.HDC, PaintTile.X * 32, PaintTile.Y * 32, 32, 32, RGB(84, 138, 150)
End Sub

'Map mouse movement
Private Sub PicMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then PicMap_Click

ShpPlace.Left = Split((X / 32), ",")(0) * 32
ShpPlace.Top = Split((Y / 32), ",")(0) * 32
End Sub

'Tileset mouse down
Private Sub PicTileset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseDownTile.X = Split((X / 32), ",")(0)
MouseDownTile.Y = Split((Y / 32), ",")(0)
End Sub

'Tileset mouse movement
Private Sub PicTileset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShpTile.Left = Split((X / 32), ",")(0) * 32
ShpTile.Top = Split((Y / 32), ",")(0) * 32
End Sub

'Tileset mouse up
Private Sub PicTileset_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseUpTile.X = Split((X / 32), ",")(0)
MouseUpTile.Y = Split((Y / 32), ",")(0)

    ToolBar.Buttons(12).Value = tbrPressed
    ToolBar.Buttons(13).Value = tbrUnpressed
    SldSize.Value = 1
    SldSize.Enabled = False

If MouseUpTile.X >= MouseDownTile.X Then
ShpBig.Width = (MouseUpTile.X - MouseDownTile.X + 1) * 32
ShpBig.Left = MouseDownTile.X * 32
End If
If MouseUpTile.X <= MouseDownTile.X Then
ShpBig.Width = (MouseDownTile.X - MouseUpTile.X + 1) * 32
ShpBig.Left = MouseUpTile.X * 32
End If
If MouseUpTile.Y >= MouseDownTile.Y Then
ShpBig.Height = (MouseUpTile.Y - MouseDownTile.Y + 1) * 32
ShpBig.Top = MouseDownTile.Y * 32
End If
If MouseUpTile.Y <= MouseDownTile.Y Then
ShpBig.Height = (MouseDownTile.Y - MouseUpTile.Y + 1) * 32
ShpBig.Top = MouseUpTile.Y * 32
End If

ShpPlace.Width = ShpBig.Width
ShpPlace.Height = ShpBig.Height
End Sub

'Map scroll
Private Sub ScrHMap_Change()
PicMap.Left = 272 - ScrHMap.Value * 32
End Sub

'Tileset scroll
Private Sub ScrTileset_Change()
PicTileset.Top = 24 - ScrTileset.Value * 32
End Sub

'Tileset scroll
Private Sub ScrTileset_Scroll()
PicTileset.Top = 24 - ScrTileset.Value * 32
End Sub

'Map scroll
Private Sub ScrVMap_Change()
PicMap.Top = 24 - ScrVMap.Value * 32
End Sub

'Size slider
Private Sub SldSize_Change()
SldSize_Scroll
End Sub

'Size slider
Private Sub SldSize_Scroll()
ShpPlace.Width = 32 * SldSize.Value
ShpPlace.Height = 32 * SldSize.Value
End Sub

'Toolbar buttons
Private Sub ToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer

If Button.Index = 1 Then Mnu_LoadTileset_Click
If Button.Index = 2 Then Mnu_Open_Click
If Button.Index = 3 Then Mnu_Save_Click
If Button.Index = 4 Then Mnu_SavePicture_Click
If Button.Index >= 6 And Button.Index <= 10 Then
    ToolBar.Buttons(6).Value = tbrUnpressed
    ToolBar.Buttons(7).Value = tbrUnpressed
    ToolBar.Buttons(8).Value = tbrUnpressed
    ToolBar.Buttons(9).Value = tbrUnpressed
    ToolBar.Buttons(10).Value = tbrUnpressed
    ToolBar.Buttons(Button.Index).Value = tbrPressed
    CurrentLayer = Button.Index - 5
    If CurrentLayer = 5 Then
        DrawLayer (1), False
        DrawLayer (2), False
        DrawLayer (3), False
        DrawLayer (4), False
    Else
        For i = 1 To CurrentLayer
            DrawLayer i, False
            If i = CurrentLayer - 1 Then GrayScale PicMap
        Next i
    End If
End If
If Button.Index = 12 Then
    ToolBar.Buttons(12).Value = tbrPressed
    ToolBar.Buttons(13).Value = tbrUnpressed
    SldSize.Value = 1
    SldSize.Enabled = False
End If
If Button.Index = 13 Then
    ToolBar.Buttons(12).Value = tbrUnpressed
    ToolBar.Buttons(13).Value = tbrPressed
    SldSize.Enabled = True
    SldSize_Scroll
End If
End Sub

'Save the level to text
Private Sub SaveLevel(Filename As String)
Dim File As Integer
Dim X As Integer
Dim Y As Integer
Dim LayerIndex As Integer
Dim IsAnObjectValue As Integer

File = FreeFile

Open Filename For Output As File
    Print #File, MapWidth & "^" & MapHeight & "^" & CurrentTilesetFile
    For LayerIndex = 1 To 4
        Print #File, "[Layer]"
        For Y = 1 To MapHeight
            For X = 1 To MapWidth
                If Layer(LayerIndex).TileCoordinates(X, Y).IsAnObject = True Then IsAnObjectValue = 1 Else IsAnObjectValue = 0
                Print #File, Layer(LayerIndex).TileCoordinates(X, Y).X & "*" & Layer(LayerIndex).TileCoordinates(X, Y).Y & "*" & IsAnObjectValue & "]"
            Next X
        Next Y
    Next LayerIndex
Close File
End Sub

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

CurrentLayer = 1

ScrHMap.Max = MapWidth - 20
ScrVMap.Max = MapHeight - 20

If ScrHMap.Max <= 0 Then ScrHMap.Enabled = False Else ScrHMap.Enabled = True
If ScrVMap.Max <= 0 Then ScrVMap.Enabled = False Else ScrVMap.Enabled = True

ScrTileset.Max = (PicTileset.Height / 32) - 20
If ScrTileset.Max <= 0 Then ScrTileset.Enabled = False Else ScrTileset.Enabled = True

DrawLayer (1), False
DrawLayer (2), False
DrawLayer (3), False
DrawLayer (4), False
End Sub

