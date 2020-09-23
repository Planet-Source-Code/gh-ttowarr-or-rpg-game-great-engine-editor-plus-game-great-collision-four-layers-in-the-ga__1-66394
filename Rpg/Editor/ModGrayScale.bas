Attribute VB_Name = "ModGrayScale"
Option Explicit

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal HDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0&

Private iDATA() As Byte
Private bDATA() As Byte
Private PicInfo As BITMAP
Private DIBInfo As BITMAPINFO
Private Speed(0 To 765) As Long

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Type RGBQUAD
   rgbBlue As Byte
   rgbGreen As Byte
   rgbRed As Byte
   rgbReserved As Byte
End Type

Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmiColors As RGBQUAD
End Type

Public Sub GrayScale(ByVal Pic As PictureBox)
Dim hdcNew As Long
Dim ret As Long
Dim BytesPerScanLine As Long
Dim PadBytesPerScanLine As Long
Dim X As Long, Y As Long
Dim R As Long, G As Long, B As Long
 
Call GetObject(Pic.Image, Len(PicInfo), PicInfo)
  hdcNew = CreateCompatibleDC(0&)
With DIBInfo.bmiHeader
    .biSize = 40
    .biWidth = PicInfo.bmWidth
    .biHeight = -PicInfo.bmHeight
    .biPlanes = 1
    .biBitCount = 32
    .biCompression = BI_RGB
    BytesPerScanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * 4)
    PadBytesPerScanLine = _
       BytesPerScanLine - (((.biWidth * .biBitCount) + 7) \ 8)
    .biSizeImage = BytesPerScanLine * Abs(.biHeight)
End With
  
ReDim iDATA(1 To 4, 1 To PicInfo.bmWidth, 1 To PicInfo.bmHeight) As Byte
ReDim bDATA(1 To 4, 1 To PicInfo.bmWidth, 1 To PicInfo.bmHeight) As Byte
  
ret = GetDIBits(hdcNew, Pic.Image, 0, PicInfo.bmHeight, iDATA(1, 1, 1), DIBInfo, DIB_RGB_COLORS)

For X = 0 To 765
    Speed(X) = X \ 3
Next X
For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
        B = iDATA(1, X, Y)
        G = iDATA(2, X, Y)
        R = iDATA(3, X, Y)
        B = Speed(R + G + B)
        iDATA(1, X, Y) = B
        iDATA(2, X, Y) = B
        iDATA(3, X, Y) = B
    Next X
    DoEvents
Next Y
DoEvents

ret = SetDIBits(hdcNew, Pic.Image, 0, PicInfo.bmHeight, iDATA(1, 1, 1), DIBInfo, DIB_RGB_COLORS)
End Sub




