Attribute VB_Name = "Module1"
Option Explicit

Public GenC As Double
Public GenC2 As Double
Public PlotX As Double
Public PlotY As Double
Public MyYscale As Double
Public ModYscale As Double
Public MyXscale As Double
Public ModXscale As Double
Public Offset As Double
Public ColInfo(256, 256) As Double
Public Currpath As String
Public X3d As Double
Public m_Angle As Long



Public Const PI As Double = 3.14159265358979
Public Const PIDEG As Double = PI / 180

Public Const DIB_RGB_COLORS = 0

Public Type BITMAPINFOHEADER
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

Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Public Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type

Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Public Sub RotateDC(DstDC As Long, DstX As Long, DstY As Long, SrcDC As Long, SrcBmp As Long, Deg As Long)

    Dim TmpDC As Long
    Dim TmpBmp As Long
    Dim OldObject As Long
    Dim BitCount As Long
    Dim LineWidth As Long
    Dim retVal As Long
    Dim Width As Long
    Dim Height As Long
    Dim h&, b&, f&, d&, i&
    Dim dx1 As Double
    dx1 = 1#
    Dim dy1 As Double
    Dim SrcBits() As Byte
    Dim TmpBits() As Byte
    
    Dim tempBIH As BITMAPINFOHEADER
    
    Dim TempAlpha As Byte
    Dim biSize As Long
    biSize = LenB(tempBIH)
    Dim Info As BITMAPINFO
    Dim Info2 As BITMAPINFO
    Info.bmiHeader.biSize = biSize
    Info2.bmiHeader.biSize = biSize
    retVal = GetDIBits(SrcDC, SrcBmp, 0, 0, 0&, Info, DIB_RGB_COLORS)
    If retVal = 0 And Info.bmiHeader.biWidth = 0 Then Exit Sub
    TmpDC = CreateCompatibleDC(SrcDC)
    Width = Info.bmiHeader.biWidth
    Height = Info.bmiHeader.biHeight
    
    Dim NewSize As Long
    NewSize = Sqr(Width * Width + Height * Height) + 2
    TmpBmp = CreateCompatibleBitmap(SrcDC, NewSize, NewSize)
    If (TmpBmp <> 0) Then
        OldObject = SelectObject(TmpDC, TmpBmp)
        BitBlt TmpDC, 0, 0, NewSize, NewSize, DstDC, DstX - NewSize / 2, DstY - NewSize / 2, vbSrcCopy
        Info.bmiHeader.biBitCount = 24
        Info.bmiHeader.biCompression = 0
        Info2.bmiHeader.biBitCount = 24
        Info2.bmiHeader.biCompression = 0
        Info2.bmiHeader.biPlanes = 1
        Info2.bmiHeader.biHeight = NewSize
        Info2.bmiHeader.biWidth = NewSize
        Dim LineWidth2
        LineWidth2 = NewSize * 3
        If (LineWidth2 Mod 4 <> 0) Then LineWidth2 = LineWidth2 + (4 - LineWidth2 Mod 4)
        Dim BitCount2 As Long
        BitCount2 = LineWidth2 * NewSize
        LineWidth = Width * 3
        If (LineWidth Mod 4 <> 0) Then LineWidth = LineWidth + (4 - LineWidth Mod 4)
        BitCount = LineWidth * Height
        ReDim SrcBits(0 To BitCount - 1) As Byte
        ReDim TmpBits(0 To BitCount2 - 1) As Byte
        GetDIBits SrcDC, SrcBmp, 0, Height, SrcBits(0), Info, DIB_RGB_COLORS
        GetDIBits TmpDC, TmpBmp, 0, NewSize, TmpBits(0), Info2, DIB_RGB_COLORS
        Dim CurOffset As Long
        Dim NewX As Double, NewY As Double
        Dim Xmm As Long, Ymm As Long
        Dim i1 As Long
        Dim v1 As Boolean
        dx1 = Cos(Deg * PIDEG)
        dy1 = Sin(Deg * PIDEG)

        For h = 0 To NewSize - 1
            CurOffset = LineWidth2 * h
            For b = 0 To NewSize - 1
                f = CurOffset + 3 * b
                NewX = Width / 2 + (b - NewSize / 2) * dx1 - (h - NewSize / 2) * dy1
                NewY = Height / 2 + (b - NewSize / 2) * dy1 + (h - NewSize / 2) * dx1
                Xmm = (NewX + 0.5)
                Ymm = (NewY + 0.5)
                If ((Xmm >= 0) And (Xmm < Width) And (Ymm >= 0) And (Ymm < Height)) Then
                    v1 = True
                    i1 = LineWidth * Ymm + 3 * Xmm
                    If v1 Then
                        For d = 0 To 2
                            TmpBits(f + d) = SrcBits(i1 + d)
                        Next d
                    End If
                End If
            Next b
        Next h

        SetDIBitsToDevice DstDC, DstX - NewSize / 2, DstY - NewSize / 2, NewSize, NewSize, 0, 0, 0, NewSize, TmpBits(0), Info2, DIB_RGB_COLORS
        DeleteObject SelectObject(TmpDC, OldObject)
    End If
    DeleteDC TmpDC
End Sub


