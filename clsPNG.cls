VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPNG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event ChunkFound(ChkName As String, ChkLength As Long, ChkData() As Byte, ChkCRC As Long)
Public Event LoadProgress(Max As Long, Value As Long)

Public Enum PNGErrorContants
    pngeSucceeded
    pngeFileNotFound
    pngeOpenError
    pngeInvalidFile
End Enum
Public Enum PNGColorTypes
    pngeGrayscale = 0
    pngeRGB = 2
    pngePalette = 3
    pngeGrayscaleAlpha = 4
    pngeRGBAlpha = 6
End Enum

Private MyChunks As New Collection
Private MyHdr As IHDR
Private MyBmpImg As BITMAPIMAGE

Private strCurrentFilename As String
Private boolGammaEnabled As Boolean
Private intGammaRed As Integer
Private intGammaGreen As Integer
Private intGammaBlue As Integer
Private lngGray As Long
Private lngBoost As Long

Public Type PNGCHUNK
    ChkLength As Long
    ChkType As String
    ChkData() As Byte
    ChkCRC As Long
End Type
Private Type IHDR
    Width As Long
    Height As Long
    BitDepth As Byte
    ColorType As Byte
    Compression As Byte
    Filter As Byte
    Interlacing As Byte
End Type
Private Type PNGRGB
    Red As Byte
    Green As Byte
    Blue As Byte
End Type
Private Type RGBQUAD
    Blue As Byte
    Green As Byte
    Red As Byte
    Reserved As Byte
End Type
Private Type BITMAPINFOHEADER
    Size As Long
    Width As Long
    Height As Long
    Planes As Integer
    BitCount As Integer
    Compression As Long
    SizeImage As Long
    XPelsPerMeter As Long
    YPelsPerMeter As Long
    ClrUsed As Long
    ClrImportant As Long
End Type
Private Type BITMAPFILEHEADER
    Type As Integer
    Size As Long
    Reserved1 As Integer
    Reserved2 As Integer
    OffBits As Long
End Type
Private Type BITMAPIMAGE
    BmpHdr As BITMAPINFOHEADER
    ClrTbl() As RGBQUAD
    BmpData() As Byte
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function zlibUncomp Lib "zlib.dll" Alias "uncompress" (Dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Const Z_OK = 0
Private Const Z_BUF_ERROR = -5
Private Const Z_DATA_ERROR = -3
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As Any, ByVal wUsage As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function filterColor Lib "wa3gamma.dll" (ByVal Color As Long, ByVal R As Long, ByVal G As Long, ByVal B As Long, ByVal Gray As Long, ByVal Boost As Long) As Long

Private Sub Class_Initialize()
    boolGammaEnabled = False
    intGammaRed = 128
    intGammaGreen = 128
    intGammaBlue = 128
    lngGray = 0
    lngBoost = 0
End Sub

Private Sub Align32(Hdr As IHDR, BitCount As Integer, Dat() As Byte)
    Dim RowBytes As Long, SrcRowBytes As Long
    Dim Y As Long, Dest() As Byte
    Dim SrcOff As Long, DestOff As Long
    
    If BitCount = 32 Then Exit Sub
    
    RowBytes = LineBytes(Hdr.Width, BitCount)
    SrcRowBytes = DataPerRow(Hdr) - 1
    
    If RowBytes = SrcRowBytes Then
        Exit Sub
    Else
        ReDim Dest(0 To RowBytes * Hdr.Height - 1)
        For Y = 0 To Hdr.Height - 1
            SrcOff = Y * SrcRowBytes
            DestOff = Y * RowBytes
            
            CopyMemory Dest(DestOff), Dat(SrcOff), SrcRowBytes
        Next Y
        Dat = Dest
    End If
End Sub

Public Property Let GammaEnabled(ByVal boolValue As Boolean)
    boolGammaEnabled = boolValue
End Property

Public Property Let GammaRed(ByVal intValue As Integer)
    intGammaRed = intValue
End Property

Public Property Let GammaGreen(ByVal intValue As Integer)
    intGammaGreen = intValue
End Property

Public Property Let GammaBlue(ByVal intValue As Integer)
    intGammaBlue = intValue
End Property

Public Property Let Gray(ByVal lngValue As Long)
    lngGray = lngValue
End Property

Public Property Let Boost(ByVal lngValue As Long)
    lngBoost = lngValue
End Property

Public Property Get BitDepth() As Long
    BitDepth = MyHdr.BitDepth
End Property

Public Property Get Interlaced() As Boolean
    Interlaced = (MyHdr.Interlacing <> 0)
End Property

Public Property Get CurrentFilename() As String
    CurrentFilename = strCurrentFilename
End Property
Private Function BitsPerPixel(Hdr As IHDR) As Long
    Select Case Hdr.ColorType
        Case 0, 3: BitsPerPixel = Hdr.BitDepth
        Case 2: BitsPerPixel = 3 * Hdr.BitDepth
        Case 6: BitsPerPixel = 4 * Hdr.BitDepth
        Case 4: BitsPerPixel = 2 * Hdr.BitDepth
    End Select
End Function

Private Function BytesToStr(ByteVals() As Byte) As String
    Dim n As Long
    
    For n = LBound(ByteVals) To UBound(ByteVals)
        BytesToStr = BytesToStr & Chr(ByteVals(n))
    Next n
End Function
Public Property Get ColorType() As PNGColorTypes
    ColorType = MyHdr.ColorType
End Property

Private Function DataPerRow(Hdr As IHDR) As Long
    DataPerRow = (MyHdr.Width * BitsPerPixel(Hdr) + 7) \ 8 + 1
End Function

Private Function GetBmpImgInterlaced() As BITMAPIMAGE
End Function

Private Function GetGrayTbl(BitDepth As Integer) As RGBQUAD()
    Dim CurLevel As Integer, n As Long, LevelDiff As Byte
    Dim Tbl() As RGBQUAD
    
    If BitDepth <> 16 Then
        ReDim Tbl(0 To 2 ^ BitDepth - 1)
    Else
        ReDim Tbl(0 To 255)
    End If
    
    LevelDiff = 255 / UBound(Tbl)
    
    For n = 0 To UBound(Tbl)
        With Tbl(n)
            .Red = CurLevel
            .Green = CurLevel
            .Blue = CurLevel
        End With
        CurLevel = CurLevel + LevelDiff
    Next n
    
    GetGrayTbl = Tbl
End Function
Private Sub Conv16To8(Dat() As Byte)
    Dim n As Long, DestDat() As Byte, DestOff As Long

    ReDim DestDat(0 To (UBound(Dat) + 1) \ 2 - 1)
    
    For n = 0 To UBound(Dat) Step 2
        DestDat(DestOff) = Dat(n)
        DestOff = DestOff + 1
    Next n
    
    Dat = DestDat
End Sub

Private Sub GrayAToRGBA(Dat() As Byte)
    Dim n As Long, DestDat() As Byte, DestOff As Long
    
    ReDim DestDat(0 To (UBound(Dat) + 1) * 2 - 1)
    
    For n = 0 To UBound(Dat) Step 2
        DestDat(DestOff) = Dat(n)
        DestDat(DestOff + 1) = Dat(n)
        DestDat(DestOff + 2) = Dat(n)
        DestDat(DestOff + 3) = Dat(n + 1)
        DestOff = DestOff + 4
    Next n
    
    Dat = DestDat
End Sub

Public Property Get Height() As Long
    Height = MyHdr.Height
End Property

Private Sub Pal2To8(Width As Long, Height As Long, Dat() As Byte, RowBytes As Long)
    Dim DestDat() As Byte, DestRowBytes As Long, n As Long
    Dim Px As Byte, DestOff As Long, X As Long, Y As Long

    DestRowBytes = LineBytes(Width, 8)
    ReDim DestDat(0 To DestRowBytes * Height - 1)

    For Y = 0 To Height - 1
        DestOff = Y * DestRowBytes
        For X = 0 To Width - 1
            n = Y * (RowBytes - 1) + X \ 4
            
            If (X Mod 4) <> 3 Then
                Px = (Dat(n) \ 4 ^ (3 - (X Mod 4))) And 3
            Else
                Px = Dat(n) And 3
            End If
            
            DestDat(DestOff) = Px
            
            DestOff = DestOff + 1
        Next X
    Next Y
    
    Dat = DestDat
End Sub
Private Function Decompress(CompStream() As Byte, UncompSize As Long) As Byte()
    Dim Buffer() As Byte
    
    ReDim Buffer(0 To UncompSize - 1)
    
    If zlibUncomp(Buffer(0), UncompSize, CompStream(0), _
        UBound(CompStream) + 1) <> Z_OK Then Exit Function

    If UncompSize > 0 Then Decompress = Buffer
End Function
Private Function DeFilter(Hdr As IHDR, Dat() As Byte) As Byte()
    Dim NewDat() As Byte, Y As Long, Ival As Long
    Dim n As Long, StartByte As Long, DestByte As Long
    Dim BPRow As Long, X As Long, RowBytes() As Byte
    Dim PrevRowBytes() As Byte
    
    If Not HasLoaded Then Exit Function
    
    Ival = Interval(Hdr)
    BPRow = DataPerRow(Hdr)
    ReDim NewDat(0 To UBound(Dat) - MyHdr.Height)
    ReDim PrevRowBytes(0 To DataPerRow(Hdr) - 2)
    ReDim RowBytes(0 To DataPerRow(Hdr) - 2)
    
    For Y = 0 To MyHdr.Height - 1
        StartByte = BPRow * Y
        DestByte = StartByte - Y
        X = 0
        CopyMemory RowBytes(0), Dat(StartByte + 1), BPRow - 1
        Select Case Dat(StartByte)
            Case 0 'None
            Case 1 'Sub
                ReverseSub RowBytes, Ival
            Case 2 'Up
                ReverseUp RowBytes, PrevRowBytes
            Case 3 'Average
                ReverseAverage RowBytes, PrevRowBytes, Ival
            Case 4 'Paeth
                ReversePaeth RowBytes, PrevRowBytes, Ival
        End Select
        CopyMemory NewDat(DestByte), RowBytes(0), BPRow - 1
        PrevRowBytes = RowBytes
    Next Y
    
    DeFilter = NewDat
End Function
Private Sub DrawAlpha(hDC As Long, X As Long, Y As Long, BmpImg As BITMAPIMAGE)
    Dim Myx As Long, Myy As Long, DatOff As Long
    Dim R As Long, G As Long, B As Long, A As Long
    Dim sR As Long, sG As Long, sB As Long
    Dim dR As Long, dG As Long, dB As Long
    Dim DestData() As Byte, BytesPerRow As Long
    Dim DestOff As Long, DestHdr As BITMAPINFOHEADER
    Dim MemDC As Long, hBmp As Long, hOldBmp As Long
    Dim SrcData() As Byte
    
    Dim lngFilterColor As Long
    
    BytesPerRow = LineBytes(MyHdr.Width, 24)
    ReDim DestData(0 To BytesPerRow * MyHdr.Height - 1)
    ReDim SrcData(0 To UBound(DestData))
    DestHdr = BmpImg.BmpHdr
    DestHdr.BitCount = 24

    MemDC = CreateCompatibleDC(hDC)
    hBmp = CreateCompatibleBitmap(hDC, MyHdr.Width, MyHdr.Height)
    hOldBmp = SelectObject(MemDC, hBmp)
    BitBlt MemDC, 0, 0, MyHdr.Width, MyHdr.Height, hDC, X, Y, vbSrcCopy
    GetDIBits MemDC, hBmp, 0, MyHdr.Height, SrcData(0), DestHdr, 0
    SelectObject hOldBmp, MemDC
    DeleteObject hBmp
    DeleteDC MemDC
    
    For Myy = 0 To MyHdr.Height - 1
        For Myx = 0 To MyHdr.Width - 1
            DestOff = Myy * BytesPerRow + Myx * 3

            sR = SrcData(DestOff + 2)
            sG = SrcData(DestOff + 1)
            sB = SrcData(DestOff)
            
            With BmpImg
                B = .BmpData(DatOff)
                G = .BmpData(DatOff + 1)
                R = .BmpData(DatOff + 2)
                A = .BmpData(DatOff + 3)
            End With
            
            If boolGammaEnabled Then
                lngFilterColor = filterColor(RGB(R, G, B), intGammaBlue, intGammaGreen, intGammaRed, lngGray, lngBoost)
                With BmpImg
                    R = (lngFilterColor And &HFF)
                    G = ((lngFilterColor And &H100FF00) / &H100)
                    B = ((lngFilterColor And &HFF0000) / &H10000)
                End With
            End If

            If A = 255 Then
                DestData(DestOff + 2) = R
                DestData(DestOff + 1) = G
                DestData(DestOff) = B
            ElseIf A = 0 Then
                DestData(DestOff + 2) = sR
                DestData(DestOff + 1) = sG
                DestData(DestOff) = sB
            Else
                dR = R * A + (255 - A) * sR + 255
                dG = G * A + (255 - A) * sG + 255
                dB = B * A + (255 - A) * sB + 255
                
                CopyMemory DestData(DestOff + 2), ByVal VarPtr(dR) + 1, 1
                CopyMemory DestData(DestOff + 1), ByVal VarPtr(dG) + 1, 1
                CopyMemory DestData(DestOff), ByVal VarPtr(dB) + 1, 1
            End If
            DatOff = DatOff + 4
        Next Myx
    Next Myy
    SetDIBitsToDevice hDC, X, Y, MyHdr.Width, MyHdr.Height, 0, 0, 0, _
        MyHdr.Height, DestData(0), DestHdr, 0
End Sub
Private Function GetR(Color As Long) As Byte
    CopyMemory GetR, Color, 1
End Function

Private Function GetG(Color As Long) As Byte
    CopyMemory GetG, ByVal VarPtr(Color) + 1, 1
End Function

Private Function GetB(Color As Long) As Byte
    CopyMemory GetB, ByVal VarPtr(Color) + 2, 1
End Function
Public Sub DrawToDC(hDC As Long, X As Long, Y As Long)
    If Not ((MyBmpImg.BmpHdr.BitCount = 32) Or KeyExists(MyChunks, "tRNS1")) Then
        SetToDC hDC, X, Y, MyBmpImg
    Else
        DrawAlpha hDC, X, Y, MyBmpImg
    End If
End Sub
Private Function FileExists(FileName As String) As Boolean
    On Error Resume Next
    FileExists = (FileLen(FileName) >= 0)
End Function

Private Function GetBmpImg() As BITMAPIMAGE
    Dim DataSize As Long, DataPerRow As Long
    Dim Dat() As Byte, BmpDat() As Byte
    Dim BmpHdr As BITMAPINFOHEADER, Pal() As PNGRGB
    Dim tRNS As PNGCHUNK
    
    If Not HasLoaded Then Exit Function
    
    DataPerRow = (MyHdr.Width * BitsPerPixel(MyHdr) + 7) \ 8 + 1
    DataSize = DataPerRow * MyHdr.Height
    Dat = Decompress(GetData, DataSize)
    RaiseEvent LoadProgress(4, 2)
    Dat = DeFilter(MyHdr, Dat)
    RaiseEvent LoadProgress(4, 3)
    With BmpHdr
        .Size = Len(BmpHdr)
        .Width = MyHdr.Width
        .Height = MyHdr.Height
        .Planes = 1
    End With
    
    If KeyExists(MyChunks, "tRNS1") Then
        tRNS = MyChunks("tRNS1")
    End If

    Select Case MyHdr.ColorType
        Case 0 'Grayscale
            GetBmpImg.ClrTbl = GetGrayTbl(CInt(MyHdr.BitDepth))
            Select Case MyHdr.BitDepth
                Case 16
                    Conv16To8 Dat
                    BmpHdr.BitCount = 8
                Case 8, 4, 1
                    BmpHdr.BitCount = MyHdr.BitDepth
                Case 2
                    BmpHdr.BitCount = 8
                    Pal2To8 MyHdr.Width, MyHdr.Height, Dat, DataPerRow
            End Select
        Case 2 'RGB
            If MyHdr.BitDepth = 16 Then Conv16To8 Dat
            BmpHdr.BitCount = 24
            ReverseRGB Dat
        Case 3 'Palette
            Pal = GetPal()
            Select Case MyHdr.BitDepth
                Case 8, 4, 1: BmpHdr.BitCount = MyHdr.BitDepth
                Case 2
                    BmpHdr.BitCount = 8
                    Pal2To8 MyHdr.Width, MyHdr.Height, Dat, DataPerRow
            End Select
            If tRNS.ChkLength > 0 Then
                If MyHdr.BitDepth <> 2 Then
                    Align32 MyHdr, BmpHdr.BitCount, Dat
                End If
                PalToRGBA MyHdr.Width, MyHdr.Height, BmpHdr.BitCount, _
                    Dat, tRNS, Pal
                BmpHdr.BitCount = 32
            End If
            GetBmpImg.ClrTbl = GetWinPal(Pal)
            GetBmpImg.BmpHdr.ClrUsed = UBound(Pal) + 1
        Case 4 'Grayscale + Alpha
            If MyHdr.BitDepth = 16 Then Conv16To8 Dat
            GrayAToRGBA Dat
            BmpHdr.BitCount = 32
        Case 6 'RGB + Alpha
            If MyHdr.BitDepth = 16 Then Conv16To8 Dat
            BmpHdr.BitCount = 32
            ReverseRGBA Dat
    End Select
    GetBmpImg.BmpHdr = BmpHdr
    If Not (((MyHdr.ColorType = 3) And (BmpHdr.BitCount = 32)) Or _
        (MyHdr.BitDepth = 2)) Then
        If MyHdr.BitDepth = 16 Then
            MyHdr.BitDepth = 8
            Align32 MyHdr, BmpHdr.BitCount, Dat
            MyHdr.BitDepth = 16
        Else
            Align32 MyHdr, BmpHdr.BitCount, Dat
        End If
    End If
    MirrorData Dat, LineBytes(BmpHdr.Width, BmpHdr.BitCount)
    GetBmpImg.BmpData = Dat
    RaiseEvent LoadProgress(4, 4)
End Function
Private Function GetChunk(FileNum As Integer) As PNGCHUNK
    Dim ChkLength As Long, ChkType(1 To 4) As Byte
    Dim ChkCRC As Long
    
    Get FileNum, , ChkLength
    GetChunk.ChkLength = SwapLng(ChkLength)
    
    Get FileNum, , ChkType
    GetChunk.ChkType = BytesToStr(ChkType)
    
    If GetChunk.ChkLength > 0 Then
        ReDim GetChunk.ChkData(0 To GetChunk.ChkLength - 1)
        Get FileNum, , GetChunk.ChkData
    End If
    
    Get FileNum, , ChkCRC
    GetChunk.ChkCRC = SwapLng(ChkCRC)
End Function
Private Function GetCRC(Stream() As Byte) As Long

End Function

Private Function GetData() As Byte()
    Dim n As Long, DatChk As PNGCHUNK, MyDat() As Byte
    
    If Not HasLoaded Then Exit Function
    
    Do
        n = n + 1
        If Not KeyExists(MyChunks, "IDAT" & n) Then Exit Do
        DatChk = MyChunks("IDAT" & n)
        If n = 1 Then MyDat = DatChk.ChkData: GoTo NextDat
        ReDim Preserve MyDat(0 To UBound(MyDat) + _
            DatChk.ChkLength)
        CopyMemory MyDat(UBound(MyDat) - DatChk.ChkLength + 1), _
            DatChk.ChkData(0), DatChk.ChkLength
NextDat:
    Loop
    
    GetData = MyDat
End Function

Private Function GetHeader(Chunk As PNGCHUNK) As IHDR
    CopyMemory GetHeader, Chunk.ChkData(0), Len(GetHeader)
End Function

Private Function GetPal() As PNGRGB()
    If Not HasLoaded Then Exit Function

    Dim PlteChk As PNGCHUNK, Plte() As PNGRGB
    
    If KeyExists(MyChunks, "PLTE1") Then
        PlteChk = MyChunks("PLTE1")
        ReDim Plte(0 To PlteChk.ChkLength \ 3 - 1)
        CopyMemory Plte(0), PlteChk.ChkData(0), PlteChk.ChkLength
        
        GetPal = Plte
    End If
End Function

Private Function GetWinPal(Pal() As PNGRGB) As RGBQUAD()
    Dim n As Long, WinPal() As RGBQUAD
    
    ReDim WinPal(0 To UBound(Pal))
    
    For n = 0 To UBound(Pal)
        With WinPal(n)
            .Red = Pal(n).Red
            .Green = Pal(n).Green
            .Blue = Pal(n).Blue
        End With
    Next n
    
    GetWinPal = WinPal
End Function

Private Function HasLoaded() As Boolean
    HasLoaded = KeyExists(MyChunks, "IHDR1")
End Function

Private Function Interval(Hdr As IHDR) As Long
    Interval = BitsPerPixel(Hdr) \ 8
    If Interval = 0 Then Interval = 1
End Function

Private Function IsASCIIAlpha(Char As Byte) As Boolean
    IsASCIIAlpha = ((65 <= Char <= 90) Or (97 <= Char <= 122))
End Function
Private Function IsCharUpper(Char As Byte) As Boolean
    If Char >= 97 Then
        Exit Function
    Else
        IsCharUpper = True
    End If
End Function
Private Function IsCriticalChunk(Chunk As PNGCHUNK) As Boolean
    IsCriticalChunk = IsCharUpper(Asc(Left(Chunk.ChkType, 1)))
End Function
Private Function IsPublicChunk(Chunk As PNGCHUNK) As Boolean
    IsPublicChunk = IsCharUpper(Asc(Mid(Chunk.ChkType, 2, 1)))
End Function

Private Function IsSafeToCopyChunk(Chunk As PNGCHUNK) As Boolean
    IsSafeToCopyChunk = IsCharUpper(Asc(Right(Chunk.ChkType, 1)))
End Function
Private Function IsValidChunk(Chunk As PNGCHUNK) As Boolean
    Dim n As Integer
    
    If Chunk.ChkLength < 0 Then Exit Function
    
    For n = 1 To 4
        If Not IsASCIIAlpha(Asc(Mid(Chunk.ChkType, n, 1))) _
            Then Exit Function
    Next n
    
    If Not IsCharUpper(Asc(Mid(Chunk.ChkType, 3, 1))) Then _
        Exit Function
    
    'If Not GetCRC(Chunk.ChkData) = Chunk.ChkCRC Then _
    '    Exit Function
        
    IsValidChunk = True
End Function
Private Function IsValidHeader(Hdr As IHDR) As Boolean
    With Hdr
        .Width = SwapLng(.Width)
        .Height = SwapLng(.Height)
        
        If .Width < 1 Then Exit Function
        If .Height < 1 Then Exit Function
        Select Case .ColorType
            Case 3
                Select Case .BitDepth
                    Case 1, 2, 4, 8
                    Case Else: Exit Function
                End Select
            Case 2, 4, 6
                Select Case .BitDepth
                    Case 8, 16
                    Case Else: Exit Function
                End Select
            Case 0
                Select Case .BitDepth
                    Case 1, 2, 4, 8, 16
                    Case Else: Exit Function
                End Select
            Case Else: Exit Function
        End Select
        If .Compression <> 0 Then Exit Function
        If .Filter <> 0 Then Exit Function
        Select Case .Interlacing
            Case 0, 1
            Case Else: Exit Function
        End Select
    End With
    
    IsValidHeader = True
End Function
Private Function IsValidSignature(FileNum As Integer) As Boolean
    Dim Signature(1 To 8) As Byte
    
    Get FileNum, , Signature
    
    If Signature(1) <> 137 Then Exit Function
    If Signature(2) <> 80 Then Exit Function
    If Signature(3) <> 78 Then Exit Function
    If Signature(4) <> 71 Then Exit Function
    If Signature(5) <> 13 Then Exit Function
    If Signature(6) <> 10 Then Exit Function
    If Signature(7) <> 26 Then Exit Function
    If Signature(8) <> 10 Then Exit Function
    
    IsValidSignature = True
End Function
Private Function KeyExists(Coll As Collection, KeyName As String) As Boolean
    On Error Resume Next
    Coll (KeyName)
    KeyExists = (Err.Number <> 5)
End Function

Private Function LineBytes(Width As Long, BitCount As Integer) As Long
    LineBytes = ((Width * BitCount + 31) \ 32) * 4
End Function
Public Function LoadPNGFile(FileName As String) As PNGErrorContants
    Dim FileNum As Integer, CurChk As PNGCHUNK
    Dim Chunks As New Collection, n As Long, m As Long
    Dim KeyName As String, Hdr As IHDR, AfterIDAT As Boolean
    
    If Not FileExists(FileName) Then
        LoadPNGFile = pngeFileNotFound
        Exit Function
    End If
    
    FileNum = FreeFile
    On Error GoTo OpenErr
    Open FileName For Binary Access Read As FileNum
    On Error GoTo 0

    If Not IsValidSignature(FileNum) Then GoTo NoValidFile
    
    RaiseEvent LoadProgress(4, 0)
    
    Do
        n = n + 1
        CurChk = GetChunk(FileNum)
        If Not IsValidChunk(CurChk) Then GoTo NoValidFile
        If KeyExists(Chunks, "IDAT1") And (CurChk.ChkType <> "IDAT") Then
            AfterIDAT = True
        End If
        Select Case CurChk.ChkType
            Case "IHDR"
                If n <> 1 Then GoTo NoValidFile
                If CurChk.ChkLength <> 13 Then GoTo NoValidFile
                Hdr = GetHeader(CurChk)
                If Not IsValidHeader(Hdr) Then GoTo NoValidFile
            Case "PLTE"
                If KeyExists(Chunks, "PLTE1") Then GoTo NoValidFile
                If KeyExists(Chunks, "IDAT1") Then GoTo NoValidFile
                If CurChk.ChkLength < 3 Then GoTo NoValidFile
                If (CurChk.ChkLength Mod 3) <> 0 Then GoTo NoValidFile
                Select Case Hdr.ColorType
                    Case 2, 3, 6
                    Case Else: GoTo NoValidFile
                End Select
                If (CurChk.ChkLength \ 3) > (2 ^ Hdr.BitDepth) Then _
                    GoTo NoValidFile
            Case "tRNS"
                If Hdr.ColorType = 3 Then
                    If Not KeyExists(Chunks, "PLTE1") Then GoTo NoValidFile
                    If CurChk.ChkLength > (Chunks("PLTE1").ChkLength \ 3) Then
                        GoTo NoValidFile
                    End If
                End If
                If KeyExists(Chunks, "IDAT1") Then GoTo NoValidFile
                Select Case Hdr.ColorType
                    Case 4, 6: GoTo NoValidFile
                End Select
            Case "bKGD"
                If Hdr.ColorType = 3 Then
                    If Not KeyExists(Chunks, "PLTE1") Then GoTo NoValidFile
                End If
                If KeyExists(Chunks, "IDAT1") Then GoTo NoValidFile
            Case "IDAT"
                If AfterIDAT Then GoTo NoValidFile
            Case "IEND"
                If Loc(FileNum) <> LOF(FileNum) Then GoTo NoValidFile
            Case Else
                If IsCriticalChunk(CurChk) Then GoTo NoValidFile
        End Select
        With CurChk
            RaiseEvent ChunkFound(.ChkType, .ChkLength, .ChkData, .ChkCRC)
        End With
        m = 0
        Do
            m = m + 1
            KeyName = CurChk.ChkType & m
            If Not KeyExists(Chunks, KeyName) Then Exit Do
        Loop
        Chunks.Add CurChk, KeyName
        If CurChk.ChkType = "IEND" Then Exit Do
    Loop

    Close FileNum
    
    If (Hdr.ColorType = 3) And (Not KeyExists(Chunks, "PLTE1")) Then _
        GoTo NoValidFile
    Set MyChunks = Chunks
    MyHdr = Hdr
    RaiseEvent LoadProgress(4, 1)
    If MyHdr.Interlacing = 0 Then
        MyBmpImg = GetBmpImg
    Else
        MyBmpImg = GetBmpImgInterlaced
    End If
    strCurrentFilename = FileName
    LoadPNGFile = pngeSucceeded
    Exit Function
NoValidFile:
    Close FileNum
    LoadPNGFile = pngeInvalidFile
    Exit Function
OpenErr:
    LoadPNGFile = pngeOpenError
End Function
Private Sub MirrorData(Dat() As Byte, RowBytes As Long)
    Dim NewDat() As Byte, Y As Long, Height As Long
    Dim StartLine As Long, DestLine As Long
    
    ReDim NewDat(0 To UBound(Dat))
    Height = (UBound(Dat) + 1) \ RowBytes
    
    For Y = 0 To Height - 1
        StartLine = Y * RowBytes
        DestLine = (Height - Y - 1) * RowBytes
        CopyMemory NewDat(DestLine), Dat(StartLine), RowBytes
    Next Y
    
    Dat = NewDat
End Sub

Private Function PaethPredictor(Left As Byte, Above As Byte, UpperLeft As Byte) As Byte
    Dim pA As Integer, pB As Integer, pC As Integer, p As Integer
    
    p = CInt(Left) + CInt(Above) - CInt(UpperLeft)
    
    pA = Abs(p - Left)
    pB = Abs(p - Above)
    pC = Abs(p - UpperLeft)
    
    If (pA <= pB) And (pA <= pC) Then
        PaethPredictor = Left
    ElseIf pB <= pC Then
        PaethPredictor = Above
    Else
        PaethPredictor = UpperLeft
    End If
End Function

Private Sub PalToRGBA(Width As Long, Height As Long, BitDepth As Integer, Dat() As Byte, tRNS As PNGCHUNK, Pal() As PNGRGB)
    Dim DestDat() As Byte, n As Long, PalEntry As Byte
    Dim DestOff As Long, TrnsBnd As Long
    Dim X As Long, Y As Long, WidthBytes As Long
    
    ReDim DestDat(0 To 4 * Width * Height - 1)
    TrnsBnd = UBound(tRNS.ChkData)
    WidthBytes = LineBytes(Width, BitDepth)

    Select Case BitDepth
        Case 8
            For Y = 0 To Height - 1
                For X = 0 To Width - 1
                    n = Y * WidthBytes + X
                    PalEntry = Dat(n)
                    With Pal(PalEntry)
                        DestDat(DestOff) = .Blue
                        DestDat(DestOff + 1) = .Green
                        DestDat(DestOff + 2) = .Red
                    End With
                    If PalEntry <= TrnsBnd Then
                        DestDat(DestOff + 3) = tRNS.ChkData(PalEntry)
                    Else
                        DestDat(DestOff + 3) = 255
                    End If
                    DestOff = DestOff + 4
                Next X
            Next Y
        Case 4
            For Y = 0 To Height - 1
                For X = 0 To Width - 1
                    n = Y * WidthBytes + X \ 2
                    
                    If (X Mod 2) = 1 Then
                        PalEntry = Dat(n) And 15
                    Else
                        PalEntry = (Dat(n) \ 16) And 15
                    End If
                    With Pal(PalEntry)
                        DestDat(DestOff) = .Blue
                        DestDat(DestOff + 1) = .Green
                        DestDat(DestOff + 2) = .Red
                    End With
                    If PalEntry <= TrnsBnd Then
                        DestDat(DestOff + 3) = tRNS.ChkData(PalEntry)
                    Else
                        DestDat(DestOff + 3) = 255
                    End If
                    DestOff = DestOff + 4
                Next X
            Next Y
        Case 1
            For Y = 0 To Height - 1
                For X = 0 To Width - 1
                    n = Y * WidthBytes + X \ 8
                    If (X Mod 8) <> 7 Then
                        PalEntry = (Dat(n) \ 2 ^ (7 - X Mod 8)) And 1
                    Else
                        PalEntry = Dat(n) And 1
                    End If
                    With Pal(PalEntry)
                        DestDat(DestOff) = .Blue
                        DestDat(DestOff + 1) = .Green
                        DestDat(DestOff + 2) = .Red
                    End With
                    If PalEntry <= TrnsBnd Then
                        DestDat(DestOff + 3) = tRNS.ChkData(PalEntry)
                    Else
                        DestDat(DestOff + 3) = 255
                    End If
                    DestOff = DestOff + 4
                Next X
            Next Y
    End Select
    
    Dat = DestDat
End Sub

Private Sub ReverseRGB(Dat() As Byte)
    Dim n As Long, Tmp As Byte
    
    For n = 0 To UBound(Dat) Step 3
        Tmp = Dat(n)
        Dat(n) = Dat(n + 2)
        Dat(n + 2) = Tmp
    Next n
End Sub

Private Sub ReverseRGBA(Dat() As Byte)
    Dim n As Long, Tmp As Byte
    
    For n = 0 To UBound(Dat) Step 4
        Tmp = Dat(n)
        Dat(n) = Dat(n + 2)
        Dat(n + 2) = Tmp
    Next n
End Sub

Private Sub ReverseSub(CurRow() As Byte, Interval As Long)
    Dim PrevOff As Long, PrevVal As Byte, BPRow As Long
    Dim n As Long, X As Integer
    
    BPRow = UBound(CurRow) + 1
   
    For n = 0 To BPRow - 1
        PrevOff = n - Interval
        If PrevOff >= 0 Then
            PrevVal = CurRow(PrevOff)
        End If
        
        X = CInt(CurRow(n)) + CInt(PrevVal)
        CopyMemory CurRow(n), X, 1
    Next n
End Sub

Private Sub ReverseAverage(CurRow() As Byte, PrevRow() As Byte, Interval As Long)
    Dim PrevOff As Long, PrevVal As Byte, BPRow As Long
    Dim n As Long, X As Integer
    
    BPRow = UBound(CurRow) + 1
   
    For n = 0 To BPRow - 1
        PrevOff = n - Interval
        If PrevOff >= 0 Then
            PrevVal = CurRow(PrevOff)
        End If
        
        X = CurRow(n) + (CInt(PrevRow(n)) + CInt(PrevVal)) \ 2
        CopyMemory CurRow(n), X, 1
    Next n
End Sub
Private Sub ReversePaeth(CurRow() As Byte, PrevRow() As Byte, Interval As Long)
    Dim BPRow As Long, n As Long, X As Integer
    Dim LeftPixOff As Long, LeftPix As Byte
    Dim UpperLeftPix As Byte
    
    BPRow = UBound(CurRow) + 1
   
    For n = 0 To BPRow - 1
        LeftPixOff = n - Interval
        If LeftPixOff >= 0 Then
            LeftPix = CurRow(LeftPixOff)
            UpperLeftPix = PrevRow(LeftPixOff)
        End If

        X = CInt(CurRow(n)) + CInt(PaethPredictor( _
            LeftPix, PrevRow(n), UpperLeftPix))
        CopyMemory CurRow(n), X, 1
    Next n
End Sub
Private Sub ReverseUp(CurRow() As Byte, PrevRow() As Byte)
    Dim PrevVal As Byte, BPRow As Long
    Dim n As Long, X As Integer
    
    BPRow = UBound(CurRow) + 1
   
    For n = 0 To BPRow - 1
        PrevVal = PrevRow(n)
        
        X = CInt(CurRow(n)) + CInt(PrevVal)
        CopyMemory CurRow(n), X, 1
    Next n
End Sub
Private Sub SetToDC(hDC As Long, X As Long, Y As Long, BmpImg As BITMAPIMAGE)
    Dim BmpInfo() As Byte, TblSize As Long

    On Error Resume Next
    TblSize = (UBound(BmpImg.ClrTbl) + 1) * 4
    On Error GoTo 0
    ReDim BmpInfo(1 To Len(BmpImg.BmpHdr) + TblSize)
    CopyMemory BmpInfo(1), BmpImg.BmpHdr, Len(BmpImg.BmpHdr)
    If TblSize > 0 Then
        CopyMemory BmpInfo(Len(BmpImg.BmpHdr) + 1), BmpImg.ClrTbl(0), _
            TblSize
    End If
    With BmpImg.BmpHdr
        SetDIBitsToDevice hDC, X, Y, .Width, .Height, 0, 0, 0, .Height, _
            BmpImg.BmpData(0), BmpInfo(1), 0
    End With
End Sub

Private Function SwapLng(ByVal SrcLng As Long) As Long
    CopyMemory ByVal VarPtr(SwapLng) + 3, SrcLng, 1
    CopyMemory ByVal VarPtr(SwapLng) + 2, ByVal VarPtr(SrcLng) + 1, 1
    CopyMemory ByVal VarPtr(SwapLng) + 1, ByVal VarPtr(SrcLng) + 2, 1
    CopyMemory SwapLng, ByVal VarPtr(SrcLng) + 3, 1
End Function
Public Property Get Width() As Long
    Width = MyHdr.Width
End Property
