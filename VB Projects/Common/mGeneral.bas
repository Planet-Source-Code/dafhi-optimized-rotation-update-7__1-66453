Attribute VB_Name = "mGeneral"
Option Explicit

'mGeneral.bas by dafhi  Sept 16, 2006

'This module contains declarations that I use alot

Public Type PointAPI
    X    As Integer
    Y    As Integer
End Type

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type RGBTriple
    Blue As Byte
    Green As Byte
    Red As Byte
End Type

Type RGBQUAD
 Blue  As Byte
 Green As Byte
 Red   As Byte
 Alpha As Byte
End Type

Type Bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(1) As SAFEARRAYBOUND
End Type

Public Type ImageHook1D
    SA            As SAFEARRAY1D
    Ary()         As Long
End Type

Dim I As Long
Dim J As Long

Public Const pi As Double = 3.14159265358979
Public Const TwoPi As Double = 2 * pi
Public Const piBy2 As Single = pi / 2
Public Const halfPi As Single = piBy2

Const HWND_TOPMOST = -1
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1

Public Const NOTE_1OF12 As Double = 2 ^ (1 / 12)

Public Const ASC_DOUBLE_QUOTE As Integer = 34

Dim LBA  As Long
Dim UBA  As Long
Dim LenA As Long

Dim mStr       As String
Dim mAsc       As Integer

Dim mStrA      As String
Dim mStrB      As String

'ARGBHSV() Function
Private Blu_&
Private Grn_&
Private Red_&
Private subt!

Public Const GrayScaleRGB As Long = 1 + 256& + 65536

Public Const MaskHIGH       As Long = &HFF0000
Public Const MaskMID        As Long = &HFF00&
Public Const MaskLOW        As Long = &HFF&
Public Const MaskRB         As Long = &HFF00FF
Public Const MaskR          As Long = &HFF0000
Public Const MaskG          As Long = &HFF00&
Public Const MaskB          As Long = &HFF&
Public Const RB_Add1        As Long = &H10001
Public Const G_Add1         As Long = &H100&
Public Const L65536         As Long = 65536
Public Const L256           As Long = 256&

'skew corner
Public g_sk_zoom   As Single
Public g_sk_angle  As Single

'CheckFPS()
Public Tick       As Long
Public FrameCount As Long
Public speed      As Single
Public sFPS       As Single

Public PrevTick   As Long
Public NextTick   As Long
Private TickSum   As Long

Private Const Interval_Micro As Long = 4

Public Const TIME_MARK   As Integer = 256

'Midi standard
Public Const NOTE_ON     As Byte = &H90
Public Const NOTE_OFF    As Byte = &H80
Public Const NOTE_GONE   As Byte = &H81

Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

'Private Type BITMAPINFO256 'from www.vbAccelerator.com
'    bmiHeader As BITMAPINFOHEADER
'    bmiColors(0 To 255) As RGBQUAD
'End Type

Declare Function StretchDIBits Lib "gdi32" _
        (ByVal hDC As Long, _
         ByVal X As Long, _
         ByVal Y As Long, _
         ByVal dx As Long, _
         ByVal dy As Long, _
         ByVal SrcX As Long, _
         ByVal SrcY As Long, _
         ByVal wSrcWidth As Long, _
         ByVal wSrcHeight As Long, _
         lpBits As Any, _
         lpBitsInfo As BITMAPINFOHEADER, _
         ByVal wUsage As Long, _
         ByVal dwRop As Long) As Long

Public Const BI_RGB         As Long = 0
Public Const DIB_RGB_COLORS As Long = 0

Private Type SliderProc
    lLefPlus   As Integer
    lRghM      As Integer
    lBotM      As Integer
    lWidM2     As Integer
    lHgtM2     As Integer
    lPos       As Integer
    lPosP      As Integer
    MousDown   As Integer
    IsVertical As Boolean
End Type

Public Type SliderControl
    lTop       As Integer
    lLeft      As Integer
    lBot       As Integer
    lRight     As Integer
    lColor     As Long
    val_0_to_1 As Single
    Proc       As SliderProc
End Type

'Type LARGE_INTEGER
'    lowpart As Long
'    highpart As Long
'End Type
'Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
'Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Declare Function timeGetTime Lib "winmm.dll" () As Long
Declare Function RedrawWindow Lib "user32" (ByVal hwnd&, lprcUpdate As RECT, ByVal hrgnUpdate&, ByVal fuRedraw&) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'SurfaceDescFromFile
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

'Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lpBits As Long, ByVal Handle As Long, ByVal dw As Long) As Long
'Public Declare Function CreateDIBitmap Lib "gdi32" (ByVal hdc As Long, _
 lpInfoHeader As BITMAPINFOHEADER, _
 ByVal dwUsage As Long, _
 lpInitBits As Any, _
 lpInitInfo As BITMAPINFO, _
 ByVal wUsage As Long) As Long

'from www.vbAccelerator.com
'Private Declare Function GetDIBits256 Lib "gdi32" Alias "GetDIBits" (ByVal aHDC _
 As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As _
 Long, lpBits As Any, lpBI As BITMAPINFO256, ByVal wUsage As Long) As Long

'Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As Bitmap) As Long
'Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, _
 ByVal nHeight As Long, _
 ByVal nPlanes As Long, _
 ByVal nBitCount As Long, _
 lpBits As Any) As Long
 
Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, _
 ByVal hBitmap As Long, _
 ByVal nStartScan As Long, _
 ByVal nNumScans As Long, _
 lpBits As Any, _
 lpBI As BITMAPINFO, _
 ByVal wUsage As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
'Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
'Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy&)

Public Type BitmapFileHeader ' 14 bytes
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

Public Const BMPFileSignature As Integer = &H4D42

'****************'
'*              *'
'*   Graphics   *'
'*              *'
'****************'

Function GetPadBytes(ByVal PelWidth As Integer, Optional ByVal BytesPixel As Integer = 3, Optional ByVal ByteAlign As Long = 4) As Long
    GetPadBytes = ByteAlign - 1 - (BytesPixel * PelWidth + ByteAlign - 1) Mod ByteAlign
End Function


Public Function RGBHSV(hue_0_To_1530!, ByVal saturation_0_To_1!, value_0_To_255!) As Long
Dim hue_and_sat As Single
Dim value1      As Single
Dim diff1       As Single
Dim maxim       As Single

 If value_0_To_255 > 0 Then
  value1 = value_0_To_255 + 0.5
  If saturation_0_To_1 > 0 Then
   maxim = hue_0_To_1530 - 1530& * Int(hue_0_To_1530 / 1530&)
   diff1 = saturation_0_To_1 * value_0_To_255
   subt = value1 - diff1
   diff1 = diff1 / 255
   If maxim <= 510 Then
    Blu_ = Int(subt)
    If maxim <= 255 Then
     hue_and_sat = maxim * diff1!
     Red_ = Int(value1)
     Grn_ = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 255) * diff1!
     Grn_ = Int(value1)
     Red_ = Int(value1 - hue_and_sat)
    End If
   ElseIf maxim <= 1020 Then
    Red_ = Int(subt)
    If maxim <= 765 Then
     hue_and_sat = (maxim - 510) * diff1!
     Grn_ = Int(value1)
     Blu_ = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 765) * diff1!
     Blu_ = Int(value1)
     Grn_ = Int(value1 - hue_and_sat)
    End If
   Else
    Grn_ = Int(subt)
    If maxim <= 1275 Then
     hue_and_sat = (maxim - 1020) * diff1!
     Blu_ = Int(value1)
     Red_ = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 1275) * diff1!
     Red_ = Int(value1)
     Blu_ = Int(value1 - hue_and_sat)
    End If
   End If
   RGBHSV = Red_ Or Grn_ * 256& Or Blu_ * 65536
  Else 'saturation_0_To_1 <= 0
   RGBHSV = Int(value1) * CLng(65793) '1 + 256 + 65536
  End If
 Else 'value_0_To_255 <= 0
  RGBHSV = 0&
 End If
End Function
Public Function ARGBHSV(ByVal hue_0_To_1530!, ByVal saturation_0_To_1!, value_0_To_255!) As Long
Dim hue_and_sat As Single
Dim value1      As Single
Dim diff1       As Single
Dim maxim       As Single

 If value_0_To_255 > 0 Then
  value1 = value_0_To_255 + 0.5
  If saturation_0_To_1 > 0 Then
   maxim = hue_0_To_1530 - 1530& * Int(hue_0_To_1530 / 1530&)
   diff1 = saturation_0_To_1 * value_0_To_255
   subt = value1 - diff1
   diff1 = diff1 / 255
   If maxim <= 510 Then
    Blu_ = Int(subt)
    If maxim <= 255 Then
     hue_and_sat = maxim * diff1!
     Red_ = Int(value1)
     Grn_ = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 255) * diff1!
     Grn_ = Int(value1)
     Red_ = Int(value1 - hue_and_sat)
    End If
   ElseIf maxim <= 1020 Then
    Red_ = Int(subt)
    If maxim <= 765 Then
     hue_and_sat = (maxim - 510) * diff1!
     Grn_ = Int(value1)
     Blu_ = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 765) * diff1!
     Blu_ = Int(value1)
     Grn_ = Int(value1 - hue_and_sat)
    End If
   Else
    Grn_ = Int(subt)
    If maxim <= 1275 Then
     hue_and_sat = (maxim - 1020) * diff1!
     Blu_ = Int(value1)
     Red_ = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 1275) * diff1!
     Red_ = Int(value1)
     Blu_ = Int(value1 - hue_and_sat)
    End If
   End If
   ARGBHSV = Red_ * 65536 Or Grn_ * 256& Or Blu_
  Else 'saturation_0_To_1 <= 0
   ARGBHSV = Int(value1) * CLng(65793) '1 + 256 + 65536
  End If
 Else 'value_0_To_255 <= 0
  ARGBHSV = 0&
 End If
End Function
Public Function FlipRB(Color_ As Long) As Long
Dim LBlu As Long
    LBlu = Color_ And &HFF&
    FlipRB = (Color_ And &HFF00&) + 256& * (LBlu * 256&) + (Color_ \ 256&) \ 256&
End Function


Sub FPS_Init() 'right before game loop
    PrevTick = timeGetTime
    NextTick = PrevTick + Interval_Micro - 1
End Sub
Function CheckFPS(Optional RetFPS, Optional ByVal speed_coefficient As Single = 1, Optional Interval_Millisec& = 200) As Boolean
    
'CODE SAMPLE
'1. Paste comments below to Form
'2. hit ctrl-h
'3. line 1 says [comment mark][1 space] (2 characters total)
'4. line 2 says nothing
'5. Replace All
'6. be sure to reference mGeneral.bas
    
' Private Sub Form_Load()
    ' FPS_Init 'initialize time variables
    ' Do While DoEvents '"very simple game loop"
        
        ' Cls
        ' Print "posx = posx + dx * speed
        ' Print "speed is smaller for faster CPU
        
        ' If CheckFPS(FPS, speed_multiplier, 200) Then
        '    Caption = "FPS: " & FPS
        ' End If
    ' Loop
' End Sub
    
    Tick = timeGetTime
    
    FrameCount = FrameCount + 1
    TickSum = Tick - PrevTick
    speed = speed_coefficient * (TickSum / FrameCount)
    
    If Tick > NextTick Then
        RetFPS = 1000 * FrameCount / TickSum
        sFPS = RetFPS
        NextTick = Tick + Interval_Millisec - 1
        If NextTick < Tick Then NextTick = Tick
        FrameCount = 0
        PrevTick = Tick
        CheckFPS = True
    Else
        CheckFPS = False
    End If

End Function


'********************'
'*                  *'
'*   String stuff   *'
'*                  *'
'********************'

Sub FillBytesFromString(Bytes1() As Byte, ByVal Str1 As String)
    LBA = LBound(Bytes1)
    UBA = UBound(Bytes1)
    mStr = Left$(Str1, UBA - LBA + 1)
    For I = LBA To UBA
        Bytes1(I) = Asc(Mid$(mStr, I + 1, 1))
    Next
End Sub
Function StringFromBytes(Bytes() As Byte) As String
    LenA = UBound(Bytes) - LBound(Bytes) + 1
    If LenA > 0 Then
        StringFromBytes = Bytes
        StringFromBytes = StringFromBytes + StringFromBytes
        For I = LBound(Bytes) To UBound(Bytes)
            Mid$(StringFromBytes, I + 1, 1) = Chr$(Bytes(I))
        Next
    End If
End Function
Function GetLine(StrInput As String, ByVal POS_ As Long, Optional RetPos As Long) As String
    If POS_ > Len(StrInput) Then
        GetLine = ""
        RetPos = POS_
        Exit Function
    End If
    For I = POS_ To Len(StrInput)
        J = Asc(Mid$(StrInput, I, 1))
        If I = 10 Or I = 13 Then Exit For
    Next
    GetLine = Mid$(StrInput, POS_, I - POS_)
    RetPos = POS_
End Function
Public Sub NumbersOnly(pTxt As TextBox, Optional pRetVal As Variant, Optional MightUseDollarSign As Boolean, Optional IntegerOnly As Boolean = False, Optional AlwaysPositive As Boolean = False)
Dim PointCount As Integer
Dim MinusCount As Integer
Dim J1 As Long
Dim I1 As Long
Dim bSignFound As Boolean

    J1 = 1
    If AlwaysPositive Then MinusCount = 1
    
    If MightUseDollarSign Then
        For I1 = J1 To 1
            mStr = Mid$(pTxt, I1, 1)
            If mStr = "$" Then
                J1 = 2
                bSignFound = True
            End If
        Next
    End If
    
    For I1 = J1 To Len(pTxt)
        
        If I1 > Len(pTxt) Then Exit For
    
        mStr = Mid$(pTxt, I1, 1)
        
        If mStr = "-" Then
            
            If MinusCount > 0 Then
                RemoveChar pTxt, I1, J1
            End If
            
            Add MinusCount, 1
        
        ElseIf mStr = "." Then
            
            If PointCount > 0 Or IntegerOnly Then
                RemoveChar pTxt, I1, J1
            End If
            
            Add PointCount, 1
        
        Else
        
            mAsc = Asc(mStr)
            
            If mAsc < 48 Or mAsc > 57 Then 'non numeric
                RemoveChar pTxt, I1, J1
            Else
                MinusCount = 1
            End If
            
        End If
        
    Next
    
    If Not IsMissing(pRetVal) Then
        If bSignFound Then
            J1 = 2
        Else
            J1 = 1
        End If
        If Len(pTxt) >= J1 Then
            If IsNumeric(pTxt) Then
                pRetVal = Mid$(pTxt, J1, I1 - J1)
            End If
        End If
    End If
    
End Sub
Private Sub RemoveChar(pTxt As TextBox, pPos As Long, pStart As Long)
Dim lLen As Long
    mStrA = Mid$(pTxt, pStart, pPos - pStart)
    pStart = pPos + 1
    If pStart > Len(pTxt) Then
        mStrB = ""
    Else
        mStrB = Mid$(pTxt, pStart, Len(pTxt) - pStart + 1)
    End If
    pTxt = mStrA + mStrB
    pStart = pPos
    
    Add pPos, -1
End Sub


' == File ==
Function IsFile(strFileSpec As String) As Boolean
    If strFileSpec = "" Then Exit Function
    If Len(Dir$(strFileSpec)) > 0 Then
        IsFile = True
    Else
        IsFile = False
    End If
End Function
Function ValidFile(strFullFileSpec As String) As Boolean
Dim FS
    Set FS = CreateObject("Scripting.FileSystemObject")
    ValidFile = FS.fileexists(strFullFileSpec)
End Function

' various math
Sub Add(Varia1 As Variant, ByVal value_ As Double)
    Varia1 = Varia1 + value_
End Sub
Sub mUL(Varia1 As Variant, ByVal value_ As Double)
    Varia1 = Varia1 * value_
End Sub
Sub LinearAlg(ret_ As Single, from_!, to_!, perc_!)
    ret_ = from_ + perc_ * (to_ - from_)
End Sub
Sub TruncVar(ByRef retVal As Variant)
    retVal = retVal - Int(retVal)
End Sub
Function Triangle(ByVal In_dbl#) As Double
    Triangle = In_dbl - Int(In_dbl)
    If Triangle > 0.75 Then
        Triangle = Triangle - 1
    ElseIf Triangle > 0.25 Then
        Triangle = 0.5 - Triangle
    End If
End Function
Function LMax(ByVal sVar1 As Single, ByVal sVar2 As Single) As Long
    If sVar1 < sVar2 Then
        LMax = Int(sVar2 + 0.5)
    Else
        LMax = Int(sVar1 + 0.5)
    End If
End Function
Function LMin(ByVal sVar1 As Single, ByVal sVar2 As Single) As Long
    If sVar1 < sVar2 Then
        LMin = Int(sVar1 + 0.5)
    Else
        LMin = Int(sVar2 + 0.5)
    End If
End Function
Sub SkewCorner(pRetSX!, pRetSY!, ByVal pRad!, ByVal pAngle_0_To_1!, Optional ByVal p_rnd_quadrant_swing_mult! = 0)
    pRad = pRad * g_sk_zoom
    pAngle_0_To_1 = (g_sk_angle + pAngle_0_To_1 + p_rnd_quadrant_swing_mult * (Rnd - 0.5) * 0.25) * TwoPi
    pRetSX = pRad * Cos(pAngle_0_To_1)
    pRetSY = pRad * Sin(pAngle_0_To_1)
End Sub
Sub SngModulus(ByRef retVal As Variant, ByVal pMod As Single)
    If pMod = 0 Then Exit Sub
    retVal = retVal - pMod * Int(retVal / pMod)
End Sub
Sub RadianModulus(ByRef retAngle As Variant)
 retAngle = retAngle - TwoPi * Int(retAngle / TwoPi)
End Sub
Function TriangleModulus(ByVal in1 As Single, ByVal modulus As Single) As Single
Dim mod4!
  
    mod4 = modulus * 4
    
    'mod operation
    TriangleModulus = in1 - mod4 * Int(in1 / mod4)
    
    'triangle constraint
    If TriangleModulus > modulus * 3 Then
        TriangleModulus = TriangleModulus - mod4
    ElseIf TriangleModulus > modulus Then
        TriangleModulus = modulus * 2 - TriangleModulus
    End If
  
End Function
Function ModulusTriPositive(ByVal in1 As Single, ByVal modulus As Single) As Single
Dim mod2!
  
    mod2 = modulus * 2
    
    'mod operation
    ModulusTriPositive = in1 - mod2 * Int(in1 / mod2)
    
    If ModulusTriPositive > modulus Then ModulusTriPositive = mod2 - ModulusTriPositive
  
End Function
Function RndPosNeg() As Long
    RndPosNeg = 2 * Int(Rnd - 0.5) + 1
End Function
Function GetAngle2(ByVal sngDX!, ByVal sngDY!) As Single
 If sngDY = 0! Then
  If sngDX < 0! Then
   GetAngle2 = pi * (3& / 2&)
  ElseIf sngDX > 0 Then
   GetAngle2 = pi / 2!
  End If
 Else
  If sngDY > 0! Then
   GetAngle2 = pi - Atn(sngDX / sngDY)
  Else
   GetAngle2 = Atn(sngDX / -sngDY)
  End If
 End If
End Function
Function GetAngle(ByVal sngDX!, ByVal sngDY!) As Single
 If sngDX = 0! Then
  If sngDY < 0! Then
   GetAngle = pi * 1.5!
  ElseIf sngDY > 0 Then
   GetAngle = pi * 0.5!
  End If
 Else
  If sngDX > 0! Then
   GetAngle = Atn(sngDY / sngDX)
  Else
   GetAngle = pi - Atn(sngDY / -sngDX)
  End If
 End If
End Function
Public Sub Swap(pVar1 As Variant, pVar2 As Variant)
Dim lVar3 As Variant

    lVar3 = pVar1
    pVar1 = pVar2
    pVar2 = lVar3

End Sub



Public Sub DrawSlider(pForm As Form, pSlidr As SliderControl, Optional ByVal pNewValue As Single = -1)

    pForm.Line (pSlidr.lLeft, pSlidr.lTop)-(pSlidr.lRight, pSlidr.lBot), pSlidr.lColor, B
    If pNewValue <> -1 Then SliderVal pForm, pSlidr, pNewValue
        
    m_DrawSlidePos pForm, pSlidr
    
End Sub
Private Sub m_DrawSlidePos(pForm As Form, pSlidr As SliderControl)
    If pSlidr.Proc.lPos <> pSlidr.Proc.lPosP Then
        If pSlidr.Proc.IsVertical Then
            pForm.Line (pSlidr.lLeft + 1, pSlidr.Proc.lPosP)-(pSlidr.lRight - 1, pSlidr.Proc.lPosP), pForm.BackColor
            pForm.Line (pSlidr.lLeft + 1, pSlidr.Proc.lPos)-(pSlidr.lRight - 1, pSlidr.Proc.lPos), pForm.BackColor
        Else
            pForm.Line (pSlidr.Proc.lPosP, pSlidr.lTop + 1)-(pSlidr.Proc.lPosP, pSlidr.lBot), pForm.BackColor
            pForm.Line (pSlidr.Proc.lPos, pSlidr.lTop + 1)-(pSlidr.Proc.lPos, pSlidr.lBot), pSlidr.lColor
        End If
        pSlidr.Proc.lPosP = pSlidr.Proc.lPos
    End If
End Sub
Public Sub SliderVal(pForm As Form, pSlidr As SliderControl, ByVal pNewValue As Single)
    If pNewValue < 0 Then
        pSlidr.val_0_to_1 = 0
    ElseIf pNewValue > 1 Then
        pSlidr.val_0_to_1 = 1
    Else
        pSlidr.val_0_to_1 = pNewValue
    End If
    If pSlidr.Proc.IsVertical Then
        pSlidr.Proc.lPos = pSlidr.lBot - 1 - pSlidr.val_0_to_1 * pSlidr.Proc.lHgtM2
    Else
        pSlidr.Proc.lPos = pSlidr.Proc.lLefPlus + pSlidr.val_0_to_1 * pSlidr.Proc.lWidM2
    End If
    m_DrawSlidePos pForm, pSlidr
End Sub
Public Sub SliderDims(pSlidr As SliderControl, Optional ByVal pLeft As Integer = -1, Optional ByVal pTop As Integer = -1, Optional ByVal pWid As Integer = -1, Optional ByVal pHgt As Integer = -1, Optional ByVal IsVertical As Boolean)
    If pLeft < 0 Then pLeft = 0
    If pTop < 0 Then pTop = 0
    If pWid < 5 Then pWid = 5
    If pHgt < 5 Then pHgt = 5
    pSlidr.lLeft = pLeft
    pSlidr.lTop = pTop
    pSlidr.lBot = pTop + pHgt - 1
    pSlidr.lRight = pLeft + pWid - 1
    pSlidr.Proc.lLefPlus = pLeft + 1
    pSlidr.Proc.lRghM = pSlidr.lRight - 1
    pSlidr.Proc.lHgtM2 = pHgt - 2
'    pSlidr.Proc.lTopM = pSlidr.lTop
    pSlidr.Proc.lWidM2 = pSlidr.Proc.lRghM - pSlidr.Proc.lLefPlus
    pSlidr.Proc.IsVertical = IsVertical
End Sub
Public Function SliderFound(pSlide As SliderControl, ByVal X As Integer, ByVal Y As Integer) As Boolean
    SliderFound = False
    If X > pSlide.lLeft And X < pSlide.lRight Then
    If Y > pSlide.lTop And Y < pSlide.lBot Then
        SliderFound = True
    End If
    End If
End Function
