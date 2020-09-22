Attribute VB_Name = "mSurfaceDesc"
Option Explicit

'+------------------+------------------------------------+'
'| mSurfaceDesc.bas | developed in Visual Basic 6.0      |'
'+---------+--------+------------------------------------+'
'| Release | Public 0.1a August 14 2006 - 060814         |'
'+---------+-------+-------------------------------------+'
'| Original author | dafhi                               |'
'+-------------+---+-------------------------------------+'
'| Description | SurfaceDescriptor UDT is a wrapper for  |'
'+-------------+ graphics array processing.              |'
'|                                                       |'
'| Blit() sub uses StretchDiBits which is generally the  |'
'| fastest method overall for processing on arrays       |'
'|                                                       |'
'+--------- Dependencies ------------+                   |'
'| mGeneral.bas                      |                   |'
'| -> FileDlg2.cls                   |                   |'
'|                                   |                   |'
'+------------------------------+----+-------------------+'
'| Contributors / Modifications |                        |'
'+------------------------------+                        |'
'|                                                       |'
'|                                                       |'
'+-------------------------------------------------------+'

' =============== Changes & Fixes ================

' - Sept 19, 2006 -
' + CreateSurfaceDesc changes "dimension" type member values
' for invalid x or y dim size request
' + removed "If ySrcE > SrcMaxY Then Exit For" in MaskBlit()

' - Sept 7, 2006 -
' + Added DescProc type to SurfaceDescriptor
' + Changed U1D member math and moved it to DescProc
' + Renamed PelCount to TotalPixels

' - Aug 23, 2006 -

' + Renamed ScanRgns to SectDelt
' + Changed Wide and High members from Long (32 bit) to Integer (16 bit)
' + Added MaskColor member

' - Aug 16, 2006 -

' + Removed one UDT for simplification
' + Simplified RunLen_Encode and MaskBlit (should be unnoticeably faster ;)

' - Aug 14, 2006 -

' + Renamed CreateMaskStructure to RunLen_Encode
' + Added ability to change safearray lowbound
' in HookRGBQ_Begin Hook1D_Begin

' ================ How to use   ==================

' 1. CreateSurfaceDesc or SurfaceDescFromFile to initialize surface

' 2. Blit() will blit to DC specified by MySurface.ToDC (you can set it in the initialization subs)

' 3. Graphics write examples TestSurfaceDesc() and ColorFill()

' 4. RunLen_Encode() will, given MaskColor > -1 (and filled image)
' create the structure necessary for the following ..

' 5. MaskBlit() is intended for sprites, and blits from one surface
' to another

' ...

Dim mSA1      As SAFEARRAY1D 'used from SurfaceDescToFile

Private Type StartAndFin
    Start As Integer   'x or y start of run length
    Delta As Integer   'delta between start and end, or length - 1
End Type

Private Type MaskProc
    ySegs      As Integer
    not_used   As Integer
    vRun()     As StartAndFin 'DeltaSE
    SectDelt() As Long
    hRun()     As StartAndFin
End Type

Private Type DescProc
    LowXM       As Long
    LowYM       As Long
    HighXP      As Long
    HighYP      As Long
    Low1D       As Long
    High1D      As Long
End Type

Type SurfaceDescriptor
    Proc        As DescProc
    ToDC        As Long    'helpful
    Wide        As Integer
    High        As Integer
    WM          As Integer 'For X = [0 To mySD.WM] or ..
                           'For X = [mySD.LowX To mySD.LowX + mySD.WM]
    HM          As Integer
    LowX        As Integer 'lowbound
    LowY        As Integer
    HighX       As Integer 'highbound
    HighY       As Integer
    TotalPixels As Long
    U1D         As Long    'helpful: ubound for safearray 1d creation
    BIH         As BITMAPINFOHEADER
    Dib32()     As Long
    MaskInfo    As MaskProc
    MaskColor   As Long
End Type

Public gSurf  As SurfaceDescriptor


'''''''''''''''''''''''''''
'                         '
'   Run-Length Blit       '
'                         '
'''''''''''''''''''''''''''

Public Sub RunLen_Encode(pSurf As SurfaceDescriptor, Optional ByVal MaskColor As Long = -1)
Dim lx        As Integer
Dim ly        As Integer
Dim IsBlit    As Boolean
Dim IsBlitP   As Boolean
Dim ScBlit    As Boolean
Dim ScBlitP   As Boolean
Dim cRgn      As Long
Dim BlitLenM  As Long
Dim vRgnPtr   As Long
Dim DimMode   As Long
Dim cRgnP     As Long
Dim vLen      As Long
Dim ScanPtr   As Long

    'SurfaceDescFromFile or CreateSurfaceDesc first!

    If pSurf.TotalPixels < 1 Then Exit Sub
    
    pSurf.MaskColor = MaskColor
    
    For DimMode = 0 To 1
    
        For ly = pSurf.LowY To pSurf.LowY + pSurf.HM
            BlitLenM = 0
            For lx = pSurf.LowX To pSurf.LowX + pSurf.WM
                IsBlit = pSurf.Dib32(lx, ly) <> MaskColor
                If IsBlit Xor IsBlitP Then
                    If IsBlit Then 'wasn't blit, now is
                        If DimMode = 1 Then
                            pSurf.MaskInfo.hRun(cRgn).Start = lx
                        End If
                    Else 'was blit, now not
                        If DimMode = 1 Then
                            pSurf.MaskInfo.hRun(cRgn).Delta = BlitLenM
                        End If
                        BlitLenM = 0
                        Add cRgn, 1
                    End If
                ElseIf IsBlit Then
                    Add BlitLenM, 1
                End If
                IsBlitP = IsBlit
            Next
            IsBlitP = False
            
            If IsBlit Then
                If DimMode = 1 Then
                    pSurf.MaskInfo.hRun(cRgn).Delta = BlitLenM
                End If
                Add cRgn, 1
            End If
            
            ScBlit = (cRgn - cRgnP) > 0
            If ScBlit Xor ScBlitP Then
                If ScBlit Then 'wasn't, now is
                    Add vRgnPtr, 1
                    If DimMode = 1 Then
                        pSurf.MaskInfo.vRun(vRgnPtr).Start = ly
                    End If
                    vLen = 0
                Else 'was, now isn't
                    If DimMode = 1 Then
                        pSurf.MaskInfo.vRun(vRgnPtr).Delta = vLen - 1
                    End If
                End If
            End If
            
            If ScBlit Then
                If DimMode = 1 Then
                    pSurf.MaskInfo.SectDelt(ScanPtr) = cRgn - 1 - cRgnP
                End If
                Add ScanPtr, 1
                cRgnP = cRgn
            End If
            
            Add vLen, 1
            ScBlitP = ScBlit
        
        Next
         
        If vRgnPtr > 0 Then
            If DimMode = 0 Then
                Erase pSurf.MaskInfo.vRun
                ReDim pSurf.MaskInfo.vRun(1 To vRgnPtr)
                Erase pSurf.MaskInfo.SectDelt
                ReDim pSurf.MaskInfo.SectDelt(ScanPtr - 1)
                pSurf.MaskInfo.ySegs = vRgnPtr
            ElseIf ScBlit Then
                pSurf.MaskInfo.vRun(vRgnPtr).Delta = vLen - 1
                vLen = 0
            End If
        End If
        
        If cRgn > 0 Then
            If DimMode = 0 Then
                Erase pSurf.MaskInfo.hRun
                ReDim pSurf.MaskInfo.hRun(0 To cRgn - 1)
            End If
            cRgn = 0
            cRgnP = 0
        End If
    
        ScBlitP = False
        IsBlit = False
        vRgnPtr = 0
        ScanPtr = 0
        
    Next

End Sub
Public Sub MaskBlit(pDest As SurfaceDescriptor, pSrc As SurfaceDescriptor, Optional ByVal pX As Single, Optional ByVal pY As Single)
Dim yPtr     As Long
Dim LenRef   As Long
Dim hPtrE    As Long
Dim hPtrS    As Long

Dim ySrcS    As Integer
Dim ySrcE    As Integer
Dim xSrcS    As Integer
Dim xSrcE    As Integer

Dim ySrcE2   As Integer
Dim xSrcE2   As Integer

Dim lTmp__   As Integer
Dim SrcBotM1 As Integer

Dim DestLeft As Integer
Dim DestBot  As Integer
Dim yDst     As Integer
Dim ySrcEP   As Integer

Dim SrcMinY  As Integer
Dim SrcMinX  As Integer
Dim SrcMaxY  As Integer
Dim SrcMaxX  As Integer

    'RunLen_Encode() contains the encode source
    
    z_GetClipRgn DestLeft, SrcMinX, SrcMaxX, pSrc.WM, pSrc.LowX, pDest.LowX, pDest.LowX + pDest.WM, pX
    z_GetClipRgn DestBot, SrcMinY, SrcMaxY, pSrc.HM, pSrc.LowY, pDest.LowY, pDest.LowY + pDest.HM, pY
    
    SrcBotM1 = SrcMinY - 1
    
    For yPtr = 1 To pSrc.MaskInfo.ySegs
    
        'vertical contiguous chunk of scanlines that have data
        ySrcS = pSrc.MaskInfo.vRun(yPtr).Start
        ySrcE = ySrcS + pSrc.MaskInfo.vRun(yPtr).Delta
        
        'z_GetClipRgn computes MaxY, MinY, etc.
        If ySrcE > SrcMaxY Then
            ySrcE2 = SrcMaxY
        Else
            ySrcE2 = ySrcE
        End If
        
        For ySrcS = ySrcS To ySrcE2 'vertical run length
        
            'with new scanline we have this recomputation
            hPtrE = hPtrS + pSrc.MaskInfo.SectDelt(LenRef)
            
            If ySrcS > SrcBotM1 Then

                yDst = ySrcS + DestBot
                
                For hPtrS = hPtrS To hPtrE
                    
                    xSrcS = pSrc.MaskInfo.hRun(hPtrS).Start
                    xSrcE = xSrcS + pSrc.MaskInfo.hRun(hPtrS).Delta
                    
                    If xSrcS < SrcMinX Then xSrcS = SrcMinX
                    
                    If xSrcE > SrcMaxX Then xSrcE = SrcMaxX
    
                    For xSrcS = xSrcS To xSrcE
                        pDest.Dib32(xSrcS + DestLeft, yDst) = pSrc.Dib32(xSrcS, ySrcS)
                    Next
                    
                Next
            
            End If
            
            LenRef = LenRef + 1
            hPtrS = hPtrE + 1
        
        Next
        
    Next

End Sub
Private Sub z_GetClipRgn(RetDest As Integer, RetSrcMin As Integer, RetSrcMax As Integer, ByVal pSrcM1 As Integer, ByVal pSrcLow As Integer, ByVal pDestLow As Integer, ByVal pDestHigh As Integer, pVal As Single)

    RetDest = Int(pVal + 0.5) 'round
    
    RetSrcMax = pSrcLow + pSrcM1
    
    If RetDest + pSrcM1 > pDestHigh Then
        RetSrcMax = RetSrcMax - (RetDest + pSrcM1 - pDestHigh)
    End If
    
    RetSrcMin = pSrcLow
    If RetDest < pDestLow Then
        RetSrcMin = RetSrcMin + pDestLow - RetDest
    End If
    
    RetDest = RetDest - pSrcLow
    
End Sub


'''''''''''''''''''''''''''
'                         '
'   Example array write   '
'                         '
'''''''''''''''''''''''''''

Sub TestSurfaceDesc(pSDESC As SurfaceDescriptor)
Dim LJ As Long, lI As Long

    If pSDESC.TotalPixels < 1 Then Exit Sub

    For LJ = pSDESC.LowY To pSDESC.LowY + pSDESC.HM
        For lI = pSDESC.LowX To pSDESC.LowX + pSDESC.WM
            pSDESC.Dib32(lI, LJ) = ARGBHSV(255, Rnd, Rnd * 255)
        Next
    Next
    
    Blit pSDESC
    
End Sub

Public Sub ColorFill(Surf As SurfaceDescriptor, pAry() As Long, pSA As SAFEARRAY1D, Optional ByVal pColor As Long = 0, Optional ByVal pLowBound As Long = 0)
Dim L1 As Long

    'how to use hook subs
    
    Hook1D_Begin Surf, pAry, pSA, pLowBound
    
    For L1 = pLowBound To pLowBound + Surf.U1D
        pAry(L1) = pColor 'now have 1d access!
    Next
    
    Hook1D_End pAry

End Sub


'''''''''''''''''''''''''''
'                         '
'   Blit                  '
'                         '
'''''''''''''''''''''''''''

Sub Blit(pSD As SurfaceDescriptor, Optional ByVal pX As Integer, Optional ByVal pY As Integer, Optional ByVal pWid As Integer = -1, Optional ByVal pHgt As Integer = -1)

    If pSD.TotalPixels < 1 Then Exit Sub

    If pWid < 0 Then pWid = pSD.Wide
    If pHgt < 0 Then pHgt = pSD.High
    
    StretchDIBits pSD.ToDC, _
      pX, pY, pWid, pHgt, _
      0, 0, pSD.Wide, pSD.High, _
      pSD.Dib32(pSD.LowX, pSD.LowY), pSD.BIH, DIB_RGB_COLORS, vbSrcCopy

End Sub


'''''''''''''''''''''''''''
'                         '
'   Create Surface        '
'                         '
'''''''''''''''''''''''''''

Sub CreateSurfaceDesc(SDesc1 As SurfaceDescriptor, lHDC As Long, ByVal Wide As Long, ByVal High As Long, Optional ByVal LowX As Integer = 0, Optional ByVal LowY As Integer = 0)

    'Example: CreateSurfaceDesc mySD, mySD.Dib32, Picture1.hDC, 640, 480, 1, 1

    If Wide = SDesc1.Wide And High = SDesc1.High Then Exit Sub
    
    SDesc1.High = High
    SDesc1.Wide = Wide
    SDesc1.WM = SDesc1.Wide - 1
    SDesc1.HM = SDesc1.High - 1
    SDesc1.TotalPixels = Wide * High
    SDesc1.U1D = SDesc1.TotalPixels - 1
    SDesc1.LowX = LowX
    SDesc1.LowY = LowY
    SDesc1.Proc.LowXM = LowX - 1
    SDesc1.Proc.LowYM = LowY - 1
    SDesc1.Proc.HighXP = LowX + Wide
    SDesc1.Proc.HighYP = LowY + High
    SDesc1.Proc.Low1D = CLng(LowY) * Wide + CLng(LowX)
    SDesc1.Proc.High1D = SDesc1.Proc.Low1D + SDesc1.U1D
    SDesc1.HighX = LowX + SDesc1.WM
    SDesc1.HighY = LowY + SDesc1.HM
    
    If Wide * High < 1 Or Wide * High > 10000000 Then Exit Sub
    SDesc1.ToDC = lHDC
    SDesc1.BIH.biHeight = High
    SDesc1.BIH.biWidth = Wide
    SDesc1.BIH.biPlanes = 1
    SDesc1.BIH.biBitCount = 32
    SDesc1.BIH.biSize = Len(SDesc1.BIH)
    SDesc1.BIH.biSizeImage = 4 * SDesc1.TotalPixels
    SDesc1.BIH.biCompression = BI_RGB
    Erase SDesc1.Dib32
    ReDim SDesc1.Dib32(LowX To LowX + SDesc1.WM, LowY To LowY + SDesc1.HM)

End Sub

Public Sub Surface_OnResize(pSurf As SurfaceDescriptor, pPic As Picture, Optional pDC As Long)
    If pPic.Height < 1 Or pPic.Width < 1 Then Exit Sub
    CreateSurfaceDesc pSurf, pDC, pPic.Width, pPic.Height
End Sub


'''''''''''''''''''''''''''
'                         '
'   Load                  '
'                         '
'''''''''''''''''''''''''''

Public Function SurfaceDescFromFile(Surf As SurfaceDescriptor, strFileName$, Optional ByVal pHDC As Long, Optional ByVal MaskColor As Long = -1, Optional ByVal pLowX As Integer = 0, Optional ByVal pLowY As Integer = 0, Optional ByVal StrFolder$ = "") As String
Dim tBM As Bitmap
Dim CDC&, lStrFile As String, tBI As BITMAPINFO
Dim lStrFileFolder As String
Dim L1 As Long

Dim ThePic As StdPicture
   
   z_SurfaceDescFileCommon lStrFileFolder, StrFolder, lStrFile, strFileName

   On Local Error GoTo FileError
   Set ThePic = LoadPicture(lStrFile) 'this will crash with invalid pictures
   
   If GetObject(ThePic, Len(tBM), tBM) = 0 Then
      Set ThePic = Nothing
      MsgBox "FILE ERROR"
      Exit Function
   End If

   CreateSurfaceDesc Surf, pHDC, tBM.bmWidth, tBM.bmHeight, pLowX, pLowY
    
   CDC = CreateCompatibleDC(0)           ' Temporary device

   DeleteObject SelectObject(CDC, ThePic)  ' Converted bitmap
   
   tBI.bmiHeader.biSize = 40
   Call GetDIBits(CDC, ThePic.Handle, 0, 0, ByVal 0&, tBI, 0)
   tBI.bmiHeader.biBitCount = 32
   
   L1 = GetDIBits(CDC, ThePic.Handle, 0, Surf.High, Surf.Dib32(pLowX, pLowY), tBI, 0)
   
   If MaskColor > -1 Then
       RunLen_Encode Surf, MaskColor
   End If
   
   DeleteDC CDC
   Set ThePic = Nothing
   
   If L1 = 0 Then
      MsgBox "DIB ERROR"
      Exit Function
   End If

    SurfaceDescFromFile = "Success!"
   Exit Function
    
FileError:
      MsgBox "FILE ERROR"

End Function
Public Function FileToSurfaceDesc(Surf As SurfaceDescriptor, strFileName$, Optional ByVal pHDC As Long, Optional ByVal MaskColor As Long = -1, Optional ByVal pLowX As Integer = 0, Optional ByVal pLowY As Integer = 0, Optional ByVal StrFolder$ = "") As String
    FileToSurfaceDesc = SurfaceDescFromFile(Surf, strFileName, pHDC, MaskColor, pLowX, pLowY, StrFolder)
End Function


'''''''''''''''''''''''''''
'                         '
'   Save                  '
'                         '
'''''''''''''''''''''''''''

Public Function SurfaceDescToFile(Surf As SurfaceDescriptor, strFileName$, Optional ByVal StrFolder$ = "") As String
Dim tBMF As BitmapFileHeader
Dim lStrFile As String, tBIH As BITMAPINFOHEADER
Dim lStrFileFolder As String
Dim L1 As Long, L2 As Long, L3 As Long, FFile As Integer
Dim PadBytes As Long, Bytes() As Byte
Dim tRGBQ() As RGBQUAD
    
    If Surf.TotalPixels < 1 Then
        SurfaceDescToFile = "Width or Height < 1!"
        Exit Function
    End If

    PadBytes = GetPadBytes(Surf.Wide, 3)
    
    tBIH = Surf.BIH
    
    tBIH.biBitCount = 24
    tBIH.biSizeImage = (Surf.Wide * tBIH.biBitCount / 8 + PadBytes) * Surf.High
    
    tBMF.bfType = BMPFileSignature
    tBMF.bfOffBits = Len(tBMF) + Len(tBIH)
    tBMF.bfSize = tBMF.bfOffBits + tBIH.biSizeImage

    z_SurfaceDescFileCommon lStrFileFolder, StrFolder, lStrFile, strFileName
    
    If Right$(lStrFile, 4) <> ".bmp" Then
        lStrFile = lStrFile & ".bmp"
    End If
    
    FFile = FreeFile
    
    Open lStrFile For Output As #FFile
        Write #FFile, "" 'erase existing data
    Close #FFile
    
    ReDim Bytes(Surf.Wide * 3 + PadBytes - 1)
    
    HookRGBQ_Begin Surf, tRGBQ, mSA1 'pointing RGB Quad to Surf.Dib32 array
     
    Open lStrFile For Binary As #FFile
    
    Seek #FFile, 1
    
    Put #FFile, , tBMF
    Put #FFile, , tBIH
    
    For L2 = 0 To Surf.HM
        For L1 = 0 To Surf.WM * 3 Step 3
            Bytes(L1) = tRGBQ(L3).Blue
            Bytes(L1 + 1) = tRGBQ(L3).Green
            Bytes(L1 + 2) = tRGBQ(L3).Red
            Add L3, 1
        Next
        Put #FFile, , Bytes
    Next
    
    Close #FFile
    
    HookRGBQ_End tRGBQ
    
    SurfaceDescToFile = "Success!"

End Function
Public Function FileFromSurfaceDesc(Surf As SurfaceDescriptor, strFileName$, Optional ByVal pHDC As Long, Optional ByVal StrFolder$ = "") As String
    FileFromSurfaceDesc = SurfaceDescToFile(Surf, strFileName, StrFolder)
End Function

Private Sub z_SurfaceDescFileCommon(pStrFileFolder As String, pStrFolder As String, pStrFile As String, strFileName As String)

    If pStrFolder <> "" Then
     pStrFileFolder = pStrFolder
    End If
    
    If Right$(pStrFileFolder, 1) = "\" Or pStrFileFolder = "" Then
        pStrFile = pStrFileFolder & strFileName
    Else
        pStrFile = pStrFileFolder & "\" & strFileName
    End If

End Sub


'''''''''''''''''''''''''''
'                         '
'   SafeArray1D Wraps     '
'                         '
'''''''''''''''''''''''''''

Sub HookRGBQ_Begin(pSource As SurfaceDescriptor, pAryFor1D() As RGBQUAD, pSA1D As SAFEARRAY1D, Optional ByVal pLowBound As Long = 0)

    'See: SurfaceDescToFile for why this is used

    If pSource.BIH.biSizeImage < 1 Then Exit Sub

    z_HookBegin_Common pSource, pSA1D, pLowBound
    
    CopyMemory ByVal VarPtrArray(pAryFor1D), VarPtr(pSA1D), 4
    
End Sub
Sub HookRGBQ_End(pAry() As RGBQUAD)
    CopyMemory ByVal VarPtrArray(pAry), 0&, 4
End Sub

Sub Hook1D_Begin(pSource As SurfaceDescriptor, pAryFor1D() As Long, pSA1D As SAFEARRAY1D, Optional ByVal pLowBound As Long = 0)

    If pSource.BIH.biSizeImage < 1 Then Exit Sub

    z_HookBegin_Common pSource, pSA1D, pLowBound
    
    CopyMemory ByVal VarPtrArray(pAryFor1D), VarPtr(pSA1D), 4
    
End Sub
Sub Hook1D_End(pAry() As Long)
    CopyMemory ByVal VarPtrArray(pAry), 0&, 4
End Sub
Private Sub z_HookBegin_Common(pSource As SurfaceDescriptor, pSA1D As SAFEARRAY1D, pLowBound As Long)

    If pSource.TotalPixels < 1 Then Exit Sub
    
    pSA1D.lLbound = pLowBound
    pSA1D.cbElements = 4
    pSA1D.cElements = pSource.TotalPixels
    pSA1D.cDims = 1
    pSA1D.pvData = VarPtr(pSource.Dib32(pSource.LowX, pSource.LowY))
    
End Sub


'''''''''''''''''''''''''''
'                         '
'   Reference             '
'                         '
'''''''''''''''''''''''''''

Private Function z_SurfaceDescFromFile_UsingGetPixel(Surf As SurfaceDescriptor, strFileName$, Optional ByVal pHDC As Long, Optional ByVal StrFolder$ = "") As String
Dim tBM As Bitmap, sPic As IPictureDisp
Dim CDC&, lStrFile As String, tBI As BITMAPINFO
Dim lStrFileFolder As String
Dim L1 As Long, L2 As Long, L3 As Long

    'This sub totally works except for invalid picture data
    'where LoadPicture() crashes

    On Local Error Resume Next
    
    z_SurfaceDescFileCommon lStrFileFolder, StrFolder, lStrFile, strFileName

    If IsFile(lStrFile) Then
        Set sPic = New StdPicture
        Set sPic = LoadPicture(lStrFile)
    Else
        z_SurfaceDescFromFile_UsingGetPixel = "File not found"
        Exit Function
    End If
    
    
    If sPic = vbEmpty Then
        z_SurfaceDescFromFile_UsingGetPixel = "Invalid image data"
        GoTo OOPS
    End If
        
    CDC = CreateCompatibleDC(0)           ' Temporary device

    DeleteObject SelectObject(CDC, sPic)  ' Converted bitmap
    
    GetObjectAPI sPic, Len(tBM), tBM
    
    CreateSurfaceDesc Surf, pHDC, tBM.bmWidth, tBM.bmHeight
    
    tBI.bmiHeader = Surf.BIH

    For L2 = 0 To Surf.HM
        L3 = Surf.HM - L2
        For L1 = 0 To Surf.WM
            Surf.Dib32(L1, L2) = FlipRB(GetPixel(CDC, L1, L3))
        Next
    Next
    
    z_SurfaceDescFromFile_UsingGetPixel = "Success!"

    Set sPic = Nothing
    
OOPS:
 
    DeleteDC CDC
    
End Function
