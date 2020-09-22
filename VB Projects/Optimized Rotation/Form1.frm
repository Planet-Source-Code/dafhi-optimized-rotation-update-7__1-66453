VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   1800
      Top             =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Optimized Rotation (Update 7) by dafhi  Oct 7, 2006

' Lots of comments + diagram.gif

' BASICS
'1. Put some points in 3-D coordinate space
'2. Position and Rotate related 3-D axis
'3. Subroutine ProjectionTransform() has the projective transformation algorithm

'Press Spacebar to switch to stereoscopic view!
'Esc - Exit
'All other keys do the same action - move points into randomized position

Private Type Point3D
    X          As Single
    Y          As Single
    Z          As Single
    hue        As Single
    satur      As Single
End Type

Private Type AxisElement
    Origin     As Point3D
    Projection As Point3D
End Type

Private Type Obj3D
    Center     As AxisElement
    X_AXIS     As AxisElement
    Y_AXIS     As AxisElement
    Z_AXIS     As AxisElement
    PT()       As Point3D
    PtCount    As Long
End Type

Dim MyScene         As Obj3D

Dim mSurfaceDesc    As SurfaceDescriptor

Dim mSA             As SAFEARRAY1D 'for access of mSurfaceDesc.Dib32(2d array)) into 1d
Dim mImage1D()      As Long

Dim m_angleX        As Single  'animate the whole object or scene
Dim m_angleY        As Single
Dim m_angleZ        As Single
Dim mRotSpeed       As Point3D

Dim mDepthAry()     As Single  'stack for pixel z-buf and erase
Dim mDepthAryRef()  As Long
Dim mDepthRefHgt    As Long
Dim mConversion1D   As Long

Dim mWid            As Long    'simple names for window dims
Dim mHgt            As Long

Dim mWidBy2         As Integer 'stereo projection variables
Dim mRtEyeX         As Integer
Dim mStereoDisplacm As Long

Dim mHueBase        As Single  'it's subtle .. hue does change :)
Dim mBoolStereo     As Boolean

Private Sub Form_Load()

    Randomize
    mHueBase = -50
    
    ScaleMode = vbPixels
    Move 200, 200
    
    Timer1.Interval = 5
    Timer1.Enabled = True
    
    O3D_Init MyScene, 12000, 400  'dimension some points and scale axis
    ArrangeObject MyScene

    mRotSpeed.X = 0.007
    mRotSpeed.Y = 0.013
    mRotSpeed.Z = 0.022
    
    m_angleX = Rnd * TwoPi 'begin at random angle
    m_angleZ = Rnd * TwoPi

    FPS_Init
    
End Sub
Private Sub Form_Resize()

    mWid = ScaleWidth
    mHgt = ScaleHeight
    
    CreateSurfaceDesc mSurfaceDesc, hDC, mWid, mHgt, Int(-mWid / 2), Int(-mHgt / 2)
    
    If mWid < 1 Or mHgt < 1 Then Exit Sub
    
    Erase mDepthAry 'store pixel depth, so don't draw over front pixels
    ReDim mDepthAry(mSurfaceDesc.Proc.Low1D To mSurfaceDesc.Proc.High1D)
    
    'small amount of plots compared with total image size,
    'quicker to erase using this reference
    Erase mDepthAryRef
    ReDim mDepthAryRef(1 To mWid * mHgt)
    
    mWidBy2 = Int(0.5 * (mWid + 1)) 'stereoscopic image split
    mRtEyeX = mWidBy2
    mStereoDisplacm = Int(0.5 * (mRtEyeX + 1))

End Sub

Private Sub Timer1_Timer()

    If mBoolStereo Then
        DrawPoints mSurfaceDesc, MyScene, m_angleX, m_angleY, _
         m_angleZ, -mStereoDisplacm, , , , , , mWidBy2
        DrawPoints mSurfaceDesc, MyScene, m_angleX, m_angleY - pi * 0.02, _
         m_angleZ, mStereoDisplacm, , , , mRtEyeX
    Else
        DrawPoints mSurfaceDesc, MyScene, m_angleX, m_angleY, m_angleZ
    End If
    
    Blit mSurfaceDesc
    
    'clear mImage1D (mSurfaceDesc.Dib32() accessed as 1d) and mDepthAry()
    Hook1D_Begin mSurfaceDesc, mImage1D, mSA, mSurfaceDesc.Proc.Low1D
    For mDepthRefHgt = mDepthRefHgt To 1 Step -1
        mConversion1D = mDepthAryRef(mDepthRefHgt)
        mDepthAry(mConversion1D) = 0
        mImage1D(mConversion1D) = vbBlack
    Next
    Hook1D_End mImage1D
    
    If CheckFPS(, 0.018) Then
        Caption = "FPS: " & Round(sFPS, 1)
    End If

    Add m_angleX, mRotSpeed.X * speed
    Add m_angleY, mRotSpeed.Y * speed
    Add m_angleZ, mRotSpeed.Z * speed
    
    Add mHueBase, speed
    SngModulus mHueBase, 1530 'hue on a leash
    
End Sub

Private Sub DrawPoints(SDESC As SurfaceDescriptor, O3D As Obj3D, _
  angleX As Single, angleY As Single, angleZ As Single, _
  Optional ByVal TransX As Single, _
  Optional ByVal TransY As Single, _
  Optional ByVal TransZ As Single = 3, _
  Optional ByVal Scalar As Single = 1, _
  Optional ByVal BlitX As Integer, Optional ByVal BlitY As Integer, _
  Optional ByVal pWid As Integer, Optional ByVal pHgt As Integer, _
  Optional ByVal TransZMultByScalar As Boolean = True)
  
Dim lDep As Single, lI As Long
Dim BoundX1 As Integer, BoundX2 As Integer
Dim BoundY1 As Integer, BoundY2 As Integer
Dim lP3D As Point3D

    RotateAxis O3D, angleX, angleY, angleZ 'non-destructive

    If Scalar = 0 Then Scalar = 1
    
    Add TransX, O3D.Center.Origin.X + 0.5 'user adjust + object pos all done here
    Add TransY, O3D.Center.Origin.Y + 0.5 'Int(.. + .5) for proper rounding operation in the loop .. convert floating point to integer x and y
    
    Scalar = Scalar * O3D.X_AXIS.Origin.X
    
    If TransZMultByScalar Then
        Add O3D.Center.Projection.Z, TransZ * Scalar * 0.55
    Else
        Add O3D.Center.Projection.Z, TransZ * 0.55
    End If
    
    'view rect
    z_DrawPoints_GetClip BoundX1, BoundX2, BlitX, pWid, SDESC.Wide + 1, SDESC.Proc.LowXM, SDESC.Proc.HighXP
    z_DrawPoints_GetClip BoundY1, BoundY2, BlitY, pHgt, SDESC.High + 1, SDESC.Proc.LowYM, SDESC.Proc.HighYP
    
    'Access SDESC.Dib32(2d array) as mImage1D()
    Hook1D_Begin SDESC, mImage1D, mSA, SDESC.Proc.Low1D
    
    For lI = 1 To O3D.PtCount
    
        'O3D.PT() is the object point, transform doesn't destroy
        ProjectionTransform lP3D, O3D.PT(lI), O3D.Center.Projection, _
          O3D.X_AXIS.Projection, _
          O3D.Y_AXIS.Projection, _
          O3D.Z_AXIS.Projection
          
        If lP3D.Z > 0.001 Then 'depth past screen .. prevent div by zero
        
            ''pinhole depth distort and point scaling
            lDep = Scalar / Sqr(lP3D.X * lP3D.X + lP3D.Y * lP3D.Y + lP3D.Z * lP3D.Z)
            
            ''flatscreen depth distort and point scaling
            'lDep = Scalar / lP3D.Z
            
            lP3D.X = Int(lP3D.X * lDep + TransX) 'X POS
            
            If lP3D.X > BoundX1 And lP3D.X < BoundX2 Then
            lP3D.Y = Int(lP3D.Y * lDep + TransY) 'Y POS
            
            If lP3D.Y > BoundY1 And lP3D.Y < BoundY2 Then
            mConversion1D = mWid * lP3D.Y + lP3D.X
            
            'only draw if "pixel blank" or ..
            If mDepthAry(mConversion1D) = 0 Then
            
                mDepthRefHgt = mDepthRefHgt + 1
                mDepthAryRef(mDepthRefHgt) = mConversion1D
                
                mDepthAry(mConversion1D) = lP3D.Z
                
                If lDep > 1 Then lDep = 1
                
                mImage1D(mConversion1D) = ARGBHSV(mHueBase + O3D.PT(lI).hue, O3D.PT(lI).satur, lDep * 255)
                
            '.. dot is "in front" of one already at this pixel
            ElseIf lP3D.Z < mDepthAry(mConversion1D) Then
            
                mDepthAry(mConversion1D) = lP3D.Z
                
                If lDep > 1 Then lDep = 1
                
                mImage1D(mConversion1D) = ARGBHSV(mHueBase + O3D.PT(lI).hue, O3D.PT(lI).satur, lDep * 255)
                
            End If 'z < depthary or depthary = 0
            End If 'Y > -1 and < mHgt
            End If 'X > -1 and < mWid
        End If 'Z > .001
    Next
    
    Hook1D_End mImage1D
    
End Sub
Private Sub ProjectionTransform(RetProjection As Point3D, Origin As Point3D, Vertex As Point3D, pX_AXIS As Point3D, pY_AXIS As Point3D, pZ_AXIS As Point3D)
''''''''''''''''''''''''''''''''''''''''''''''''
' This is a vectorial addition and multiplication
' algorithm.
'
' An unrotated 3D axis:
'
' Y Vector (0,1,0)
' |
' |
' | Z Vector (0,0,1)
' +------- X Vector (1,0,0)
' Vertex (0,0,0)
'
' When you know the x,y,z for each of the four
' axis parts, it is easy to compute by hand,
' a point projection associated with the axis.
'
' Consider a point (0.5,0.25,1.0)
' What you do is start at the axis vertex,
' (0,0,0) + 0.5 [point x] * (1,0,0) = (0.5,0,0)
'
' From there, go 0.25 * Axis Y Vector:
' (0.5,0,0) + 0.25 [point y] * (0,1,0) = (0.5,0.25,0)
'
' Lastly,
' (0.5,0.25,0) + 1.0 [point z] * (0,0,1) = (0.5,0.25,1)

    RetProjection.X = Origin.X * pX_AXIS.X + Origin.Y * pY_AXIS.X + Origin.Z * pZ_AXIS.X
    RetProjection.Y = Origin.X * pX_AXIS.Y + Origin.Y * pY_AXIS.Y + Origin.Z * pZ_AXIS.Y
    RetProjection.Z = Origin.X * pX_AXIS.Z + Origin.Y * pY_AXIS.Z + Origin.Z * pZ_AXIS.Z + Vertex.Z

' You may notice I haven't used Vertex X and Y.
' They are combined with user-spec x and y translation
' and depth distort calculated after the call to this sub.

End Sub
Private Sub RotateAxis(O3D As Obj3D, Optional ByVal angleX As Single, Optional ByVal angleY As Single, Optional ByVal angleZ As Single)
    RotatePoint_Test O3D.Center.Projection, angleX, angleY, angleZ, O3D.Center.Origin
    RotatePoint_Test O3D.X_AXIS.Projection, angleX, angleY, angleZ, O3D.X_AXIS.Origin
    RotatePoint_Test O3D.Z_AXIS.Projection, angleX, angleY, angleZ, O3D.Z_AXIS.Origin
    RotatePoint_Test O3D.Y_AXIS.Projection, angleX, angleY, angleZ, O3D.Y_AXIS.Origin
End Sub
Private Sub z_DrawPoints_GetClip(RetLimLo As Integer, RetLimHi As Integer, InPOS As Integer, ByVal InDIM As Integer, ByVal Surf_DimPLUS1 As Integer, ByVal Surf_LimLoMINUS1 As Integer, ByVal Surf_LimHiPLUS1 As Long)

    RetLimLo = InPOS + Surf_LimLoMINUS1 'convert pos traditional (low bound = 0) to pos low bound = surf low bound
    
    If InDIM > 0 Then
        Surf_DimPLUS1 = InDIM + 1
    End If
    
    RetLimHi = RetLimLo + Surf_DimPLUS1
    
    If RetLimLo < Surf_LimLoMINUS1 Then RetLimLo = Surf_LimLoMINUS1
    If RetLimHi > Surf_LimHiPLUS1 Then RetLimHi = Surf_LimHiPLUS1

End Sub


Private Sub RotatePoint_Test(P3D As Point3D, angleX As Single, angleY As Single, angleZ As Single, P3D_SRC As Point3D)
Dim l_tmp As Single

    'changed so that y rotates last for stereoscopy
    
    l_tmp = P3D_SRC.X * Cos(angleZ) - P3D_SRC.Y * Sin(angleZ)
    P3D.Y = P3D_SRC.Y * Cos(angleZ) + P3D_SRC.X * Sin(angleZ)
    P3D.X = l_tmp

    l_tmp = P3D_SRC.Z * Cos(angleX) - P3D.Y * Sin(angleX)
    P3D.Y = P3D.Y * Cos(angleX) + P3D_SRC.Z * Sin(angleX)
    P3D.Z = l_tmp

    l_tmp = P3D.X * Cos(angleY) - P3D.Z * Sin(angleY)
    P3D.Z = P3D.Z * Cos(angleY) + P3D.X * Sin(angleY)
    P3D.X = l_tmp
    
End Sub
Private Sub ArrangeObject(O3D As Obj3D, _
 Optional ByVal NumPoints As Long = -1, _
 Optional ByVal centerX As Single, _
 Optional ByVal centerY As Single, Optional ByVal centerZ As Single, _
 Optional ByVal AryStart As Long = 1, _
 Optional ByVal ReScale As Single = 0.5, _
 Optional ByVal angleX As Single, _
 Optional ByVal angleY As Single, Optional ByVal angleZ As Single)
 
Dim localPT As Point3D
Dim lI As Long
Dim lfreq As Single
    
    If O3D.PtCount < 1 Or AryStart < 1 Or AryStart > O3D.PtCount Then Exit Sub
    
    If NumPoints < 0 Then NumPoints = O3D.PtCount
    
    lI = AryStart + NumPoints - 1
    If lI > O3D.PtCount Then lI = O3D.PtCount
    
    lfreq = 1.25 * (0.65 + Rnd)
    
    For lI = AryStart To lI 'random points into some kind of shape (cube for now)
    
        If Rnd < 0.1 Then
            localPT.X = 0.5
            localPT.Z = 0
            localPT.Y = 0
            RotatePoint_Test localPT, Rnd * pi, Rnd * pi * 3, Rnd * piBy2, localPT
            localPT.satur = Sqr( _
             (4 * Triangle(localPT.X * lfreq)) ^ 2 + _
             (4 * Triangle(localPT.Y * lfreq)) ^ 2 + _
             (4 * Triangle(localPT.Z * lfreq)) ^ 2) - 0.3
            localPT.hue = 400 * Sqr( _
             ((localPT.X + 0#)) ^ 2 + _
             ((localPT.Y + 3#)) ^ 2 + _
             ((localPT.Z + 2#)) ^ 2) + mHueBase * 9
        Else
            localPT.X = 0.45 * RndPosNeg
            localPT.Z = Rnd - 0.5
            localPT.Y = Rnd - 0.5
            localPT.satur = 1 - (Sqr( _
             (4 * Triangle(localPT.X * lfreq)) ^ 2 + _
             (4 * Triangle(localPT.Y * lfreq)) ^ 2 + _
             (4 * Triangle(localPT.Z * lfreq)) ^ 2) - _
             3.3 * Sqr(localPT.X ^ 2 + localPT.Z ^ 2 + localPT.Y ^ 2) + 1.3)
            localPT.hue = 400 * Sqr( _
             ((localPT.X + 2)) ^ 2 + _
             ((localPT.Y + 0.8)) ^ 2 + _
             ((localPT.Z + 0.75)) ^ 2) + mHueBase * 9
        End If
        
        If localPT.satur > 1 Then
            localPT.satur = 1
        ElseIf localPT.satur < 0 Then
            localPT.satur = 0
        End If
        
        RotatePoint_Test localPT, angleX, angleY, angleZ, localPT
        
        localPT.Z = centerZ + localPT.Z * ReScale
        localPT.Y = centerY + localPT.Y * ReScale
        localPT.X = centerX + localPT.X * ReScale
        
        O3D.PT(lI) = localPT
    
    Next

End Sub

Private Sub O3D_Init(O3D As Obj3D, Optional ByVal PointCount As Long = 100, Optional ByVal pSize As Single = 1)

    If PointCount < 1 Or PointCount = O3D.PtCount Then Exit Sub

    O3D.PtCount = PointCount
    
    Erase O3D.PT
    ReDim O3D.PT(1 To PointCount)
    
    O3D.Center.Origin.X = 0
    O3D.Center.Origin.Y = 0
    O3D.Center.Origin.Z = 0
    
    O3D.X_AXIS.Origin.X = pSize
    O3D.X_AXIS.Origin.Y = 0
    O3D.X_AXIS.Origin.Z = 0
    
    O3D.Z_AXIS.Origin.X = 0
    O3D.Z_AXIS.Origin.Y = 0
    O3D.Z_AXIS.Origin.Z = pSize
    
    O3D.Y_AXIS.Origin.X = 0
    O3D.Y_AXIS.Origin.Y = pSize
    O3D.Y_AXIS.Origin.Z = 0
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    ElseIf KeyCode = vbKeySpace Then
        mBoolStereo = Not mBoolStereo
    Else
        ArrangeObject MyScene, Int(Rnd * MyScene.PtCount + 1) / 2, Rnd - 0.5, Rnd - 0.5, Rnd - 0.5, Int(Rnd * MyScene.PtCount + 1), 0.2 * (Rnd + 1), Rnd * pi * 3, Rnd * piBy2, Rnd * pi
    End If
End Sub
