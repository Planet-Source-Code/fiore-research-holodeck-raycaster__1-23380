VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   750
   ClientTop       =   720
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim binit As Boolean 'A simple flag (true/false) that states whether we've initialised or not. If the initialisation is successful this changes to true, the program also checks before doing any drawing if this flag is true. If the initialisation failed and we try and draw things we'll get lots of errors...

Dim dx As New DirectX7 'This is the root object. DirectDraw is created from this
Dim dd As DirectDraw7 'This is DirectDraw, all things DirectDraw come from here
Dim Mainsurf As DirectDrawSurface7 'This holds our bitmap
Dim primary As DirectDrawSurface7 'This surface represents the screen
Dim backbuffer As DirectDrawSurface7
Dim ddsd1 As DDSURFACEDESC2 'this describes the primary surface
Dim ddsd2 As DDSURFACEDESC2 'this describes the bitmap that we load
Dim ddsd3 As DDSURFACEDESC2 'this describes the size of the screen

Dim TexName As String
Dim Dirpath As String
Dim brunning As Boolean 'this is another flag that states whether or not the main game loop is running.
Dim CurModeActiveStatus As Boolean 'This checks that we still have the correct display mode
Dim bRestore As Boolean 'If we don't have the correct display mode then this flag states that we need to restore the display mode
Dim LastTimeChecked, FPSString

Sub Init()
    On Local Error GoTo errOut 'If there is an error we end the program.
    
    Set dd = dx.DirectDrawCreate("") 'the ("") means that we want the default driver
    Me.Show 'maximises the form and makes sure it's visible
    
    
    'The first line links the DirectDraw object to our form, It also sets the parameters
    'that are to be used - the important ones being DDSCL_FULLSCREEN and DDCSL_EXCLUSIVE. Making it
    'exclusive is important, it means that while our application is running nothing else can
    'use DirectDraw, and it makes windows give us more time/attention
    Call dd.SetCooperativeLevel(Me.hWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE)
    'This is where we actually see a change. It states that we want a display mode
    'of 640x480 with 16 bit colour (65526 colours). the fourth argument ("0") is the
    'refresh rate. leave this to 0 and DirectX will sort out the best refresh rate. It is advised
    'that you don't mess about with this variable. the fifth variable is only used when you
    'want to use the more advanced resolutions (usually the lower, older ones)...
    Call dd.SetDisplayMode(320, 200, 16, 0, DDSDM_DEFAULT)
    
    
    'get the screen surface and create a back buffer too
    ddsd1.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    ddsd1.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
    ddsd1.lBackBufferCount = 1
    Set primary = dd.CreateSurface(ddsd1)
    
    'Get the backbuffer
    Dim caps As DDSCAPS2
    caps.lCaps = DDSCAPS_BACKBUFFER
    Set backbuffer = primary.GetAttachedSurface(caps)
    backbuffer.GetSurfaceDesc ddsd3
    
    
    ' init the surfaces
    InitSurfaces
    
    'This is the main loop. It only runs whilst brunning=true
    binit = True
    brunning = True
    Do While brunning
        ViewAngleWalk = ViewAngle + Angle90
        If ViewAngleWalk > Angle360 Then ViewAngleWalk = ViewAngleWalk - Angle360
        If ViewAngleWalk < Angle0 Then ViewAngleWalk = Angle360 + ViewAngleWalk
        TempX = 0
        TempY = 0
        If Leftb = True Then
            ViewAngle = ViewAngle - Angle6
            If ViewAngle < Angle0 Then ViewAngle = Angle360 + ViewAngle
        End If
        If Rightb = True Then
            ViewAngle = ViewAngle + Angle6
            If ViewAngle > Angle360 Then ViewAngle = ViewAngle - Angle360
        End If
        If Downb = True Then
            TempX = -CosTable(ViewAngleWalk) * Stride
            TempY = SinTable(ViewAngleWalk) * Stride
        End If
        If Upb = True Then
            TempX = CosTable(ViewAngleWalk) * Stride
            TempY = -SinTable(ViewAngleWalk) * Stride
        End If
        
        PlayerX = PlayerX + TempX
        PlayerY = PlayerY + TempY
        FPS = FPS + 1
        blt
        
        DoEvents 'you MUST have a doevents in the loop, otherwise you'll overflow the
        'system (which is bad). All your application does is keep sending messages to DirectX
        'and windows, if you dont give them time to complete the operation they'll crash.
        'adding doevents allows windows to finish doing things that its doing.
    Loop
    
    
errOut:
    'If there is an error we want to close the program down straight away.
    EndIt
End Sub
Sub InitSurfaces()
    
    Set Mainsurf = Nothing
    'load the bitmap textures into a surface
    ddsd2.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsd2.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    ddsd2.lWidth = 640
    ddsd2.lHeight = 64
    Set Mainsurf = dd.CreateSurfaceFromFile(Dirpath & TexturePak, ddsd2)
    'Now it builds all the look-up tables, these tables use a lot of memory
    'but, will make the program run much faster, because it only has to calculate
    'these values once, insted of hundreds of times a second
    Dim RadAngle As Double
    For i = Angle0 To Angle360
        RadAngle = 0.0003272 + i * 3.27249234791667E-03
        TanTable(i) = Tan(RadAngle)
        CoTanTable(i) = 1 / TanTable(i)
        CosTable(i) = Cos(RadAngle)
        SinTable(i) = Sin(RadAngle)
    Next
    For i = 1 To 200
        PixelYTable(i) = 32 / i * FocalDist
    Next
    For i = 0 To 10
        TextureTable(i) = i * 64
    Next
    For i = 1 To 100000
        WallHeightTable(i) = 64 / i * FocalDist
    Next
End Sub
Sub blt()
    'DirectX 7 stuff
    On Local Error GoTo errOut
    If binit = False Then Exit Sub
    Dim ddrval As Long
    Dim rBack As RECT
    bRestore = False
    Do Until ExModeActive
        DoEvents
        bRestore = True
    Loop
    DoEvents
    If bRestore Then
        bRestore = False
        dd.RestoreAllSurfaces
        InitSurfaces
    End If
    Dim FillRect As RECT
    FillRect.Bottom = 200
    FillRect.Right = 320
    If Floors = False Then
        ddrval = backbuffer.BltColorFill(FillRect, RGB(0, 0, 0))
    End If
    'Begin the raycasting loop, this block of code sends out 320 rays for the walls,
    'and then one ray for every 2 remaining pixels(one for floor, ceiling uses same
    'data)
    VAngle = ViewAngle - Angle30 'start 30 degrees to the left of the player
    If VAngle > Angle360 Then VAngle = VAngle - Angle360 'simple error checking
    If VAngle < Angle0 Then VAngle = Angle360 + VAngle
    For Ray = 0 To 319 'begin casting the 320 rays
        RayLength = 0 'set the length of the new ray at 0
        XStep = SinTable(VAngle) * 1 'calculate the x and y difference for a ray with a
        YStep = CosTable(VAngle) * 1 'length of 1 pixel
        RayX = PlayerX 'start the ray from the player
        RayY = PlayerY
        WallHit = 12 'more error stoping code
        Do
            RayX = RayX - XStep 'subtract the difference in x and y
            RayY = RayY - YStep
            RayLength = RayLength + 1 'make the ray longer(used later for wall height)
            'check to see if the ray hit a wall
            WallHit = World(Int(RayX / CellX), Int(RayY / CellY))
        Loop Until WallHit < 10 'Exit loop if ray hit wall
        'adjust the raylength to correct for 'fishbowl' effect
        RayLength = Int(RayLength * CosTable(Abs(VAngle - ViewAngle)))
        'figure out the height of the walls
        WallHeight = WallHeightTable(RayLength) '(64 / Raylength) * FocalDist
        addY = RayY Mod 64 'finds the x and y offset for texture mapping
        addX = RayX Mod 64
        If addX > 64 Then addX = 64 'again, more bug reducing code
        If addY > 64 Then addY = 64
        If addX < 1 Then addX = 1
        If addY < 1 Then addY = 1
        'The brilliant code that checks to see which side of the wall the ray hit
        If Int(addY) >= 32 Then
            If Int(addX) > Abs(Int(addY - 64)) Then
                x = TextureTable(WallHit) + addX
            Else
                x = TextureTable(WallHit) + addY
            End If
        Else
            If Int(addX) > Int(addY) Then
                If WallHit >= 0 Then
                    x = TextureTable(WallHit) + addX
                End If
            Else
                x = TextureTable(WallHit) + addY
            End If
        End If
        If Floors = True Then
        'this code casts the rays for the floor and ceiling
        PixelY = WallHeight / 2 'start with the pixel directly below the last wall
        'this is expressed in distance from center(100)
        Do
            'using simple triangle equations we can find the straight distance to
            'the pixel on the floor
            StraightFloorDist = PixelYTable(PixelY) '* FocalDist
            'this straight distance is not what we want however, we want the
            'actual distance, but since we know the straight distance we can compute
            'the actual distance, you can see this in the figure below. The angle we
            'use to compute this is the angle relative to the player viewing angle
            '     \    |
            '      \   |
            '     AD\  |SD
            '        \ |
            '         \P
            FloorRayDist = StraightFloorDist / CosTable(Abs(ViewAngle - VAngle))
            'we can then use the actual distance(see above) to determine the x and
            'y offsets for the floor pixel
            XFloorDist = SinTable(VAngle) * FloorRayDist
            YFloorDist = CosTable(VAngle) * FloorRayDist
            'adding these offsets to the player position gives us the exact pixel
            'that the ray hits
            FloorX = PlayerX + XFloorDist
            FloorY = PlayerY + YFloorDist
            'we then use this information to find the exact pixel on the texture map
            FloorPixelX = FloorX Mod 64
            FloorPixelY = FloorY Mod 64
            If FloorPixelX > 64 Then FloorPixelX = 64 'more bug prevention code
            If FloorPixelX < 1 Then FloorPixelX = 1
            'then all we have to do is draw the texture map pixel on to the pixel on
            'the screen which we started from
            rBack.Left = FloorPixelX + TextureTable(FloorTex)
            rBack.Top = FloorPixelY
            rBack.Bottom = rBack.Top + 1
            rBack.Right = rBack.Left + 1
            ddrval = backbuffer.BltFast(Ray, PixelY + 100, Mainsurf, rBack, DDBLTFAST_DONOTWAIT)
            'since the ceiling is the exact same distance away from the player, the
            'computer just recalculates the texture map pixel, and draws it to the screen
            'this way drawing ceilings, however, will have to be changed in later
            'versions so the player can look up and down, crouch, and fly, but for now,
            'in the interest of speed, it will do
            rBack.Left = FloorPixelX + TextureTable(CeilingTex)
            rBack.Top = FloorPixelY
            rBack.Bottom = rBack.Top + 1
            rBack.Right = rBack.Left + 1
            ddrval = backbuffer.BltFast(Ray, 100 - PixelY, Mainsurf, rBack, DDBLTFAST_DONOTWAIT)
            PixelY = PixelY + 1 'then go to the next pixel below the previous
        Loop Until PixelY > 200 'stop this when the pixel reaches the bottom of the screen
        End If
        'Draw the texture on the screen
        'using the information about the texture computed before the floor casting routine
        'we can now draw the wall
        HeightFix = 0
        If WallHeight > 200 Then
        HeightFix = WallHeight / 2 - 100
        End If
        Dim rBack2 As RECT
        rBack.Top = 0 + HeightFix
        rBack.Left = x
        rBack.Bottom = 64 - HeightFix
        rBack.Right = x + 1
        rBack2.Left = Ray
        rBack2.Top = 100 - (WallHeight / 2)
        rBack2.Bottom = 100 + (WallHeight / 2)
        rBack2.Right = Ray + 1
        ddrval = backbuffer.blt(rBack2, Mainsurf, rBack, DDBLT_WAIT)
        VAngle = VAngle + 1 'move the ray angle one to the right
        If VAngle > Angle360 Then VAngle = VAngle - Angle360 'again, more bug controlling
        If VAngle < Angle0 Then VAngle = Angle360 + VAngle 'code(you'll see a lot of this)
    Next Ray 'repeat this entire process 319 more times
    'now print the framerate in the upper-left corner
    backbuffer.SetForeColor RGB(0, 0, 255)
    Call backbuffer.DrawText(0, 0, FPSString, False)
    If dx.TickCount - LastTimeChecked >= 1000 Then
        'every second reset the FPS
        LastTimeChecked = dx.TickCount
        FPSString = FPS
        FPS = 0
    End If
    'since we are drawing this to a back buffer, we need to now draw every thing to
    'the primary surface(also called the screen)
    primary.Flip Nothing, DDFLIP_WAIT
    'now the computer will do this again, and again as many times as possible
errOut:
End Sub
Sub EndIt() 'directX 7 end program sub
    Call dd.RestoreDisplayMode
    Call dd.SetCooperativeLevel(Me.hWnd, DDSCL_NORMAL)
    End
End Sub

Private Sub Form_Click()
    EndIt
End Sub

Private Sub Form_Load()
    'Starts the whole program.
    Init
    Dirpath = App.Path
    If Right(Dirpath, 1) <> "\" Then
        Dirpath = Dirpath + "\"
    End If
End Sub

Private Sub Form_Paint()
    blt
End Sub


Function ExModeActive() As Boolean
    Dim TestCoopRes As Long
    TestCoopRes = dd.TestCooperativeLevel
    If (TestCoopRes = DD_OK) Then
        ExModeActive = True
    Else
        ExModeActive = False
    End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyRight
            Rightb = True
        Case vbKeyLeft
            Leftb = True
        Case vbKeyUp
            Upb = True
        Case vbKeyDown
            Downb = True
        Case vbKeySpace
            Fireb = True
        Case vbKeyQ
            EndIt
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyRight
            Rightb = False
        Case vbKeyLeft
            Leftb = False
        Case vbKeyUp
            Upb = False
        Case vbKeyDown
            Downb = False
        Case vbKeySpace
            Fireb = False
    End Select
End Sub
