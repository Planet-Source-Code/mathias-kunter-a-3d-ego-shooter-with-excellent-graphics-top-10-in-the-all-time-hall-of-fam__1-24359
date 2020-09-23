Attribute VB_Name = "Game"
'General
Private Const PlayT% = 90                       'Time of the game
Private Const BlendT% = 2                       'Blend-Time
Private Const TextBlendT% = 2                   'Blend-Time for text
Private Const WorldG! = 1                       'g (used for patrons)

'Private
Private Const EyesHeight! = 2.8                 'height of the eyes
Private Const WalkSpeed% = 6                    'your walk-speed
Private Const PixelPer360% = 1024               'pixel to move with the mouse until a 360

'MG
Private Const MGTimePer360% = 3                 'seconds of the MG until a 360
Private Const MGCollT% = 20                     'seconds until the MG comes again

'MG Holding
Private Const MGLoadT! = 0.4                    'time of taking the MG
Private Const MGLoadDiff! = 1.2                 'start height of taking the MG
Private Const MGPatronsPerSec% = 10             'fire rate of the MG per second (has no effect on the bot, is only used to count down the bullets)
Private Const MGPointsPerSec% = 10              'the points which are subtracted per second while using the MG
Private Const MGPatronsShowPerSec% = 15         'patrons shown per second

'Mann
Private Const ManTimePer360% = 2                'time until the bot makes a 360
Private Const ManFallT% = 1                     'time until bot falls down when fragged
Private Const ManDieT! = 0.5                    'time until bot is fragged from the MG





Private Landscape As New Mk3dObject
Private MG As New Mk3dObject, MGHolding As New Mk3dObject
Private ManAnim As New Mk3dAnimatedObject, ManCalced As New Mk3dObject
Private MGBullets(MGPatronsShowPerSec - 1) As New Mk3dObject
Private MGBulletsDesc(MGPatronsShowPerSec - 1) As MGBulletType
Private Blood As New Mk3dEffectObject


Private MapArea!(1, 1)

Private yPos As D3DVECTOR, yEyes As D3DVECTOR, yAngle As D3DVECTOR

Private MGPos As D3DVECTOR, MGAngle As D3DVECTOR, MGWaitT!

Private MGPatrons!, MGHoldingUseT!, MGHoldingState As MGHoldingStateEnum
Private MGHoldingPos As D3DVECTOR, MGHoldingAngle As D3DVECTOR, MGHoldingLightIndex%
Private ActMGPatrons%, MGPatronsWaitT!

Private ManPos As D3DVECTOR, ManAngle As D3DVECTOR, ManState As ManStateEnum
Private ManRotTo!, ManGoLen!, ManWentLen!, ManWalkDir As D3DVECTOR
Private ManGoT!, ManActionWaitT!, ManShotT!


Private MenuBackgr As DirectDrawSurface7
Private ActPlayT!
Private FrameT!
Private GamePoints!
Private CollDetCount%, CollDet() As D3DVECTOR
Private HitSpecialW(1) As Boolean
Private TextBlendWaitT!, ShowText As Boolean, TextToShow As String
Public GameFont As IFont

Private Type MGBulletType
    MGBulletDir As D3DVECTOR
    MGStartT As Single
    MGFallSpeed As Single
End Type

Private Enum ManStateEnum
    MAN_BLENDIN
    MAN_BLENDOUT
    MAN_ROTATE
    MAN_GO
    MAN_DIE
End Enum

Private Enum MGHoldingStateEnum
    MG_NONE
    MG_BLENDIN
    MG_NORMAL
    MG_FIRE
    MG_BLENDOUT
End Enum

Private Enum MGPatronHitEnum
    MGPATRON_HITNOTHING
    MGPATRON_HITLANDSCAPE
    MGPATRON_HITMAN
End Enum





Public Sub Load(ByVal FirstLoad As Boolean)
    Dim i%, StartT!
    Dim LoadObj(3) As New Mk3dObject, Failed As Boolean

    On Local Error GoTo Failed
    
    'init the landscape
    If Not Landscape.CreateFromFile(App.Path & "\Objects\Landscape.obj", FirstLoad) Then Failed = True
    'init the MG
    If Not MG.CreateFromFile(App.Path & "\Objects\Weapon.obj", False) Then Failed = True
    'init your personal MG
    If Not MGHolding.CreateFromFile(App.Path & "\Objects\Weapon.obj", False) Then Failed = True
    MGHoldingLightIndex = Mk3d.LightAdd(MGHolding.GetLight(0))
    'init patrons
    For i = 0 To UBound(MGBullets)
        If Not MGBullets(i).CreateFromFile(App.Path & "\Objects\Bullet.obj", False) Then Failed = True
    Next i
    'init bot
    If Not LoadObj(0).CreateFromFile(App.Path & "\Objects\Bot Stand.obj", False) Then Failed = True
    LoadObj(0).Central True, False, True
    If Not LoadObj(1).CreateFromFile(App.Path & "\Objects\Bot StepRight.obj", False) Then Failed = True
    LoadObj(1).Central True, False, True
    If Not LoadObj(2).CreateFromFile(App.Path & "\Objects\Bot StepLeft.obj", False) Then Failed = True
    LoadObj(2).Central True, False, True
    If Not LoadObj(3).CreateFromFile(App.Path & "\Objects\Bot Died.obj", False) Then Failed = True
    LoadObj(3).Central True, False, True
    If Not ManAnim.CreateFromObjects(4, LoadObj()) Then Failed = True
    'init blood
    If FirstLoad Then
        Blood.Initsialize 5400               'init vertices for maximum 100 frags
        Blood.EffectFileLoad App.Path & "\Objects\Blood.obj"
        Blood.TextureSet App.Path & "\Textures\Blood.bmp"
        Blood.MaterialSet App.Path & "\Materials\Global.mat", 0
    End If
    
    If Failed Then
        MsgBox "There was an error while loading the 3d-objects.", vbCritical
        Mk3d.ExitDX
        End
    End If
    Exit Sub
    
Failed:
    MsgBox "There was an error while loading the 3d-coordinates.", vbCritical
    Mk3d.ExitDX
    End
End Sub

Public Sub Initsialize()
    Dim i%

    'init startup position
    Open App.Path & "\Data\Startup.dat" For Input As #1
    Input #1, yPos.x
    Input #1, yPos.y
    Input #1, yPos.z
    Input #1, yAngle.x
    Input #1, yAngle.y
    Input #1, MGPos.x
    Input #1, MGPos.y
    Input #1, MGPos.z
    Input #1, MGAngle.x
    Input #1, MGAngle.y
    Input #1, ManPos.x
    Input #1, ManPos.y
    Input #1, ManPos.z
    Input #1, ManAngle.x
    Input #1, ManAngle.y
    Close #1
    
    'init collision detection
    Open App.Path & "\Data\Size.dat" For Input As #1
    Input #1, MapArea(0, 0)                         'x min
    Input #1, MapArea(0, 1)                         'x max
    Input #1, MapArea(1, 0)                         'z min
    Input #1, MapArea(1, 1)                         'z max
    Close #1
    
    Open App.Path & "\Data\Collission.dat" For Input As #1
    Input #1, CollDetCount
    If CollDetCount <> 0 Then ReDim CollDet(CollDetCount - 1)
    For i = 0 To CollDetCount - 1
        Input #1, CollDet(i).x
        Input #1, CollDet(i).y
        Input #1, CollDet(i).z
    Next i
    Close #1
    
    'general
    FrameT = 0
    GamePoints = 0
    HitSpecialW(0) = False
    HitSpecialW(1) = False
    ShowText = True
    TextToShow = "WELCOME TO SHOT IT"
    Mk3d.PrimarySurf.SetForeColor vbWhite
    Mk3d.BackBufferSurf.SetForeColor vbWhite
    'private
    yEyes = yPos
    yEyes.y = yEyes.y + EyesHeight
    'MG
    MG.Central True, True, True
    MG.MoveTo MGPos
    MG.Rotate MGAngle
    MGWaitT = MGCollT
    'MG-Holding
    MGHolding.Central True, True, True
    MGHolding.Rotate Mk3d.VectorMake(0, 1.57075, 0)
    MGPatrons = 0
    ActMGPatrons = 0
    MGPatronsWaitT = 0
    MGHoldingUseT = 0
    MGHoldingPos = Mk3d.VectorMake(0, 0, 0)
    MGHoldingState = MG_NONE
    MGHoldingAngle = Mk3d.VectorMake(0, 0, 0)
    MGHoldingLightIndex = 0
    'blood
    Blood.EffectVcnt = 0
    Blood.EffectFileCentral True, True, True
    'bot
    ManAnim.MoveTo ManPos
    ManAnim.Rotate ManAngle
    
    ManState = MAN_BLENDIN
    ManRotTo = 0
    ManGoLen = 0
    ManWentLen = 0
    ManWalkDir = Mk3d.VectorMake(0, 0, 0)
    ManGoT = 0
    ManActionWaitT = 0
    ManShotT = 0
End Sub

Public Function Run() As Integer
    Dim i%, j%, cnt&, StartT!
    Dim MinLeft%, SecLeft%
    Dim yLookAt As D3DVECTOR, yLookDir As D3DVECTOR, yLookRefer As D3DVECTOR
    Dim MGHoldingRefer As D3DVECTOR
    
    'last settings before the game starts
    'general
    On Local Error Resume Next
    StartT = Timer
    yLookRefer = Mk3d.VectorMake(0, 0, 1)
    MGHoldingRefer = Mk3d.VectorMake(1.1, -0.9, 1.3)            'refers to the camera-position
    Set ManCalced = ManAnim.GetKeyFrameObj(0)
    DoEvents
    
    
    'Game-Loop
    Do
        'general
        cnt = cnt + 1                                           'get the frame-time
        ActPlayT = GetTimeDiff(StartT, Timer)
        FrameT = ActPlayT / cnt
        MinLeft = Int((PlayT - ActPlayT) / 60)
        SecLeft = Int(PlayT - ActPlayT - MinLeft * 60) + 1
        TextBlendWaitT = TextBlendWaitT + FrameT
        If TextBlendWaitT > TextBlendT Then ShowText = False
        yLookDir = Mk3d.VectorRotate(yLookRefer, yAngle)
        DoEvents
        
        'MG-Holding
        GameMGHolding MGHoldingRefer, yLookDir, cnt
        For i = 0 To ActMGPatrons - 1
            With MGBulletsDesc(i)
                .MGFallSpeed = .MGFallSpeed + WorldG * FrameT
                MGBullets(i).Move Mk3d.VectorMake(.MGBulletDir.x, -.MGFallSpeed, .MGBulletDir.z)
            End With
            If MGBullets(i).GetPosition.y < 0 Then
                For j = i To ActMGPatrons - 2
                    MGBullets(j).MoveTo MGBullets(j + 1).GetPosition
                    MGBulletsDesc(j) = MGBulletsDesc(j + 1)
                Next j
                ActMGPatrons = ActMGPatrons - 1
            End If
        Next i
        
        'keyboard
        If GameKeyboard(yLookDir) Then
            Run = Int(GamePoints)
            Exit Function        'exit
        End If
                
        'mouse
        GameMouse
        
        'MG
        GameMG
                
        'bot
        GameMan
        
        'render the szene
        Mk3d.dx.VectorAdd yLookAt, yEyes, yLookDir                    'set the camera
        Mk3d.SetCamera yEyes, yLookAt
        Mk3d.d3dDevice.BeginScene                                     'render the szene
        Mk3d.d3dDevice.Clear 1, Mk3d.d3drcViewport(), D3DCLEAR_TARGET, Mk3d.dx.CreateColorRGB(0, 0, 0), 0, 0
        Mk3d.d3dDevice.Clear 1, Mk3d.d3drcViewport(), D3DCLEAR_ZBUFFER, 0, 1, 0
        Mk3d.Render Landscape
        If Not MGHoldingState = MG_NONE Then Mk3d.Render MGHolding
        For i = 0 To ActMGPatrons - 1
            Mk3d.Render MGBullets(i)
        Next i
        If MGWaitT >= MGCollT Then Mk3d.Render MG
        Mk3d.RenderEffect Blood
        Mk3d.Render ManCalced
        Mk3d.d3dDevice.EndScene
        
        'info
        Mk3d.BackBufferSurf.DrawText 10, 10, "Time: " & MinLeft & ":" & SecLeft, False
        Mk3d.BackBufferSurf.DrawText 10, 40, "Score: " & Int(GamePoints), False
        If ShowText Then Mk3d.BackBufferSurf.DrawText TextCentralX(Len(TextToShow)), Mk3d.VPSize(1) / 2 - 200, TextToShow, False
        If Not MGHoldingState = MG_NONE Then
            Mk3d.BackBufferSurf.DrawText 10, 70, "Bullets: " & Int(MGPatrons), False
            'cross
            Mk3d.BackBufferSurf.DrawLine Mk3d.VPSize(0) / 2, Mk3d.VPSize(1) / 2 - 10, Mk3d.VPSize(0) / 2, Mk3d.VPSize(1) / 2 + 10
            Mk3d.BackBufferSurf.DrawLine Mk3d.VPSize(0) / 2 - 10, Mk3d.VPSize(1) / 2, Mk3d.VPSize(0) / 2 + 10, Mk3d.VPSize(1) / 2
        End If
        
        Mk3d.PrimarySurf.Flip Nothing, DDFLIP_DONOTWAIT
    Loop While ActPlayT < PlayT
    Run = Int(GamePoints)
End Function

Public Sub Menu(ByVal YName As String)
    Dim i%, j%, FirstPlay As Boolean, PressState As Boolean, ReDraw As Boolean
    Dim ActSel%, MaxSel%, ActChoice%
    Dim Score%, RecNames$(4), RecScores%(4), RecInd%
    Dim SurfaceDesc As DDSURFACEDESC2
    Dim KeybState As DIKEYBOARDSTATE
    
    On Local Error Resume Next
    MaxSel = 3
    ReDraw = True
    FirstPlay = True
    
    SurfaceDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Set MenuBackgr = Mk3d.dd.CreateSurfaceFromFile(App.Path & "\Data\Background.bmp", SurfaceDesc)
    
    Do
        Mk3d.diDeviceKeyb.GetDeviceStateKeyboard KeybState
        ActChoice = -1
        DoEvents
        
        If Not KeybState.Key(200) = 0 Then
            If Not PressState And Not ActSel = 0 Then
                ActSel = ActSel - 1
                ReDraw = True
            End If
            PressState = True
        ElseIf KeybState.Key(208) = 0 Then
            PressState = False
        End If
        
        If Not KeybState.Key(208) = 0 Then
            If Not PressState And Not ActSel = MaxSel Then
                ActSel = ActSel + 1
                ReDraw = True
            End If
            PressState = True
        ElseIf KeybState.Key(200) = 0 Then
            PressState = False
        End If
        
        'Return
        If Not KeybState.Key(28) = 0 Or Not KeybState.Key(156) = 0 Then
            ActChoice = ActSel
        End If
        
        If ReDraw Then
            ClearMenu
            SetMenuColor ActSel, 0
            Mk3d.PrimarySurf.DrawText TextCentralX(16), TextCentralY(0, 4, 40), "Start a new game", False
            SetMenuColor ActSel, 1
            Mk3d.PrimarySurf.DrawText TextCentralX(9), TextCentralY(1, 4, 40), "Highscore", False
            SetMenuColor ActSel, 2
            Mk3d.PrimarySurf.DrawText TextCentralX(7), TextCentralY(2, 4, 40), "Credits", False
            SetMenuColor ActSel, 3
            Mk3d.PrimarySurf.DrawText TextCentralX(9), TextCentralY(3, 4, 40), "Exit game", False
        End If
        ReDraw = False
        Mk3d.PrimarySurf.SetForeColor vbBlack
        
        If Not ActChoice = -1 Then
            ClearMenu
            Select Case ActChoice
                Case 0
                    'start the game
                    Mk3d.PrimarySurf.SetForeColor vbBlack
                    Mk3d.PrimarySurf.DrawText TextCentralX(24), TextCentralY(0, 1, 40), "Loading - please wait...", False
                    Game.Load FirstPlay
                    Game.Initsialize
                    Mk3d.diDeviceMouse.Acquire
                    Score = Game.Run
                    ClearMenu
                    Mk3d.diDeviceMouse.Unacquire
                    dsWalkSound.Stop
                    dsShootSound.Stop
                    FirstPlay = False
                    'check if the scrore is a record
                    Open App.Path & "\Data\Highscore.dat" For Random As #1 Len = 16
                    For i = 0 To 4
                        Get #1, i * 2 + 1, RecNames(i)
                        Get #1, i * 2 + 2, RecScores(i)
                    Next i
                    Close #1
                    RecInd = -1
                    For i = 0 To 4
                        If Score > RecScores(i) Then
                            'in the highscore!
                            For j = 4 To i + 1 Step -1
                                RecNames(j) = RecNames(j - 1)
                                RecScores(j) = RecScores(j - 1)
                            Next j
                            RecNames(i) = YName
                            RecScores(i) = Score
                            RecInd = i
                            Exit For
                        End If
                    Next i
                    Open App.Path & "\Data\Highscore.dat" For Random As #1 Len = 16
                    For i = 0 To 4
                        Put #1, i * 2 + 1, RecNames(i)
                        Put #1, i * 2 + 2, RecScores(i)
                    Next i
                    Close #1
                    ClearMenu
                    Mk3d.PrimarySurf.SetForeColor vbBlack
                    ShowHighscore Score, True, RecInd
                    ReDraw = True
                Case 1
                    'show highscore
                    ShowHighscore 0, False, 0
                    ReDraw = True
                Case 2
                    'show credits
                    Mk3d.PrimarySurf.DrawText TextCentralX(7), TextCentralY(0, 6, 40), "CREDITS", False
                    Mk3d.PrimarySurf.DrawText TextCentralX(26), TextCentralY(2, 6, 40), "Programmer: Mathias Kunter", False
                    Mk3d.PrimarySurf.DrawText TextCentralX(28), TextCentralY(3, 6, 40), "Mail: mathiaskunter@yahoo.de", False
                    Mk3d.PrimarySurf.DrawText TextCentralX(15), TextCentralY(5, 6, 40), "ESC to continue", False
                    WaitForESC
                    ReDraw = True
                Case 3
                    'exit game
                    Mk3d.ExitDX
                    DoEvents
                    End
            End Select
            ReDraw = True
        End If
    Loop While Not ActChoice = MaxSel
End Sub

Private Sub ShowHighscore(ByVal YScore As Integer, ByVal ShowYScore As Boolean, ByVal RecInd As Integer)
    Dim i%, RecNames$(4), RecScores%(4)
    
    Open App.Path & "\Data\Highscore.dat" For Random As #1 Len = 16
    For i = 0 To 4
        Get #1, i * 2 + 1, RecNames(i)
        Get #1, i * 2 + 2, RecScores(i)
    Next i
    Close #1
    
    Mk3d.PrimarySurf.DrawText TextCentralX(9), TextCentralY(0, 10, 40), "HIGHSCORE", False
    For i = 0 To 4
        If Not RecNames(i) = "" Then
            If ShowYScore And i = RecInd Then
                Mk3d.PrimarySurf.SetForeColor vbRed
            Else
                Mk3d.PrimarySurf.SetForeColor vbBlack
            End If
            Mk3d.PrimarySurf.DrawText TextCentralX(Len(i + 1 & ".: " & RecScores(i) & " points of " & RecNames(i))), TextCentralY(i + 2, 10, 40), i + 1 & ".: " & RecScores(i) & " points of " & RecNames(i), False
        End If
    Next i
    Mk3d.PrimarySurf.SetForeColor vbBlack
    If ShowYScore Then
        Mk3d.PrimarySurf.DrawText TextCentralX(Len("This game: " & YScore & " points")), TextCentralY(8, 10, 40), "This game: " & YScore & " points", False
        Mk3d.PrimarySurf.DrawText TextCentralX(15), TextCentralY(9, 10, 40), "ESC to continue", False
    Else
        Mk3d.PrimarySurf.DrawText TextCentralX(15), TextCentralY(9, 10, 40), "ESC to continue", False
    End If
    
    WaitForESC
End Sub

Private Sub WaitForESC()
    Dim KeybState As DIKEYBOARDSTATE
    
    Do
        DoEvents
        Mk3d.diDeviceKeyb.GetDeviceStateKeyboard KeybState
    Loop While KeybState.Key(1) = 0
End Sub

Private Sub ClearMenu()
    Dim EmptyRect As RECT

    Mk3d.PrimarySurf.Blt EmptyRect, MenuBackgr, EmptyRect, DDBLT_DONOTWAIT
    Mk3d.d3dDevice.BeginScene
    Mk3d.d3dDevice.Clear 1, Mk3d.d3drcViewport(), D3DCLEAR_TARGET, Mk3d.dx.CreateColorRGB(1, 1, 1), 0, 0
    Mk3d.d3dDevice.EndScene
    Mk3d.PrimarySurf.Flip Nothing, DDFLIP_DONOTWAIT
End Sub

Private Sub SetMenuColor(ByVal ActSel As Integer, NowSel As Integer)
    If ActSel = NowSel Then
        Mk3d.PrimarySurf.SetForeColor vbRed
    ElseIf Mk3d.PrimarySurf.GetForeColor = vbRed Then
        Mk3d.PrimarySurf.SetForeColor vbBlack
    End If
End Sub

Private Function TextCentralX(ByVal TextLen As Integer) As Integer
    TextCentralX = Mk3d.VPSize(0) / 2 - (GameFont.Size - 4) * TextLen / 2
End Function

Private Function TextCentralY(ByVal ActLine As Integer, ByVal NrLines As Integer, ByVal LinesDff As Integer) As Integer
    TextCentralY = Mk3d.VPSize(1) / 2 - LinesDff * (NrLines / 2 - ActLine)
End Function





Private Function GameKeyboard(yLookDir As D3DVECTOR) As Boolean
    Dim KeybState As DIKEYBOARDSTATE
    Dim yWalkDir As D3DVECTOR, yPosBef As D3DVECTOR, yCollDet As D3DVECTOR
    Dim RotAngle!, cntEnable%
    
    Mk3d.diDeviceKeyb.GetDeviceStateKeyboard KeybState
    If Not KeybState.Key(1) = 0 Then
        'ESC, go to the menu
        Do
            DoEvents
            Mk3d.diDeviceKeyb.GetDeviceStateKeyboard KeybState
        Loop While Not KeybState.Key(1) = 0                     'wait until ESC is no longer pressed
        GameKeyboard = True
        Exit Function
    End If
    RotAngle = GetMoveAngle(KeybState)
    If Not RotAngle = -1 Then
        yWalkDir = yLookDir
        yWalkDir.y = 0
        Mk3d.dx.VectorNormalize yWalkDir
        yWalkDir = Mk3d.VectorRotate(yWalkDir, Mk3d.VectorMake(0, RotAngle, 0))
        Mk3d.dx.VectorScale yWalkDir, yWalkDir, WalkSpeed * FrameT
        
        'Collision Detection
        yPosBef = yPos
        Mk3d.dx.VectorAdd yPos, yPos, yWalkDir
        yCollDet = GetCollDet(yPosBef, yPos, 2, False, True)
        If yCollDet.x = 0 Then
            yWalkDir.x = 0
            yPos = yPosBef
            cntEnable = cntEnable + 1
        End If
        If yCollDet.z = 0 Then
            yWalkDir.z = 0
            yPos = yPosBef
            cntEnable = cntEnable + 1
        End If
        If cntEnable = 1 Then
            'only if of the two values is zero. if both are zero, you can't walk in ANY direction
            Mk3d.dx.VectorNormalize yWalkDir
            Mk3d.dx.VectorScale yWalkDir, yWalkDir, WalkSpeed * FrameT
            Mk3d.dx.VectorAdd yPos, yPos, yWalkDir
        End If
        Mk3d.dx.VectorAdd yEyes, yEyes, yWalkDir
        If Not cntEnable = 2 Then
            'start playing the walk sound
            dsWalkSound.Play DSBPLAY_LOOPING
        Else
            'stop playing the walk sound
            dsWalkSound.Stop
            dsWalkSound.SetCurrentPosition 0
        End If
        
        'move also MG
        If Not MGHoldingState = MG_NONE Then
            Mk3d.dx.VectorAdd MGHoldingPos, MGHoldingPos, yWalkDir
            MGHolding.MoveTo MGHoldingPos
        End If
    Else
        'stop playing the walk sound
        dsWalkSound.Stop
        dsWalkSound.SetCurrentPosition 0
    End If
End Function


Private Sub GameMouse()
    Dim MouseState As DIMOUSESTATE, yRot As D3DVECTOR

    Mk3d.diDeviceMouse.GetDeviceStateMouse MouseState
    yRot.x = MouseState.y * 6.283 / PixelPer360
    yRot.y = MouseState.x * 6.283 / PixelPer360
    yAngle.x = yAngle.x - yRot.x
    yAngle.y = yAngle.y - yRot.y
    If yAngle.x > 1 Then yAngle.x = 1
    If yAngle.x < -1 Then yAngle.x = -1
    If yAngle.y > 6.283 Then yAngle.y = yAngle.y - 6.283
    If yAngle.y < 0 Then yAngle.y = yAngle.y + 6.283
    
    If MGHoldingState = MG_NORMAL Or MGHoldingState = MG_FIRE Then
        If Not MouseState.buttons(0) = 0 And MGHoldingState = MG_NORMAL Then
            MGHoldingState = MG_FIRE
            'start playing the MG sound
            dsShootSound.SetCurrentPosition 0
            dsShootSound.Play DSBPLAY_LOOPING
        ElseIf MouseState.buttons(0) = 0 Then
            MGHoldingState = MG_NORMAL
            Mk3d.LightSetState MGHoldingLightIndex, False
            yEyes.y = EyesHeight
            'stop playing the MG sound
            dsShootSound.Stop
        End If
    End If
End Sub


Private Sub GameMG()
    Dim MGRot As D3DVECTOR
    
    MGWaitT = MGWaitT + FrameT
    MGRot.y = 6.283 * FrameT / MGTimePer360
    MG.Rotate MGRot
    MGAngle.y = MGAngle.y + MGRot.y
End Sub


Private Sub GameMGHolding(MGHoldingRefer As D3DVECTOR, yLookDir As D3DVECTOR, FrCnt As Long)
    Dim MGHoldingLookDir As D3DVECTOR, Corrx!, NewMGPos As D3DVECTOR
    Dim MGPatronState As MGPatronHitEnum, MGPatronPos As D3DVECTOR, MGPatronPosBef As D3DVECTOR
    Dim RdLight As D3DLIGHT7, RdLightStat As Boolean

    If MGHoldingState = MG_NONE Then                       'take MG
        If InArea(yPos, MGPos, 1) And MGWaitT > MGCollT Then
            MGPatrons = 100
            MGWaitT = 0
            MGHoldingUseT = 0
            MGHoldingState = MG_BLENDIN
            NewMGPos = GetRandomPos
            MGPos = Mk3d.VectorMake(NewMGPos.x, MGPos.y, NewMGPos.z)
            MG.MoveTo MGPos
        End If
        Exit Sub
    End If
    
    MGHoldingUseT = MGHoldingUseT + FrameT
    MGPatronsWaitT = MGPatronsWaitT + FrameT
    If MGPatronsWaitT > 1 / MGPatronsShowPerSec Then MGPatronsWaitT = 0
    
    If MGHoldingState = MG_BLENDIN Then
        If MGHoldingUseT < MGLoadT Then
            Corrx = MGLoadDiff - MGHoldingUseT / MGLoadT * MGLoadDiff
        Else
            MGHoldingState = MG_NORMAL
        End If
    ElseIf MGHoldingState = MG_BLENDOUT Then
        If MGHoldingUseT < MGLoadT Then
            'stop playing the sound
            dsShootSound.Stop
            Corrx = MGHoldingUseT / MGLoadT * MGLoadDiff
        Else
            MGHoldingState = MG_NONE
            Exit Sub
        End If
    ElseIf MGHoldingState = MG_FIRE Then
        'turn the the light source of the MG on and off
        RdLight = MGHolding.GetLight(0)
        RdLight.Position = MGHoldingPos
        If FrCnt Mod 2 = 0 Then
            yEyes.y = EyesHeight + 0.05
            RdLightStat = True
        Else
            yEyes.y = EyesHeight
        End If
        Mk3d.LightUpdate MGHoldingLightIndex, RdLight
        Mk3d.LightSetState MGHoldingLightIndex, RdLightStat
        
        If Not Int(MGPatrons) = 0 Then
            'count down the bullets
            MGPatrons = MGPatrons - MGPatronsPerSec * FrameT
            If Int(MGPatrons) <= 0 Then
                MGHoldingState = MG_BLENDOUT
                MGHoldingUseT = 0
                MGPatrons = 0
            Else
                'show the patrons and subtract some points for shooting with the MG
                If Not ActMGPatrons = MGPatronsShowPerSec And MGPatronsWaitT = 0 Then
                    'another bullet can be shown
                    MGBullets(ActMGPatrons).MoveTo MGHoldingPos
                    MGBullets(ActMGPatrons).Move Mk3d.VectorRotate(Mk3d.VectorMake(-0.05, 0, 0.3), yAngle)
                    With MGBulletsDesc(ActMGPatrons)
                        .MGBulletDir = Mk3d.VectorMake(GetRandom(-0.025, 0.025), 0, GetRandom(-0.025, 0.025))
                        .MGStartT = Timer
                        .MGFallSpeed = 0
                    End With
                    ActMGPatrons = ActMGPatrons + 1
                End If
                GamePoints = GamePoints - MGPointsPerSec * FrameT
                If GamePoints < 0 Then GamePoints = 0
            End If
            
            'Collision Detection of the shots
            If ManState = MAN_BLENDIN Or ManState = MAN_GO Or ManState = MAN_ROTATE Then
                MGPatronState = MGPATRON_HITNOTHING
                MGPatronPos = yEyes
                Do
                    MGPatronState = GetPatronCollDet(MGPatronPos, 0.4)
                    MGPatronPosBef = MGPatronPos
                    Mk3d.dx.VectorAdd MGPatronPos, MGPatronPos, yLookDir
                Loop While MGPatronState = MGPATRON_HITNOTHING
                
                If MGPatronState = MGPATRON_HITMAN Then
                    ManShotT = ManShotT + FrameT
                    If ManShotT > ManDieT Then
                        ShowText = True
                        If ManState = MAN_BLENDIN Then
                            GamePoints = GamePoints + 150
                            TextToShow = "BOT FAST FRAG: + 150"
                        Else
                            GamePoints = GamePoints + 100
                            TextToShow = "BOT FRAG: + 100"
                        End If
                        ManState = MAN_DIE
                        ManActionWaitT = 0
                        TextBlendWaitT = 0
                    End If
                End If
            End If
        End If
    End If
    
    'calculate the position of the MG which you are holding
    MGHoldingLookDir = Mk3d.VectorRotate(MGHoldingRefer, yAngle)
    MGHoldingLookDir.y = MGHoldingLookDir.y - Corrx
    Mk3d.dx.VectorAdd MGHoldingPos, yEyes, MGHoldingLookDir
    MGHolding.MoveTo MGHoldingPos
    'calculate the angle of the MG which you are holding
    MGHolding.Rotate Mk3d.VectorMake(0, -MGHoldingAngle.y, 0)
    MGHolding.Rotate Mk3d.VectorMake(-MGHoldingAngle.x, 0, 0)
    MGHolding.Rotate Mk3d.VectorMake(yAngle.x, yAngle.y, 0)
    MGHoldingAngle = yAngle
    MGHoldingAngle.x = MGHoldingAngle.x
End Sub

Private Sub GameMan()
    Dim ManWalkSpeed!, ManRot!, ManAddDir As D3DVECTOR
    Dim ManPosBef As D3DVECTOR, ManGetColl As D3DVECTOR
    
    ManGoT = ManGoT + FrameT
    If ManGoT > 0.5 Then ManGoT = 0
    ManActionWaitT = ManActionWaitT + FrameT
    
    If ManState = MAN_BLENDIN Then
        'bot is shown
        If ManActionWaitT > BlendT Then
            ManState = MAN_ROTATE
            ManRotTo = Rnd * 6.283
            ManGoLen = Rnd * 30 + 5
            ManWentLen = 0
        End If
    ElseIf ManState = MAN_ROTATE Then
        'bot changes angle
        ManRot = GetAngleDiff(ManAngle.y, ManRotTo)
        If Abs(ManRot) < 0.1745 Then        '10 degree
            'bot starts walking
            ManState = MAN_GO
            ManGoT = 0
            ManWalkDir = Mk3d.VectorRotate(Mk3d.VectorMake(0, 0, -1), ManAngle)
        Else
            'bot needs still rotation
            ManRot = 6.283 * Sgn(ManRot) * FrameT / ManTimePer360
            ManAnim.Rotate Mk3d.VectorMake(0, ManRot, 0)
            'ManCalced is also changed because it's only a pointer
            ManAngle.y = ManAngle.y + ManRot
        End If
    ElseIf ManState = MAN_GO Then
        'bot walks
        ManWalkSpeed = WalkSpeed * FrameT
        ManWentLen = ManWentLen + ManWalkSpeed
        Mk3d.dx.VectorScale ManAddDir, ManWalkDir, ManWalkSpeed
        ManPosBef = ManPos
        Mk3d.dx.VectorAdd ManPos, ManPos, ManAddDir

        ManGetColl = GetCollDet(ManPosBef, ManPos, 1.5, True, False)
        If ManGetColl.x = 0 Or ManGetColl.z = 0 Then
            If ManGetColl.x = 0 And ManGetColl.z = 0 Then
                ManRotTo = ManAngle.y + 3.1415
            ElseIf ManGetColl.x = 0 Then
                'bot collids with something, in x-direction
                If Sgn(ManAddDir.x) = 1 Then
                    ManRotTo = GetRandom(3.1415, 6.283)
                Else
                    ManRotTo = GetRandom(0, 3.1415)
                End If
            ElseIf ManGetColl.z = 0 Then
                'bot collids with something, in z-direction
                If Sgn(ManAddDir.z) = 1 Then
                    ManRotTo = GetRandom(-1.57075, 1.57075)
                Else
                    ManRotTo = GetRandom(1.57075, 4.71225)
                End If
            End If
            ManPos = ManPosBef
            ManState = MAN_ROTATE
            ManGoLen = Rnd * 30 + 5
            ManWentLen = 0
            Set ManCalced = ManAnim.GetKeyFrameObj(0)
        ElseIf ManWentLen > ManGoLen Then
            'bot reaches his destination point
            ManState = MAN_ROTATE
            ManRotTo = Rnd * 6.283
            ManGoLen = Rnd * 30 + 5
            ManWentLen = 0
            Set ManCalced = ManAnim.GetKeyFrameObj(0)
        Else
            ManAnim.MoveTo ManPos
            If ManGoT < 0.125 Then
                Set ManCalced = ManAnim.CalcAnimObject(0, 1, ManGoT * 800)
            ElseIf ManGoT < 0.25 Then
                Set ManCalced = ManAnim.CalcAnimObject(1, 0, ManGoT * 800 - 100)
            ElseIf ManGoT < 0.375 Then
                Set ManCalced = ManAnim.CalcAnimObject(0, 2, ManGoT * 800 - 200)
            Else
                Set ManCalced = ManAnim.CalcAnimObject(2, 0, ManGoT * 800 - 300)
            End If
        End If
    ElseIf ManState = MAN_DIE Then
        'bot is dieing
        Set ManCalced = ManAnim.CalcAnimObject(0, 3, ManActionWaitT / ManFallT * 100)
        If ManActionWaitT > ManFallT Then
            Set ManCalced = ManAnim.GetKeyFrameObj(3)
            ManState = MAN_BLENDOUT
            ManActionWaitT = 0
            
            'create the blood
            If Not Blood.EffectVcnt = Blood.EffectVmax Then
                Blood.EffectFileAdd Mk3d.VectorMake(ManPos.x, 0.1, ManPos.z)
            End If
        End If
    Else
        'bot is blended out
        If ManActionWaitT > BlendT Then
            Set ManCalced = ManAnim.GetKeyFrameObj(0)
            ManState = MAN_BLENDIN
            ManActionWaitT = 0
            ManShotT = 0
            ManPos = GetRandomPos
            ManAnim.MoveTo ManPos
        End If
    End If
End Sub




Private Function GetRandomPos() As D3DVECTOR
    Dim Ready As Boolean, StPos As D3DVECTOR, GotColl As D3DVECTOR

    Do
        StPos.x = GetRandom(4, 54)
        StPos.z = GetRandom(10, 60)
        GotColl = GetCollDet(StPos, StPos, 2, True, False)
        If Not GotColl.x = 0 And Not GotColl.z = 0 Then
            Ready = True
            GetRandomPos = StPos
        End If
    Loop While Not Ready
End Function

Private Function GetTimeDiff(ByVal StartTime As Single, ByVal EndTime As Single) As Single
    If EndTime < StartTime Then
        EndTime = EndTime + 86400
    End If
    GetTimeDiff = EndTime - StartTime
End Function

Private Function GetAngleDiff(ByVal Angle1 As Single, ByVal Angle2 As Single) As Single
    Dim Result!
    
    Result = Angle2 - Angle1
    If Result > 3.1415 Then Result = Result - 6.283
    If Result < -3.1415 Then Result = Result + 6.283
    GetAngleDiff = Result
End Function

Private Function GetMoveAngle(KeybState As DIKEYBOARDSTATE) As Single
    'Key 1: ESC
    'Key 28, 156: Enter
    'Key 200: Cursor up
    'Key 203: Cursor left
    'Key 205: Cursor right
    'Key 208: Cursor down
    
    GetMoveAngle = -1
    If Not KeybState.Key(200) = 0 And Not KeybState.Key(203) = 0 Then
        'Cursor up and Cursor left
        GetMoveAngle = 0.785375
        Exit Function
    End If
    If Not KeybState.Key(203) = 0 And Not KeybState.Key(208) = 0 Then
        'Cursor left and Cursor down
        GetMoveAngle = 2.356125
        Exit Function
    End If
    If Not KeybState.Key(208) = 0 And Not KeybState.Key(205) = 0 Then
        'Cursor down and Cursor right
        GetMoveAngle = 3.926875
        Exit Function
    End If
    If Not KeybState.Key(205) = 0 And Not KeybState.Key(200) = 0 Then
        'Cursor right and Cursor up
        GetMoveAngle = 5.497625
        Exit Function
    End If
        
    If Not KeybState.Key(200) = 0 Then
        'Cursor up
        GetMoveAngle = 0
    End If
    If Not KeybState.Key(203) = 0 Then
        'Cursor left
        GetMoveAngle = 1.57075
    End If
    If Not KeybState.Key(208) = 0 Then
        'Cursor down
        GetMoveAngle = 3.1415
    End If
    If Not KeybState.Key(205) = 0 Then
        'Cursor right
        GetMoveAngle = 4.71225
    End If
End Function

Private Function GetCollDet(ActPos As D3DVECTOR, NewPos As D3DVECTOR, ByVal MaxDiff As Single, ByVal CheckYou As Boolean, ByVal CheckMan As Boolean) As D3DVECTOR
    Dim i%, Diff As D3DVECTOR, ChPosX As D3DVECTOR, ChPosZ As D3DVECTOR
    
    'only x and z have to be checked
    GetCollDet = Mk3d.VectorMake(1, 1, 1)
    
    'check if the point is inside the world
    If NewPos.x < MapArea(0, 0) + MaxDiff Or NewPos.x > MapArea(0, 1) - MaxDiff Then
        GetCollDet.x = 0
    End If
    If NewPos.z < MapArea(1, 0) + MaxDiff Or NewPos.z > MapArea(1, 1) - MaxDiff Then
        GetCollDet.z = 0
    End If
    If GetCollDet.x = 0 And GetCollDet.z = 0 Then Exit Function
    
    'checks if the point comes too close to a tree or to a wall
    Mk3d.dx.VectorSubtract Diff, NewPos, ActPos
    ChPosX = ActPos
    ChPosX.x = ChPosX.x + Diff.x
    ChPosZ = ActPos
    ChPosZ.z = ChPosZ.z + Diff.z
    For i = 0 To CollDetCount - 1
        If InArea(NewPos, CollDet(i), MaxDiff) Then
            If Not GetCollDet.x = 0 Then
                If InArea(ChPosX, CollDet(i), MaxDiff) Then GetCollDet.x = 0
            End If
            If Not GetCollDet.z = 0 Then
                If InArea(ChPosZ, CollDet(i), MaxDiff) Then GetCollDet.z = 0
            End If
            If GetCollDet.x = 0 And GetCollDet.z = 0 Then Exit Function
        End If
    Next i
    
    'checks if the point is too close to your positon
    If CheckYou Then
        If InArea(ChPosX, yPos, MaxDiff) Then GetCollDet.x = 0
        If InArea(ChPosZ, yPos, MaxDiff) Then GetCollDet.z = 0
    End If
    If GetCollDet.x = 0 And GetCollDet.z = 0 Then Exit Function
    
    'checks if the point is too close to the bot's position
    If CheckMan And Not ManState = MAN_BLENDOUT Then
        If InArea(ChPosX, ManPos, MaxDiff) Then GetCollDet.x = 0
        If InArea(ChPosZ, ManPos, MaxDiff) Then GetCollDet.z = 0
    End If
End Function

Private Function GetPatronCollDet(PatronPos As D3DVECTOR, ByVal MaxDiff As Single) As MGPatronHitEnum
    Dim i%
    Dim HitNow As Boolean
    
    If InArea(PatronPos, Mk3d.VectorMake(3.5, 0, 12.5), 1) And Not HitSpecialW(0) Then
        'hit Special Window 1
        HitSpecialW(0) = True
        HitNow = True
    End If
    If InArea(PatronPos, Mk3d.VectorMake(3.5, 0, 57.5), 1) And Not HitSpecialW(1) Then
        'hit Special Window 2
        HitSpecialW(1) = True
        HitNow = True
    End If
    If HitSpecialW(0) And HitSpecialW(1) And HitNow And ActPlayT < 10 Then
        GamePoints = GamePoints + 100
        TextBlendWaitT = 0
        ShowText = True
        TextToShow = "WINDOW SPECIAL: + 100"
    End If
    
    If InArea(PatronPos, ManPos, MaxDiff) And PatronPos.y < 3 Then
        GetPatronCollDet = MGPATRON_HITMAN
        Exit Function
    End If
    If PatronPos.x < MapArea(0, 0) Or PatronPos.x > MapArea(0, 1) Or PatronPos.y < 0 Or PatronPos.y > 3 Or PatronPos.z < MapArea(1, 0) Or PatronPos.z > MapArea(1, 1) Then
        GetPatronCollDet = MGPATRON_HITLANDSCAPE
        Exit Function
    End If
    For i = 0 To CollDetCount - 1
        If InArea(PatronPos, CollDet(i), MaxDiff) And PatronPos.y < 6 Then
            GetPatronCollDet = MGPATRON_HITLANDSCAPE
            Exit Function
        End If
    Next i
    GetPatronCollDet = MGPATRON_HITNOTHING
End Function

Private Function GetRandom(ByVal Min As Single, ByVal Max As Single) As Single
    GetRandom = Rnd * (Max - Min) + Min
End Function

Private Function InArea(Pos As D3DVECTOR, AreaPos As D3DVECTOR, ByVal AreaSize As Single) As Boolean
    If Pos.x > AreaPos.x - AreaSize And Pos.x < AreaPos.x + AreaSize And Pos.z > AreaPos.z - AreaSize And Pos.z < AreaPos.z + AreaSize Then
        InArea = True
    End If
End Function
