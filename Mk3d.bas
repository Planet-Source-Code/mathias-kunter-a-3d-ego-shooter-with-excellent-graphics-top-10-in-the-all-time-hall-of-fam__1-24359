Attribute VB_Name = "Mk3d"
Option Explicit

Public dx As New DirectX7

Public dd As DirectDraw7
Public PrimarySurf As DirectDrawSurface7
Public BackBufferSurf As DirectDrawSurface7

Public d3d As Direct3D7
Public d3dDevice As Direct3DDevice7
Public d3drcViewport(0) As D3DRECT
Public RenderState As RenderEnum
Public VPAbleSize%()
Public VPSize%(1)

Public di As DirectInput
Public diDeviceKeyb As DirectInputDevice, diDeviceMouse As DirectInputDevice

Public ds As DirectSound
Public dsWalkSound As DirectSoundBuffer
Public dsShootSound As DirectSoundBuffer


Private ActRenderState$
Private d3dLightCount%

Public Enum RenderEnum
    RGB
    HAL
    NONE
End Enum

Public Type Mk3dTriangle
    p(2) As D3DVERTEX
End Type

Public Type Mk3dPolygon
    TextureIndex As Integer
    MaterialIndex As Integer
    CullMode As CONST_D3DCULL
    Tri(1) As Mk3dTriangle
End Type

Public Type Mk3dTOC
    StartIndex As Integer
    EndIndex As Integer
    TexturIndex As Integer
    MaterialIndex As Integer
    CullMode As CONST_D3DCULL
End Type

Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long




'***************************************INIT-ROUTINES***************************************

Public Function InitDX() As Boolean
    Dim i%, j%, cnt%
    Dim DevEnum As Direct3DEnumDevices
    Dim DisplayModes As DirectDrawEnumModes, DisplayDesc As DDSURFACEDESC2
    Dim ExAlr As Boolean, TmpSize%()
    
    On Local Error GoTo Failed
    Set dd = dx.DirectDrawCreate("")
    Set d3d = dd.GetDirect3D
    Set di = dx.DirectInputCreate
    
    'Software or Hardware Rendering
    Set DevEnum = d3d.GetDevicesEnum
    RenderState = NONE
    For i = 1 To DevEnum.GetCount
        If DevEnum.GetGuid(i) = "IID_IDirect3DRGBDevice" Then
            RenderState = RGB
        ElseIf DevEnum.GetGuid(i) = "IID_IDirect3DHALDevice" Then
            RenderState = HAL
            Exit For
        End If
    Next i
    If RenderState = HAL Then
        ActRenderState = "IID_IDirect3DHALDevice"
    ElseIf RenderState = RGB Then
        ActRenderState = "IID_IDirect3DRGBDevice"
    Else
        GoTo Failed
    End If
    
    'get possible resolutions
    Set DisplayModes = dd.GetDisplayModesEnum(DDEDM_DEFAULT, DisplayDesc)
    ReDim TmpSize(DisplayModes.GetCount - 1, 1)
    For i = 1 To DisplayModes.GetCount
        DisplayModes.GetItem i, DisplayDesc
        ExAlr = False
        For j = 0 To UBound(TmpSize) - 1
            If TmpSize(j, 0) = DisplayDesc.lWidth And TmpSize(j, 1) = DisplayDesc.lHeight Then
                ExAlr = True
                Exit For
            End If
        Next j
        
        If Not ExAlr And DisplayDesc.lHeight >= 480 Then
            TmpSize(cnt, 0) = DisplayDesc.lWidth
            TmpSize(cnt, 1) = DisplayDesc.lHeight
            cnt = cnt + 1
        End If
    Next i
    ReDim VPAbleSize(cnt - 1, 1)
    For i = 0 To cnt - 1
        VPAbleSize(i, 0) = TmpSize(i, 0)
        VPAbleSize(i, 1) = TmpSize(i, 1)
    Next i
    InitDX = True
    Exit Function
    
Failed:
    InitDX = False
End Function

Public Function InitDDraw(GameFont As IFont) As Boolean
    Dim i%
    Dim DescPrim As DDSURFACEDESC2, CapsRB As DDSCAPS2, DescBack As DDSURFACEDESC2
    Dim ZEnum As Direct3DEnumPixelFormats, FoundZB As Boolean, DescZ As DDSURFACEDESC2, ZBufferPixelF As DDPIXELFORMAT
    Dim ZBufferSurf As DirectDrawSurface7
    
    On Local Error GoTo Failed
    dd.SetCooperativeLevel RenderForm.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE
    dd.SetDisplayMode VPSize(0), VPSize(1), 16, 0, DDSDM_DEFAULT
    
    'primary surface
    DescPrim.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    DescPrim.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_3DDEVICE Or DDSCAPS_COMPLEX Or DDSCAPS_FLIP
    DescPrim.lBackBufferCount = 1
    Set PrimarySurf = dd.CreateSurface(DescPrim)
    PrimarySurf.SetFont GameFont
    
    'backbuffer surface
    CapsRB.lCaps = DDSCAPS_BACKBUFFER
    Set BackBufferSurf = PrimarySurf.GetAttachedSurface(CapsRB)
    BackBufferSurf.GetSurfaceDesc DescBack
    BackBufferSurf.SetFont GameFont
    
    'Z-Buffer Surface
    Set ZEnum = d3d.GetEnumZBufferFormats(ActRenderState)
    For i = 1 To ZEnum.GetCount()
        ZEnum.GetItem i, ZBufferPixelF
        If ZBufferPixelF.lFlags = DDPF_ZBUFFER Then
            FoundZB = True
            Exit For
        End If
    Next i
    If FoundZB Then
        DescZ.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT Or DDSD_PIXELFORMAT
        DescZ.ddsCaps.lCaps = DDSCAPS_ZBUFFER
        DescZ.lWidth = VPSize(0)
        DescZ.lHeight = VPSize(1)
        DescZ.ddpfPixelFormat = ZBufferPixelF
        Set ZBufferSurf = dd.CreateSurface(DescZ)
        BackBufferSurf.AddAttachedSurface ZBufferSurf
    Else
        GoTo Failed
    End If
    InitDDraw = True
    Exit Function
    
Failed:
    InitDDraw = False
End Function

Public Function InitD3D() As Boolean
    Dim vPort As D3DVIEWPORT7
    
    On Local Error GoTo Failed
    Set d3dDevice = d3d.CreateDevice(ActRenderState, BackBufferSurf)
    With vPort
        .lX = 0
        .lY = 0
        .lWidth = VPSize(0)
        .lHeight = VPSize(1)
        .minz = 0
        .maxz = 1
    End With
    d3dDevice.SetViewport vPort
    With d3drcViewport(0)
        .X1 = 0
        .Y1 = 0
        .X2 = VPSize(0)
        .Y2 = VPSize(1)
    End With
    
    With d3dDevice
        .SetRenderTarget BackBufferSurf
        
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTFG_LINEAR
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTFG_LINEAR
        
        .SetRenderState D3DRENDERSTATE_AMBIENT, dx.CreateColorRGB(0.2, 0.2, 0.2)
    End With
    InitD3D = True
    Exit Function
    
Failed:
    InitD3D = False
End Function

Public Function InitDInput() As Boolean
    Dim diDevEnum As DirectInputEnumDevices
    
    On Local Error GoTo Failed
    
    'init mouse
    Set diDevEnum = di.GetDIEnumDevices(DIDEVTYPE_MOUSE, DIEDFL_ATTACHEDONLY)
    If diDevEnum.GetCount = 0 Then
        MsgBox "The system couldn't find a mouse.", vbCritical
        Mk3d.ExitDX
        End
    End If
    Set diDeviceMouse = di.CreateDevice("GUID_SysMouse")
    diDeviceMouse.SetCommonDataFormat DIFORMAT_MOUSE
    diDeviceMouse.SetCooperativeLevel RenderForm.hWnd, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND
    
    'init keyboard
    Set diDevEnum = di.GetDIEnumDevices(DIDEVTYPE_KEYBOARD, DIEDFL_ATTACHEDONLY)
    If diDevEnum.GetCount = 0 Then
        MsgBox "The system couldn't find a keyboard.", vbCritical
        Mk3d.ExitDX
        End
    End If
    Set diDeviceKeyb = di.CreateDevice("GUID_SysKeyboard")
    diDeviceKeyb.SetCommonDataFormat DIFORMAT_KEYBOARD
    diDeviceKeyb.SetCooperativeLevel RenderForm.hWnd, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND
    diDeviceKeyb.Acquire
    InitDInput = True
    Exit Function
    
Failed:
    InitDInput = False
End Function

Public Function InitDSound() As Boolean
    Dim SoundDesc As DSBUFFERDESC, WavFormat As WAVEFORMATEX
    
    On Local Error GoTo Failed
       
    Set ds = dx.DirectSoundCreate("")
    ds.SetCooperativeLevel RenderForm.hWnd, DSSCL_NORMAL
    
    'Shooting sound
    Set dsShootSound = ds.CreateSoundBufferFromFile(App.Path & "\Sounds\Shoot.wav", SoundDesc, WavFormat)
    
    'Walking sound
    Set dsWalkSound = ds.CreateSoundBufferFromFile(App.Path & "\Sounds\Walk.wav", SoundDesc, WavFormat)
    
    InitDSound = True
    Exit Function
    
Failed:
    InitDSound = False
End Function

Public Sub ExitDX()
    On Local Error Resume Next
    ShowCursor True
    diDeviceKeyb.Unacquire
    diDeviceMouse.Unacquire
    dd.RestoreDisplayMode
    dd.SetCooperativeLevel RenderForm.hWnd, DDSCL_NORMAL Or DDSCL_EXCLUSIVE
    Set dx = Nothing
End Sub




Public Sub SetClipPlane(ByVal Near As Single, ByVal Far As Single)
    Dim matProj As D3DMATRIX
    
    dx.IdentityMatrix matProj
    dx.ProjectionMatrix matProj, Near, Far, 1.570795
    d3dDevice.SetTransform D3DTRANSFORMSTATE_PROJECTION, matProj
End Sub

Public Sub SetCamera(CameraPosition As D3DVECTOR, CameraLookAt As D3DVECTOR)
    Dim matView As D3DMATRIX
    
    dx.IdentityMatrix matView
    dx.ViewMatrix matView, CameraPosition, CameraLookAt, VectorMake(0, 1, 0), 0
    d3dDevice.SetTransform D3DTRANSFORMSTATE_VIEW, matView
End Sub


'***************************************LIGHT-ROUTINES***************************************

Public Function LightAdd(Light As D3DLIGHT7) As Integer
    d3dDevice.SetLight d3dLightCount, Light
    d3dDevice.LightEnable d3dLightCount, True
    LightAdd = d3dLightCount
    d3dLightCount = d3dLightCount + 1
End Function

Public Sub LightDel(ByVal Index As Integer)
    Dim i%
    Dim LightBef As D3DLIGHT7
    
    For i = Index To d3dLightCount - 2
        d3dDevice.GetLight i + 1, LightBef
        d3dDevice.SetLight i, LightBef
        d3dDevice.LightEnable i, d3dDevice.GetLightEnable(i + 1)
    Next i
    d3dDevice.LightEnable d3dLightCount - 1, False
    d3dLightCount = d3dLightCount - 1
End Sub

Public Sub LightUpdate(ByVal Index As Integer, Light As D3DLIGHT7)
    d3dDevice.SetLight Index, Light
End Sub

Public Sub LightSetState(ByVal Index As Integer, ByVal State As Boolean)
    d3dDevice.LightEnable Index, State
End Sub



'***************************************VECTOR AND VERTEX-ROUTINES***************************************

Public Function VectorMake(ByVal x As Single, ByVal y As Single, ByVal z As Single) As D3DVECTOR
    With VectorMake
        .x = x
        .y = y
        .z = z
    End With
End Function

Public Function VertexToVector(Vertex As D3DVERTEX) As D3DVECTOR
    With VertexToVector
        .x = Vertex.x
        .y = Vertex.y
        .z = Vertex.z
    End With
End Function

Public Function VectorToVertex(Vector As D3DVECTOR) As D3DVERTEX
    With VectorToVertex
        .x = Vector.x
        .y = Vector.y
        .z = Vector.z
    End With
End Function

Public Function VectorToFilledVertex(Vector As D3DVECTOR, Vertex As D3DVERTEX) As D3DVERTEX
    VectorToFilledVertex = Vertex
    With VectorToFilledVertex
        .x = Vector.x
        .y = Vector.y
        .z = Vector.z
    End With
End Function

Public Function VectorRotate(Vector As D3DVECTOR, RotAngle As D3DVECTOR) As D3DVECTOR
    Dim i%
    Dim RMat As D3DMATRIX, PMat As D3DMATRIX, DstMat As D3DMATRIX
    
    Mk3d.dx.IdentityMatrix PMat
    Mk3d.dx.IdentityMatrix DstMat
    For i = 0 To 2
        Mk3d.dx.IdentityMatrix RMat
        If i = 0 Then
            Mk3d.dx.RotateXMatrix RMat, RotAngle.x
        ElseIf i = 1 Then
            Mk3d.dx.RotateYMatrix RMat, RotAngle.y
        Else
            Mk3d.dx.RotateZMatrix RMat, RotAngle.z
        End If
        Mk3d.dx.MatrixMultiply DstMat, DstMat, RMat
    Next i
    
    dx.IdentityMatrix PMat
    PMat.rc41 = Vector.x
    PMat.rc42 = Vector.y
    PMat.rc43 = Vector.z
    
    dx.MatrixMultiply PMat, PMat, DstMat
    VectorRotate.x = PMat.rc41
    VectorRotate.y = PMat.rc42
    VectorRotate.z = PMat.rc43
End Function




'***************************************RENDER-ROUTINE***************************************

Public Sub Render(Obj As Mk3dObject)
    Dim i%, j%, k%
    Dim RenderV() As D3DVERTEX, RenderStart&, RenderLen&
    Dim LastTex%, LastMat%, LastCull As CONST_D3DCULL
    Dim ReadTex%, ReadMat%, ReadCull As CONST_D3DCULL
    
    RenderV = Obj.GetRenderVertexWorld
    For i = 0 To Obj.VertexTOCCount - 1
        'get the texture
        ReadTex = Obj.GetTOCTex(i)
        If Not ReadTex = LastTex Or i = 0 Then
            d3dDevice.SetTexture 0, Obj.GetTexture(ReadTex)
            LastTex = ReadTex
        End If
        
        'get the material
        ReadMat = Obj.GetTOCMat(i)
        If Not ReadMat = LastMat Or i = 0 Then
            d3dDevice.SetMaterial Obj.GetMaterial(ReadMat)
            LastMat = ReadMat
        End If
        
        'get the cullmode
        ReadCull = Obj.GetTOCCull(i)
        If Not ReadCull = LastCull Or i = 0 Then
            d3dDevice.SetRenderState D3DRENDERSTATE_CULLMODE, ReadCull
            LastCull = ReadCull
        End If
        
        RenderStart = Obj.GetTOCStart(i)
        RenderLen = Obj.GetTOCEnd(i) - RenderStart + 1
        d3dDevice.DrawPrimitive D3DPT_TRIANGLELIST, D3DFVF_VERTEX, RenderV(RenderStart), RenderLen, D3DDP_DEFAULT
    Next i
End Sub

Public Sub RenderEffect(Effect As Mk3dEffectObject)
    Dim EffectV() As D3DVERTEX
    
    EffectV = Effect.GetEffectVertex
    d3dDevice.SetTexture 0, Effect.EffectTex
    d3dDevice.SetMaterial Effect.MaterialGet
    d3dDevice.SetRenderState D3DRENDERSTATE_CULLMODE, D3DCULL_NONE
    d3dDevice.DrawPrimitive D3DPT_TRIANGLELIST, D3DFVF_VERTEX, EffectV(0), Effect.EffectVcnt, D3DDP_DEFAULT
End Sub
