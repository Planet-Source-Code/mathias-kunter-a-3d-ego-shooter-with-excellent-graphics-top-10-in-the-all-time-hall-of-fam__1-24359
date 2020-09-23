VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mk3dObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public VertexCount&, VertexTOCCount%, LightCount%, TextureCount%, MaterialCount%
Public VertexFile$, LightFile$, TextureFile$, MaterialFile$
Private RenderV() As D3DVERTEX
Private PolyTOC() As Mk3dTOC
Private Lights() As D3DLIGHT7, LightsInd%()
Private Textures() As DirectDrawSurface7
Private Materials() As D3DMATERIAL7

Private WorldCoord As D3DVECTOR


Public Function CreateFromFile(ByVal FileName As String, LightEnable As Boolean) As Boolean
    Dim i%, j%, k%, cnt%
    Dim Polys() As Mk3dPolygon, PolyCount%
    Dim ReadPath$, StaPos%, ActTex%, ActMat%, ActCull As CONST_D3DCULL
    Dim SurfaceDesc As DDSURFACEDESC2, TextureEnum As Direct3DEnumPixelFormats
    
    On Local Error GoTo Failed
    
    'load general object data
    Open FileName For Input As #1
    Input #1, PolyCount
    Input #1, TextureFile
    Input #1, MaterialFile
    If PolyCount <> 0 Then ReDim Polys(PolyCount - 1)
    'load object data
    For i = 0 To PolyCount - 1
        With Polys(i)
            Input #1, .CullMode
            Input #1, .TextureIndex
            Input #1, .MaterialIndex
            For j = 0 To 1
                For k = 0 To 2
                    Input #1, .Tri(j).p(k).x
                    Input #1, .Tri(j).p(k).y
                    Input #1, .Tri(j).p(k).z
                    Input #1, .Tri(j).p(k).tu
                    Input #1, .Tri(j).p(k).tv
                Next k
            Next j
        End With
    Next i
    Close #1
    VertexFile = FileName
    TextureFile = App.Path & "\Textures\" & TextureFile
    MaterialFile = App.Path & "\Materials\" & MaterialFile
    
    'sort polys by 1) textur, 2) material, 3) cull
    'this is done because of the performance
    PolySort Polys, 0, PolyCount - 1, 1                                     'sort by textures
    If Not PolyCount = 0 Then ActTex = Polys(0).TextureIndex                'sort by materials
    For i = 0 To PolyCount - 1
        If i = PolyCount - 1 Then
            PolySort Polys, StaPos, i, 2
        ElseIf Not Polys(i).TextureIndex = ActTex Then
            'Neuer CullMode beginnt hier
            PolySort Polys, StaPos, i - 1, 2
            StaPos = i
            ActTex = Polys(i).TextureIndex
        End If
    Next i
    If Not PolyCount = 0 Then                                               'sort by cull
        ActTex = Polys(0).TextureIndex
        ActMat = Polys(0).MaterialIndex
    End If
    StaPos = 0
    For i = 0 To PolyCount - 1
        If i = PolyCount - 1 Then
            PolySort Polys, StaPos, i, 3
        ElseIf Not Polys(i).TextureIndex = ActTex Or Not Polys(i).MaterialIndex = ActMat Then
            'new textur or new cull starts here
            PolySort Polys, StaPos, i - 1, 3
            StaPos = i
            ActTex = Polys(i).TextureIndex
            ActMat = Polys(i).MaterialIndex
        End If
    Next i
    
    'set up the render-vertices
    VertexCount = PolyCount * 6
    If VertexCount <> 0 Then ReDim RenderV(VertexCount - 1)
    For i = 0 To PolyCount - 1
        For j = 0 To 1
            For k = 0 To 2
                RenderV(cnt) = Polys(i).Tri(j).p(k)
                cnt = cnt + 1
            Next k
        Next j
    Next i
    
    'create a TOC of the render-vertices
    'get the size of the TOC
    If Not PolyCount = 0 Then
        ActTex = Polys(0).TextureIndex
        ActMat = Polys(0).MaterialIndex
        ActCull = Polys(0).CullMode
    End If
    StaPos = 0
    VertexTOCCount = 0
    For i = 0 To PolyCount - 1
        If Not Polys(i).TextureIndex = ActTex Or Not Polys(i).MaterialIndex = ActMat Or Not Polys(i).CullMode = ActCull Then
            VertexTOCCount = VertexTOCCount + 1
            ActTex = Polys(i).TextureIndex
            ActMat = Polys(i).MaterialIndex
            ActCull = Polys(i).CullMode
        End If
    Next i
    VertexTOCCount = VertexTOCCount + 1
    'write TOC
    If Not PolyCount = 0 Then
        ActTex = Polys(0).TextureIndex
        ActMat = Polys(0).MaterialIndex
        ActCull = Polys(0).CullMode
    End If
    cnt = 0
    If Not VertexTOCCount = 0 Then ReDim PolyTOC(VertexTOCCount - 1)
    For i = 0 To PolyCount - 1
        If i = PolyCount - 1 Then
            PolyTOC(cnt).StartIndex = StaPos * 6
            PolyTOC(cnt).EndIndex = i * 6 + 5
            PolyTOC(cnt).TexturIndex = Polys(StaPos).TextureIndex
            PolyTOC(cnt).MaterialIndex = Polys(StaPos).MaterialIndex
            PolyTOC(cnt).CullMode = Polys(StaPos).CullMode
        ElseIf Not Polys(i).TextureIndex = ActTex Or Not Polys(i).MaterialIndex = ActMat Or Not Polys(i).CullMode = ActCull Then
            PolyTOC(cnt).StartIndex = StaPos * 6
            PolyTOC(cnt).EndIndex = i * 6 - 1                  'from (i - 1) * 6 + 5
            PolyTOC(cnt).TexturIndex = Polys(StaPos).TextureIndex
            PolyTOC(cnt).MaterialIndex = Polys(StaPos).MaterialIndex
            PolyTOC(cnt).CullMode = Polys(StaPos).CullMode
            cnt = cnt + 1
            StaPos = i
            ActTex = Polys(i).TextureIndex
            ActMat = Polys(i).MaterialIndex
            ActCull = Polys(i).CullMode
        End If
    Next i
    
    'load lights
    LightFile = Left(FileName, Len(FileName) - 3) & "lig"
    Open LightFile For Input As #1
    Input #1, LightCount
    If LightCount <> 0 Then
        ReDim Lights(LightCount - 1)
        ReDim LightsInd(LightCount - 1)
    End If
    For i = 0 To LightCount - 1
        With Lights(i)
            Input #1, .Ambient.a
            Input #1, .Ambient.r
            Input #1, .Ambient.g
            Input #1, .Ambient.b
            Input #1, .attenuation0
            Input #1, .attenuation1
            Input #1, .attenuation2
            Input #1, .diffuse.a
            Input #1, .diffuse.r
            Input #1, .diffuse.g
            Input #1, .diffuse.b
            Input #1, .direction.x
            Input #1, .direction.y
            Input #1, .direction.z
            Input #1, .dltType
            Input #1, .falloff
            Input #1, .phi
            Input #1, .Position.x
            Input #1, .Position.y
            Input #1, .Position.z
            Input #1, .range
            Input #1, .specular.a
            Input #1, .specular.r
            Input #1, .specular.g
            Input #1, .specular.b
            Input #1, .theta
        End With
        If LightEnable Then LightsInd(i) = Mk3d.LightAdd(Lights(i))
    Next i
    Close #1
    
    'load textures
    Open TextureFile For Input As #1
    Input #1, TextureCount
    If TextureCount <> 0 Then ReDim Textures(TextureCount - 1)
    For i = 0 To TextureCount - 1
        Input #1, ReadPath
        ReadPath = App.Path & "\Textures\" & ReadPath
        Set TextureEnum = Mk3d.d3dDevice.GetTextureFormatsEnum()
        TextureEnum.GetItem 1, SurfaceDesc.ddpfPixelFormat
        SurfaceDesc.ddsCaps.lCaps = DDSCAPS_TEXTURE
        Set Textures(i) = Mk3d.dd.CreateSurfaceFromFile(ReadPath, SurfaceDesc)
    Next i
    Close #1
    
    'load materials
    Open MaterialFile For Input As #1
    Input #1, MaterialCount
    If MaterialCount <> 0 Then ReDim Materials(MaterialCount - 1)
    For i = 0 To MaterialCount - 1
        With Materials(i)
            Input #1, .Ambient.a
            Input #1, .Ambient.r
            Input #1, .Ambient.g
            Input #1, .Ambient.b
            
            Input #1, .diffuse.a
            Input #1, .diffuse.r
            Input #1, .diffuse.g
            Input #1, .diffuse.b
            
            Input #1, .emissive.a
            Input #1, .emissive.r
            Input #1, .emissive.g
            Input #1, .emissive.b
            
            Input #1, .power
            
            Input #1, .specular.a
            Input #1, .specular.r
            Input #1, .specular.g
            Input #1, .specular.b
        End With
    Next i
    Close #1
    CreateFromFile = True
    Exit Function
    
Failed:
    CreateFromFile = False
    Close
End Function


Private Sub PolySort(Polys() As Mk3dPolygon, ByVal StartIndex As Integer, ByVal EndIndex As Integer, ByVal Criteria As Integer)
    Dim i%, j%
    
    For i = StartIndex To EndIndex
        For j = i To EndIndex
            Select Case Criteria
                Case 1          'Textures
                    If Polys(j).TextureIndex < Polys(i).TextureIndex Then
                        PolySwap Polys, i, j
                    End If
                Case 2          'Materials
                    If Polys(j).MaterialIndex < Polys(i).MaterialIndex Then
                        PolySwap Polys, i, j
                    End If
                Case 3          'CullMode
                    If Polys(j).CullMode < Polys(i).CullMode Then
                        PolySwap Polys, i, j
                    End If
            End Select
        Next j
    Next i
End Sub


Private Sub PolySwap(Polys() As Mk3dPolygon, ByVal Index1 As Integer, ByVal Index2 As Integer)
    Dim SwapPoly As Mk3dPolygon
    
    SwapPoly = Polys(Index1)
    Polys(Index1) = Polys(Index2)
    Polys(Index2) = SwapPoly
End Sub



'CHANGING THE OBJECT

Public Sub Move(MoveDelta As D3DVECTOR)
    dx.VectorAdd WorldCoord, WorldCoord, MoveDelta
End Sub

Public Sub MoveTo(Position As D3DVECTOR)
    WorldCoord = Position
End Sub

Public Sub VertexMove(MoveDelta As D3DVECTOR)
    Dim i&
    
    For i = 0 To VertexCount - 1
        With RenderV(i)
            .x = .x + MoveDelta.x
            .y = .y + MoveDelta.y
            .z = .z + MoveDelta.z
        End With
    Next i
End Sub

Public Sub Central(ByVal CenX As Boolean, ByVal CenY As Boolean, ByVal CenZ As Boolean)
    Dim i&, Decr As D3DVECTOR, Incr As D3DVECTOR, MP As D3DVECTOR
    
    For i = 0 To VertexCount - 1
        If RenderV(i).x < Decr.x Or i = 0 Then Decr.x = RenderV(i).x
        If RenderV(i).y < Decr.y Or i = 0 Then Decr.y = RenderV(i).y
        If RenderV(i).z < Decr.z Or i = 0 Then Decr.z = RenderV(i).z
        If RenderV(i).x > Incr.x Or i = 0 Then Incr.x = RenderV(i).x
        If RenderV(i).y > Incr.y Or i = 0 Then Incr.y = RenderV(i).y
        If RenderV(i).z > Incr.z Or i = 0 Then Incr.z = RenderV(i).z
    Next i
    Mk3d.dx.VectorAdd MP, Decr, Incr
    Mk3d.dx.VectorScale MP, MP, 0.5
    For i = 0 To VertexCount - 1
        If CenX Then RenderV(i).x = RenderV(i).x - MP.x
        If CenY Then RenderV(i).y = RenderV(i).y - MP.y
        If CenZ Then RenderV(i).z = RenderV(i).z - MP.z
    Next i
End Sub

Public Sub Rotate(RotAngle As D3DVECTOR)
    Dim i&, j%, k%
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
    
    For i = 0 To VertexCount - 1
        With RenderV(i)
            dx.IdentityMatrix PMat
            PMat.rc41 = .x
            PMat.rc42 = .y
            PMat.rc43 = .z
            
            dx.MatrixMultiply PMat, PMat, DstMat
            .x = PMat.rc41
            .y = PMat.rc42
            .z = PMat.rc43
        End With
    Next i
End Sub




'CHECK-ROUTINES

Public Function GetVertex(ByVal Index As Integer) As D3DVERTEX
    GetVertex = RenderV(Index)
End Function

Public Function GetVertexWorld(ByVal Index As Integer) As D3DVERTEX
    GetVertexWorld = RenderV(Index)
    With GetVertexWorld
        .x = .x + WorldCoord.x
        .y = .y + WorldCoord.y
        .z = .z + WorldCoord.z
    End With
End Function

Public Function GetRenderVertexWorld() As D3DVERTEX()
    Dim RetV() As D3DVERTEX, i&
    
    RetV = RenderV
    For i = 0 To VertexCount - 1
        With RetV(i)
            .x = .x + WorldCoord.x
            .y = .y + WorldCoord.y
            .z = .z + WorldCoord.z
        End With
    Next i
    GetRenderVertexWorld = RetV
End Function

Public Function GetRenderVertex() As D3DVERTEX()
    GetRenderVertex = RenderV
End Function

Public Sub PutVertex(ByVal Index As Integer, WrVertex As D3DVERTEX)
    RenderV(Index) = WrVertex
End Sub

Public Sub PutVector(ByVal Index As Integer, WrVector As D3DVECTOR)
    With RenderV(Index)
        .x = WrVector.x
        .y = WrVector.y
        .z = WrVector.z
    End With
End Sub


Public Function GetPosition() As D3DVECTOR
    GetPosition = WorldCoord
End Function


Public Function GetLight(ByVal Index As Integer) As D3DLIGHT7
    GetLight = Lights(Index)
End Function

Public Function PutLight(ByVal Index As Integer, WrLight As D3DLIGHT7)
    Lights(Index) = WrLight
End Function


Public Function GetTexture(ByVal Index As Integer) As DirectDrawSurface7
    Set GetTexture = Textures(Index)
End Function


Public Function GetMaterial(ByVal Index As Integer) As D3DMATERIAL7
    GetMaterial = Materials(Index)
End Function


Public Function GetTOCStart(ByVal Index As Integer) As Integer
    GetTOCStart = PolyTOC(Index).StartIndex
End Function

Public Function GetTOCEnd(ByVal Index As Integer) As Integer
    GetTOCEnd = PolyTOC(Index).EndIndex
End Function

Public Function GetTOCTex(ByVal Index As Integer) As Integer
    GetTOCTex = PolyTOC(Index).TexturIndex
End Function

Public Function GetTOCMat(ByVal Index As Integer) As Integer
    GetTOCMat = PolyTOC(Index).MaterialIndex
End Function

Public Function GetTOCCull(ByVal Index As Integer) As CONST_D3DCULL
    GetTOCCull = PolyTOC(Index).CullMode
End Function
