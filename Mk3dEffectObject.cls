VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mk3dEffectObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private EffectV() As D3DVERTEX, FileV() As D3DVERTEX
Public EffectVcnt&, EffectVmax&, FileVcnt%
Public EffectTex As DirectDrawSurface7
Private EffectMat As D3DMATERIAL7


Public Sub Initsialize(ByVal MaxVertex As Long)
    EffectVmax = MaxVertex
    ReDim EffectV(MaxVertex - 1)
End Sub

Public Sub TextureSet(ByVal TexturePath As String)
    Dim SurfaceDesc As DDSURFACEDESC2, TextureEnum As Direct3DEnumPixelFormats
    
    On Local Error Resume Next
    Set TextureEnum = Mk3d.d3dDevice.GetTextureFormatsEnum()
    TextureEnum.GetItem 1, SurfaceDesc.ddpfPixelFormat
    SurfaceDesc.ddsCaps.lCaps = DDSCAPS_TEXTURE
    Set EffectTex = Mk3d.dd.CreateSurfaceFromFile(TexturePath, SurfaceDesc)
End Sub

Public Sub MaterialSet(ByVal MaterialPath As String, ByVal MatIndex As Integer)
    Dim TmpCnt%, i%, RdMat As D3DMATERIAL7
    
    Open MaterialPath For Input As #1
    Input #1, TmpCnt
    For i = 0 To TmpCnt - 1
        With RdMat
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
        If i = MatIndex Then Exit For
    Next i
    Close #1
    EffectMat = RdMat
End Sub

Public Function MaterialGet() As D3DMATERIAL7
    MaterialGet = EffectMat
End Function

Public Sub EffectVertexSet(ByVal Index As Long, EffectVertex As D3DVERTEX)
    EffectV(Index) = EffectVertex
    EffectVcnt = EffectVcnt + 1
End Sub

Public Sub EffectVertexAdd(EffectVertex As D3DVERTEX)
    EffectV(EffectVcnt) = EffectVertex
    EffectVcnt = EffectVcnt + 1
End Sub

Public Sub EffectFloorAdd(MP As D3DVECTOR, r As Single)
    With EffectV(EffectVcnt)
        .x = MP.x - r
        .y = MP.y
        .z = MP.z + r
    End With
    EffectVcnt = EffectVcnt + 1
    With EffectV(EffectVcnt)
        .x = MP.x + r
        .y = MP.y
        .z = MP.z + r
        .tu = 1
    End With
    EffectVcnt = EffectVcnt + 1
    With EffectV(EffectVcnt)
        .x = MP.x + r
        .y = MP.y
        .z = MP.z - r
        .tu = 1
        .tv = 1
    End With
    EffectVcnt = EffectVcnt + 1
    EffectV(EffectVcnt) = EffectV(EffectVcnt - 1)
    EffectVcnt = EffectVcnt + 1
    With EffectV(EffectVcnt)
        .x = MP.x - r
        .y = MP.y
        .z = MP.z - r
        .tv = 1
    End With
    EffectVcnt = EffectVcnt + 1
    EffectV(EffectVcnt) = EffectV(EffectVcnt - 5)
    EffectVcnt = EffectVcnt + 1
End Sub

Public Sub EffectFileLoad(ByVal FileName As String)
    Dim i%, j%, k%, cnt%
    Dim RdPolyCount%, RdString$, RdPoly As Mk3dPolygon
    
    'load general object data
    Open FileName For Input As #1
    Input #1, RdPolyCount
    Input #1, RdString
    Input #1, RdString
    If RdPolyCount <> 0 Then ReDim FileV(RdPolyCount * 6 - 1)
    FileVcnt = RdPolyCount * 6
    'load object data
    For i = 0 To RdPolyCount - 1
        With RdPoly
            Input #1, .CullMode
            Input #1, .TextureIndex
            Input #1, .MaterialIndex
            For j = 0 To 1
                For k = 0 To 2
                    Input #1, FileV(cnt).x
                    Input #1, FileV(cnt).y
                    Input #1, FileV(cnt).z
                    Input #1, FileV(cnt).tu
                    Input #1, FileV(cnt).tv
                    cnt = cnt + 1
                Next k
            Next j
        End With
    Next i
    Close #1
    
    EffectFileCentral True, True, True
End Sub

Public Sub EffectFileAdd(MP As D3DVECTOR)
    Dim i%
    
    For i = 0 To FileVcnt - 1
        EffectV(EffectVcnt + i) = FileV(i)
        With EffectV(EffectVcnt + i)
            .x = .x + MP.x
            .y = .y + MP.y
            .z = .z + MP.z
        End With
    Next i
    EffectVcnt = EffectVcnt + FileVcnt
End Sub

Public Sub EffectFileCentral(ByVal CenX As Boolean, ByVal CenY As Boolean, ByVal CenZ As Boolean)
    Dim i&, Decr As D3DVECTOR, Incr As D3DVECTOR, MP As D3DVECTOR
    
    For i = 0 To FileVcnt - 1
        If FileV(i).x < Decr.x Or i = 0 Then Decr.x = FileV(i).x
        If FileV(i).y < Decr.y Or i = 0 Then Decr.y = FileV(i).y
        If FileV(i).z < Decr.z Or i = 0 Then Decr.z = FileV(i).z
        If FileV(i).x > Incr.x Or i = 0 Then Incr.x = FileV(i).x
        If FileV(i).y > Incr.y Or i = 0 Then Incr.y = FileV(i).y
        If FileV(i).z > Incr.z Or i = 0 Then Incr.z = FileV(i).z
    Next i
    Mk3d.dx.VectorAdd MP, Decr, Incr
    Mk3d.dx.VectorScale MP, MP, 0.5
    For i = 0 To FileVcnt - 1
        If CenX Then FileV(i).x = FileV(i).x - MP.x
        If CenY Then FileV(i).y = FileV(i).y - MP.y
        If CenZ Then FileV(i).z = FileV(i).z - MP.z
    Next i
End Sub

Public Sub EffectVertexDelete(ByVal Index As Long)
    Dim i&
    
    For i = Index To EffectVcnt - 2
        EffectV(i) = EffectV(i + 1)
    Next i
    EffectVcnt = EffectVcnt - 1
End Sub


Public Function GetEffectVertex() As D3DVERTEX()
    GetEffectVertex = EffectV
End Function
