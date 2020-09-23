VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mk3dAnimatedObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private KeyObjCount%
Private ObjKey() As New Mk3dObject
Private ObjAct As New Mk3dObject


Public Function CreateFromObjects(ByVal ObjCount As Integer, Obj() As Mk3dObject) As Boolean
    Dim i%
    
    On Local Error GoTo Failed
    
    ReDim ObjKey(ObjCount - 1)
    For i = 0 To ObjCount - 1
        Set ObjKey(i) = Obj(i)
    Next i
    KeyObjCount = ObjCount
    CreateFromObjects = ObjAct.CreateFromFile(ObjKey(0).VertexFile, False)
    Exit Function
    
Failed:
    CreateFromObjects = False
End Function


Public Function CalcAnimObject(ByVal KeyFrameIndex As Integer, ByVal NewFrameIndex As Integer, ByVal StatePercent As Single) As Mk3dObject
    Dim i&
    Dim RenderVKey() As D3DVERTEX, RenderVNew() As D3DVERTEX, ActVecKey As D3DVECTOR, ActVecNew As D3DVECTOR
    Dim ScaleFac!, InterpVec As D3DVECTOR
    
    If ObjKey(KeyFrameIndex).VertexCount <> ObjKey(NewFrameIndex).VertexCount Then Exit Function
    
    RenderVKey = ObjKey(KeyFrameIndex).GetRenderVertex
    RenderVNew = ObjKey(NewFrameIndex).GetRenderVertex
    ScaleFac = StatePercent / 100
    
    'interpolate polygons
    For i = 0 To ObjKey(KeyFrameIndex).VertexCount - 1
        'calculate the distance of two points and interpolate it
        ActVecKey = Mk3d.VertexToVector(RenderVKey(i))
        ActVecNew = Mk3d.VertexToVector(RenderVNew(i))
        Mk3d.dx.VectorSubtract InterpVec, ActVecNew, ActVecKey
        Mk3d.dx.VectorScale InterpVec, InterpVec, ScaleFac
        'calculate the new position
        Mk3d.dx.VectorAdd InterpVec, ActVecKey, InterpVec
        'write the position back to the vertex
        ObjAct.PutVector i, InterpVec
    Next i
    Set CalcAnimObject = ObjAct
End Function

Public Function GetAnimObj() As Mk3dObject
    Set GetAnimObj = ObjAct
End Function

Public Function GetKeyFrameObj(ByVal Index As Integer) As Mk3dObject
    Set GetKeyFrameObj = ObjKey(Index)
End Function

Public Sub Move(MoveDelta As D3DVECTOR)
    Dim i%
    
    For i = 0 To KeyObjCount - 1
        ObjKey(i).Move MoveDelta
    Next i
    ObjAct.Move MoveDelta
End Sub

Public Sub MoveTo(Position As D3DVECTOR)
    Dim i%
    
    For i = 0 To KeyObjCount - 1
        ObjKey(i).MoveTo Position
    Next i
    ObjAct.MoveTo Position
End Sub

Public Sub Rotate(RotAngle As D3DVECTOR)
    Dim i%
    
    For i = 0 To KeyObjCount - 1
        ObjKey(i).Rotate RotAngle
    Next i
    ObjAct.Rotate RotAngle
End Sub
