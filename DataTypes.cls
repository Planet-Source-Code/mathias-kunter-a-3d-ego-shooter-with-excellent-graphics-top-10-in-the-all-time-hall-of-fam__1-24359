VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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
    Coord(1) As Mk3dTriangle
End Type

