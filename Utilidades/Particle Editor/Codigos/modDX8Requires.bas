Attribute VB_Name = "modDX8Requires"
Option Explicit

Public vertList(3) As TLVERTEX

Public Type D3D8Textures

    texture As Direct3DTexture8
    texwidth As Integer
    texheight As Integer

End Type

Public dX        As DirectX8
Public D3D       As Direct3D8
Public D3DDevice As Direct3DDevice8
Public D3DX      As D3DX8

Public Type TLVERTEX

    x As Single
    Y As Single
    Z As Single
    rhw As Single
    color As Long
    Specular As Long
    tu As Single
    tv As Single

End Type

Public Const PI   As Single = 3.14159265358979
Public base_light As Long
