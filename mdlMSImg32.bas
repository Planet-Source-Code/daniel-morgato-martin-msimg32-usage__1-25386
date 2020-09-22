Attribute VB_Name = "mdlMSImg32"
Option Explicit

Public Declare Function AlphaBlend Lib "msimg32" (ByVal hdcDest As Long, _
    ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, _
    ByVal nWidthDest As Long, ByVal hHeightDest As Long, _
    ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, _
    ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, _
    ByVal nHeightSrc As Long, ByVal blendFunc As Long) As Boolean

Public Declare Function TransparentBlt Lib "msimg32" (ByVal hdcDest As Long, _
    ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, _
    ByVal nWidthDest As Long, ByVal hHeightDest As Long, _
    ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, _
    ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, _
    ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Boolean
 
Public Declare Function GradientFill Lib "msimg32" _
    (ByVal Desthdc As Long, pVertex As TRIVERTEX, _
    ByVal dwNumVertex As Long, pMesh As Any, _
    ByVal dwNumMesh As Long, ByVal dwMode As Long) As Boolean
 
Type TRIVERTEX
    x As Long
    y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type

Type GRADIENT_TRIANGLE
    Vertex1 As Long
    Vertex2 As Long
    Vertex3 As Long
End Type

Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type
 
Global Const GRADIENT_FILL_TRIANGLE = 2
Global Const GRADIENT_FILL_RECT_H = 0
Global Const GRADIENT_FILL_RECT_V = 1

