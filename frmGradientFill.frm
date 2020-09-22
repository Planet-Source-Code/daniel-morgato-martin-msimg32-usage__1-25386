VERSION 5.00
Begin VB.Form frmGradientFill 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GradientFill example"
   ClientHeight    =   4095
   ClientLeft      =   2100
   ClientTop       =   1335
   ClientWidth     =   4560
   Icon            =   "frmGradientFill.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   273
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   Begin VB.CommandButton Command2 
      Caption         =   "Shaded square"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Shaded triangle"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   1575
   End
End
Attribute VB_Name = "frmGradientFill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Private Sub Command1_Click()
Dim Vert(3) As TRIVERTEX
Dim GTri As GRADIENT_TRIANGLE
    
    Vert(0).x = 150
    Vert(0).y = 20
    Vert(0).Red = 0
    Vert(0).Green = &HFF00
    Vert(0).Blue = 0
    Vert(0).Alpha = 0
    
    Vert(1).x = 40
    Vert(1).y = 220
    Vert(1).Red = 0
    Vert(1).Green = 0
    Vert(1).Blue = &HFF00
    Vert(1).Alpha = 0
    
    Vert(2).x = 260
    Vert(2).y = 220
    Vert(2).Red = &HFF00
    Vert(2).Green = 0
    Vert(2).Blue = 0
    Vert(2).Alpha = 0
    
    GTri.Vertex1 = 0
    GTri.Vertex2 = 1
    GTri.Vertex3 = 2
    
    Me.Cls
    Call GradientFill(Me.hDC, Vert(0), 3, GTri, 1, GRADIENT_FILL_TRIANGLE)
    Me.Refresh
End Sub


Private Sub Command2_Click()
Dim Vert(2) As TRIVERTEX
Dim gRect As GRADIENT_RECT
    
    Vert(0).x = 20
    Vert(0).y = 20
    Vert(0).Red = 0
    Vert(0).Green = 0
    Vert(0).Blue = 0
    Vert(0).Alpha = 0
    
    Vert(1).x = 270
    Vert(1).y = 200
    Vert(1).Red = 0
    Vert(1).Green = 0
    Vert(1).Blue = &HFF00
    Vert(1).Alpha = 0
    
    gRect.UpperLeft = 0
    gRect.LowerRight = 1
    
    Me.Cls
    Call GradientFill(Me.hDC, Vert(0), 2, gRect, 1, GRADIENT_FILL_RECT_H)
    Me.Refresh
End Sub


