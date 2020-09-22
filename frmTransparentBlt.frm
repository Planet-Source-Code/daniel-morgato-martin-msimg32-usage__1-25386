VERSION 5.00
Begin VB.Form frmTransparentBlt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TransparentBlt example"
   ClientHeight    =   4455
   ClientLeft      =   4365
   ClientTop       =   2130
   ClientWidth     =   3285
   Icon            =   "frmTransparentBlt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   297
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   219
   Begin VB.CommandButton Command1 
      Caption         =   "Apply TransparentBlt"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   1815
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      Height          =   1560
      Index           =   1
      Left            =   120
      Picture         =   "frmTransparentBlt.frx":0E42
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   1
      Top             =   2280
      Width           =   3060
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      Height          =   1560
      Index           =   0
      Left            =   120
      Picture         =   "frmTransparentBlt.frx":FCE4
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   360
      Width           =   3060
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Source:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   555
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Destination:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "frmTransparentBlt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Pic(0).Cls
    Call TransparentBlt(Pic(0).hdc, 0, 0, Pic(0).ScaleWidth, _
        Pic(0).ScaleHeight, Pic(1).hdc, 0, 0, Pic(1).ScaleWidth, _
        Pic(1).ScaleHeight, Pic(1).Point(10, 10))
    Pic(0).Refresh
End Sub

