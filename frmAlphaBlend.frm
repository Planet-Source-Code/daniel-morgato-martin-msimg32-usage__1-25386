VERSION 5.00
Begin VB.Form frmAlphaBlend 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AlphaBlend example"
   ClientHeight    =   5295
   ClientLeft      =   3630
   ClientTop       =   525
   ClientWidth     =   4800
   Icon            =   "frmAlphaBlend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   353
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.CommandButton Command2 
      Caption         =   "Apply AlphaBlend"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   4755
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate gradient"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4755
      Width           =   1815
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      Height          =   1980
      Index           =   1
      Left            =   120
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   1
      Top             =   2640
      Width           =   4560
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      Height          =   1980
      Index           =   0
      Left            =   120
      Picture         =   "frmAlphaBlend.frx":0E42
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   360
      Width           =   4560
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Source:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   555
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Destination:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "frmAlphaBlend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim i As Long, j As Long

    Pic(1).Cls
    For i = 0 To Pic(1).ScaleWidth - 1
        For j = 0 To Pic(1).ScaleHeight - 1
            Pic(1).PSet (i, j), RGB(Fix(i * 255 / Pic(1).ScaleWidth), _
            0, 255 - Fix(j * 255 / Pic(1).ScaleHeight))
        Next
        Pic(1).Refresh
    Next
End Sub

Private Sub Command2_Click()
Dim SourceConstantAlpha As Long, r As Byte, StrRes As String

    StrRes = InputBox("Give a number from 0 to 255 (the greater the " + _
        "value the farest you get from the clouds):", _
        "Alpha blend example...", 100)
        
    If StrRes = "" Then Exit Sub
    
    r = CLng(StrRes) Mod 256

    SourceConstantAlpha = r * 65536
    Pic(0).Cls
    Call AlphaBlend(Pic(0).hDC, 0, 0, Pic(0).ScaleWidth, Pic(0).ScaleHeight, _
        Pic(1).hDC, 0, 0, Pic(1).ScaleWidth, Pic(1).ScaleHeight, _
        SourceConstantAlpha)
    Pic(0).Refresh
End Sub


