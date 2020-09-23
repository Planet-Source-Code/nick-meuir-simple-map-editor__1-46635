VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tiles"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   16.325
   ScaleMode       =   0  'User
   ScaleWidth      =   8.361
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picSelected 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4080
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox picTileSet 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15360
      Left            =   120
      Picture         =   "frmTile.frx":0000
      ScaleHeight     =   1024
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   120
      Width           =   3840
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private intTileX As Integer
Private intTileY As Integer

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If unloadfrm Then
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

Private Sub picTileSet_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SelectTile intTileX, intTileY, 1
    TileSelected = 1 + intTileX + (intTileY * 8) '0 is the blank tile
    BitBlt picSelected.hDC, 0, 0, 32, 32, picTileSet.hDC, intTileRight, intTileBottom, vbSrcCopy
    picSelected.Refresh
End Sub

Private Sub picTileSet_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Let intTileX = Fix(x / 32)
    Let intTileY = Fix(y / 32)
    Label1 = 1 + intTileX + (intTileY * 8)
    Shape1.Left = intTileX * 32
    Shape1.Top = intTileY * 32
End Sub

