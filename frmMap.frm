VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Map"
   ClientHeight    =   10425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11130
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10425
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   2640
   End
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   0
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   9975
      Left            =   120
      ScaleHeight     =   661
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   637
      TabIndex        =   6
      Top             =   120
      Width           =   9615
      Begin VB.PictureBox picMap 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   11640
         Left            =   0
         ScaleHeight     =   774
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   927
         TabIndex        =   7
         Top             =   0
         Width           =   13935
         Begin VB.Shape Shape1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            Height          =   495
            Left            =   600
            Top             =   480
            Width           =   495
         End
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   10080
      Width           =   9615
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9975
      Left            =   9720
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   10080
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   10080
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   10080
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Save file name, without the .map:"
      Height          =   855
      Left            =   10080
      TabIndex        =   8
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   10080
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private intTileX As Integer
Private intTileY As Integer
Private bMouseD As Boolean
Private ff As Long
Private MapW As Integer
Private MapH As Integer
Private bRunning As Boolean

Private Sub cmdSave_Click()
    If Text1.Text = "" Then
        MsgBox "File name is empty."
        Exit Sub
    End If
    Dim a As Integer, b As Integer
    Open App.Path & "\Maps\" & Text1.Text & ".map" For Binary As ff
        MapW = 100
        MapH = 100
        Put ff, , MapW
        Put ff, , MapH
        Put ff, , Map
    Close ff
End Sub

Private Sub cmdLoad_Click()
    dlg1.Filter = "MAP Files (.map)|*.map|"
    dlg1.ShowOpen
    If dlg1.FileName = "" Then Exit Sub
    Dim sX As Integer
    Dim sY As Integer
    Open dlg1.FileName For Binary As ff
        Get ff, , MapW
        Get ff, , MapH
        Get ff, , Map
    Close ff
    Dim a As Integer, b As Integer
    For a = 0 To MapW
        For b = 0 To MapH
            'MsgBox Map(a, b)
            If Not Map(a, b) = 0 Then
                sY = Fix((Map(a, b) - 1) / 8)
                sX = ((Map(a, b) - 1) - (sY * 8)) * 32
                sY = sY * 32
                BitBlt picMap.hDC, a * 32, b * 32, 32, 32, Form1.picTileSet.hDC, sX, sY, vbSrcCopy
            End If
        Next
    Next
    picMap.Refresh
End Sub

Private Sub Form_Load()
    ff = FreeFile
    picMap.Width = 100 * 32
    picMap.Height = 100 * 32
End Sub

Private Sub HScroll1_Change()
    picMap.Left = 0 - (HScroll1.Value)
End Sub

Private Sub HScroll1_Scroll()
    picMap.Left = 0 - (HScroll1.Value)
End Sub

Private Sub picMap_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    bMouseD = True
End Sub

Private Sub picMap_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    SelectTile intTileX, intTileY, 1
    Let intTileX = Fix(x \ 32)
    Let intTileY = Fix(y \ 32)
    Label1 = intTileX & ", " & intTileY
    Shape1.Left = intTileX * 32
    Shape1.Top = intTileY * 32
End Sub

Private Sub picMap_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    bMouseD = False
End Sub

Private Sub Timer1_Timer()
    If bMouseD Then
        BitBlt picMap.hDC, intTileRight, intTileBottom, 32, 32, Form1.picSelected.hDC, intTileRightMap, intTileBottomMap, vbSrcCopy
        Map(Fix(intTileRight / 32), Fix(intTileBottom / 32)) = TileSelected
        picMap.Refresh
    End If
End Sub

Private Sub VScroll1_Change()
picMap.Top = 0 - (VScroll1.Value)
End Sub

Private Sub VScroll1_Scroll()
picMap.Top = 0 - (VScroll1.Value * 2)
End Sub
