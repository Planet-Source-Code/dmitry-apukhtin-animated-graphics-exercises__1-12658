VERSION 5.00
Object = "{61E1E365-AC60-11D4-AA88-00A0CC334D72}#22.0#0"; "atOcx.ocx"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   7905
   ClientLeft      =   1800
   ClientTop       =   555
   ClientWidth     =   9525
   LinkTopic       =   "Form4"
   ScaleHeight     =   527
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   635
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000001&
      DrawStyle       =   2  'Dot
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   90
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   0
      Top             =   1965
      Width           =   1875
      Begin VB.HScrollBar sb 
         Height          =   255
         Left            =   60
         TabIndex        =   1
         Top             =   90
         Width           =   1215
      End
   End
   Begin atOcx.AniTrans AniTrans1 
      Height          =   1800
      Left            =   90
      Top             =   105
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   3175
      UserDataSize    =   16
      HitTesting      =   1
      CollisionChecking=   2
      Animate         =   1
      AnimDelay       =   100
      CurFrame        =   10
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form4.frx":0000
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    sb.Max = AniTrans1.FrameCount
    Set pic.Picture = Form3.Picture
    Unload Form3
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    AniTrans1.Move 0, 0
    pic.Move 0, AniTrans1.Height, ScaleWidth, ScaleHeight - AniTrans1.Height
    pic_Paint
End Sub

Private Sub pic_Paint()
    Dim i As Long, x As Long
    i = sb.Value + 1
    pic.Cls
    With AniTrans1
        Do While x < pic.ScaleWidth And i <= .FrameCount
            pic.PaintPicture .FrameMask(i), x, 0, , , , , , , vbSrcPaint
            pic.PaintPicture .FramePicture(i), x, 0, , , , , , , vbSrcAnd
            
            pic.PaintPicture .FramePicture(i), x, .Height + 1
            pic.PaintPicture .FrameMask(i), x, .Height * 2 + 2
            
            pic.CurrentX = x + 1: pic.CurrentY = 1
            pic.Print i
            x = x + .Width + 1
            i = i + 1
        Loop
    End With
End Sub

Private Sub pic_Resize()
    On Error Resume Next
    sb.Move 0, pic.ScaleHeight - sb.Height, pic.ScaleWidth
    sb.LargeChange = pic.ScaleWidth \ AniTrans1.Width
End Sub

Private Sub sb_Change()
    pic_Paint
    pic.Refresh
    pic.SetFocus
End Sub

Private Sub sb_Scroll()
    sb_Change
End Sub
