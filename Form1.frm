VERSION 5.00
Object = "{61E1E365-AC60-11D4-AA88-00A0CC334D72}#22.0#0"; "atOcx.ocx"
Begin VB.Form Form1 
   BackColor       =   &H80000001&
   Caption         =   "at Test Form 1 - Stress and CollisionTest. Drag things around and use Right-Click"
   ClientHeight    =   7200
   ClientLeft      =   1575
   ClientTop       =   885
   ClientWidth     =   9585
   FillColor       =   &H0000FF00&
   FillStyle       =   0  'Solid
   ForeColor       =   &H000000FF&
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   639
   WindowState     =   2  'Maximized
   Begin VB.CommandButton bForm4 
      Caption         =   "4 >"
      Height          =   315
      Left            =   1350
      TabIndex        =   2
      Top             =   6825
      Width           =   585
   End
   Begin VB.CommandButton bForm3 
      Caption         =   "3 >"
      Height          =   315
      Left            =   690
      TabIndex        =   1
      Top             =   6825
      Width           =   585
   End
   Begin VB.CommandButton bMore 
      Caption         =   "2 >"
      Height          =   315
      Left            =   45
      TabIndex        =   0
      Top             =   6825
      Width           =   585
   End
   Begin atOcx.AniTrans ag 
      Height          =   1800
      Index           =   0
      Left            =   180
      Tag             =   "(embedded gif)"
      Top             =   120
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   3175
      UserDataSize    =   0
      HitTesting      =   1
      CollisionChecking=   3
      Animate         =   1
      AnimDelay       =   -1
      CurFrame        =   1
      FlipMode        =   0
      Stretch         =   0   'False
      MouseIcon       =   "Form1.frx":000C
      MousePointer    =   99
      aB              =   "Form1.frx":016E
   End
   Begin VB.Menu mnuPOP 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuPFront 
         Caption         =   "Bring to Front"
      End
      Begin VB.Menu mnuPBack 
         Caption         =   "Send to Back"
      End
      Begin VB.Menu mnuPS0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPPaused 
         Caption         =   "Paused"
      End
      Begin VB.Menu mnuPPlayF 
         Caption         =   "Play Forward"
      End
      Begin VB.Menu mnuPPlayB 
         Caption         =   "Play Backward"
      End
      Begin VB.Menu mnuPS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPFlipH 
         Caption         =   "Flip Horz"
      End
      Begin VB.Menu mnuPFlipV 
         Caption         =   "Flip Vert"
      End
      Begin VB.Menu mnuPS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPFN 
         Caption         =   "File Names"
      End
      Begin VB.Menu mnuPLoad 
         Caption         =   "Load"
      End
      Begin VB.Menu mnuPDelete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private selIndex As Long, CollisionAmount As Single
Private hBrush As Long

Private Sub Form_Activate()
    If ag.UBound = 0 Then loadGifs
End Sub

Private Sub Form_Load()
    hBrush = CreateSolidBrush(vbGreen)
    selIndex = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DeleteObject hBrush
End Sub

Private Sub bMore_Click()
    Hide
    DoEvents
    Form2.Show vbModal
    Show
End Sub

Private Sub bForm3_Click()
    Hide
    DoEvents
    Form3.Show vbModal
    Show
End Sub

Private Sub bForm4_Click()
    Hide
    DoEvents
    Form4.Show vbModal
    Show
End Sub

Private Sub loadGifs()
    On Error Resume Next
    Dim ts As String, i As Long, tp As String
    tp = App.Path & "\Gifs\"
    ts = Dir$(tp & "*.GIF")
    If ts = "" Then
        tp = App.Path & "..\..\..\Gifs\"
        ts = Dir$(tp & "*.GIF")
        If ts = "" Then Exit Sub
    End If
    Dim x As Long, Y As Long, maxY As Long
    ag(0).Move 0, 0
    x = ag(0).Width
    i = ag.LBound + 1
    Do Until ts = "" Or i > 40
        Dim A() As Byte
        Open tp & ts For Binary Access Read As #1
        ReDim A(LOF(1) - 1) As Byte
        Get #1, , A
        Close #1
        If i Then Load ag(i)
        With ag(i)
            .ParseBytes (A)
            .Tag = ts
            .AnimDelay = 100
            .Animate = AniForward
            If x + .Width > ScaleWidth Then
                x = 0: Y = Y + maxY: maxY = .Height
            End If
            .Move x, Y ', ag(0).Width, ag(0).Height
            .MousePointer = ag(0).MousePointer
            Set .MouseIcon = ag(0).MouseIcon
            .ZOrder 0
            .Visible = True
            DoEvents
            
            If .Height > maxY Then maxY = .Height
            If Y > ScaleHeight Then Exit Do
            x = x + .Width
            i = i + 1
            ts = Dir$
        End With
    Loop
End Sub

Private Sub ag_OnPaint(Index As Integer, ByVal hDC As Long)
    With ag(Index)
        Dim rc As RECT
        If Index = 0 Then
            Dim hOB As Long
            Rectangle hDC, 0, 0, .Width, 8
            rc.Left = 1
            rc.Top = 1
            rc.Right = .Width * CollisionAmount
            rc.Bottom = 7
            FillRect hDC, rc, hBrush
        End If
        
        If Index = selIndex Then Rectangle hDC, 0, 0, .Width, .Height
        
        If mnuPFN.Checked Then
            rc.Right = .Width
            DrawText hDC, .Tag, -1, rc, &H121
        End If
    End With
End Sub

Private Sub doPopUp()
    With ag(selIndex)
        mnuPPaused.Checked = .Animate = AniFreeze
        mnuPPlayF.Checked = .Animate And AniForward
        mnuPPlayB.Checked = .Animate And AniBackward
        
        mnuPFlipH.Checked = .FlipMode And FlipHorizontal
        mnuPFlipV.Checked = .FlipMode And FlipVertical
        
        PopupMenu mnuPOP, 2
    End With
End Sub

Private Sub ag_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button Then
        If selIndex <> Index Then
            selNone
            selIndex = Index
            ag(selIndex).Refresh
        End If
        If Button = 2 Then doPopUp
    End If
End Sub

Private Sub ag_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Static dX As Single, dY As Single
    If Index = selIndex Then
        If Button Then
            With ag(Index)
                If dX Then
                    .Move .Left - dX + x, .Top - dY + Y
                    If Index Then
                        CollisionAmount = .TestCollision(ag(0)) / 500
                    Else
                        CollisionAmount = 0
                    End If
                    ag(0).Refresh
                Else
                    dX = x
                    dY = Y
                End If
            End With
        Else
            dX = 0
        End If
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    selNone
End Sub

Private Sub selNone()
    Dim i As Long
    If selIndex >= 0 Then
        i = selIndex
        selIndex = -1
        ag(i).Refresh
    End If
End Sub

Private Sub mnuPDelete_Click()
    Unload ag(selIndex)
    selIndex = -1
End Sub

Private Sub mnuPFlipH_Click()
    With ag(selIndex)
        .FlipMode = .FlipMode Xor FlipHorizontal
    End With
End Sub

Private Sub mnuPFlipV_Click()
    With ag(selIndex)
        .FlipMode = .FlipMode Xor FlipVertical
    End With
End Sub

Private Sub mnuPFN_Click()
    mnuPFN.Checked = Not mnuPFN.Checked
End Sub

Private Sub mnuPFront_Click()
    ag(selIndex).ZOrder 0
End Sub

Private Sub mnuPBack_Click()
    ag(selIndex).ZOrder 1
End Sub

Private Sub mnuPLoad_Click()
    ag(selIndex).About
End Sub

Private Sub mnuPPaused_Click()
    With ag(selIndex)
        .Animate = AniFreeze
    End With
End Sub

Private Sub mnuPPlayB_Click()
    With ag(selIndex)
        .Animate = .Animate Xor AniBackward
    End With
End Sub

Private Sub mnuPPlayF_Click()
    With ag(selIndex)
        .Animate = .Animate Xor AniForward
    End With
End Sub

