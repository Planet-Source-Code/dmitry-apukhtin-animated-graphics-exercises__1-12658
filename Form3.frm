VERSION 5.00
Object = "{61E1E365-AC60-11D4-AA88-00A0CC334D72}#22.0#0"; "atOcx.ocx"
Begin VB.Form Form3 
   Caption         =   "at Test 3 - Collision Controlled Motion"
   ClientHeight    =   7230
   ClientLeft      =   1755
   ClientTop       =   480
   ClientWidth     =   9285
   LinkTopic       =   "Form3"
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   482
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   619
   Begin VB.Timer gTimer 
      Interval        =   1000
      Left            =   8700
      Top             =   135
   End
   Begin atOcx.AniTrans aTarget 
      Height          =   645
      Index           =   2
      Left            =   4875
      Tag             =   "2"
      Top             =   5640
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1138
      UserDataSize    =   16
      HitTesting      =   1
      CollisionChecking=   1
      Animate         =   1
      AnimDelay       =   100
      CurFrame        =   8
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form3.frx":A368
   End
   Begin atOcx.AniTrans aTarget 
      Height          =   795
      Index           =   1
      Left            =   2595
      Tag             =   "3"
      Top             =   5685
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   1402
      UserDataSize    =   16
      HitTesting      =   1
      CollisionChecking=   1
      Animate         =   1
      AnimDelay       =   100
      CurFrame        =   2
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form3.frx":B6EC
   End
   Begin atOcx.AniTrans aBlow3 
      Height          =   1245
      Left            =   1365
      Top             =   1455
      Visible         =   0   'False
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   3201
      UserDataSize    =   16
      HitTesting      =   1
      CollisionChecking=   2
      Animate         =   0
      AnimDelay       =   100
      CurFrame        =   1
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form3.frx":CF83
   End
   Begin atOcx.AniTrans aBlow2 
      Height          =   1170
      Left            =   2715
      Top             =   1455
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   2064
      UserDataSize    =   16
      HitTesting      =   1
      CollisionChecking=   2
      Animate         =   0
      AnimDelay       =   100
      CurFrame        =   1
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form3.frx":15723
   End
   Begin atOcx.AniTrans aBlow1 
      Height          =   1500
      Left            =   150
      Top             =   1470
      Visible         =   0   'False
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   2646
      UserDataSize    =   16
      HitTesting      =   1
      CollisionChecking=   1
      Animate         =   0
      AnimDelay       =   100
      CurFrame        =   9
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form3.frx":16F4F
   End
   Begin atOcx.AniTrans aBlow 
      Height          =   930
      Index           =   0
      Left            =   5820
      Top             =   855
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1640
      UserDataSize    =   16
      HitTesting      =   1
      CollisionChecking=   0
      Animate         =   0
      AnimDelay       =   0
      CurFrame        =   0
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form3.frx":1AFDE
   End
   Begin atOcx.AniTrans aTarget 
      Height          =   1500
      Index           =   0
      Left            =   180
      Tag             =   "1"
      Top             =   5670
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   2646
      UserDataSize    =   16
      HitTesting      =   1
      CollisionChecking=   1
      Animate         =   1
      AnimDelay       =   100
      CurFrame        =   6
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form3.frx":1AFF6
   End
   Begin atOcx.AniTrans aBomb 
      Height          =   615
      Index           =   0
      Left            =   4815
      Top             =   915
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   1085
      UserDataSize    =   16
      HitTesting      =   1
      CollisionChecking=   2
      Animate         =   0
      AnimDelay       =   0
      CurFrame        =   0
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form3.frx":1DD37
   End
   Begin atOcx.AniTrans aUFO 
      Height          =   675
      Index           =   0
      Left            =   4965
      Top             =   60
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   1191
      UserDataSize    =   0
      HitTesting      =   1
      CollisionChecking=   2
      Animate         =   0
      AnimDelay       =   0
      CurFrame        =   0
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form3.frx":1DD4F
   End
   Begin atOcx.AniTrans aUFO2 
      Height          =   375
      Left            =   1725
      Top             =   270
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1588
      UserDataSize    =   0
      HitTesting      =   1
      CollisionChecking=   1
      Animate         =   1
      AnimDelay       =   100
      CurFrame        =   1
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form3.frx":1DD67
   End
   Begin atOcx.AniTrans aUFO1 
      Height          =   1125
      Left            =   315
      Top             =   150
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1984
      UserDataSize    =   0
      HitTesting      =   1
      CollisionChecking=   1
      Animate         =   1
      AnimDelay       =   100
      CurFrame        =   2
      FlipMode        =   0
      Stretch         =   -1  'True
      aB              =   "Form3.frx":1EFD5
   End
   Begin atOcx.AniTrans aBomb1 
      Height          =   375
      Left            =   300
      Top             =   825
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      UserDataSize    =   0
      HitTesting      =   1
      CollisionChecking=   1
      Animate         =   1
      AnimDelay       =   50
      CurFrame        =   3
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form3.frx":209D4
   End
   Begin atOcx.AniTrans aBomb2 
      Height          =   360
      Left            =   1695
      Top             =   780
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   635
      UserDataSize    =   0
      HitTesting      =   1
      CollisionChecking=   1
      Animate         =   0
      AnimDelay       =   50
      CurFrame        =   7
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form3.frx":213C2
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const iDX = 0
Const iDY = 1
Const iINDEX = 2
Const iY = 3
Const iTYPE = 4
Const iOWNER = 5
Const iLIFE = 6
Const iBLOWDELAY = 7

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

Private hBrush As Long

Private Sub Form_Load()
    Randomize
    
    Dim i As Long
    For i = 0 To aTarget.UBound: initTarget i: Next
    
    aTarget(0).UD(iDX) = 0
    aTarget(0).Left = (ScaleWidth - aTarget(0).Width) \ 2
    
    hBrush = CreateSolidBrush(vbGreen)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DeleteObject hBrush
End Sub

Private Sub gTimer_Timer()
    addUFO
    gTimer.Interval = 1000 + Rnd * 2000
End Sub

Private Sub addUFO()
    Load aUFO(aUFO.UBound + 1)
    With aUFO(aUFO.UBound)
        If Int(Rnd * 2) Then
            .CloneFrames aUFO1
        Else
            .CloneFrames aUFO2
        End If
        .UD(iDX) = 10 * Rnd
        If Int(Rnd * 2) Then
            .Move -.Width, Rnd * 100
        Else
            .Move ScaleWidth, Rnd * 100
            .UD(iDX) = -.UD(iDX)
        End If
        .Visible = True
        If Int(Rnd * 2) Then .Animate = AniForward Else .Animate = AniBackward
    End With
End Sub

Private Sub blowUFO(ByVal Index As Long)
    aUFO(Index).CollisionChecking = ccNone
    Load aBlow(aBlow.UBound + 1)
    With aBlow(aBlow.UBound)
        On Error GoTo E
        .CloneFrames aBlow1
        .UD(iTYPE) = 0
        .UD(iINDEX) = Index
        .UD(iDX) = aUFO(Index).UD(iDX)
        .Move aUFO(Index).Left, aUFO(Index).Top - 32
        .ZOrder 0
        .Visible = True
        .Animate = AniForward
        DoEvents
E:
    End With
End Sub

Private Sub aUFO_OnFrameTick(Index As Integer)
    With aUFO(Index)
        On Error GoTo E
        .Move .Left + .UD(iDX)
        If .Left < -.Width Or .Left > ScaleWidth Then
            Unload aUFO(Index)
        Else
            Dim A As AniTrans
            For Each A In aUFO
                If .TestCollision(A.Object) Then
                    blowUFO A.Index
                    blowUFO Index
                    Exit Sub
                End If
            Next
            If Int(Rnd * 200) < 10 Then
                addBomb .Left + .Width \ 2 - 16, .Top + .Height \ 2 - 8, .UD(iDX), Index
            End If
        End If
E:
    End With
End Sub

Private Sub addBomb(ByVal x As Long, ByVal Y As Long, ByVal dX As Long, ByVal Owner As Long)
    Load aBomb(aBomb.UBound + 1)
    With aBomb(aBomb.UBound)
        If Int(Rnd * 2) Then
            .CloneFrames aBomb1
            .UD(iINDEX) = 1
        Else
            .CloneFrames aBomb2
            .UD(iINDEX) = 2
        End If
        .UD(iOWNER) = Owner
        .UD(iDX) = dX \ 2
        .UD(iY) = Y * 10
        .UD(iDY) = 1
        .Move x, Y
        .Animate = AniForward
        .Visible = True
    End With
End Sub

Private Sub aBomb_OnFrameTick(Index As Integer)
    With aBomb(Index)
        On Error GoTo E
        .Move .Left + .UD(iDX), .UD(iY) \ 10
        .UD(iY) = .UD(iY) + .UD(iDY)
        .UD(iDY) = .UD(iDY) + 2
        .UD(iDX) = .UD(iDX) * 8 / 10
        If .UD(iBLOWDELAY) Then
            .UD(iBLOWDELAY) = .UD(iBLOWDELAY) - 1
            If .UD(iBLOWDELAY) = 0 Then
                Load aBlow(aBlow.UBound + 1)
                With aBlow(aBlow.UBound)
                    .CloneFrames aBlow1
                    .UD(iTYPE) = 2
                    .UD(iINDEX) = Index
                    .UD(iDX) = aBomb(Index).UD(iDX)
                    .Move aBomb(Index).Left - 20, aBomb(Index).Top - 48
                    .ZOrder 0
                    .Animate = AniForward
                    .Visible = True
                End With
            End If
        End If
        If .Top > ScaleHeight Then
            Unload aBomb(Index)
        Else
            Dim i As Long
            On Error GoTo 0
            For i = 0 To aTarget.UBound
                If testBombHit(aBomb(Index), aTarget(i)) Then Exit Sub
            Next
        End If
        Dim A As AniTrans
        For Each A In aUFO
            If A.Index <> .UD(iOWNER) Then
                If .TestCollision(A.Object) Then
                    If .UD(iDY) > 0 Then
                        .UD(iDY) = -.UD(iDY) \ 2
                    Else
                        .UD(iDY) = -.UD(iDY)
                    End If
                    blowUFO A.Index
                    Exit For
                End If
            End If
        Next
E:
    End With
End Sub

Private Function testBombHit(testBomb As AniTrans, testTarget As AniTrans) As Boolean
    With testBomb
        If .TestCollision(testTarget.Object) Then
            On Error GoTo E
            .CollisionChecking = ccNone
            .UD(iDY) = -.UD(iDY) \ 2
            .UD(iBLOWDELAY) = 20 + Rnd * 20
            Load aBlow(aBlow.UBound + 1)
            aBlow(aBlow.UBound).UD(iTYPE) = 1
            If .UD(iINDEX) = 1 Then
                testTarget.UD(iLIFE) = testTarget.UD(iLIFE) - 10
                aBlow(aBlow.UBound).CloneFrames aBlow3
                aBlow(aBlow.UBound).Move testTarget.Left + (testTarget.Width - aBlow3.Width) \ 2, testTarget.Top + (testTarget.Height - aBlow3.Height)
            Else
                testTarget.UD(iLIFE) = testTarget.UD(iLIFE) - 2
                aBlow(aBlow.UBound).CloneFrames aBlow2
                aBlow(aBlow.UBound).Move testTarget.Left + (testTarget.Width - aBlow2.Width) \ 2, testTarget.Top + (testTarget.Height - aBlow2.Height)
            End If
            If testTarget.UD(iLIFE) > 0 Then
                testTarget.CurFrame = 12
            Else
                initTarget testTarget.Index
            End If
            aBlow(aBlow.UBound).Animate = AniForward
            aBlow(aBlow.UBound).ZOrder 0
            aBlow(aBlow.UBound).Visible = True
E:
            testBombHit = True
        End If
    End With
End Function

Private Sub aBlow_OnFrameTick(Index As Integer)
    With aBlow(Index)
        On Error GoTo E
        Select Case .UD(iTYPE)
            Case 0
                .Left = .Left + .UD(iDX)
                If .CurFrame = 1 Then Unload aUFO(.UD(iINDEX))
                If .CurFrame = 7 Then Unload aBlow(Index)
            Case 1
                If .CurFrame = .FrameCount Then Unload aBlow(Index)
            Case 2
                .Left = .Left + .UD(iDX)
                If .CurFrame = 11 Then Unload aBomb(.UD(iINDEX))
                If .CurFrame = 7 Then Unload aBlow(Index)
        End Select
E:
    End With
End Sub

Private Sub aTarget_OnFrameTick(Index As Integer)
    With aTarget(Index)
        If .Left < 0 Or .Left > ScaleWidth - .Width Then
            .UD(iDX) = -.UD(iDX)
        End If
        .Left = .Left + .UD(iDX)
    End With
End Sub

Private Sub aTarget_OnPaint(Index As Integer, ByVal hDC As Long)
    With aTarget(Index)
        Dim rc As RECT, hOB As Long
        Rectangle hDC, 0, 0, .Width, 8
        rc.Left = 1
        rc.Top = 1
        rc.Right = .Width * .UD(iLIFE) / 100
        rc.Bottom = 7
        FillRect hDC, rc, hBrush
    End With
End Sub

Private Sub initTarget(ByVal Index As Integer)
    With aTarget(Index)
        .UD(iLIFE) = 100
        .UD(iDX) = .Tag
        If Int(Rnd * 2) Then
            .Move 0, ScaleHeight - .Height
        Else
            .Move ScaleWidth - .Width, ScaleHeight - .Height
            .UD(iDX) = -.UD(iDX)
        End If
    End With
End Sub
