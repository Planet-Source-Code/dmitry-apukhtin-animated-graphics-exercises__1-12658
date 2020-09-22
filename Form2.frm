VERSION 5.00
Object = "{61E1E365-AC60-11D4-AA88-00A0CC334D72}#22.0#0"; "atOcx.ocx"
Begin VB.Form Form2 
   Caption         =   "at Test 2 - animation and ZOrder"
   ClientHeight    =   6225
   ClientLeft      =   1800
   ClientTop       =   1590
   ClientWidth     =   9495
   ClipControls    =   0   'False
   HasDC           =   0   'False
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   6225
   ScaleWidth      =   9495
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   8940
      Top             =   75
   End
   Begin atOcx.AniTrans aMove 
      Height          =   1500
      Index           =   6
      Left            =   300
      Top             =   1980
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   2646
      UserDataSize    =   0
      HitTesting      =   1
      CollisionChecking=   1
      Animate         =   1
      AnimDelay       =   100
      CurFrame        =   2
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form2.frx":2DDD
   End
   Begin atOcx.AniTrans aMove 
      Height          =   1440
      Index           =   5
      Left            =   60
      Top             =   2970
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2646
      UserDataSize    =   0
      HitTesting      =   1
      CollisionChecking=   1
      Animate         =   1
      AnimDelay       =   100
      CurFrame        =   3
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form2.frx":6227
   End
   Begin atOcx.AniTrans aMove 
      Height          =   1500
      Index           =   0
      Left            =   7905
      Top             =   15
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2646
      UserDataSize    =   0
      HitTesting      =   1
      CollisionChecking=   1
      Animate         =   1
      AnimDelay       =   100
      CurFrame        =   5
      FlipMode        =   0
      Stretch         =   0   'False
      MousePointer    =   15
      aB              =   "Form2.frx":B57F
   End
   Begin atOcx.AniTrans aFly 
      Height          =   1095
      Index           =   0
      Left            =   2175
      Tag             =   "30"
      Top             =   1575
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1931
      UserDataSize    =   0
      HitTesting      =   1
      CollisionChecking=   1
      Animate         =   1
      AnimDelay       =   100
      CurFrame        =   2
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form2.frx":CBF2
   End
   Begin atOcx.AniTrans aFly 
      Height          =   1200
      Index           =   1
      Left            =   5760
      Tag             =   "-30"
      Top             =   2610
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   2117
      UserDataSize    =   0
      HitTesting      =   1
      CollisionChecking=   1
      Animate         =   1
      AnimDelay       =   100
      CurFrame        =   3
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form2.frx":DCE1
   End
   Begin atOcx.AniTrans aFire 
      Height          =   1500
      Index           =   0
      Left            =   1275
      Top             =   -45
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2646
      UserDataSize    =   0
      HitTesting      =   1
      CollisionChecking=   1
      Animate         =   1
      AnimDelay       =   100
      CurFrame        =   6
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form2.frx":FE57
   End
   Begin atOcx.AniTrans aFire 
      Height          =   1500
      Index           =   1
      Left            =   6120
      Top             =   120
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2646
      UserDataSize    =   0
      HitTesting      =   1
      CollisionChecking=   1
      Animate         =   1
      AnimDelay       =   100
      CurFrame        =   9
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form2.frx":13F75
   End
   Begin atOcx.AniTrans aMove 
      Height          =   900
      Index           =   4
      Left            =   1695
      Tag             =   "0"
      Top             =   3570
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   1588
      UserDataSize    =   0
      HitTesting      =   1
      CollisionChecking=   1
      Animate         =   1
      AnimDelay       =   100
      CurFrame        =   5
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form2.frx":18093
   End
   Begin atOcx.AniTrans aMove 
      Height          =   900
      Index           =   3
      Left            =   1995
      Tag             =   "0"
      Top             =   2715
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1588
      UserDataSize    =   0
      HitTesting      =   1
      CollisionChecking=   1
      Animate         =   1
      AnimDelay       =   100
      CurFrame        =   4
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form2.frx":19150
   End
   Begin atOcx.AniTrans aEye 
      Height          =   195
      Left            =   3720
      Top             =   1920
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   344
      UserDataSize    =   0
      HitTesting      =   1
      CollisionChecking=   1
      Animate         =   1
      AnimDelay       =   -1
      CurFrame        =   2
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form2.frx":19C1A
   End
   Begin atOcx.AniTrans aMove 
      Height          =   480
      Index           =   1
      Left            =   5445
      Top             =   3135
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      UserDataSize    =   0
      HitTesting      =   1
      CollisionChecking=   1
      Animate         =   1
      AnimDelay       =   100
      CurFrame        =   13
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form2.frx":19EBA
   End
   Begin atOcx.AniTrans aMouse 
      Height          =   1170
      Left            =   135
      Top             =   165
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2064
      UserDataSize    =   0
      HitTesting      =   1
      CollisionChecking=   1
      Animate         =   1
      AnimDelay       =   100
      CurFrame        =   27
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form2.frx":1A762
   End
   Begin atOcx.AniTrans aTiger 
      Height          =   750
      Left            =   2700
      Top             =   690
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1323
      UserDataSize    =   0
      HitTesting      =   1
      CollisionChecking=   1
      Animate         =   1
      AnimDelay       =   100
      CurFrame        =   3
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form2.frx":28CE8
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   -15
      X2              =   9940
      Y1              =   1365
      Y2              =   1365
   End
   Begin atOcx.AniTrans aRabbit 
      Height          =   1935
      Left            =   7755
      Tag             =   "-150"
      Top             =   1245
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   3413
      UserDataSize    =   0
      HitTesting      =   1
      CollisionChecking=   1
      Animate         =   1
      AnimDelay       =   100
      CurFrame        =   2
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form2.frx":2A631
   End
   Begin atOcx.AniTrans aJump 
      Height          =   780
      Left            =   3525
      Tag             =   "-75"
      Top             =   2130
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   1376
      UserDataSize    =   0
      HitTesting      =   1
      CollisionChecking=   1
      Animate         =   1
      AnimDelay       =   100
      CurFrame        =   6
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form2.frx":2E012
   End
   Begin atOcx.AniTrans aMove 
      Height          =   915
      Index           =   2
      Left            =   3045
      Top             =   240
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   1905
      UserDataSize    =   0
      HitTesting      =   1
      CollisionChecking=   1
      Animate         =   3
      AnimDelay       =   -1
      CurFrame        =   5
      FlipMode        =   0
      Stretch         =   0   'False
      aB              =   "Form2.frx":3073E
   End
   Begin VB.Image Image2 
      Height          =   1800
      Left            =   5100
      Picture         =   "Form2.frx":31BAA
      Top             =   4440
      Width           =   1995
   End
   Begin atOcx.AniTrans aTim 
      Height          =   1530
      Left            =   3375
      Tag             =   "60"
      Top             =   4635
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   2699
      UserDataSize    =   0
      HitTesting      =   1
      CollisionChecking=   1
      Animate         =   1
      AnimDelay       =   100
      CurFrame        =   3
      FlipMode        =   1
      Stretch         =   0   'False
      MousePointer    =   12
      aB              =   "Form2.frx":32712
   End
   Begin VB.Image Image1 
      Height          =   3675
      Left            =   3315
      Picture         =   "Form2.frx":35F94
      Top             =   2565
      Width           =   3780
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const PI = 3.1415926

Private Sub aEye_OnFrameTick()
    If aEye.CurFrame = 1 Then
      With aEye
        If .Tag <> "" Then
            .Tag = ""
            .Move Rnd * (ScaleWidth - .Width), Rnd * (ScaleHeight - .Height)
        Else
            .Tag = "x"
        End If
      End With
    End If
End Sub

Private Sub aFire_OnFrameTick(Index As Integer)
    With aFire(Index)
        If .CurFrame = .FrameCount Then
            .Move Rnd * (ScaleWidth - .Width), Rnd * (ScaleHeight - .Height)
        End If
    End With
End Sub

Private Sub aFly_OnFrameTick(Index As Integer)
    With aFly(Index)
        Dim x As Long
        x = .Left + .Tag
        If x < -.Width Or x > ScaleWidth Then .Tag = -.Tag
        .Move x, ScaleHeight \ 2 + ScaleHeight * Sin(PI / 180 * (Timer * 100 + x / 10)) / 4
    End With
End Sub

Private Sub aJump_OnFrameTick()
    With aJump
        Dim x As Long
        If .Animate = AniForward Then x = 90 Else x = -90
        .Left = .Left + x
        If .Left > ScaleWidth Then .Animate = AniBackward
        If .Left < -.Width Then .Animate = AniForward
    End With
End Sub

Private Sub aMouse_OnFrameTick()
    With aMouse
        If aTiger.Left > 4900 Then
            .AnimDelay = 150
            If .CurFrame > 2 Then
                .CurFrame = 1
                .FlipMode = .FlipMode Xor FlipHorizontal
            End If
        Else
            .FlipMode = FlipNone
            .AnimDelay = 50
        End If
        If .CurFrame = .FrameCount Then .Visible = False
    End With
End Sub

Private Sub aMove_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Static dX As Long, dY As Long
    If Button Then
      With aMove(Index)
        If dX Then
            .Move .Left - dX + x, .Top - dY + Y
            If .Tag <> "" Then .Tag = .Left
        Else
            dX = x
            dY = Y
        End If
      End With
    Else
        dX = 0
    End If
End Sub

Private Sub aMove_OnFrameTick(Index As Integer)
    If Index Then
        With aMove(Index)
            If .Tag <> "" Then
                .Left = .Tag + Rnd * 300 - 150
                If .CurFrame Mod 3 = 0 Then .Visible = False Else .Visible = True
            End If
        End With
    End If
End Sub

Private Sub aRabbit_OnFrameTick()
    With aRabbit
        .Left = .Left + .Tag
        If .Left < -.Width Or .Left > ScaleWidth Then .Tag = -.Tag: .Animate = .Animate Xor AniBounce
    End With
End Sub

Private Sub Form_Load()
    aTiger.Left = ScaleWidth
    aTim.Left = -aTim.Width
    Dim i As Long
    For i = 1 To aMove.UBound
        aMove(i).MousePointer = aMove(0).MousePointer
        If aMove(i).Tag <> "" Then aMove(i).Tag = aMove(i).Left
        aMove(i).CurFrame = i
    Next
End Sub

Private Sub Timer1_Timer()
    Dim x As Long, i As Long
    x = aTiger.Left - 60
    If x < -aTiger.Width Then
        x = ScaleWidth
        aMouse.CurFrame = 1
        aMouse.Animate = AniForward
        aMouse.Visible = True
    End If
    aTiger.Left = x
    
    With aTim
        x = .Tag
        If .FlipMode = 0 Then x = -x
        .Left = .Left + x
        If .Left > ScaleWidth Then .FlipMode = FlipNone
        If .Left < -.Width Then .FlipMode = FlipHorizontal
    End With
    
    With aEye
        .Move .Left + Rnd * 30 - 15, .Top + Rnd * 30 - 15
    End With
End Sub
