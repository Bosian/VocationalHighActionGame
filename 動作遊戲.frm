VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '單線固定
   Caption         =   "動作遊戲 v 3.1"
   ClientHeight    =   8520
   ClientLeft      =   1935
   ClientTop       =   1695
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "動作遊戲.frx":0000
   ScaleHeight     =   568
   ScaleMode       =   3  '像素
   ScaleWidth      =   794
   Begin VB.Timer Timer35 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   2760
      Top             =   2760
   End
   Begin VB.Timer Timer18 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   5040
      Top             =   5520
   End
   Begin VB.Timer Timer34 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   1080
      Top             =   5640
   End
   Begin VB.Timer Timer33 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   100
      Left            =   12960
      Top             =   8160
   End
   Begin VB.Timer Timer32 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   100
      Left            =   12960
      Top             =   7680
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Index           =   7
      Interval        =   50
      Left            =   2760
      Top             =   1920
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Index           =   6
      Interval        =   160
      Left            =   2400
      Top             =   1920
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Index           =   5
      Interval        =   100
      Left            =   2040
      Top             =   1920
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Index           =   4
      Interval        =   10
      Left            =   1680
      Top             =   1920
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   100
      Left            =   1200
      Top             =   1920
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   30
      Left            =   840
      Top             =   1920
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   100
      Left            =   480
      Top             =   1920
   End
   Begin VB.Timer Timer20 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   200
      Left            =   8520
      Top             =   1560
   End
   Begin VB.Timer Timer21 
      Index           =   1
      Interval        =   100
      Left            =   9120
      Top             =   1560
   End
   Begin VB.Timer Timer31 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   480
      Top             =   5640
   End
   Begin VB.Timer Timer30 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1000
      Left            =   10080
      Top             =   0
   End
   Begin VB.Timer Timer29 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   3840
      Top             =   5040
   End
   Begin VB.Timer Timer28 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   30
      Left            =   3480
      Top             =   3480
   End
   Begin VB.Timer Timer27 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   30
      Left            =   3480
      Top             =   2880
   End
   Begin VB.Timer Timer26 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   240
      Top             =   360
   End
   Begin VB.Timer Timer25 
      Interval        =   500
      Left            =   2760
      Top             =   840
   End
   Begin VB.Timer Timer24 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   500
      Left            =   5880
      Top             =   3000
   End
   Begin VB.Timer Timer23 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   5400
      Top             =   3000
   End
   Begin VB.Timer Timer22 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   100
      Left            =   10560
      Top             =   6240
   End
   Begin VB.Timer Timer21 
      Index           =   0
      Interval        =   100
      Left            =   11040
      Top             =   1680
   End
   Begin VB.Timer Timer20 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   200
      Left            =   10560
      Top             =   1680
   End
   Begin VB.Timer Timer19 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   9120
      Top             =   4800
   End
   Begin VB.Timer Timer13 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   5040
      Top             =   4440
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   80
      Left            =   9600
      Top             =   5280
   End
   Begin VB.Timer Timer17 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   500
      Left            =   10440
      Top             =   4200
   End
   Begin VB.Timer Timer16 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   50
      Left            =   9960
      Top             =   4200
   End
   Begin VB.Timer Timer15 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   11160
      Top             =   2400
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   120
      Top             =   1920
   End
   Begin VB.Timer Timer12 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   30
      Left            =   11160
      Top             =   5280
   End
   Begin VB.Timer Timer11 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   100
      Left            =   12840
      Top             =   11280
   End
   Begin VB.Timer Timer10 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   7560
      Top             =   4800
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   100
      Left            =   6960
      Top             =   5280
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   100
      Left            =   6000
      Top             =   5280
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6360
      Top             =   3960
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   100
      Left            =   10560
      Top             =   11520
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   100
      Left            =   8520
      Top             =   11280
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   100
      Left            =   6600
      Top             =   11280
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   7560
      Top             =   1680
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   100
      Left            =   4920
      Top             =   11160
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   7200
      Top             =   5760
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   6720
      Top             =   5760
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   6240
      Top             =   5760
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   5760
      Top             =   5760
   End
   Begin VB.Shape Shape16 
      BorderWidth     =   10
      Height          =   255
      Index           =   0
      Left            =   2760
      Shape           =   2  '橢圓形
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   0
      Visible         =   0   'False
      X1              =   40
      X2              =   64
      Y1              =   248
      Y2              =   264
   End
   Begin VB.Image Image24 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   1
      Left            =   12960
      Picture         =   "動作遊戲.frx":240042
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image24 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   0
      Left            =   11520
      Picture         =   "動作遊戲.frx":240303
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image23 
      Height          =   645
      Left            =   2640
      Picture         =   "動作遊戲.frx":2405C5
      Top             =   1200
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label10 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "0 X"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Left            =   2025
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Image Image22 
      Height          =   330
      Index           =   0
      Left            =   1080
      Picture         =   "動作遊戲.frx":240B13
      Top             =   5160
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label9 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3060
      TabIndex        =   9
      Top             =   0
      Width           =   135
   End
   Begin VB.Image Image21 
      BorderStyle     =   1  '單線固定
      Height          =   705
      Index           =   7
      Left            =   14520
      Picture         =   "動作遊戲.frx":240F99
      Top             =   7080
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Image21 
      BorderStyle     =   1  '單線固定
      Height          =   705
      Index           =   6
      Left            =   15240
      Picture         =   "動作遊戲.frx":2414EC
      Top             =   7200
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Image21 
      BorderStyle     =   1  '單線固定
      Height          =   705
      Index           =   5
      Left            =   15960
      Picture         =   "動作遊戲.frx":241A1A
      Top             =   7080
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Image21 
      BorderStyle     =   1  '單線固定
      Height          =   705
      Index           =   4
      Left            =   16680
      Picture         =   "動作遊戲.frx":241F95
      Top             =   7080
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Image21 
      BorderStyle     =   1  '單線固定
      Height          =   705
      Index           =   3
      Left            =   14400
      Picture         =   "動作遊戲.frx":2424E2
      Top             =   6240
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Image21 
      BorderStyle     =   1  '單線固定
      Height          =   705
      Index           =   2
      Left            =   15120
      Picture         =   "動作遊戲.frx":242A48
      Top             =   6360
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Image21 
      BorderStyle     =   1  '單線固定
      Height          =   705
      Index           =   1
      Left            =   15840
      Picture         =   "動作遊戲.frx":242F87
      Top             =   6240
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Image21 
      BorderStyle     =   1  '單線固定
      Height          =   705
      Index           =   0
      Left            =   16680
      Picture         =   "動作遊戲.frx":2434D5
      Top             =   6240
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Image20 
      Height          =   645
      Index           =   0
      Left            =   12120
      Picture         =   "動作遊戲.frx":243A25
      Top             =   7440
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Image Image19 
      Height          =   1005
      Left            =   2160
      Picture         =   "動作遊戲.frx":243F75
      Top             =   6480
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image Image18 
      Height          =   270
      Left            =   720
      Picture         =   "動作遊戲.frx":245524
      Top             =   6840
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image Image17 
      Height          =   2775
      Left            =   480
      Picture         =   "動作遊戲.frx":245C24
      Top             =   5880
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Shape Shape14 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      Height          =   375
      Left            =   480
      Shape           =   3  '圓形
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image16 
      Height          =   735
      Left            =   600
      Top             =   9480
      Width           =   495
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   0
      Visible         =   0   'False
      X1              =   560
      X2              =   528
      Y1              =   248
      Y2              =   256
   End
   Begin VB.Label Label6 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "↑A"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   8730
      TabIndex        =   8
      Top             =   0
      Width           =   540
   End
   Begin VB.Shape delay 
      BorderColor     =   &H80000005&
      FillColor       =   &H0080FF80&
      FillStyle       =   0  '實心
      Height          =   855
      Index           =   1
      Left            =   8880
      Top             =   360
      Width           =   240
   End
   Begin VB.Image Image14 
      Height          =   315
      Index           =   1
      Left            =   3960
      Picture         =   "動作遊戲.frx":2480C1
      Top             =   10080
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image14 
      Height          =   315
      Index           =   0
      Left            =   3960
      Picture         =   "動作遊戲.frx":24820C
      Top             =   10440
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image13 
      Height          =   315
      Index           =   0
      Left            =   3960
      Picture         =   "動作遊戲.frx":248358
      Top             =   4560
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   6
      Index           =   0
      Visible         =   0   'False
      X1              =   288
      X2              =   288
      Y1              =   200
      Y2              =   224
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "Game Over..."
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5160
      TabIndex        =   7
      Top             =   4200
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label Label7 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "野鳥"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5700
      TabIndex        =   6
      Top             =   0
      Width           =   720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      Index           =   0
      Visible         =   0   'False
      X1              =   320
      X2              =   320
      Y1              =   184
      Y2              =   208
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   255
      Left            =   4440
      Shape           =   2  '橢圓形
      Top             =   2640
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label6 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   10800
      TabIndex        =   5
      Top             =   0
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "1000/1000"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   1620
   End
   Begin VB.Shape delay 
      BorderColor     =   &H80000005&
      FillColor       =   &H0080FF80&
      FillStyle       =   0  '實心
      Height          =   855
      Index           =   0
      Left            =   10800
      Top             =   360
      Width           =   240
   End
   Begin VB.Shape mp 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '實心
      Height          =   375
      Left            =   240
      Top             =   840
      Width           =   3000
   End
   Begin VB.Shape Shape8 
      FillStyle       =   0  '實心
      Height          =   135
      Index           =   0
      Left            =   9120
      Shape           =   2  '橢圓形
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  '實心
      Height          =   255
      Index           =   0
      Left            =   9120
      Shape           =   3  '圓形
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   1200
      Left            =   6000
      Picture         =   "動作遊戲.frx":2484A4
      Top             =   3840
      Width           =   1200
   End
   Begin VB.Label Label4 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Index           =   0
      Left            =   10515
      TabIndex        =   3
      Top             =   4680
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label5 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "1000/1000"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5280
      TabIndex        =   4
      Top             =   360
      Width           =   1620
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "G O →"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   9960
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "1000/1000"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   1620
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FF8080&
      FillStyle       =   0  '實心
      Height          =   375
      Left            =   240
      Top             =   840
      Width           =   3000
   End
   Begin VB.Shape hp 
      BorderColor     =   &H80000005&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '實心
      Height          =   375
      Left            =   240
      Top             =   360
      Width           =   3000
   End
   Begin VB.Shape xhp 
      BorderColor     =   &H80000005&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '實心
      Height          =   375
      Left            =   4560
      Top             =   360
      Width           =   3000
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   375
      Left            =   4560
      Top             =   360
      Width           =   3000
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   375
      Left            =   240
      Top             =   360
      Width           =   3000
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   40
      X2              =   728
      Y1              =   104
      Y2              =   104
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   7
      Left            =   12000
      Picture         =   "動作遊戲.frx":248683
      Top             =   10440
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   6
      Left            =   12120
      Picture         =   "動作遊戲.frx":2489A1
      Top             =   10320
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   5
      Left            =   12240
      Picture         =   "動作遊戲.frx":248C85
      Top             =   10200
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   4
      Left            =   12360
      Picture         =   "動作遊戲.frx":248F53
      Top             =   10080
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   3
      Left            =   12480
      Picture         =   "動作遊戲.frx":249214
      Top             =   9960
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   2
      Left            =   12600
      Picture         =   "動作遊戲.frx":2494E2
      Top             =   9840
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   1
      Left            =   12720
      Picture         =   "動作遊戲.frx":2497CB
      Top             =   9720
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   0
      Left            =   12840
      Picture         =   "動作遊戲.frx":249AE9
      Top             =   9600
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image6 
      Height          =   960
      Index           =   0
      Left            =   10080
      Picture         =   "動作遊戲.frx":249DE9
      Top             =   4920
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   1
      Left            =   12960
      Picture         =   "動作遊戲.frx":24A0E6
      Top             =   11640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   2
      Left            =   12840
      Picture         =   "動作遊戲.frx":24A403
      Top             =   11760
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   0
      Left            =   13080
      Picture         =   "動作遊戲.frx":24A6E3
      Top             =   11520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   3
      Left            =   12720
      Picture         =   "動作遊戲.frx":24A9E0
      Top             =   11880
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   4
      Left            =   12600
      Picture         =   "動作遊戲.frx":24ACB2
      Top             =   12000
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   6
      Left            =   12360
      Picture         =   "動作遊戲.frx":24AF74
      Top             =   12240
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   7
      Left            =   12240
      Picture         =   "動作遊戲.frx":24B250
      Top             =   12360
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   5
      Left            =   12480
      Picture         =   "動作遊戲.frx":24B56D
      Top             =   12120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   5
      Left            =   9840
      Picture         =   "動作遊戲.frx":24B83F
      Top             =   10320
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   4
      Left            =   9960
      Picture         =   "動作遊戲.frx":24BA66
      Top             =   10200
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   3
      Left            =   10080
      Picture         =   "動作遊戲.frx":24BC9F
      Top             =   10080
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   2
      Left            =   10200
      Picture         =   "動作遊戲.frx":24BECF
      Top             =   9960
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   1
      Left            =   10440
      Picture         =   "動作遊戲.frx":24C108
      Top             =   9720
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image9 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   3
      Left            =   8160
      Picture         =   "動作遊戲.frx":24C32F
      Top             =   12120
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image9 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   2
      Left            =   8280
      Picture         =   "動作遊戲.frx":24C514
      Top             =   12000
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image9 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   1
      Left            =   8400
      Picture         =   "動作遊戲.frx":24C74F
      Top             =   11880
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   4
      Left            =   6480
      Picture         =   "動作遊戲.frx":24C934
      Top             =   10320
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   3
      Left            =   6600
      Picture         =   "動作遊戲.frx":24CB17
      Top             =   10200
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   2
      Left            =   6720
      Picture         =   "動作遊戲.frx":24CD11
      Top             =   10080
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   1
      Left            =   6840
      Picture         =   "動作遊戲.frx":24CF0B
      Top             =   9960
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   5
      Left            =   4560
      Picture         =   "動作遊戲.frx":24D111
      Top             =   9960
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   4
      Left            =   4680
      Picture         =   "動作遊戲.frx":24D407
      Top             =   9840
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   3
      Left            =   4800
      Picture         =   "動作遊戲.frx":24D6FF
      Top             =   9720
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   2
      Left            =   4920
      Picture         =   "動作遊戲.frx":24D9DC
      Top             =   9600
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   1
      Left            =   5040
      Picture         =   "動作遊戲.frx":24DCD4
      Top             =   9480
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   0
      Left            =   10560
      Picture         =   "動作遊戲.frx":24DFCA
      Top             =   9600
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image9 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   0
      Left            =   8520
      Picture         =   "動作遊戲.frx":24E1F8
      Top             =   11760
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   0
      Left            =   6960
      Picture         =   "動作遊戲.frx":24E4B5
      Top             =   9840
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   0
      Left            =   5160
      Picture         =   "動作遊戲.frx":24E6E9
      Top             =   9360
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   3
      Left            =   8160
      Picture         =   "動作遊戲.frx":24E9E1
      Top             =   10080
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   2
      Left            =   8280
      Picture         =   "動作遊戲.frx":24EC96
      Top             =   9960
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   1
      Left            =   8400
      Picture         =   "動作遊戲.frx":24EE75
      Top             =   9840
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   0
      Left            =   8520
      Picture         =   "動作遊戲.frx":24F0AE
      Top             =   9720
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   0
      Left            =   10320
      Picture         =   "動作遊戲.frx":24F28D
      Top             =   11640
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   1
      Left            =   10200
      Picture         =   "動作遊戲.frx":24F4B4
      Top             =   11760
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   2
      Left            =   10080
      Picture         =   "動作遊戲.frx":24F6E2
      Top             =   11880
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   3
      Left            =   9960
      Picture         =   "動作遊戲.frx":24F911
      Top             =   12000
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   4
      Left            =   9840
      Picture         =   "動作遊戲.frx":24FB41
      Top             =   12120
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   5
      Left            =   9720
      Picture         =   "動作遊戲.frx":24FD70
      Top             =   12240
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   0
      Left            =   6480
      Picture         =   "動作遊戲.frx":24FF9E
      Top             =   11880
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   1
      Left            =   6360
      Picture         =   "動作遊戲.frx":2501D0
      Top             =   12000
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   2
      Left            =   6240
      Picture         =   "動作遊戲.frx":2503D4
      Top             =   12120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   3
      Left            =   6120
      Picture         =   "動作遊戲.frx":2505D2
      Top             =   12240
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   4
      Left            =   6000
      Picture         =   "動作遊戲.frx":2507CC
      Top             =   12360
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   0
      Left            =   4800
      Picture         =   "動作遊戲.frx":2509AB
      Top             =   11520
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   1
      Left            =   4680
      Picture         =   "動作遊戲.frx":250CA1
      Top             =   11640
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   2
      Left            =   4560
      Picture         =   "動作遊戲.frx":250F9A
      Top             =   11760
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   3
      Left            =   4440
      Picture         =   "動作遊戲.frx":251293
      Top             =   11880
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   4
      Left            =   4320
      Picture         =   "動作遊戲.frx":251572
      Top             =   12000
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   5
      Left            =   4200
      Picture         =   "動作遊戲.frx":25186B
      Top             =   12120
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   1215
      Index           =   0
      Left            =   4320
      Shape           =   2  '橢圓形
      Top             =   4560
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  '實心
      Height          =   255
      Index           =   0
      Left            =   10440
      Shape           =   2  '橢圓形
      Top             =   5880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  '實心
      Height          =   255
      Left            =   6360
      Shape           =   2  '橢圓形
      Top             =   5280
      Width           =   495
   End
   Begin VB.Image Image12 
      Height          =   13365
      Left            =   14400
      Picture         =   "動作遊戲.frx":251B64
      Top             =   -240
      Visible         =   0   'False
      Width           =   18270
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      Height          =   255
      Index           =   0
      Left            =   3960
      Shape           =   2  '橢圓形
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Shape12 
      FillColor       =   &H00400040&
      FillStyle       =   0  '實心
      Height          =   2535
      Left            =   4080
      Shape           =   2  '橢圓形
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape13 
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  '實心
      Height          =   615
      Index           =   0
      Left            =   4560
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      Visible         =   0   'False
      X1              =   256
      X2              =   136
      Y1              =   768
      Y2              =   824
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      Visible         =   0   'False
      X1              =   16
      X2              =   136
      Y1              =   768
      Y2              =   824
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   5
      Visible         =   0   'False
      X1              =   136
      X2              =   256
      Y1              =   744
      Y2              =   800
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      Visible         =   0   'False
      X1              =   16
      X2              =   136
      Y1              =   800
      Y2              =   744
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      Visible         =   0   'False
      X1              =   14
      X2              =   262
      Y1              =   800
      Y2              =   800
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      Visible         =   0   'False
      X1              =   16
      X2              =   256
      Y1              =   768
      Y2              =   768
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   0
      Shape           =   2  '橢圓形
      Top             =   11160
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Shape Shape15 
      FillStyle       =   0  '實心
      Height          =   135
      Index           =   0
      Left            =   12240
      Shape           =   2  '橢圓形
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ☆自創的副刊式●程式本身就有的副程式★計時器

Dim s(19) As Integer '0)跳躍計數 1)地圖計數 2)普攻回原位計數 3)火柱範圍 4)GO的閃爍 5)顯示傷害 6)A delay閃爍 7)每隻鳥動畫的不同 8)子彈產生數 9)制裁之光計數 10)流星散彈雨產生計數 11)絕招集氣 12)超必殺(集氣)之計秒關閉 13)四魂之玉計數 14)破魔連珠之計數 15) 流星散彈雨過右邊邊界之計數 16)破魔連珠 17)殘月漣漪 18)殘月漣漪集氣用 19)超念動之震波控制
Dim sic(10) As Integer '四魂之玉消失
Dim zxcv(5) As Integer '0)限制跳躍時不可上下移動 1)限制普攻不可按住 2)限制A絕招 3)判斷角色往左或往右 4)限制只能讀一次移除程序 5)限制鍵盤不能按(用於絕招集氣
Dim x As Integer '跳躍方向
Dim xy() As Integer '鳥方向
Dim ttop '暫存top
Dim k(3) As Integer '動畫1～３)人物動畫
Dim sk(10) '四魂之玉轉動
Dim dx
Dim sh As Integer '影子距離
Dim ax(1) As Integer '跑步判斷
Dim run(1) As Integer '跑步增值器
Dim ok As Integer 'ok為０則地圖移動為１則人物移動
Dim ma As Integer '數量
Dim pps As Integer '74被擊中的那一顆
Dim ppm() As Integer '0)血 1)已擊破 2)鳥閃爍 4)鳥動畫 5)停止判斷 6)計算時間 時間到則call子彈產生演算法
Dim n As Integer '擊破數
Dim Y '傷害值
Dim akiz As Integer '真實寬度放大率
Dim amax As Integer '真實寬度放大率(最大值)
Dim iss As Integer '動畫公式
Dim ball(1, 30) As Integer '0子彈方向儲存 1物件已移除
Dim ddd As Integer '掛掉數值
Dim dzxc As Integer '流星散彈雨的上下判斷
Dim stat() As Integer '流星散彈雨的上下判斷陣列
Dim ell(0) As Integer '0) 流星散彈雨(正負)決定往右或左移動
Dim up As Integer '觸發第二絕招之一按鍵
Dim slow(1, 0) As Integer 'delay之陣列
Dim ven(15) As Integer '1~15 為 六芒星的頂點 第一隻人物的絕招二
Dim hall(1, 20) As Integer '破魔連珠左右判斷0)左右 1)上下
Dim stong As Integer '爆發狀態
Dim vnn(4, 1) As Integer '制裁之光的光環指定
Dim wmy As Integer '殘月漣漪集氣之後打出的數量
Dim cir(10) As Integer '超念動之放大系數
Private Sub Timer26_Timer() '掛掉延遲
Timer26.Enabled = False
DoEvents
Call desd(1)
End Sub
Private Sub ex(a) '移除程序☆
If zxcv(4) = 1 Then Exit Sub
zxcv(4) = 1
If a = 1 Then '掛掉時
    Label8.Visible = True: Image1.Visible = False: Shape1.Visible = False: Timer26.Enabled = True '啟動延遲
Else '按F2時
    Call desd(1)
End If
End Sub
Public Sub desd(a) '移除共用
For i = 1 To s(8) '移除子彈系統
    If ball(1, i) <> 1 Then '如果那個物件存在則
        Unload Timer19(i)
        Unload Shape7(i)
        Unload Shape8(i)
    End If
Next
For f = 1 To ma '移除鳥的系統
    If ppm(1, f) <> 1 Then '如果那個物件存在則
        Unload Timer12(f)
        Unload Timer11(f)
        Unload Shape2(f)
        Unload Image6(f)
    End If
    Unload Timer22(f)
Next
n = 0
ma = 0
iss = 0
aikz = 0
amax = 0
Erase slow
Erase ven
Erase ball
Erase s
Erase zxcv
Erase xy
Erase k
Erase ax
Erase run
iss = 0
Y = 0
ok = 1
If a = 1 Then
    dx = 0
    gz = 1
    Unload Me
    Form2.Show
Else
    Call Form_Load
End If
End Sub
Private Sub Form_Load() '表單載入●
'限制動作遊戲只能讀"一次"的Load程式碼
If reset(1) = 0 Then
    reset(1) = 1
    
    Form1.Left = 1845
    Form1.Top = 660
    Form1.Width = 12000 '(800*600)
    Form1.Height = 9000
End If

ok = 1




Timer5(kikyou).Enabled = True
Select Case kikyou
    Case 0
        Image1.Picture = Image2(0).Picture
        sh = 70
        Label6(0).Caption = "制裁之光"
        Label6(1).Caption = "桔梗的意志"
    Case 1
        Image1.Picture = Image3(0).Picture
        sh = 53
        Label6(0).Caption = "神聖火柱"
        Label6(1).Caption = "崩裂"
    Case 2
        Image1.Picture = Image4(0).Picture
        sh = 70
        Label6(0).Caption = "極光暴雷"
        Label6(1).Caption = "殘月漣漪"
    Case 3
        Image1.Picture = Image5(0).Picture
        sh = 60
        Label6(0).Caption = "流星散彈雨"
        Label6(1).Caption = "超念動"
End Select
Shape1.Top = Image1.Top + sh
Shape1.Left = Image1.Left + Image1.Width \ 2 - 20
x = 1
akiz = 5 '(5)*200=amax
amax = 1000

hp.Width = 200
mp.Width = 200
Label1.Caption = hp.Width * akiz & "/" & amax
Label2.Caption = mp.Width * akiz & "/" & amax
Label1.Left = Shape4.Left + Shape4.Width \ 2 - Label1.Width \ 2
Label2.Left = Shape4.Left + Shape4.Width \ 2 - Label2.Width \ 2

For i = 0 To 1
    run(i) = 1
Next

For f = 1 To 10
    Call nnee
Next

ReDim xy(ma), ppm(6, ma)
For f = 1 To ma
    If f Mod 8 = 0 Then iss = iss + 1
    ppm(4, f) = f - 8 * iss
    
    xy(f) = 1
    Load Timer22(f)
    Load Timer11(f)
    Timer11(f).Enabled = True '鳥動畫
Next
'If music = 0 Then
'    music = 1
'    sndPlaySound "背景.wav", 9
'    '奇摩知識
'    '    Set dsz = dxa.DirectSoundCreate("")
'    '    dsz.SetCooperativeLevel Me.hWnd, DSSCL_NORMAL
'    '    bu.lFlags = DSBCAPS_STATIC
'    '    Set dsbz = dsz.CreateSoundBufferFromFile(("背景.wav"), bu, wf)
'    '    dsbz.Play DSBPLAY_LOOPING '''''''''''''''''''''''預設播放(單次).
'    '奇摩知識
'End If
End Sub
Public Sub nnee() '產生演算法☆
If ma < 20 Then
    ma = ma + 1
    
    Randomize
    
    Load Image6(ma)
    Image6(ma).Top = Int(Rnd * (Form1.ScaleHeight - Image6(ma).Height - Line1.Y1 - 10)) + Line1.Y1
    Image6(ma).Left = Int(Rnd * (Form1.ScaleWidth * 2 - Form1.ScaleWidth)) + Form1.ScaleWidth
    Image6(ma).Visible = True
    
    Load Shape2(ma)
    Shape2(ma).Top = Image6(ma).Top + 65
    Shape2(ma).Left = Image6(ma).Left + Image6(ma).Width \ 2
    Shape2(ma).Visible = True
    Shape2(ma).ZOrder 1
    
    Load Timer12(ma)
    Timer12(ma).Enabled = True '移動
End If
End Sub
Public Sub spdd(a) '子彈產生演算法
If s(8) >= 30 Then s(8) = 0
s(8) = s(8) + 1
Load Timer19(s(8)) '產生子彈移動
Load Shape7(s(8)) '產生子彈
Load Shape8(s(8)) '產生子彈影子

ball(0, s(8)) = xy(a) '存入發射子彈的那隻鳥的方向

If ball(0, s(8)) = 1 Then Shape7(s(8)).Left = Image6(a).Left - Shape7(s(8)).Width Else Shape7(s(8)).Left = Image6(a).Left + Image6(a).Width '指定子彈初始位置
Shape7(s(8)).Top = Image6(a).Top + Image6(a).Height \ 2 '指定子彈初始位置
Shape8(s(8)).Left = Shape7(s(8)).Left '指定影子初始位置
Shape8(s(8)).Top = Shape7(s(8)).Top + 32 '指定影子初始位置
Shape7(s(8)).Visible = True
Shape8(s(8)).Visible = True

Timer19(s(8)).Enabled = True '啟動子彈移動
End Sub
Private Sub Timer19_Timer(Index As Integer) '敵人子彈移動
    If Timer6.Enabled = True Then Shape7(Index).Left = Shape7(Index).Left - 5 '地圖移動
    
    Shape7(Index).Left = Shape7(Index).Left - 10 * ball(0, Index) '移動
    Shape8(Index).Left = Shape7(Index).Left
    
    If coll2(Shape8(Index), Shape1) Then '判斷有無打中角色
        If coll(Shape7(Index), Image1) Then
            Call kizzs
            If hp.Visible = True Then ball(1, Index) = 1: Unload Timer19(Index): Unload Shape7(Index): Unload Shape8(Index): Exit Sub '如果打中則角色損血
        End If
    End If
    
    If gz = 1 Then Exit Sub: Timer19(Index).Enabled = False
    If con1(Shape7(Index)) Or con2(Shape7(Index)) Then '如果超出邊界則移除
        ball(1, Index) = 1
        Unload Timer19(Index)
        Unload Shape7(Index)
        Unload Shape8(Index)
    End If
    DoEvents
End Sub
Private Sub Timer12_Timer(Index As Integer) '移動★

    If Timer6.Enabled = True Then Image6(Index).Left = Image6(Index).Left - 10 * run(1)
    
    If ppm(5, Index) = 0 Then Image6(Index).Left = Image6(Index).Left - 10 * xy(Index) '允許移動(沒有被暴雷打到)
    
    If xy(Index) = 1 Then
        Shape2(Index).Left = Image6(Index).Left + Image6(Index).Width \ 2
    Else
        Shape2(Index).Left = (Image6(Index).Left + Image6(Index).Width \ 2) - 20
    End If
    
    If ppm(5, Index) = 0 Then '沒有被暴雷打到時偵測子彈發射條件
        If xy(Index) = 1 And Image1.Left < Image6(Index).Left Or xy(Index) = -1 And Image1.Left + Image1.Width > Image6(Index).Left + Image6(Index).Width Then '發射條件(鳥看到角色)
            If coll2(Shape2(Index), Shape1) Or coll2(Shape1, Shape2(Index)) Then
                ppm(6, Index) = ppm(6, Index) + 1
                If ppm(6, Index) Mod 10 = 0 Then ppm(6, Index) = 0: Call spdd(Index) '呼叫產生子彈程序
            Else
                ppm(6, Index) = 0
            End If
        End If
    End If
    
    If Timer6.Enabled = True Then Exit Sub '地圖移動時，不判斷怪物的邊界反彈
    
    If Image6(Index).Left < 0 Then xy(Index) = -1: Shape2(Index).Left = Image6(Index).Left + Image6(Index).Width \ 2 - 20 '反向判斷
    If Image6(Index).Left + Image6(Index).Width > Form1.ScaleWidth Then xy(Index) = 1
    DoEvents
End Sub
Private Sub Timer11_Timer(Index As Integer) '鳥動畫★
If ppm(4, Index) = 7 Then
    ppm(4, Index) = 0
Else
    ppm(4, Index) = ppm(4, Index) + 1
End If

If xy(Index) = 1 Then
    Image6(Index).Picture = Image15(ppm(4, Index)).Picture
Else
    Image6(Index).Picture = Image11(ppm(4, Index)).Picture
End If
End Sub
Private Sub Timer7_Timer(Index As Integer) '鳥消失動畫★
ppm(2, Index) = ppm(2, Index) + 1
If ppm(2, Index) Mod 2 = 0 Then
    Image6(Index).Visible = True
    If ppm(2, Index) >= 12 Then
    
        '四魂之玉掉落
        Call sico(Index)
        
        Unload Shape2(Index)
        Unload Image6(Index)
        Unload Timer7(Index)
        Unload Timer11(Index)
        Unload Timer12(Index)
        ppm(2, Index) = 0
        If n = ma Then ok = 0: Timer15.Enabled = True '顯示GO下一關
        Exit Sub
    End If
Else
    Image6(Index).Visible = False
End If

If Timer6.Enabled = True Then '如果背景移動則被擊破的鳥也一起移動
    Image6(Index).Left = Image6(Index).Left - 20 * run(1)
    If xy(Index) = 1 Then
        Shape2(Index).Left = Image6(Index).Left + Image6(Index).Width \ 2
    Else
        Shape2(Index).Left = (Image6(Index).Left + Image6(Index).Width \ 2) - 20
    End If
End If
DoEvents
End Sub
Private Sub Timer22_Timer(Index As Integer) '鳥的後退
If ppm(1, Index) <> 1 Then ppm(5, Index) = 0: Timer11(Index).Enabled = True '啟動被打到而停止的特效
Timer22(Index).Enabled = False '停止自己
DoEvents
End Sub
Private Sub timer30_Timer(Index As Integer) '絕招集氣
Image16.Picture = Form1.Picture
Select Case Index
    Case 0
        zxcv(5) = 1 '鍵盤鎖定
        For f = 1 To 9
            Load Timer30(f) '載入集氣動畫
            If f >= 2 Then Timer30(f).Enabled = False: Timer30(f).Interval = 50
        Next
        Timer30(1).Interval = 10
        Timer30(1).Enabled = True
        Timer30(0).Enabled = False
        Exit Sub
    Case 1 '集氣動畫
        For f = 1 To 8 '讀8次
            s(11) = s(11) + 1
            Load Line10(s(11))
            Load Shape13(s(11))
            Select Case s(11)
                Case 1 To 3 '左上
                    Shape13(s(11)).Left = Image1.Left - Image1.Width
                    Shape13(s(11)).Top = Image1.Top + (-Image1.Height * -(f - 2)) + Image1.Height \ 2
                    'Shape13(s(11)).Visible = True
                    
                    Line10(s(11)).X1 = Shape13(s(11)).Left
                    If s(11) = 1 Then '註：我實在想不出來他們的共通點，只好如此...(想到快發瘋.....= ="")
                        
                        Line10(s(11)).Y1 = Shape13(s(11)).Top
                        Line10(s(11)).X2 = Shape13(s(11)).Left + Shape13(s(11)).Width
                        Line10(s(11)).Y2 = Shape13(s(11)).Top + Shape13(s(11)).Height
                    Else
                        If s(11) = 2 Then
                            Line10(s(11)).Y1 = Shape13(s(11)).Top + Shape13(s(11)).Height \ 2
                            Line10(s(11)).X2 = Shape13(s(11)).Left + Shape13(s(11)).Width
                            Line10(s(11)).Y2 = Line10(s(11)).Y1
                        Else
                            Line10(s(11)).Y1 = Shape13(s(11)).Top + Shape13(s(11)).Height
                            Line10(s(11)).X2 = Shape13(s(11)).Left + Shape13(s(11)).Width
                            Line10(s(11)).Y2 = Shape13(s(11)).Top
                        End If
                    End If
                    
                    Line10(s(11)).Visible = True
                Case 4 To 6
                    Shape13(s(11)).Left = Image1.Left + Image1.Width * 1.5
                    Shape13(s(11)).Top = Image1.Top + (-Image1.Height * -(f - 5)) + Image1.Height \ 2
                    'Shape13(s(11)).Visible = True
                    
                    Line10(s(11)).X1 = Shape13(s(11)).Left + Shape13(s(11)).Width
                    If s(11) = 4 Then
                        
                        Line10(s(11)).Y1 = Shape13(s(11)).Top
                        Line10(s(11)).X2 = Shape13(s(11)).Left
                        Line10(s(11)).Y2 = Shape13(s(11)).Top + Shape13(s(11)).Height
                    Else
                        If s(11) = 5 Then
                            Line10(s(11)).Y1 = Shape13(s(11)).Top + Shape13(s(11)).Height \ 2
                            Line10(s(11)).X2 = Shape13(s(11)).Left
                            Line10(s(11)).Y2 = Line10(s(11)).Y1
                        Else
                            Line10(s(11)).Y1 = Shape13(s(11)).Top + Shape13(s(11)).Height
                            Line10(s(11)).X2 = Shape13(s(11)).Left
                            Line10(s(11)).Y2 = Shape13(s(11)).Top
                        End If
                    End If
                    
                    Line10(s(11)).Visible = True
                Case 7 To 8
                    Shape13(s(11)).Left = (Image1.Left + Image1.Width \ 2) - Shape13(s(11)).Width \ 2
                    If s(11) = 8 Then Shape13(s(11)).Top = Image1.Top + Image1.Height * 1.5 Else Shape13(s(11)).Top = Image1.Top - Image1.Height \ 2
                    'Shape13(s(11)).Visible = True
                    
                    Line10(s(11)).X1 = Shape13(s(11)).Left + Shape13(s(11)).Width \ 2
                    Line10(s(11)).Y1 = Shape13(s(11)).Top
                    Line10(s(11)).X2 = Shape13(s(11)).Left + Shape13(s(11)).Width \ 2
                    Line10(s(11)).Y2 = Shape13(s(11)).Top + Shape13(s(11)).Height
                    
                    Line10(s(11)).Visible = True
            End Select
            DoEvents
        Next
        For f = 2 To 9
            Timer30(f).Enabled = True '移動
        Next
        Unload Timer30(1)
        Exit Sub
    Case 2 '以下為移動
        Line10(Index - 1).X1 = Line10(Index - 1).X1 + 10
        Line10(Index - 1).Y1 = Line10(Index - 1).Y1 + 10
        Line10(Index - 1).X2 = Line10(Index - 1).X2 + 10
        Line10(Index - 1).Y2 = Line10(Index - 1).Y2 + 10
    Case 3
        Line10(Index - 1).X1 = Line10(Index - 1).X1 + 10
        Line10(Index - 1).X2 = Line10(Index - 1).X2 + 10
    Case 4
        Line10(Index - 1).X1 = Line10(Index - 1).X1 + 10
        Line10(Index - 1).Y1 = Line10(Index - 1).Y1 - 10
        Line10(Index - 1).X2 = Line10(Index - 1).X2 + 10
        Line10(Index - 1).Y2 = Line10(Index - 1).Y2 - 10
    Case 5
        Line10(Index - 1).X1 = Line10(Index - 1).X1 - 10
        Line10(Index - 1).Y1 = Line10(Index - 1).Y1 + 10
        Line10(Index - 1).X2 = Line10(Index - 1).X2 - 10
        Line10(Index - 1).Y2 = Line10(Index - 1).Y2 + 10
    Case 6
        Line10(Index - 1).X1 = Line10(Index - 1).X1 - 10
        Line10(Index - 1).X2 = Line10(Index - 1).X2 - 10
    Case 7
        Line10(Index - 1).X1 = Line10(Index - 1).X1 - 10
        Line10(Index - 1).Y1 = Line10(Index - 1).Y1 - 10
        Line10(Index - 1).X2 = Line10(Index - 1).X2 - 10
        Line10(Index - 1).Y2 = Line10(Index - 1).Y2 - 10
        
        If Line10(1).X1 > Image1.Left Then
            For f = 1 To 8
                s(11) = 0
                Unload Line10(f)
                Unload Shape13(f)
                Unload Timer30(f + 1)
            Next
            Call superx
            Exit Sub
        End If
    Case 8
        Line10(Index - 1).Y1 = Line10(Index - 1).Y1 + 10
        Line10(Index - 1).Y2 = Line10(Index - 1).Y2 + 10
    Case 9
        Line10(Index - 1).Y1 = Line10(Index - 1).Y1 - 10
        Line10(Index - 1).Y2 = Line10(Index - 1).Y2 - 10
End Select
End Sub
Private Sub superx() '超必殺(集氣絕招)
Form1.Picture = LoadPicture()
Shape14.Width = 25 '初始化
Shape14.Height = 25 '初始化
Shape14.Left = Image1.Left + Image1.Width \ 2 '設定位置
Shape14.Top = Image1.Top + Image1.Height \ 2 '設定位置
Shape14.Visible = True
Timer31.Enabled = True '啟動效果
End Sub
Private Sub Timer31_Timer() '超必殺之特效
s(12) = s(12) + 1
If s(12) = 25 Then
    Timer31.Enabled = False
    s(12) = 0
    Shape14.Visible = False
    Form1.Picture = Image16.Picture
    dx = 0
    zxcv(5) = 0
    stong = 1
    Label10.Caption = Val(Label10.Caption) - 1 & " X"
    If Val(Label10.Caption) = 0 Then Label10.Visible = False: Image23.Visible = False
    For f = 0 To 1
        delay(f).FillColor = &H80FFFF '爆發狀態'顏色改變
    Next
    Exit Sub
End If
Shape14.Left = Shape14.Left - 20
Shape14.Width = Shape14.Width + 40
Shape14.Top = Shape14.Top - 20
Shape14.Height = Shape14.Height + 40
DoEvents
End Sub
Private Sub Timer27_Timer(Index As Integer) '制裁之光
If vnn(Index, 0) = (Index + 1) * 15 Then '如果已經執行完畢則
    If vnn(Index, 1) <= 12 And vnn(Index, 1) Mod 3 = 0 Or stong = 1 Then
        For f = 1 To ma
            If ppm(1, f) <> 1 Then
                If coll(Image6(f), Shape3(Index)) Then pps = Image6(f).Index: Call power(0)
            End If
        Next
    End If
    Unload Shape11(vnn(Index, 1)) '慢慢消失
    Unload Line3(vnn(Index, 1)) '慢慢消失
    vnn(Index, 1) = vnn(Index, 1) + 1 '指定消失的那顆
    
    If vnn(Index, 1) = (Index + 1) * 15 + 1 Then '結束
        If Index <> 0 Then '如果不是本尊計時器則
            Unload Timer27(Index) '移除分身計時器
            Unload Shape3(Index) '移除分身大光圈
        Else
            Timer27(Index).Enabled = False
            Shape3(Index).Visible = False
            Call sg(0) '呼叫絕招移除共用
        End If
    End If
Else '剛開始執行
    vnn(Index, 0) = vnn(Index, 0) + 1
    
    Shape11(vnn(Index, 0)).Left = Shape11(vnn(Index, 0) - 1).Left '下一個光圈位置
    Shape11(vnn(Index, 0)).Top = Shape11(vnn(Index, 0) - 1).Top + Shape11(vnn(Index, 0) - 1).Height '下一個光圈位置
    Shape11(vnn(Index, 0)).Visible = True
    
    '設定位置
        Shape3(Index).Left = Shape11(vnn(Index, 0)).Left - Shape3(Index).Width \ 2 + Shape11(vnn(Index, 0)).Width \ 2 '範圍大小
        Shape3(Index).Top = Shape11(vnn(Index, 0)).Top - Shape3(Index).Height \ 2 + Shape11(vnn(Index, 0)).Height \ 2 '範圍大小
        Shape3(Index).Visible = True
    '/設定位置
    
    Line3(vnn(Index, 0)).X1 = Shape11(vnn(Index, 0)).Left + Shape11(vnn(Index, 0)).Width \ 2 '中間光柱的左上角位置
    Line3(vnn(Index, 0)).Y1 = Shape11(vnn(Index, 0)).Top '光柱的左上角
    Line3(vnn(Index, 0)).X2 = Shape11(vnn(Index, 0)).Left + Shape11(vnn(Index, 0)).Width \ 2 '光柱的右下角
    Line3(vnn(Index, 0)).Y2 = Shape11(vnn(Index, 0)).Top + Shape11(vnn(Index, 0)).Height '光柱的右下
    Line3(vnn(Index, 0)).Visible = True
    
    DoEvents
End If
End Sub
Private Sub Timer34_Timer(Index As Integer) '破魔連珠
Image22(Index).Left = Image22(Index).Left - 10 * hall(0, Index)
Image22(Index).Top = Image22(Index).Top - 10 * hall(1, Index)

For f = 1 To ma
    If ppm(1, f) <> 1 Then
        If coll(Image22(Index), Image6(f)) Then pps = Image6(f).Index: Call power(1)
    End If
Next
s(16) = s(16) + 1
If s(16) = 100 Then Unload Image22(Index): Unload Timer34(Index): s(16) = 0: Exit Sub
If Image22(Index).Left < 0 Then hall(0, Index) = -1
If Image22(Index).Left + Image22(Index).Width > Form1.ScaleWidth Then hall(0, Index) = 1
If Image22(Index).Top < 0 Then hall(1, Index) = -1
If Image22(Index).Top + Image22(Index).Height > Form1.ScaleHeight Then hall(1, Index) = 1
DoEvents
End Sub
Private Sub Timer13_Timer(Index As Integer) '火柱動畫
Shape3(Index).Top = Shape3(Index).Top - 8
If Timer1.Enabled = True Then Shape3(Index).Top = Shape3(Index).Top - 6.5: Shape3(0).Top = Shape3(0).Top - 2.5 '上固定:傷害固定
If Timer2.Enabled = True Then Shape3(0).Top = Shape3(0).Top + 2.5 ' 傷害固定
If Timer3.Enabled = True Or Timer4.Enabled = True Then Shape3(Index).Left = Image1.Left - Shape3(Index).Width \ 2 + Image1.Width \ 2: Shape3(0).Left = Image1.Left - Shape3(0).Width \ 2 + Image1.Width \ 2 '左右固定:傷害固定
If Shape3(Index).Top <= Image1.Top - Image1.Height * 2 Then Unload Shape3(Index): Unload Timer13(Index) '移除程序
DoEvents
End Sub
Private Sub Timer23_Timer(Index As Integer) '暴雷
If Line2(Index).Y2 >= Shape9.Top + Shape9.Height * 17 Then
    Unload Line2(Index)
    Unload Timer23(Index)
Else
    Line2(Index).Y2 = Line2(Index).Y2 + 9
End If
DoEvents
End Sub
Private Sub Timer18_Timer(Index As Integer) '殘月漣漪
Shape3(Index).Visible = True '水波顯現
Shape3(Index).Height = Shape3(Index).Height + 4 '水波範圍變大
Shape3(Index).Width = Shape3(Index).Width + 16 '水波範圍變大
Shape3(Index).Left = Shape3(Index).Left - 8 '水波範圍變大
Shape3(Index).Top = Shape3(Index).Top - 2 '水波範圍變大
If Shape3(Index).Width Mod 100 <= 5 Or stong = 1 Then
    For f = 1 To ma
        If ppm(1, f) <> 1 Then
            If coll(Image6(f), Shape3(Index)) Or coll(Shape3(Index), Image6(f)) Then pps = Image6(f).Index: Call power(0)
        End If
    Next
End If
DoEvents
If Shape3(Index).Width >= 500 Then
    Unload Shape3(Index): Unload Timer18(Index)
    If stong = 1 Then
        s(18) = s(18) + 1
        If s(18) = 4 + stong * wmy Then s(17) = 0: Call sg(6): s(18) = 0
    End If
End If
End Sub
Private Sub Timer29_Timer(Index As Integer) '流星散彈雨
Image13(Index).Left = Image13(Index).Left + 10 * ell(0)
Image13(Index).Top = Image13(Index).Top + 3 * stat(Index)
For f = 1 To ma
    If ppm(1, f) <> 1 Then
        If coll(Image13(Index), Image6(f)) Then
            pps = Image6(f).Index: Call power(3)
            If stong = 0 Then Unload Image13(Index): Unload Timer29(Index): Exit Sub
        End If
    End If
    DoEvents
Next
If con1(Image13(Index)) Or con2(Image13(Index)) Then
    Unload Image13(Index): Unload Timer29(Index)
    If stong = 1 Then
        s(15) = s(15) + 1
        If s(15) = 16 Then s(10) = 0: s(15) = 0: Call sg(3): Shape12.Visible = False '判斷集氣狀態的流星散彈雨發射完畢
    End If
End If
DoEvents
End Sub
Private Sub Timer35_Timer(Index As Integer) '超念動之震波

Shape16(Index).Left = Shape16(Index).Left - 4
Shape16(Index).Width = Shape16(Index).Width + 8
Shape16(Index).Top = Shape16(Index).Top - 2
Shape16(Index).Height = Shape16(Index).Height + 4
cir(3) = cir(3) + 1
If cir(3) Mod 10 = 0 Then
    If Shape16(Index).BorderWidth > 1 Then
        Shape16(Index).BorderWidth = Shape16(Index).BorderWidth - 1
    Else
        Unload Shape16(Index)
        Unload Timer35(Index)
    End If
End If
DoEvents
End Sub
Private Sub Timer14_Timer(Index As Integer) '絕招★
Select Case Index
    Case 0 '制裁之光
        For f = 0 To 4 * stong '如果為集氣狀態則產生1~4的計時器
            If f <> 0 Then
                Load Shape3(f)
                Load Timer27(f)
            End If
            Timer27(f).Enabled = True
            s(9) = 15 * (f + 1) '數量
            vnn(f, 0) = 15 * f + 1
            vnn(f, 1) = 15 * f + 1
        Next
        For f = 1 To s(9) '產生制裁之光所需的材料
            Load Shape11(f) '產生光圈
            Load Line3(f)
        Next
        For f = 1 To 4 * stong '亂數決定2~5(4個)制裁之光的位置
            Randomize
            Shape11(15 * f + 1).Top = Int(Rnd * (Form1.ScaleHeight - Shape3(0).Height * 1.5)) - Shape3(0).Height * 1.5
            Shape11(15 * f + 1).Left = Int(Rnd * Form1.ScaleWidth) + 1
        Next
        Timer14(Index).Enabled = False
    Case 1 '火柱
        s(3) = s(3) + 1
        Load Shape3(s(3)) '產生火柱
        Shape3(s(3)).Left = Image1.Left - Shape3(s(3)).Width \ 2 + Image1.Width \ 2 '設定火柱的位置
        Shape3(s(3)).Top = Image1.Top + Image1.Height \ 2 '設定火柱的位置
        Shape3(s(3)).Visible = True
        Load Timer13(s(3)) '產生火柱計時器
        Timer13(s(3)).Enabled = True '啟動火柱計時器
        If s(3) Mod 2 = 0 Then '五連段
            For f = 1 To ma
                If ppm(1, f) <> 1 Then
                    If coll2(Shape2(f), Shape3(0)) Then
                        If coll(Image6(f), Shape3(0)) Then pps = Image6(f).Index: Call power(1)
                    End If
                End If
            Next
        End If
        If s(3) = 10 Then s(3) = 0: Call sg(Index) '火柱已放完
    Case 2
        s(3) = s(3) + 1
        Load Line2(s(3)) '產生雷電
        Line2(0).Y2 = Shape1.Top + 6
        
        Line2(s(3)).X1 = Int(Rnd * Shape9.Width) + Shape9.Left
        Line2(s(3)).X2 = Line2(s(3)).X1
        
        Line2(s(3)).Y1 = Int(Rnd * Shape9.Height) + Shape9.Top
        Line2(s(3)).Y2 = Line2(s(3)).Y1
                
        Line2(s(3)).Visible = True
        Load Timer23(s(3)) '產生暴雷計時器
        Timer23(s(3)).Enabled = True
        
        If s(3) Mod 5 = 0 Then '三連段
            For f = 1 To ma
                If ppm(1, f) <> 1 Then
                    If coll2(Shape2(f), Shape3(0)) Then
                        If coll(Image6(f), Shape3(0)) Then pps = Image6(f).Index: Call power(2)
                    End If
                End If
            Next
        End If
        
        If s(3) = 15 Then s(3) = 0: Call sg(Index): Shape9.Visible = False '暴雷已放完
    Case 3 '流星散彈雨
        Randomize
        s(10) = s(10) + 1
        Load Image13(s(10))
        Load Timer29(s(10))
        Image13(s(10)).Left = Int(Rnd * Shape12.Width) + Shape12.Left
        Image13(s(10)).Top = Int(Rnd * Shape12.Height) + Shape12.Top
        Image13(s(10)).Visible = True
        stat(s(10)) = dzxc
        Image13(s(10)).ZOrder 0
        Timer29(s(10)).Enabled = True
        If s(10) = 16 Then
            Timer14(Index).Enabled = False
            If stong = 0 Then s(10) = 0: Call sg(Index): Shape12.Visible = False
        End If
    Case 4 '桔梗的意志
        
    '第一條主線....
    
        If ven(14) = 0 Then '畫六芒星
            If ven(0) = 0 Then
                Line4(0).X1 = Line4(0).X1 - 4
                If Line4(0).X1 <= Shape10.Left + 16 Then Line4(0).X1 = Shape10.Left + 16: ven(0) = 1: Line4(2).X1 = Line4(0).X1: Line4(2).X2 = Line4(0).X1
            End If
            If ven(1) = 0 Then
                Line4(0).Y1 = Line4(0).Y1 - 2
                If Line4(0).Y1 <= Shape10.Top + Shape10.Height \ 4 Then Line4(0).Y1 = Shape10.Top + Shape10.Height \ 4: ven(1) = 1: Line4(2).Y1 = Line4(0).Y1: Line4(2).Y2 = Line4(2).Y1: ven(5) = 1
            End If
            
        '第一條分線
            If ven(5) = 1 Then
                If ven(4) = 0 Then
                    Line4(2).X2 = Line4(2).X2 + 4
                    If Line4(2).X2 >= Shape10.Left + Shape10.Width - 17 Then Line4(2).X2 = Shape10.Left + Shape10.Width - 17: ven(4) = 1: ven(8) = 1: Line4(1).X1 = Line4(2).X2: Line4(1).Y1 = Line4(2).Y2: Line4(1).Y2 = Line4(1).Y1: Line4(1).X2 = Line4(1).X1
                End If
            End If
        '第二條分線
            If ven(8) = 1 Then
                If ven(9) = 0 Then
                    Line4(1).X2 = Line4(1).X2 - 4
                    If Line4(1).X2 <= Shape10.Left + Shape10.Width \ 2 Then Line4(1).X2 = Shape10.Left + Shape10.Width \ 2: ven(9) = 1
                End If
                If ven(10) = 0 Then
                    Line4(1).Y2 = Line4(1).Y2 + 2
                    If Line4(1).Y2 >= Shape10.Top + Shape10.Height Then Line4(1).Y2 = Shape10.Top + Shape10.Height: ven(10) = 1
                End If
            End If
            
    '第二條主線....
        
            If ven(2) = 0 Then
                Line4(5).X2 = Line4(5).X2 + 4
                If Line4(5).X2 >= Shape10.Left + Shape10.Width - 17 Then Line4(5).X2 = Shape10.Left + Shape10.Width - 17: ven(2) = 1: Line4(4).X2 = Line4(5).X2: Line4(4).X1 = Line4(4).X2
            End If
            If ven(3) = 0 Then
                Line4(5).Y2 = Line4(5).Y2 + 2
                If Line4(5).Y2 >= Shape10.Top + (Shape10.Height \ 4) * 3 Then Line4(5).Y2 = Shape10.Top + (Shape10.Height \ 4) * 3: ven(3) = 1: Line4(4).Y2 = Line4(5).Y2: Line4(4).Y1 = Line4(4).Y2: ven(7) = 1
            End If
        
        '第一條分線
            If ven(7) = 1 Then
                If ven(6) = 0 Then
                    Line4(4).X1 = Line4(4).X1 - 4
                    If Line4(4).X1 <= Shape10.Left + 16 Then Line4(4).X1 = Shape10.Left + 16: ven(6) = 1: ven(11) = 1: Line4(3).X2 = Line4(4).X1: Line4(3).X1 = Line4(3).X2: Line4(3).Y2 = Line4(4).Y1: Line4(3).Y1 = Line4(3).Y2
                End If
            End If
        '第二條分線
            If ven(11) = 1 Then
                If ven(12) = 0 Then
                    Line4(3).X1 = Line4(3).X1 + 4
                    If Line4(3).X1 >= Shape10.Left + Shape10.Width \ 2 Then Line4(3).X1 = Shape10.Left + Shape10.Width \ 2: ven(12) = 1
                End If
                
                If ven(13) = 0 Then
                    Line4(3).Y1 = Line4(3).Y1 - 2
                    If Line4(3).Y1 <= Shape10.Top Then Line4(3).Y1 = Shape10.Top: ven(13) = 1: ven(14) = 1 '結尾
                End If
            End If
                
        '強力一擊
        Else
            If ven(15) = 0 Then
                Image17.Left = Image1.Left - Image17.Width \ 2 + Image1.Width \ 2 '桔梗位置
                Image17.Top = Image1.Top - Image17.Height \ 2 + Image1.Height \ 2
                Image17.Visible = True
                
                Image18.Top = Image17.Top + 64 '箭矢位置
                Image18.Left = Image17.Left + 16
                Image18.Visible = True
                
                Image19.Top = Image18.Top - Image19.Height \ 2 + Image18.Height \ 2 '破魔之氣
                
                For f = 0 To 1 '破魔連珠初始之方向指定
                    For ff = 0 To 20
                        If ff Mod 2 = 0 Then hall(f, ff) = -1 Else hall(f, ff) = 1
                    Next
                Next
                ven(15) = 1
            Else
                Image17.Left = Image17.Left - 10
                If Image17.Left >= 1 Then
                    Image18.Left = Image17.Left + 16
                    Image18.Top = Image17.Top + 64
                Else '當桔梗到達邊界則
                
                    '打到判斷
                    For f = 1 To ma
                        If ppm(1, f) <> 1 Then
                            If coll(Image6(f), Image18) Or coll(Image18, Image6(f)) Then pps = Image6(f).Index: Call power(1)
                        End If
                    Next
                    
                    If stong = 1 Then '集氣判斷
                        If Image19.Left + Image19.Width > Form1.ScaleWidth Then
                            Image18.Left = Form1.ScaleWidth - Image18.Width
                            Image19.Top = Image18.Top - Image19.Height \ 2 + Image18.Height \ 2
                            If s(14) = 20 Then
                                Image18.Visible = False
                                Image19.Visible = False
                                
                                s(14) = 0
                                Call sg(Index) '絕招放完
                            Else
                                Image19.Left = Form1.ScaleWidth - Image19.Width
                                
                                '破魔連珠
                                s(14) = s(14) + 1
                                Load Timer34(s(14))
                                Load Image22(s(14))
                                
                                Randomize
                                Image22(s(14)).Left = Form1.ScaleWidth - Image22(s(14)).Width '設定位置
                                Image22(s(14)).Top = Int(Rnd * Image19.Height) + Image19.Top '設定位置
                                
                                Image22(s(14)).Visible = True
                                Timer34(s(14)).Enabled = True '啟動破魔連珠移動
                            End If
                        Else
                            Call kiker '呼叫桔梗的意志共用
                        End If
                    Else
                        Call kiker '呼叫桔梗的意志共用
                        If Image18.Left > Form1.ScaleWidth Then Call sg(Index) '絕招放完
                    End If
                End If
            End If
        End If
    Case 5 '崩裂
        For f = 1 To 8
            aaf = f Mod 8
            Select Case aaf
                Case 1
                    Line5(f).X2 = Line5(f).X2 + 10
                Case 2
                    Line5(f).X2 = Line5(f).X2 + 10
                    Line5(f).Y2 = Line5(f).Y2 - 5
                Case 3
                    Line5(f).X2 = Line5(f).X2 + 10
                    Line5(f).Y2 = Line5(f).Y2 - 10
                Case 4
                Case 5
                Case 6
                Case 7
                Case 0
            End Select
            DoEvents
        Next
        'Call sg(Index)
    Case 6 '殘月漣漪
        s(17) = s(17) + 1
        Timer18(s(17)).Enabled = True
        If s(17) = 4 + stong * wmy Then
            If stong = 0 Then
                s(17) = 0: Call sg(Index)
            Else
                Timer14(Index).Enabled = False
            End If
        End If
    Case 7 '超念動
        Form1.FillStyle = 0
        Form1.AutoRedraw = False
        Form1.Refresh
        
        If cir(0) >= 50 + 50 * stong Then
            If cir(1) >= 50 Then
                If cir(2) >= 150 Then '4產生震波
                    Timer14(7).Interval = 100
                    s(19) = s(19) + 1
                    Shape16(s(19)).Visible = True
                    Timer35(s(19)).Enabled = True
                    If s(19) = 4 Then Call sg(Index)
                Else '3改變計時器速度 & 向下丟
                    cir(2) = cir(2) + 20
                    Timer14(7).Interval = 10
                End If
            Else '2向上移
                cir(1) = cir(1) + 5
            End If
        Else '1變大
            cir(0) = cir(0) + 5
        End If
        If cir(2) < 150 And stong = 0 Or stong = 1 Then Form1.Circle (Image1.Left + Image1.Width \ 2, ((Image1.Top - Image1.Height * 1.5) - cir(1)) + cir(2)), cir(0)
End Select
DoEvents
End Sub
Public Sub mpss(a As Integer) '損耗MP演算法☆
Shape3(0).Left = Image1.Left - Shape3(0).Width \ 2 + Image1.Width \ 2 '火柱基底
Shape3(0).Top = Image1.Top + Image1.Height \ 2
If mp.Width >= 40 Then
    Timer25.Enabled = False 'mp增加停止
    If kikyou = 3 Or a = 6 Then Timer1.Enabled = False: Timer2.Enabled = False: Timer3.Enabled = False: Timer4.Enabled = False: Timer6.Enabled = False: zxcv(2) = 1
    If a = kikyou + 4 Then
        Timer21(1).Enabled = False '停止A delay閃爍(絕招二)
        delay(1).Visible = False: delay(1).Height = 1: delay(1).Top = 81
    Else
        Timer21(0).Enabled = False '停止A delay閃爍
        delay(0).Visible = False: delay(0).Height = 1: delay(0).Top = 81
    End If
    mp.Width = mp.Width - 40
    Label2.Caption = mp.Width * akiz & "/" & amax
    Image12.Visible = True: Image12.ZOrder 1
    Select Case a
        Case 0
            Shape3(0).BorderColor = RGB(255, 255, 0) '指定範圍的顏色值
        
            Shape11(0).Left = Image1.Left + Image1.Width \ 2 - Shape11(0).Width \ 2
            Shape11(0).Top = Image1.Top - 190
            
            'Shape11(0).Left = Image1.Left - 175
            'Shape11(0).Top = Image1.Top - 100
            Line3(0).X1 = Shape11(0).Left + Shape11(0).Width \ 4  '中間光柱的左上角位置
            Line3(0).Y1 = Shape11(0).Top '光柱的左上角
            Line3(0).X2 = Shape11(0).Left + Shape11(0).Width \ 2  '光柱的右下角
            Line3(0).Y2 = Shape11(0).Top + Shape11(0).Height \ 2  '光柱的右下角
        Case 1 '火柱
            '------奇摩知識
            Set ds = dxa.DirectSoundCreate("")
            ds.SetCooperativeLevel Me.hWnd, DSSCL_NORMAL
            bu.lFlags = DSBCAPS_STATIC
            Set dsb = ds.CreateSoundBufferFromFile(("火柱.wav"), bu, wf)
            dsb.Play DSBPLAY_DEFAULT '''''''''''''''''''''''預設播放(單次).
            '------奇摩知識
        Case 2 '暴雷
            Shape9.Left = Image1.Left - 80 '上面的黑洞
            Shape9.Top = Image1.Top - 180 '上面的黑洞
            Shape9.Visible = True
        
            Set dscc = dxa.DirectSoundCreate("")
            dscc.SetCooperativeLevel Me.hWnd, DSSCL_NORMAL
            bu.lFlags = DSBCAPS_STATIC
            Set dsbcc = dscc.CreateSoundBufferFromFile(("暴雷.wav"), bu, wf)
            dsbcc.Play DSBPLAY_DEFAULT '''''''''''''''''''''''預設播放(單次).
        Case 3
            ReDim stat(20)
            dzxc = 0
            If zxcv(3) = 1 Then Image13(0).Picture = Image14(1).Picture: ell(0) = -1: Shape12.Left = Image1.Left + Image1.Width - Shape12.Width Else Image13(0).Picture = Image14(0).Picture: ell(0) = 1: Shape12.Left = Image1.Left '0為右1為左
            Shape12.Top = Image1.Top + Image1.Height \ 2 - Shape12.Height \ 2
            Shape12.Visible = True
        Case 4 '絕招二
            Shape10.Left = Image1.Left - Shape10.Width \ 2 + Image1.Width \ 2
            Shape10.Top = Image1.Top + Image1.Height \ 2
            Shape10.Visible = True
            
            Line4(5).X1 = Shape10.Left + Shape10.Width \ 2 '設定初始化頂點
            Line4(5).Y1 = Shape10.Top
            Line4(0).X2 = Shape10.Left + Shape10.Width \ 2 '設定初始化頂點
            Line4(0).Y2 = Shape10.Top + Shape10.Height
            
            Line4(0).X1 = Line4(0).X2
            Line4(0).Y1 = Line4(0).Y2
            
            Line4(5).X2 = Line4(5).X1
            Line4(5).Y2 = Line4(5).Y1
            
            For f = 0 To 5
                Line4(f).Visible = True
            Next
        Case 5 '崩裂
            For f = 1 To 8
                Load Line5(f)
                Line5(f).X1 = Image1.Left + Image1.Width \ 2
                Line5(f).Y1 = Image1.Top + Image1.Height \ 2
                Line5(f).X2 = Line5(f).X1
                Line5(f).Y2 = Line5(f).Y1
                
                Line5(f).Visible = True
            Next
            
            
        '-----奇摩知識
            Set dscd = dxa.DirectSoundCreate("")
            dscd.SetCooperativeLevel Me.hWnd, DSSCL_NORMAL
            bu.lFlags = DSBCAPS_STATIC
            Set dsbd = dscd.CreateSoundBufferFromFile(("崩裂.wav"), bu, wf)
            dsbd.Play DSBPLAY_DEFAULT '''''''''''''''''''''''預設播放(單次).
        '-----奇摩知識
        Case 6 '殘月漣漪
            wmy = 2
            For f = 1 To 4 + stong * wmy
                Load Shape3(f)
                Shape3(f).Width = 60 '水波原始大小
                Shape3(f).Height = 30 '水波原始大小
                
                Load Timer18(f)
                
                If stong = 0 Then
                    Shape3(f).Left = Image1.Left - Shape3(f).Width \ 2 + Image1.Width \ 2 '水波原始位置
                    Shape3(f).Top = Image1.Top + Image1.Height - 15 '水波原始位置
                    Shape3(f).BorderColor = &HFF0000 '水波顏色
                Else
                    Randomize
                    Shape3(f).Left = Int(Rnd * (Form1.ScaleWidth - Shape3(f).Width * 2)) + Shape3(f).Width '水波原始位置
                    Shape3(f).Top = Int(Rnd * (Form1.ScaleHeight - Shape3(f).Height * 2)) + Shape3(f).Height '水波原始位置
                    Shape3(f).BorderColor = &HFF0000 '水波顏色
                End If
            Next
        Case 7 '超念動
            For f = 1 To 4
                Load Shape16(f)
                Load Timer35(f)
                Shape16(f).Left = Image1.Left + Image1.Width \ 2 - Shape16(f).Width \ 2
                Shape16(f).Top = Image1.Top + Image1.Height
            Next
    End Select
    For f = 0 To 1
        delay(f).FillColor = &H80FF80
    Next
    Timer14(a).Enabled = True
End If
End Sub
Private Sub kiker() '桔梗的意志共用

Image18.Left = Image18.Left + 30 '箭矢射出
Image19.Left = Image18.Left + Image18.Width - Image19.Width \ 2 '破魔之氣
Image19.Visible = True

If Image17.Left + Image1.Width < 0 Then '當桔梗超出邊界則
    Image17.Visible = False
End If

End Sub
Public Sub sg(a As Integer) '絕招共用

Timer14(a).Enabled = False

Select Case a '絕招delay
    Case Is <= 3 '絕招一
        If a = 1 Then Erase vnn
        delay(0).Visible = True: Timer20(0).Enabled = True
Case Else '絕招二
    delay(1).Visible = True: Timer20(1).Enabled = True
End Select
If a = 4 Then '桔梗的意志
    Shape10.Visible = False
    Image18.Visible = False
    Image19.Visible = False
    Image19.Left = -11111
    For f = 0 To 5
        Line4(f).Visible = False
        Line4(f).X1 = -11
        Line4(f).X2 = -11
        Line4(f).Y1 = -11
        Line4(f).Y2 = -11
    Next
    For f = 0 To 15
        ven(f) = 0
    Next
Else
    If a = 5 Then '崩裂
        For f = 1 To 10
            Unload Line5(f)
        Next
    End If
End If

stong = 0

zxcv(2) = 0 '絕招放完
End Sub
Private Sub Timer20_Timer(Index As Integer) '絕招delay
delay(Index).Height = delay(Index).Height + 5
delay(Index).Top = delay(Index).Top - 5
If delay(Index).Height >= 57 Then delay(Index).Height = 57: delay(Index).Top = 24:: Timer20(Index).Enabled = False: Timer21(Index).Enabled = True
DoEvents
End Sub
Private Sub Timer21_Timer(Index As Integer) 'A delay閃爍
slow(Index, s(6)) = slow(Index, s(6)) + 1
If slow(Index, s(6)) Mod 2 = 0 Then
    delay(Index).Visible = True
Else
    delay(Index).Visible = False
End If
DoEvents
End Sub
Private Sub Timer10_Timer() '普攻★
If (Timer3.Enabled = True And Timer8.Enabled = True) Or (Timer4.Enabled = True And Timer8.Enabled = True) Then '跳躍攻擊
    If Timer3.Enabled = True Then '左
        Image1.Left = Image1.Left - 20
        Shape1.Left = Shape1.Left - 20
    Else '右
        Image1.Left = Image1.Left + 20
        Shape1.Left = Shape1.Left + 20
    End If
    s(2) = s(2) + 1
    Call sppr(5) '判斷是否打中
    
    If s(2) >= 5 Then s(2) = 0: Timer10.Enabled = False
Else
    If zxcv(3) = 0 Then '右
        Image1.Left = Image1.Left + 10
        Shape1.Left = Shape1.Left + 10
        s(2) = s(2) + 1
        If s(2) >= 5 Then
            Call sppr(5) '判斷是否打中
            
            Image1.Left = Image1.Left - 50
            Shape1.Left = Shape1.Left - 50
            s(2) = 0
            Timer10.Enabled = False
        End If
    Else '左
        Image1.Left = Image1.Left - 10
        Shape1.Left = Shape1.Left - 10
        s(2) = s(2) + 1
        If s(2) >= 5 Then
            Call sppr(5) '判斷是否打中
            
            Image1.Left = Image1.Left + 50
            Shape1.Left = Shape1.Left + 50
            s(2) = 0
            Timer10.Enabled = False
        End If
    End If
End If
DoEvents
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer) '按鍵按下●
If zxcv(5) = 1 Then Exit Sub
If KeyCode = keyup Then dzxc = -1
If KeyCode = keydown Then dzxc = 1 '流星上下判斷
If zxcv(2) = 0 Then  'A絕招
    If (zxcv(0) = 1 And KeyCode = keyup) Or (zxcv(0) = 1 And KeyCode = keydown) Then Exit Sub
    Select Case KeyCode
        Case keyleft '左37
            If ax(0) = 1 Then run(0) = 2
            
            Timer3.Enabled = True '移動
            zxcv(3) = 1 '動畫左右判斷
        Case keyup '上38
            up = 1
            If zxcv(0) = 0 Then
                Timer1.Enabled = True '移動
            End If
        Case keyright '右39
            If ax(1) = 1 Then run(1) = 2
            
            Timer4.Enabled = True '移動
            zxcv(3) = 0 '動畫左右判斷
        Case keydown '下40
            If zxcv(0) = 0 Then
                Timer2.Enabled = True '移動
            End If
        Case keya 'A65
            If Val(Label10.Caption) > 0 Then Timer30(0).Enabled = True '如果有四魂之玉則絕招集氣
            If up = 1 Then
                If Timer21(1).Enabled = True Then Call mpss(kikyou + 4) '讀取損耗MP演算法(絕招二)
            Else
                If Timer21(0).Enabled = True Then Call mpss(kikyou)  '讀取損耗MP演算法
            End If
        Case keys 'S 普攻83
            zxcv(1) = zxcv(1) + 1
            If zxcv(1) = 1 Then Timer10.Enabled = True
        Case keyd 'D 跳躍68
            zxcv(0) = zxcv(0) + 1
            
            Timer1.Enabled = False '上 停止
            Timer2.Enabled = False '下 停止
            
            If zxcv(0) = 1 Then
                ttop = Image1.Top
                Timer8.Enabled = True
            End If
        Case 113 'F2
            Call ex(0)
    End Select
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer) '按鍵彈起●
Select Case KeyCode
    Case keyleft '左
        ax(0) = 1
        Timer9(0).Enabled = True
        ax(1) = 0 '取消右跑
        Timer9(1).Enabled = False '取消右跑
        
        run(0) = 1 '走路狀態
        
        Timer3.Enabled = False
    Case keyup '上
        up = 0
        Timer1.Enabled = False
        dzxc = 0
    Case keyright '右
        ax(1) = 1
        Timer9(1).Enabled = True
        ax(0) = 0 '取消左跑
        Timer9(0).Enabled = False '取消左跑
        
        run(1) = 1 '走路狀態
        
        Timer6.Enabled = False '地圖移動關閉
        Timer4.Enabled = False
    Case keydown '下
        Timer2.Enabled = False
        dzxc = 0
    Case keys 's普攻
        zxcv(1) = 0
    Case keya 'A
        Timer30(0).Enabled = False '絕招停止集氣
        Timer25.Enabled = True 'MP增加
End Select
End Sub
Private Sub Timer9_Timer(Index As Integer) '跑步判斷★
If Index = 0 Then
    ax(0) = 0
Else
    ax(1) = 0
End If
Timer9(Index).Enabled = False
End Sub
Private Sub Timer8_Timer() '跳躍★

s(0) = s(0) + 1
Image1.Top = Image1.Top - (12 - s(0)) * x
Shape1.Width = Shape1.Width - (3 - s(0) \ 4) * x
Shape1.Left = Shape1.Left + (1.5 - s(0) \ 6) * x
If s(0) = 12 Then
    x = -1
    s(0) = 0
    If Image1.Top = ttop Then
        Timer8.Enabled = False
        x = 1
        zxcv(0) = 0
    End If
End If
DoEvents
End Sub
Private Sub Timer1_Timer() '上★
Image1.Top = Image1.Top - 10
Shape1.Top = Image1.Top + sh
If Image1.Top < 0 + Line1.Y1 Then Image1.Top = 0 + Line1.Y1: Shape1.Top = Image1.Top + sh
End Sub
Private Sub Timer2_Timer() '下★
Image1.Top = Image1.Top + 10
Shape1.Top = Image1.Top + sh
If Image1.Top + Image1.Height > Form1.ScaleHeight - 10 Then Image1.Top = Form1.ScaleHeight - Image1.Height - 10: Shape1.Top = Image1.Top + sh
End Sub
Private Sub Timer3_Timer() '左★
Image1.Left = Image1.Left - 10 * run(0)
Shape1.Left = Image1.Left + Image1.Width \ 2
Timer5(kikyou).Enabled = True '啟動動畫
If Image1.Left < 0 + 10 Then Image1.Left = 0 + 10
End Sub
Private Sub Timer4_Timer() '右★
If Image1.Left >= Form1.ScaleWidth \ 2 Then '地圖是否移動
    If ok = 0 Then
        Timer6.Enabled = True
        Exit Sub
    Else
        If Image1.Left + Image1.Width > Form1.ScaleWidth - 10 Then Image1.Left = Form1.ScaleWidth - Image1.Width - 10
    End If
End If
If Timer6.Enabled = True Then
    Image1.Left = Image1.Left + 10 * run(1) - 15 * run(1)
Else
    Image1.Left = Image1.Left + 10 * run(1)
End If
Shape1.Left = Image1.Left + Image1.Width \ 2 - 20
Timer5(kikyou).Enabled = True '啟動動畫
End Sub
Private Sub Timer5_Timer(Index As Integer) '角色動畫★
k(kikyou) = k(kikyou) + 1
Select Case kikyou
    Case 0
        If k(kikyou) = 6 Then k(kikyou) = 0
        If zxcv(3) = 0 Then
            Image1.Picture = Image2(k(kikyou)).Picture
        Else
            Image1.Picture = Image7(k(kikyou)).Picture
        End If
    Case 1
        If k(kikyou) = 5 Then k(kikyou) = 0
        If zxcv(3) = 0 Then
            Image1.Picture = Image3(k(kikyou)).Picture
        Else
            Image1.Picture = Image8(k(kikyou)).Picture
        End If
    Case 2
        If k(kikyou) = 4 Then k(kikyou) = 0
        If zxcv(3) = 0 Then
            Image1.Picture = Image4(k(kikyou)).Picture
        Else
            Image1.Picture = Image9(k(kikyou)).Picture
        End If
    Case 3
        If k(kikyou) = 6 Then k(kikyou) = 0
        If zxcv(3) = 0 Then
            Image1.Picture = Image5(k(kikyou)).Picture
        Else
            Image1.Picture = Image10(k(kikyou)).Picture
        End If
End Select
End Sub
Private Sub Timer32_Timer(Index As Integer) '四魂之玉轉動動畫

If sk(Index) = 7 Then
    sk(Index) = 0
Else
    sk(Index) = sk(Index) + 1
End If

Image20(Index).Picture = Image21(sk(Index)).Picture

If Timer6.Enabled = True Then Image20(Index).Left = Image20(Index).Left - 20 * run(1): Shape15(Index).Left = Shape15(Index).Left - 20 * run(1) '地圖移動
If coll(Image20(Index), Image1) Then
    Label10.Visible = True: Image23.Visible = True: Label10.Caption = Val(Label10.Caption) + 1 & " X": Image20(Index).Left = -1111: Shape15(Index).Left = -1111
    If hp.Width >= 180 Then '補血
        hp.Width = 200
    Else
        hp.Width = hp.Width + 20
    End If
    Label1.Caption = hp.Width * akiz & "/" & amax
    
    If mp.Width >= 180 Then '補魔
        mp.Width = 200
    Else
        mp.Width = mp.Width + 20
    End If
    Label2.Caption = mp.Width * akiz & "/" & amax
End If
End Sub
Private Sub Timer15_Timer() 'GO顯示★
s(4) = s(4) + 1
If s(4) Mod 2 = 0 Then
    Label3.Visible = False
Else
    Label3.Visible = True
End If
If ok = 1 Then Timer15.Enabled = False: Label3.Visible = False: Call desd(0)
DoEvents
End Sub
Private Sub Timer25_Timer() 'MP增加演算法
mp.Width = mp.Width + 5
If mp.Width >= 200 Then mp.Width = 200
Label2.Caption = mp.Width * akiz & "/" & amax
End Sub
Public Sub kizzs() '角色損血演算法☆
If s(10) = 0 And hp.Visible = True Then
    'Call voice
    Randomize
    ym = Int(Rnd * 8) + 5
    If hp.Width - ym <= 0 Then
        hp.Visible = False
        Label1.Caption = "0/" & amax
        zxcv(2) = 1
        Call ex(1): Exit Sub
    Else
        hp.Width = hp.Width - ym
        Label1.Caption = hp.Width * akiz & "/" & amax
    End If
    DoEvents
End If
End Sub
Private Sub voice() '碰撞聲音
'-----奇摩知識
    Set dsx = dxa.DirectSoundCreate("")
    dsx.SetCooperativeLevel Me.hWnd, DSSCL_NORMAL
    bu.lFlags = DSBCAPS_STATIC
    Set dsbx = dsx.CreateSoundBufferFromFile(("碰撞.wav"), bu, wf)
    dsbx.Play DSBPLAY_DEFAULT '''''''''''''''''''''''預設播放(單次).
'-----奇摩知識
End Sub
Public Sub sppr(a As Integer) '判斷是否打中☆
For f = 1 To ma
    If ppm(1, f) <> 1 Then
        If coll2(Shape1, Shape2(f)) Or coll2(Shape2(f), Shape1) Then
            If coll(Image1, Image6(f)) Or coll(Image6(f), Image1) Then pps = Image6(f).Index: Call power(a): Exit Sub
        End If
    End If
    DoEvents
Next
End Sub
Public Sub power(a As Integer) '敵人損血演算法☆
'-------------------------------------'0～３範圍技,4 普攻,5 跳躍攻擊
ppm(5, pps) = 1 '停止時的特效
Timer11(pps).Enabled = False

Randomize
sdj = Int(Rnd * 3) + 1
If a = 2 And sdj = 2 Then Timer22(pps).Interval = 5000 Else Timer22(pps).Interval = 200 '暴雷的特殊效果
    
If a <= 3 Then a = 10 Else a = 0
If ppm(0, pps) = 0 Then
    xhp.Width = 200
    xhp.Visible = True
    ppm(0, pps) = xhp.Width
Else
    xhp.Visible = True
    xhp.Width = ppm(0, pps)
End If
'Call voice
Randomize
Y = Int(Rnd * 10 + a) + 5
If xhp.Width - Y <= 0 Then
    ppm(0, pps) = 0
    xhp.Visible = False
    Label5.Caption = ppm(0, pps) & "/" & amax
    Load Timer7(pps)
    ppm(1, pps) = 1
    Image6(pps).Top = Image6(pps).Top + 15 '鳥掉下來
    Timer7(pps).Enabled = True '啟動鳥消失動畫
    n = n + 1
Else
    xhp.Width = xhp.Width - Y
    ppm(0, pps) = xhp.Width
    Label5.Caption = ppm(0, pps) * akiz & "/" & amax
    
    If te = 1 Then Call news '是否跳傷害
    
End If
'擊退
If zxcv(3) = 0 Then
    Image6(pps).Left = Image6(pps).Left + 30
    Shape2(pps).Left = Image6(pps).Left + Image6(pps).Width \ 2
Else
    Image6(pps).Left = Image6(pps).Left - 30
    Shape2(pps).Left = (Image6(pps).Left + Image6(pps).Width \ 2) - 20
End If

'分數
Label9.Caption = Val(Label9.Caption) + 50

Timer22(pps).Enabled = True
DoEvents
End Sub
Public Sub news() '顯示傷害演算法☆
'產生物件
s(5) = s(5) + 1
If s(5) = 700 Then s(5) = 0
Load Label4(s(5))
'分配位置
Label4(s(5)).Top = Image6(pps).Top - Label4(s(5)).Height
Label4(s(5)).Left = Image6(pps).Left + Image6(pps).Width \ 2
'設定屬性
Label4(s(5)).Caption = Y * akiz
Label4(s(5)).Visible = True
Load Timer16(s(5)) '產生一個個別移動    計時器
Load Timer17(s(5)) '產生一個個別移動移除計時器
Timer16(s(5)).Enabled = True '開啟移動計時器
Timer17(s(5)).Enabled = True '開啟移除計時器
DoEvents
End Sub
Private Sub Timer16_Timer(Index As Integer) '顯示傷害的移動★
Label4(Index).Top = Label4(Index).Top - 10
DoEvents
End Sub
Private Sub Timer17_Timer(Index As Integer) '移除傷害顯示★
Unload Timer16(Index)
Unload Label4(Index)
Unload Timer17(Index)
DoEvents
End Sub
Private Sub Timer6_Timer() '主背景移動★
dx = dx + 20 * run(1)
If dx >= 1024 Then dx = 0
If dx <= 100 Then '=====大於564後就超圖了=====
    Form1.PaintPicture Form1.Picture, 0, 0, 924, 708, dx, 0, 924, 708, vbSrcCopy
Else '上面form1截圖 複製到          0,0,表寬,表高,dx,0,表寬,表高
    Form1.PaintPicture Form1.Picture, 0, 0, 1024 - dx, 708, dx, 0, 1024 - dx, 708, vbSrcCopy
    '上面form1截圖 複製到           0,0,圖寬-dx,表高,dx,0,圖寬-dx,表高
    Form1.PaintPicture Form1.Picture, 1024 - dx, 0, dx - 100, 708, 0, 0, dx - 100, 708, vbSrcCopy
    '上面form1截圖 複製到           圖寬-dx,0,表寬+(dx-圖寬),表高,0,0,表寬+(dx+圖寬),表高
    
    s(1) = s(1) + 1
    If s(1) = 10 Then s(1) = 0: ok = 1: Timer6.Enabled = False '====圖寬=====
End If
DoEvents
End Sub
Private Sub sico(a As Integer) '四魂之玉掉落演算法
Randomize
f = Int(Rnd * 3) + 1
If f = 2 Then
    If s(13) = 10 Then
        s(13) = 0
    Else
        s(13) = s(13) + 1
    End If
    Load Image20(s(13)) '四魂之玉本體
    Load Shape15(s(13)) '四魂之玉影子
    Load Timer32(s(13)) '四魂之玉轉動動畫
    Load Timer33(s(13)) '四魂之玉消失
    
    Image20(s(13)).Top = Image6(a).Top + Image6(a).Height \ 4
    Image20(s(13)).Left = Image6(a).Left + Image6(a).Width \ 4
    Shape15(s(13)).Left = Image20(s(13)).Left + Image20(s(13)).Width \ 4
    Shape15(s(13)).Top = Image20(s(13)).Top + 48
        
    Image20(s(13)).Visible = True
    Shape15(s(13)).Visible = True
    Timer32(s(13)).Enabled = True
    Timer33(s(13)).Enabled = True
    
    Exit Sub
End If
End Sub
Private Sub Timer33_Timer(Index As Integer) '四魂之玉消失
sic(Index) = sic(Index) + 1
If sic(Index) >= 10 Then
    If sic(Index) Mod 2 = 0 Then
        Image20(Index).Visible = True
    Else
        Image20(Index).Visible = False
    End If
    
    If sic(Index) = 20 Then
        Image20(Index).Visible = False
        sic(Index) = 0
        sk(Index) = 0
        Unload Image20(Index)
        Unload Shape15(Index)
        Unload Timer32(Index)
        Unload Timer33(Index)
    End If
End If
End Sub
