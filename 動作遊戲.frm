VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "動作遊戲 v 1.0.2"
   ClientHeight    =   12720
   ClientLeft      =   1890
   ClientTop       =   1095
   ClientWidth     =   14475
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "動作遊戲.frx":0000
   ScaleHeight     =   848
   ScaleMode       =   3  '像素
   ScaleWidth      =   965
   Begin VB.Timer Timer17 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   800
      Left            =   11400
      Top             =   4560
   End
   Begin VB.Timer Timer16 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   70
      Left            =   11400
      Top             =   4080
   End
   Begin VB.Timer Timer15 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   11400
      Top             =   2400
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   20
      Left            =   1440
      Top             =   2040
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   20
      Left            =   1080
      Top             =   2040
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   100
      Left            =   720
      Top             =   2040
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   20
      Left            =   360
      Top             =   2040
   End
   Begin VB.Timer Timer13 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   11400
      Top             =   1560
   End
   Begin VB.Timer Timer12 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   11400
      Top             =   1080
   End
   Begin VB.Timer Timer11 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   12960
      Top             =   10800
   End
   Begin VB.Timer Timer10 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   3600
      Top             =   1560
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   100
      Left            =   2880
      Top             =   1560
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   100
      Left            =   2400
      Top             =   1560
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   1560
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   100
      Left            =   9120
      Top             =   11400
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   100
      Left            =   7080
      Top             =   11160
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   100
      Left            =   5160
      Top             =   11160
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   6240
      Top             =   1560
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   100
      Left            =   3480
      Top             =   11040
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   1680
      Top             =   1560
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   1320
      Top             =   1560
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   960
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   480
      Top             =   1560
   End
   Begin VB.Label Label4 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Index           =   0
      Left            =   9720
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "209"
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
      Left            =   9360
      TabIndex        =   4
      Top             =   120
      Width           =   540
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "靈力值"
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
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "生命值"
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
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1080
   End
   Begin VB.Shape mp 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '實心
      Height          =   375
      Left            =   240
      Top             =   600
      Width           =   3135
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000005&
      FillColor       =   &H00C00000&
      FillStyle       =   0  '實心
      Height          =   375
      Left            =   240
      Top             =   600
      Width           =   3135
   End
   Begin VB.Shape hp 
      BorderColor     =   &H80000005&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '實心
      Height          =   375
      Left            =   240
      Top             =   120
      Width           =   3135
   End
   Begin VB.Shape xhp 
      BorderColor     =   &H80000005&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '實心
      Height          =   375
      Left            =   8040
      Top             =   120
      Width           =   3135
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   375
      Left            =   8040
      Top             =   120
      Width           =   3135
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   375
      Left            =   240
      Top             =   120
      Width           =   3135
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   40
      X2              =   728
      Y1              =   72
      Y2              =   72
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  '實心
      Height          =   135
      Index           =   0
      Left            =   9840
      Shape           =   2  '橢圓形
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   7
      Left            =   10200
      Picture         =   "動作遊戲.frx":240042
      Top             =   12240
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   6
      Left            =   10320
      Picture         =   "動作遊戲.frx":240360
      Top             =   12120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   5
      Left            =   10440
      Picture         =   "動作遊戲.frx":240644
      Top             =   12000
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   4
      Left            =   10560
      Picture         =   "動作遊戲.frx":240912
      Top             =   11880
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   3
      Left            =   10680
      Picture         =   "動作遊戲.frx":240BD3
      Top             =   11760
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   2
      Left            =   10800
      Picture         =   "動作遊戲.frx":240EA1
      Top             =   11640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   1
      Left            =   10920
      Picture         =   "動作遊戲.frx":24118A
      Top             =   11520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   0
      Left            =   11040
      Picture         =   "動作遊戲.frx":2414A8
      Top             =   11400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image6 
      Height          =   960
      Index           =   0
      Left            =   9360
      Picture         =   "動作遊戲.frx":2417A8
      Top             =   4320
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   1
      Left            =   12960
      Picture         =   "動作遊戲.frx":241AA5
      Top             =   11640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   2
      Left            =   12840
      Picture         =   "動作遊戲.frx":241DC2
      Top             =   11760
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   0
      Left            =   13080
      Picture         =   "動作遊戲.frx":2420A2
      Top             =   11520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   3
      Left            =   12720
      Picture         =   "動作遊戲.frx":24239F
      Top             =   11880
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   4
      Left            =   12600
      Picture         =   "動作遊戲.frx":242671
      Top             =   12000
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   6
      Left            =   12360
      Picture         =   "動作遊戲.frx":242933
      Top             =   12240
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   7
      Left            =   12240
      Picture         =   "動作遊戲.frx":242C0F
      Top             =   12360
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   5
      Left            =   12480
      Picture         =   "動作遊戲.frx":242F2C
      Top             =   12120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   5
      Left            =   8400
      Picture         =   "動作遊戲.frx":2431FE
      Top             =   10200
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   4
      Left            =   8520
      Picture         =   "動作遊戲.frx":243425
      Top             =   10080
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   3
      Left            =   8640
      Picture         =   "動作遊戲.frx":24365E
      Top             =   9960
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   2
      Left            =   8760
      Picture         =   "動作遊戲.frx":24388E
      Top             =   9840
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   1
      Left            =   9000
      Picture         =   "動作遊戲.frx":243AC7
      Top             =   9600
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image9 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   3
      Left            =   6720
      Picture         =   "動作遊戲.frx":243CEE
      Top             =   12000
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image9 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   2
      Left            =   6840
      Picture         =   "動作遊戲.frx":243ED3
      Top             =   11880
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image9 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   1
      Left            =   6960
      Picture         =   "動作遊戲.frx":24410E
      Top             =   11760
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   4
      Left            =   5040
      Picture         =   "動作遊戲.frx":2442F3
      Top             =   10200
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   3
      Left            =   5160
      Picture         =   "動作遊戲.frx":2444D6
      Top             =   10080
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   2
      Left            =   5280
      Picture         =   "動作遊戲.frx":2446D0
      Top             =   9960
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   1
      Left            =   5400
      Picture         =   "動作遊戲.frx":2448CA
      Top             =   9840
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   5
      Left            =   3120
      Picture         =   "動作遊戲.frx":244AD0
      Top             =   9840
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   4
      Left            =   3240
      Picture         =   "動作遊戲.frx":244DC6
      Top             =   9720
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   3
      Left            =   3360
      Picture         =   "動作遊戲.frx":2450BE
      Top             =   9600
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   2
      Left            =   3480
      Picture         =   "動作遊戲.frx":24539B
      Top             =   9480
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   1
      Left            =   3600
      Picture         =   "動作遊戲.frx":245693
      Top             =   9360
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   0
      Left            =   9120
      Picture         =   "動作遊戲.frx":245989
      Top             =   9480
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image9 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   0
      Left            =   7080
      Picture         =   "動作遊戲.frx":245BB7
      Top             =   11640
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   0
      Left            =   5520
      Picture         =   "動作遊戲.frx":245E74
      Top             =   9720
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   0
      Left            =   3720
      Picture         =   "動作遊戲.frx":2460A8
      Top             =   9240
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image1 
      Height          =   1200
      Left            =   1080
      Picture         =   "動作遊戲.frx":2463A0
      Top             =   6720
      Width           =   1200
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   3
      Left            =   6720
      Picture         =   "動作遊戲.frx":24657F
      Top             =   9960
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   2
      Left            =   6840
      Picture         =   "動作遊戲.frx":246834
      Top             =   9840
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   1
      Left            =   6960
      Picture         =   "動作遊戲.frx":246A13
      Top             =   9720
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   0
      Left            =   7080
      Picture         =   "動作遊戲.frx":246C4C
      Top             =   9600
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   0
      Left            =   8880
      Picture         =   "動作遊戲.frx":246E2B
      Top             =   11520
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   1
      Left            =   8760
      Picture         =   "動作遊戲.frx":247052
      Top             =   11640
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   2
      Left            =   8640
      Picture         =   "動作遊戲.frx":247280
      Top             =   11760
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   3
      Left            =   8520
      Picture         =   "動作遊戲.frx":2474AF
      Top             =   11880
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   4
      Left            =   8400
      Picture         =   "動作遊戲.frx":2476DF
      Top             =   12000
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   5
      Left            =   8280
      Picture         =   "動作遊戲.frx":24790E
      Top             =   12120
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   0
      Left            =   5040
      Picture         =   "動作遊戲.frx":247B3C
      Top             =   11760
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   1
      Left            =   4920
      Picture         =   "動作遊戲.frx":247D6E
      Top             =   11880
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   2
      Left            =   4800
      Picture         =   "動作遊戲.frx":247F72
      Top             =   12000
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   3
      Left            =   4680
      Picture         =   "動作遊戲.frx":248170
      Top             =   12120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   4
      Left            =   4560
      Picture         =   "動作遊戲.frx":24836A
      Top             =   12240
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   0
      Left            =   3360
      Picture         =   "動作遊戲.frx":248549
      Top             =   11400
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   1
      Left            =   3240
      Picture         =   "動作遊戲.frx":24883F
      Top             =   11520
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   2
      Left            =   3120
      Picture         =   "動作遊戲.frx":248B38
      Top             =   11640
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   3
      Left            =   3000
      Picture         =   "動作遊戲.frx":248E31
      Top             =   11760
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   4
      Left            =   2880
      Picture         =   "動作遊戲.frx":249110
      Top             =   11880
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   5
      Left            =   2760
      Picture         =   "動作遊戲.frx":249409
      Top             =   12000
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  '實心
      Height          =   135
      Left            =   1440
      Shape           =   2  '橢圓形
      Top             =   7800
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s(6) As Integer '0)跳躍計數 1)地圖計數 2)普攻回原位計數 3)火柱範圍 4)GO的閃爍 5)顯示傷害 6)火柱計數
Dim zxcv(3) As Integer '0)限制跳躍時不可上下移動 1)限制普攻不可按住 2)限制A絕招 3)判斷角色往左或往右
Dim x As Integer '跳躍方向
Dim xy() As Integer '鳥方向
Dim ttop '暫存top
Dim k(4) As Integer '動畫1～３)人物動畫 4)敵人鳥動畫
Dim dx
Dim sh As Integer '影子距離
Dim ax(1) As Integer '跑步判斷
Dim run(1) As Integer '跑步增值器
Dim ok As Integer 'ok為０則地圖移動為１則人物移動
Dim ma As Integer '數量
Dim pps As Integer '被擊中的那一顆
Dim ppm() As Integer '0)血 1)命
Dim n As Integer '擊破數
Dim y '傷害值
Private Sub ex() '移除程序
Timer13.Enabled = False
Erase s
Erase zxcv
Erase xy
Erase k
Erase ax
Erase run
dx = 0
ok = 0
Unload Me
Form2.Show
End Sub
Private Sub Form_Load() '表單載入
Form1.Left = 1845
Form1.Top = 660
Form1.Width = 12000 '(800*600)
Form1.Height = 9000
Label1.Left = hp.Left + hp.Width \ 2 - Label1.Width \ 2
Label2.Left = Label1.Left
Timer5(kikyou).Enabled = True
Select Case kikyou
    Case 0
        Image1.Picture = Image2(0).Picture
        sh = 80
    Case 1
        Image1.Picture = Image3(0).Picture
        sh = 70
    Case 2
        Image1.Picture = Image4(0).Picture
        sh = 75
    Case 3
        Image1.Picture = Image5(0).Picture
        sh = 65
End Select
Shape1.Top = Image1.Top + sh
Shape1.Left = Image1.Left + Image1.Width \ 2 - 20
x = 1
For i = 0 To 1
    run(i) = 1
Next

ma = 20

Timer12.Enabled = True '產生
End Sub
Private Sub Timer12_Timer() '產生
ReDim xy(ma)
For f = 1 To ma
    xy(f) = 1
Next

ReDim ppm(0, ma)

Randomize
For f = 1 To ma
    Load Image6(f)
    Image6(f).Top = Int(Rnd * (Form1.ScaleHeight - Image6(f).Height - Line1.Y1)) + Line1.Y1
    Image6(f).Left = Int(Rnd * (Form1.ScaleWidth * 2 - Form1.ScaleWidth)) + Form1.ScaleWidth
    Image6(f).Visible = True
    
    Load Shape2(f)
    Shape2(f).Top = Image6(f).Top + 75
    Shape2(f).Left = Image6(f).Left + Image6(f).Width \ 2
    Shape2(f).Visible = True
    Shape2(f).ZOrder 0
Next
Timer12.Enabled = False
Timer11.Enabled = True
Timer13.Enabled = True
End Sub
Private Sub Timer13_Timer() '敵人鳥移動
For f = 1 To ma
    If Timer13.Enabled = False Then Exit Sub
    If Image6(f).Visible = True Then
        If Timer6.Enabled = True Then
            Image6(f).Left = (Image6(f).Left - 10 * xy(f)) - 15 * run(1)
        Else
            Image6(f).Left = Image6(f).Left - 10 * xy(f)
        End If
        Shape2(f).Left = Image6(f).Left + Image6(f).Width \ 2
        
        If Image6(f).Left < 0 Then xy(f) = -1: Shape2(f).Left = Image6(f).Left + Image6(f).Width \ 2 - 20
        If Image6(f).Left + Image6(f).Width > Form1.ScaleWidth Then xy(f) = 1
        If coll(Image6(f), Image1) Or coll(Image1, Image6(f)) Then pps = Image6(f).Index: Call kizzs '損血演算法
    End If
Next
End Sub
Private Sub Timer14_Timer(Index As Integer) '絕招
Select Case Index
    Case 0
        zxcv(2) = 0
        Timer14(kikyou).Enabled = False
    Case 1 '火柱
        Form1.AutoRedraw = False
        s(3) = s(3) + 50
        r = Image1.Height + s(3)
        Form1.Circle (Image1.Left + Image1.Width \ 2, Image1.Top + Image1.Height \ 2), r, RGB(255, 0, 0)
        s(6) = s(6) + 1
        If s(6) Mod 8 = 0 Then Call power(1): s(6) = 0
        If s(3) >= 800 Then s(3) = 0: Timer14(kikyou).Enabled = False: zxcv(2) = 0: Form1.AutoRedraw = True: Form1.Refresh
    Case 2
        zxcv(2) = 0
        Timer14(kikyou).Enabled = False
    Case 3
        zxcv(2) = 0
        Timer14(kikyou).Enabled = False
End Select
End Sub
Private Sub Timer10_Timer() '普攻
If (Timer3.Enabled = True And Timer8.Enabled = True) Or (Timer4.Enabled = True And Timer8.Enabled = True) Then '跳躍攻擊
    If Timer3.Enabled = True Then '左
        Image1.Left = Image1.Left - 20
        Shape1.Left = Shape1.Left - 20
    Else '右
        Image1.Left = Image1.Left + 20
        Shape1.Left = Shape1.Left + 20
    End If
    s(2) = s(2) + 1
    For f = 1 To ma
        If coll(Image1, Image6(f)) Or coll(Image6(f), Image1) Then pps = Image6(f).Index: Call power(5)
    Next
    If s(2) >= 5 Then s(2) = 0: Timer10.Enabled = False
Else
    If Timer5(kikyou).Enabled = True Then '右
        Image1.Left = Image1.Left + 10
        Shape1.Left = Shape1.Left + 10
        s(2) = s(2) + 1
        If s(2) >= 5 Then
            For f = 1 To ma
                If coll(Image1, Image6(f)) Or coll(Image6(f), Image1) Then pps = Image6(f).Index: Call power(4)
            Next
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
            For f = 1 To ma
                If coll(Image1, Image6(f)) Or coll(Image6(f), Image1) Then pps = Image6(f).Index: Call power(4)
            Next
            Image1.Left = Image1.Left + 50
            Shape1.Left = Shape1.Left + 50
            s(2) = 0
            Timer10.Enabled = False
        End If
    End If
End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer) '按鍵
If zxcv(2) = 0 Then 'A絕招
    If (zxcv(0) = 1 And KeyCode = 38) Or (zxcv(0) = 1 And KeyCode = 40) Then Exit Sub
    Select Case KeyCode
        Case 37 '左
            If ax(0) = 1 Then run(0) = 2
            
            Timer3.Enabled = True '移動
            zxcv(3) = 1 '動畫左右判斷
        Case 38 '上
            If zxcv(0) = 0 Then
                Timer1.Enabled = True '移動
            End If
        Case 39 '右
            If ax(1) = 1 Then run(1) = 2
            
            Timer4.Enabled = True '移動
            zxcv(3) = 0 '動畫左右判斷
        Case 40 '下
            If zxcv(0) = 0 Then
                Timer2.Enabled = True '移動
            End If
        Case 65 'A
            zxcv(2) = 1
            Timer1.Enabled = False
            Timer2.Enabled = False
            Timer3.Enabled = False '停止移動防止出錯
            Timer4.Enabled = False
            Timer6.Enabled = False
            Select Case kikyou
                Case 0
                Case 1
                    sndPlaySound "火柱.wav", 1
                Case 2
                Case 3
            End Select
            Timer14(kikyou).Enabled = True
        Case 83 'S 普攻
            zxcv(1) = zxcv(1) + 1
            If zxcv(1) = 1 Then Timer10.Enabled = True
        Case 68 'D 跳躍
            zxcv(0) = zxcv(0) + 1
            
            Timer1.Enabled = False '上 停止
            Timer2.Enabled = False '下 停止
            
            If zxcv(0) = 1 Then
                ttop = Image1.Top
                Timer8.Enabled = True
            End If
        Case 113 'F2
            Call ex
    End Select
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer) '按鍵彈起
Select Case KeyCode
    Case 37 '左
        ax(0) = 1
        Timer9(0).Enabled = True
        ax(1) = 0 '取消右跑
        Timer9(1).Enabled = False '取消右跑
        
        run(0) = 1 '走路狀態
        
        Timer3.Enabled = False
    Case 38 '上
        Timer1.Enabled = False
    Case 39 '右
        ax(1) = 1
        Timer9(1).Enabled = True
        ax(0) = 0 '取消左跑
        Timer9(0).Enabled = False '取消左跑
        
        run(1) = 1 '走路狀態
            
        Timer6.Enabled = False '地圖移動關閉
        Timer4.Enabled = False
    Case 40 '下
        Timer2.Enabled = False
    Case 83 's普攻
        zxcv(1) = 0
End Select
End Sub
Private Sub Timer11_Timer() '敵人鳥動畫
k(4) = k(4) + 1
If k(4) = 8 Then k(4) = 0
For f = 1 To ma
    If xy(f) = 1 Then
        Image6(f).Picture = Image15(k(4)).Picture
    Else
        Image6(f).Picture = Image11(k(4)).Picture
    End If
Next
End Sub
Private Sub Timer9_Timer(Index As Integer) '跑步判斷
If Index = 0 Then
    ax(0) = 0
Else
    ax(1) = 0
End If
Timer9(Index).Enabled = False
End Sub
Private Sub Timer8_Timer() '跳躍

s(0) = s(0) + 1
Image1.Top = Image1.Top - (12 - s(0)) * x
Shape1.Width = Shape1.Width - (3 - s(0) \ 4) * x
Shape1.Left = Shape1.Left + (1.5 - s(0) \ 6) * x
If s(0) = 12 Then
    s(0) = 0
    x = -1
    If Image1.Top = ttop Then
        Timer8.Enabled = False
        x = 1
        zxcv(0) = 0
    End If
End If

End Sub
Private Sub Timer1_Timer() '上
Image1.Top = Image1.Top - 10
Shape1.Top = Image1.Top + sh
If Image1.Top < 0 + 30 Then Image1.Top = 0 + 30
End Sub
Private Sub Timer2_Timer() '下
Image1.Top = Image1.Top + 10
Shape1.Top = Image1.Top + sh
If Image1.Top + Image1.Height > Form1.ScaleHeight - 10 Then Image1.Top = Form1.ScaleHeight - Image1.Height - 10
End Sub
Private Sub Timer3_Timer() '左
Image1.Left = Image1.Left - 10 * run(0)
Shape1.Left = Image1.Left + Image1.Width \ 2
Timer5(kikyou).Enabled = True '啟動動畫
If Image1.Left < 0 + 10 Then Image1.Left = 0 + 10
End Sub
Private Sub Timer4_Timer() '右
If Image1.Left >= Form1.ScaleWidth \ 2 Then '地圖是否移動
    If ok = 0 Then
        Timer6.Enabled = True
        Exit Sub
    Else
        If n = ma Then
            Timer15.Enabled = True '顯示GO下一關●
        Else
            If Image1.Left + Image1.Width > Form1.ScaleWidth - 10 Then Image1.Left = Form1.ScaleWidth - Image1.Width - 10
        End If
    End If
End If
If Timer6.Enabled = True Then
    Image1.Left = (Image1.Left + 10 * run(1)) - 15 * run(1)
Else
    Image1.Left = Image1.Left + 10 * run(1)
End If
Shape1.Left = Image1.Left + Image1.Width \ 2 - 20
Timer5(kikyou).Enabled = True '啟動動畫
End Sub
Private Sub Timer5_Timer(Index As Integer) '角色動畫
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
Private Sub Timer15_Timer() 'GO顯示
s(4) = s(4) + 1
If s(4) Mod 2 = 0 Then
    Label3.Visible = False
Else
    Label3.Visible = True
End If
If Image1.Left + Image1.Width > Form1.ScaleWidth Then Timer15.Enabled = False: MsgBox "第一關全破", 64, "訊息": Call ex
End Sub
Public Sub kizzs() '角色損血演算法●
If Timer14(kikyou).Enabled = False Then
    Randomize
    ym = Int(Rnd * 3) + 1
    If hp.Width - ym <= 0 Then
        hp.Visible = False
        MsgBox "你掛了～請按確定重新遊戲", 64, "訊息"
        Call ex
    Else
        'hp.Width = hp.Width - ym
    End If
End If
End Sub
Public Sub power(a) '敵人損血演算法● 0～３範圍技,4 普攻,5 跳躍攻擊
If ppm(0, pps) = 0 Then
    xhp.Width = 209
    xhp.Visible = True
    ppm(0, pps) = xhp.Width
Else
    xhp.Visible = True
    xhp.Width = ppm(0, pps)
End If
Randomize
y = Int(Rnd * 10) + 1
If xhp.Width - y <= 0 Then
    ppm(0, pps) = 0
    xhp.Visible = False
    Image6(pps).Visible = False
    Shape2(pps).Visible = False
    Image6(pps).Left = 0 - 11111
    n = n + 1
Else
    xhp.Width = xhp.Width - y
    ppm(0, pps) = xhp.Width
    Label5.Caption = ppm(0, pps)
    Call news
End If
'擊退
If Timer5(kikyou).Enabled = True Then
    Image6(pps).Left = Image6(pps).Left + 30
Else
    Image6(pps).Left = Image6(pps).Left - 30
End If
End Sub
Public Sub news() '顯示傷害演算法
'產生物件
s(5) = s(5) + 1
If s(5) = 1024 Then s(5) = 0
Load Label4(s(5))
'分配位置
Label4(s(5)).Top = Image6(pps).Top - Label4(s(5)).Height
Label4(s(5)).Left = Image6(pps).Left + Image6(pps).Width \ 2
'設定屬性
Label4(s(5)).Caption = y
Label4(s(5)).Visible = True
Load Timer16(s(5)) '產生一個個別移動    計時器
Load Timer17(s(5)) '產生一個個別移動移除計時器
Timer16(s(5)).Enabled = True '開啟移動計時器
Timer17(s(5)).Enabled = True '開啟移除計時器
End Sub
Private Sub Timer16_Timer(Index As Integer) '顯示傷害的移動
Label4(Index).Top = Label4(Index).Top - 15
End Sub
Private Sub Timer17_Timer(Index As Integer) '移除傷害顯示
Unload Timer16(Index)
Unload Label4(Index)
Unload Timer17(Index)
End Sub
Private Sub Timer6_Timer() '主背景移動
dx = dx + 25 * run(1)
If dx >= 1024 Then
    dx = 0
    
    s(1) = s(1) + 1
    If s(1) = 2 Then ok = 1: Timer6.Enabled = False '====圖寬=====
End If

If dx <= 100 Then '=====大於564後就超圖了=====
    Form1.PaintPicture Form1.Picture, 0, 0, 924, 708, dx, 0, 924, 708, vbSrcCopy
Else '上面form1截圖 複製到          0,0,表寬,表高,dx,0,表寬,表高
    Form1.PaintPicture Form1.Picture, 0, 0, 1024 - dx, 708, dx, 0, 1024 - dx, 708, vbSrcCopy
    '上面form1截圖 複製到           0,0,圖寬-dx,表高,dx,0,圖寬-dx,表高
    Form1.PaintPicture Form1.Picture, 1024 - dx, 0, dx - 100, 708, 0, 0, dx - 100, 708, vbSrcCopy
    '上面form1截圖 複製到           圖寬-dx,0,表寬+(dx-圖寬),表高,0,0,表寬+(dx+圖寬),表高
End If
Form1.Refresh
End Sub
