VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  '單線固定
   Caption         =   "遊戲設定"
   ClientHeight    =   4875
   ClientLeft      =   4770
   ClientTop       =   4200
   ClientWidth     =   7215
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7215
   Begin VB.CommandButton Command1 
      Caption         =   "預設值"
      Height          =   495
      Index           =   0
      Left            =   5640
      TabIndex        =   12
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "取消(ESC)"
      Height          =   495
      Index           =   1
      Left            =   5640
      TabIndex        =   11
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定(Enter)"
      Default         =   -1  'True
      Height          =   495
      Index           =   2
      Left            =   5640
      TabIndex        =   10
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "影像、聲音"
      Height          =   4575
      Left            =   2640
      TabIndex        =   8
      Top             =   120
      Width           =   2655
      Begin VB.CheckBox Check1 
         Caption         =   "跳傷害效果"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Value           =   1  '核取
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "鍵盤設定"
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton Command2 
         Caption         =   "D"
         Height          =   495
         Index           =   6
         Left            =   1080
         Style           =   1  '圖片外觀
         TabIndex        =   19
         Top             =   3840
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "S"
         Height          =   495
         Index           =   5
         Left            =   1080
         Style           =   1  '圖片外觀
         TabIndex        =   18
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "A"
         Height          =   495
         Index           =   4
         Left            =   1080
         Style           =   1  '圖片外觀
         TabIndex        =   17
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "→"
         Height          =   495
         Index           =   3
         Left            =   1080
         Style           =   1  '圖片外觀
         TabIndex        =   16
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "←"
         Height          =   495
         Index           =   2
         Left            =   1080
         Style           =   1  '圖片外觀
         TabIndex        =   15
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "↓"
         Height          =   495
         Index           =   1
         Left            =   1080
         Style           =   1  '圖片外觀
         TabIndex        =   14
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "↑"
         Height          =   495
         Index           =   0
         Left            =   1080
         Style           =   1  '圖片外觀
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "跳躍"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   3960
         Width           =   660
      End
      Begin VB.Label Label6 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "攻擊"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   3360
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "絕招"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   2760
         Width           =   660
      End
      Begin VB.Label Label4 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "右"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   4
         Top             =   2160
         Width           =   330
      End
      Begin VB.Label Label3 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "左"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   3
         Top             =   1560
         Width           =   330
      End
      Begin VB.Label Label2 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "下"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Top             =   960
         Width           =   330
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "上"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   330
      End
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      TabIndex        =   20
      Top             =   240
      Width           =   180
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As Integer '鍵盤設定的那一顆
Dim m(6) As Integer '預設值

Private Sub Check1_Click()
If Check1.Value = 1 Then te = 1 Else te = 0
End Sub
Private Sub Command1_Click(Index As Integer) '0預設值、1取消、0確定
Select Case Index
    Case 0
        Check1.Value = 1: te = 1 ' 跳傷害效果
        
        Call start
        Call system(0)
    Case 1
        Call system(1)
        Call imvo
        Unload Me
    Case 2
        Open "iv" For Output As #1
            Write #1, te
        Close #1
    
        Open "keycode" For Output As #1
            Write #1, keyup, keydown, keyleft, keyright, keya, keys, keyd
        Close #1
        Unload Me
End Select
End Sub
Private Sub Command2_Click(Index As Integer) '鍵盤設定
If s <> Index Then Command2(s).BackColor = &H8000000F
s = Index
Command2(Index).BackColor = RGB(128, 255, 250)
End Sub
Private Sub Command2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    Case 0 '上
        keyup = KeyCode
    Case 1 '下
        keydown = KeyCode
    Case 2 '左
        keyleft = KeyCode
    Case 3 '右
        keyright = KeyCode
    Case 4 '絕招
        keya = KeyCode
    Case 5 '攻擊
        keys = KeyCode
    Case 6 '跳躍
        keyd = KeyCode
End Select
Command2(Index).Caption = Chr(KeyCode) '大寫為UCase 小寫為LCase Chr為轉成英文

If KeyCode = 37 Then Command2(Index).Caption = "←"
If KeyCode = 38 Then Command2(Index).Caption = "↑"
If KeyCode = 39 Then Command2(Index).Caption = "→"
If KeyCode = 40 Then Command2(Index).Caption = "↓"

Command2(Index).BackColor = &H8000000F
End Sub
Private Sub Form_Load()
Call start

Label8.Caption = " 如要調成 " & vbCrLf & " ↑↓←→ " & vbCrLf & "請按預設值"

Call system(1)
Call imvo

Open "first" For Output As #1
    Write #1, 1
Close #1
End Sub
Public Sub system(a As Integer) '自動還原
If a = 0 Then
    For f = 0 To 6
        Call Command2_KeyDown((f), m(f), 0)
    Next
Else
    Open "keycode" For Append As #1 '保護機制
    Close #1 '保護機制
    Open "keycode" For Input As #1 '讀取鍵盤的存檔
        For f = 0 To 6
            If Not EOF(1) Then Input #1, m(f)
            Call Command2_KeyDown((f), m(f), 0)
        Next
    Close #1
End If
End Sub
Public Sub imvo() '影像、聲音設定讀取共用
Open "iv" For Append As #1
Close #1
Open "iv" For Input As #1
    If Not EOF(1) Then Input #1, te
    If te = 1 Then Check1.Value = 1 Else Check1.Value = 0
Close #1
End Sub
Public Sub start() '存讀檔之用
m(0) = 38 '初始化
m(1) = 40
m(2) = 37
m(3) = 39
m(4) = 65
m(5) = 83
m(6) = 68
End Sub
Private Sub Form_Unload(Cancel As Integer) '按X時
Call Command1_Click(1)
End Sub
