VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  '單線固定
   Caption         =   "遊戲設定"
   ClientHeight    =   4875
   ClientLeft      =   6450
   ClientTop       =   4560
   ClientWidth     =   7215
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7215
   Begin VB.Frame Frame2 
      Caption         =   "2P 鍵盤設定"
      Height          =   4575
      Left            =   2640
      TabIndex        =   21
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton Command2 
         Caption         =   "T"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   1080
         Style           =   1  '圖片外觀
         TabIndex        =   35
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   1080
         Style           =   1  '圖片外觀
         TabIndex        =   34
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   1080
         Style           =   1  '圖片外觀
         TabIndex        =   33
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   10
         Left            =   1080
         Style           =   1  '圖片外觀
         TabIndex        =   32
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Z"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   11
         Left            =   1080
         Style           =   1  '圖片外觀
         TabIndex        =   31
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   12
         Left            =   1080
         Style           =   1  '圖片外觀
         TabIndex        =   30
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   13
         Left            =   1080
         Style           =   1  '圖片外觀
         TabIndex        =   29
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label15 
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
         TabIndex        =   28
         Top             =   360
         Width           =   330
      End
      Begin VB.Label Label14 
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
         TabIndex        =   27
         Top             =   960
         Width           =   330
      End
      Begin VB.Label Label13 
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
         TabIndex        =   26
         Top             =   1560
         Width           =   330
      End
      Begin VB.Label Label12 
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
         TabIndex        =   25
         Top             =   2160
         Width           =   330
      End
      Begin VB.Label Label11 
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
         TabIndex        =   24
         Top             =   2760
         Width           =   660
      End
      Begin VB.Label Label10 
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
         TabIndex        =   23
         Top             =   3360
         Width           =   660
      End
      Begin VB.Label Label9 
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
         TabIndex        =   22
         Top             =   3960
         Width           =   660
      End
   End
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
   Begin VB.Frame Frame3 
      Caption         =   "影像、聲音"
      Height          =   855
      Left            =   5040
      TabIndex        =   8
      Top             =   120
      Width           =   1935
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
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Value           =   1  '核取
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "1P 鍵盤設定"
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton Command2 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   1080
         Style           =   1  '圖片外觀
         TabIndex        =   19
         Top             =   3840
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   1080
         Style           =   1  '圖片外觀
         TabIndex        =   18
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   ","
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Left            =   5040
      TabIndex        =   20
      Top             =   1080
      Width           =   180
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As Integer '鍵盤設定的那一顆
Dim m(13) As Integer '預設值

Private Sub Check1_Click()
If Check1.Value = 1 Then te = 1 Else te = 0
End Sub
Private Sub Command1_Click(Index As Integer) '0預設值、1取消、2確定
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
            For f = 0 To 1
                Write #1, keyup(f), keydown(f), keyleft(f), keyright(f), keya(f), keys(f), keyd(f)
            Next
        Close #1
        Unload Me
End Select
End Sub
Private Sub Command2_Click(Index As Integer) '1P 鍵盤設定
If s <> Index Then Command2(s).BackColor = &H8000000F
s = Index
Command2(Index).BackColor = RGB(128, 255, 250)
End Sub
Private Sub Command2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If Shift <> 0 Then MsgBox "不可為 Shift、Crtl、Alt", 48, "訊息": Exit Sub
If Index >= 7 Then a = 1
Select Case Index
    Case 0 + 7 * a '上
        keyup(a) = KeyCode
    Case 1 + 7 * a '下
        keydown(a) = KeyCode
    Case 2 + 7 * a '左
        keyleft(a) = KeyCode
    Case 3 + 7 * a '右
        keyright(a) = KeyCode
    Case 4 + 7 * a '絕招
        keya(a) = KeyCode
    Case 5 + 7 * a '攻擊
        keys(a) = KeyCode
    Case 6 + 7 * a '跳躍
        keyd(a) = KeyCode
        
    'Case 7 '上
    '    keyup2 = KeyCode
    'Case 8 '下
    '    keydown2 = KeyCode
    'Case 9 '左
    '    keyleft2 = KeyCode
    'Case 10 '右
    '    keyright2 = KeyCode
    'Case 11 '絕招
    '    keya2 = KeyCode
    'Case 12 '攻擊
    '    keys2 = KeyCode
    'Case 13 '跳躍
    '    keyd2 = KeyCode
End Select
Command2(Index).Caption = Chr(KeyCode) '大寫為UCase 小寫為LCase Chr為轉成英文

Select Case KeyCode
    Case 37
        Command2(Index).Caption = "←"
    Case 38
        Command2(Index).Caption = "↑"
    Case 39
        Command2(Index).Caption = "→"
    Case 40
        Command2(Index).Caption = "↓"
    Case 188
        Command2(Index).Caption = ","
    Case 190
        Command2(Index).Caption = "."
    Case 191
        Command2(Index).Caption = "/"
    Case 32
        Command2(Index).Caption = "空白鍵"
End Select

Command2(Index).BackColor = &H8000000F
End Sub
Private Sub Form_Load()
Form3.Left = Screen.Width \ 2 - Form3.Width \ 2
Form3.Top = Screen.Height \ 2 - Form3.Height \ 2

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
    For f = 0 To 13
        Call Command2_KeyDown((f), m(f), 0)
    Next
Else
    Open "keycode" For Append As #1 '保護機制
    Close #1 '保護機制
    Open "keycode" For Input As #1 '讀取鍵盤的存檔
        For f = 0 To 13
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
m(4) = 188
m(5) = 190
m(6) = 191

m(7) = 84 '初始化
m(8) = 71
m(9) = 70
m(10) = 72
m(11) = 90
m(12) = 88
m(13) = 67
End Sub
Private Sub Form_Unload(Cancel As Integer) '按"關閉 (X)"時
Call Command1_Click(1)
End Sub
