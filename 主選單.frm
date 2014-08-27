VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "動作遊戲 v 1.0"
   ClientHeight    =   8490
   ClientLeft      =   1770
   ClientTop       =   1110
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   100
      Left            =   8280
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   100
      Left            =   6240
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   100
      Left            =   4560
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   100
      Left            =   2400
      Top             =   720
   End
   Begin VB.Line Line2 
      X1              =   960
      X2              =   10440
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   210
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "方向鍵選擇角色 Enter確定 用A S D來控制角色"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1170
      TabIndex        =   0
      Top             =   240
      Width           =   8820
   End
   Begin VB.Image Image27 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   5
      Left            =   3120
      Picture         =   "主選單.frx":0000
      Top             =   9120
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image27 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   4
      Left            =   3240
      Picture         =   "主選單.frx":02F9
      Top             =   9000
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image27 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   3
      Left            =   3360
      Picture         =   "主選單.frx":05F2
      Top             =   8880
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image27 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   2
      Left            =   3480
      Picture         =   "主選單.frx":08D1
      Top             =   8760
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image27 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   1
      Left            =   3600
      Picture         =   "主選單.frx":0BCA
      Top             =   8640
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image27 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   0
      Left            =   3720
      Picture         =   "主選單.frx":0EC3
      Top             =   8520
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image26 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   4
      Left            =   4920
      Picture         =   "主選單.frx":11B9
      Top             =   9360
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image26 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   3
      Left            =   5040
      Picture         =   "主選單.frx":1398
      Top             =   9240
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image26 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   2
      Left            =   5160
      Picture         =   "主選單.frx":1592
      Top             =   9120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image26 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   1
      Left            =   5280
      Picture         =   "主選單.frx":1790
      Top             =   9000
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image26 
      BorderStyle     =   1  '單線固定
      Height          =   1020
      Index           =   0
      Left            =   5400
      Picture         =   "主選單.frx":1994
      Top             =   8880
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image23 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   5
      Left            =   8640
      Picture         =   "主選單.frx":1BC6
      Top             =   9240
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image23 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   4
      Left            =   8760
      Picture         =   "主選單.frx":1DF4
      Top             =   9120
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image23 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   3
      Left            =   8880
      Picture         =   "主選單.frx":2023
      Top             =   9000
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image23 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   2
      Left            =   9000
      Picture         =   "主選單.frx":2253
      Top             =   8880
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image23 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   1
      Left            =   9120
      Picture         =   "主選單.frx":2482
      Top             =   8760
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image23 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   0
      Left            =   9240
      Picture         =   "主選單.frx":26B0
      Top             =   8640
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image24 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   0
      Left            =   6720
      Picture         =   "主選單.frx":28D7
      Top             =   9120
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image24 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   1
      Left            =   6840
      Picture         =   "主選單.frx":2AB6
      Top             =   9000
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image24 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   2
      Left            =   6960
      Picture         =   "主選單.frx":2CEF
      Top             =   8880
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image24 
      BorderStyle     =   1  '單線固定
      Height          =   1260
      Index           =   3
      Left            =   7080
      Picture         =   "主選單.frx":2ECE
      Top             =   8760
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Line Line1 
      X1              =   960
      X2              =   10440
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Image Image1 
      Height          =   1200
      Index           =   0
      Left            =   2160
      Picture         =   "主選單.frx":3183
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Image Image1 
      Height          =   960
      Index           =   1
      Left            =   4320
      Picture         =   "主選單.frx":3479
      Top             =   1680
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   1200
      Index           =   2
      Left            =   6000
      Picture         =   "主選單.frx":36AB
      Top             =   1560
      Width           =   1200
   End
   Begin VB.Image Image1 
      Height          =   1200
      Index           =   3
      Left            =   7920
      Picture         =   "主選單.frx":3960
      Top             =   1560
      Width           =   1200
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s '哪一個角色
Dim k(3) '動畫
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer) '按鍵
Select Case KeyCode
    Case 37 '左
        If s <> 0 Then
            Call Image1_Click(s - 1)
        End If
    Case 39 '右
        If s <> 3 Then
            Call Image1_Click(s + 1)
        End If
    Case 13 'Enter
        kikyou = s
        Unload Me
        Form1.Show
End Select
End Sub
Private Sub Form_Load()
Form2.Left = 1845
Form2.Top = 660
Form2.Width = 12000 '(800*600)
Form2.Height = 9000

Call Image1_Click(0)
End Sub
Private Sub Image1_Click(Index As Integer)
s = Index
Timer1(s).Enabled = True
Call one(s) '呼叫角色說明的狀態
For i = 0 To 3
    If i <> s Then
        Timer1(i).Enabled = False
    End If
Next
End Sub
Private Sub Timer1_Timer(Index As Integer)
k(s) = k(s) + 1
Select Case s
    Case 0
        If k(s) = 6 Then k(s) = 0
        Image1(s).Picture = Image27(k(s)).Picture
    Case 1
        If k(s) = 5 Then k(s) = 0
        Image1(s).Picture = Image26(k(s)).Picture
    Case 2
        If k(s) = 4 Then k(s) = 0
        Image1(s).Picture = Image24(k(s)).Picture
    Case 3
        If k(s) = 6 Then k(s) = 0
        Image1(s).Picture = Image23(k(s)).Picture
End Select
End Sub
Private Sub one(a)
Select Case a
    Case 0
        Label2 = "                 神鹿" & vbCrLf & "屬性：土系" & vbCrLf & "絕招：天崩地裂" & vbCrLf & "簡介："
    Case 1
        Label2 = "                 火鳥" & vbCrLf & "屬性：火系" & vbCrLf & "絕招：火柱" & vbCrLf & "簡介：由四大元素中的火元素凝聚而成，" & vbCrLf & "       以自己為中心發出範圍性的二連環傷害。"
    Case 2
        Label2 = "                 雷鳥" & vbCrLf & "屬性：風系" & vbCrLf & "絕招：暴雷" & vbCrLf & "簡介：由極微小的負離子所構成的強力閃電，" & vbCrLf & "       "
    Case 3
        Label2 = "                 北極熊" & vbCrLf & "屬性：水系" & vbCrLf & "絕招：明鏡止水" & vbCrLf & "簡介："
End Select
End Sub
