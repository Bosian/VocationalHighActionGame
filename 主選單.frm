VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   1  '��u�T�w
   Caption         =   "���~���I v 4.1 Bata�p��s��"
   ClientHeight    =   8190
   ClientLeft      =   3990
   ClientTop       =   2880
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  '�S���ؽu
      Height          =   8175
      Left            =   0
      ScaleHeight     =   8175
      ScaleWidth      =   11895
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   11895
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   11160
         Top             =   1200
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   11040
         Top             =   480
      End
      Begin VB.Label Label5 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "GAME START"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   36
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   720
         Index           =   4
         Left            =   4320
         TabIndex        =   11
         Top             =   6480
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label Label5 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�ާ@�]�w"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   20.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   3
         Left            =   5250
         TabIndex        =   10
         Top             =   5520
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  '��u�T�w
         Height          =   600
         Index           =   7
         Left            =   9960
         Picture         =   "�D���.frx":0000
         Top             =   6840
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  '��u�T�w
         Height          =   600
         Index           =   6
         Left            =   10080
         Picture         =   "�D���.frx":079C
         Top             =   6720
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  '��u�T�w
         Height          =   600
         Index           =   5
         Left            =   10200
         Picture         =   "�D���.frx":0F42
         Top             =   6600
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  '��u�T�w
         Height          =   600
         Index           =   4
         Left            =   10320
         Picture         =   "�D���.frx":16D6
         Top             =   6480
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  '��u�T�w
         Height          =   600
         Index           =   3
         Left            =   10440
         Picture         =   "�D���.frx":1E69
         Top             =   6360
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  '��u�T�w
         Height          =   600
         Index           =   2
         Left            =   10560
         Picture         =   "�D���.frx":2609
         Top             =   6240
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  '��u�T�w
         Height          =   600
         Index           =   1
         Left            =   10680
         Picture         =   "�D���.frx":2DBE
         Top             =   6120
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Image Image3 
         Height          =   540
         Left            =   4560
         Picture         =   "�D���.frx":355A
         Top             =   2760
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  '��u�T�w
         Height          =   600
         Index           =   0
         Left            =   10800
         Picture         =   "�D���.frx":3CE2
         Top             =   6000
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Label6 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "���~���I"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   72
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   1440
         Left            =   3120
         TabIndex        =   9
         Top             =   840
         Width           =   5775
      End
      Begin VB.Label Label5 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "1P VS 2P"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   20.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   2
         Left            =   5280
         TabIndex        =   8
         Top             =   4680
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "2 Players"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   20.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   1
         Left            =   5280
         TabIndex        =   7
         Top             =   3720
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Label Label5 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "1 Player"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   20.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   0
         Left            =   5280
         TabIndex        =   6
         Top             =   2760
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Timer Timer2 
      Index           =   0
      Interval        =   30
      Left            =   9960
      Top             =   3000
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   100
      Left            =   8280
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   100
      Left            =   6240
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   100
      Left            =   4440
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   100
      Left            =   2400
      Top             =   960
   End
   Begin VB.Timer Timer2 
      Index           =   1
      Interval        =   30
      Left            =   10560
      Top             =   3000
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   1
      Left            =   6000
      TabIndex        =   13
      Top             =   3360
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label3 
      Alignment       =   2  '�m�����
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   1
      Left            =   6000
      TabIndex        =   12
      Top             =   2880
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label4 
      Alignment       =   2  '�m�����
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "2P"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   1
      Left            =   1800
      TabIndex        =   4
      Top             =   6480
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label Label4 
      Alignment       =   2  '�m�����
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "1P"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   6480
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   1200
      Index           =   1
      Left            =   1440
      Shape           =   1  '�����
      Top             =   6960
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1200
      Index           =   0
      Left            =   0
      Shape           =   1  '�����
      Top             =   6960
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�ΤW�U���k��H���@Enter�T�w F2���s�C��"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   7980
   End
   Begin VB.Label Label3 
      Alignment       =   2  '�m�����
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   2880
      Width           =   210
   End
   Begin VB.Line Line2 
      X1              =   960
      X2              =   10440
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   3360
      Width           =   210
   End
   Begin VB.Image Image27 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   5
      Left            =   3120
      Picture         =   "�D���.frx":446A
      Top             =   9120
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image27 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   4
      Left            =   3240
      Picture         =   "�D���.frx":4763
      Top             =   9000
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image27 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   3
      Left            =   3360
      Picture         =   "�D���.frx":4A5C
      Top             =   8880
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image27 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   2
      Left            =   3480
      Picture         =   "�D���.frx":4D3B
      Top             =   8760
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image27 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   1
      Left            =   3600
      Picture         =   "�D���.frx":5034
      Top             =   8640
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image27 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   0
      Left            =   3720
      Picture         =   "�D���.frx":532D
      Top             =   8520
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image26 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   4
      Left            =   4920
      Picture         =   "�D���.frx":5623
      Top             =   9360
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image26 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   3
      Left            =   5040
      Picture         =   "�D���.frx":5802
      Top             =   9240
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image26 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   2
      Left            =   5160
      Picture         =   "�D���.frx":59FC
      Top             =   9120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image26 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   1
      Left            =   5280
      Picture         =   "�D���.frx":5BFA
      Top             =   9000
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image26 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   0
      Left            =   5400
      Picture         =   "�D���.frx":5DFE
      Top             =   8880
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image23 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   5
      Left            =   8640
      Picture         =   "�D���.frx":6030
      Top             =   9240
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image23 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   4
      Left            =   8760
      Picture         =   "�D���.frx":625E
      Top             =   9120
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image23 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   3
      Left            =   8880
      Picture         =   "�D���.frx":648D
      Top             =   9000
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image23 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   2
      Left            =   9000
      Picture         =   "�D���.frx":66BD
      Top             =   8880
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image23 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   1
      Left            =   9120
      Picture         =   "�D���.frx":68EC
      Top             =   8760
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image23 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   0
      Left            =   9240
      Picture         =   "�D���.frx":6B1A
      Top             =   8640
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image24 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   0
      Left            =   6720
      Picture         =   "�D���.frx":6D41
      Top             =   9120
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image24 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   1
      Left            =   6840
      Picture         =   "�D���.frx":6F20
      Top             =   9000
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image24 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   2
      Left            =   6960
      Picture         =   "�D���.frx":7159
      Top             =   8880
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image24 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   3
      Left            =   7080
      Picture         =   "�D���.frx":7338
      Top             =   8760
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Line Line1 
      X1              =   960
      X2              =   10440
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Image Image1 
      Height          =   1200
      Index           =   0
      Left            =   2160
      Picture         =   "�D���.frx":75ED
      Top             =   1560
      Width           =   1200
   End
   Begin VB.Image Image1 
      Height          =   960
      Index           =   1
      Left            =   4200
      Picture         =   "�D���.frx":78E3
      Top             =   1800
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   1200
      Index           =   2
      Left            =   6000
      Picture         =   "�D���.frx":7B15
      Top             =   1680
      Width           =   1200
   End
   Begin VB.Image Image1 
      Height          =   1200
      Index           =   3
      Left            =   7920
      Picture         =   "�D���.frx":7DCA
      Top             =   1680
      Width           =   1200
   End
   Begin VB.Menu �]�w 
      Caption         =   "�C������"
      Begin VB.Menu game 
         Caption         =   "�ާ@�]�w"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s(1) As Integer '���@�Ө���
Dim k(4) As Integer '�ʵe
Dim x(1) As Integer '���ʥ��k�]��
Dim nam, m(5) '�H������
Dim p(1) As Integer '�H�����ذ{�{
Dim dptr(3) As Integer '0)GAME START�{�{ 1)Enter���� 2)1P��w���� 3)2P��w����

Private Sub Form_Unload(Cancel As Integer) '����
Erase dptr
End Sub
Private Sub Form_Load()

Open "first" For Append As #1 '�����O�_���Ĥ@������{��
Close #1
Open "first" For Input As #1
    If Not EOF(1) Then Input #1, a
Close #1
If Val(a) = 0 Then
    te = 1
    Open "iv" For Output As #1
        Write #1, te
    Close #1
    Open "keycode" For Append As #1
    Close #1
    Open "keycode" For Output As #1
        Write #1, 38, 40, 37, 39, 188, 190, 191, _
                  84, 71, 70, 72, 90, 88, 67
    Close #1
    Form3.Show 1
Else
    Open "iv" For Input As #1
        If Not EOF(1) Then Input #1, te
    Close #1
    
    Open "keycode" For Input As #1 'Ū����L���s��
        For f = 0 To 1
            If Not EOF(1) Then Input #1, keyup(f), keydown(f), keyleft(f), keyright(f), keya(f), keys(f), keyd(f)
        Next
    Close #1
End If


If reset(0) = 0 Then
    reset(0) = 1
    '����u��Ū�@�����{���X
    Form2.Width = 12000 '(800*600)
    Form2.Height = 9000
    
    ReDim smp(7) As Integer
    
    
    akiz = 5 '��j����(5)*200=amax
    amax = 1000 '��j���᪺�e��
    
    smp(0) = 40 '�������
    smp(1) = 60 '���t���W
    smp(2) = 40 '�����ɹp
    smp(3) = 60 '�y�P���u�B
    
    smp(4) = 60 '�ܱ𪺷N��
    smp(5) = 40 '�Y��
    smp(6) = 60 '�ݤ뺧��
    smp(7) = 40 '�W����
End If
Form2.Left = Screen.Width \ 2 - Form2.Width \ 2
Form2.Top = Screen.Height \ 2 - Form2.Height \ 2

'��Player�e����l
    Picture1.Visible = True
    Label6.Left = Me.ScaleWidth \ 2 - Label6.Width \ 2 '���D
    Label6.Top = Label6.Height \ 2 '���D
    For f = 0 To 4
        If f > 0 And f < 4 Then Label5(f).Top = Label5(f - 1).Top + Label5(f - 1).Height + Image3.Height
    Next
    Label5(4).Top = (Me.ScaleHeight \ 3) * 2 - Label5(4).Height \ 2 'GAME START
    Label5(4).Visible = True
    Timer4.Enabled = True
    Image3.Left = Label5(0).Left - Image3.Width * 1.5
    Image3.Top = Label5(0).Top + Label5(0).Height \ 2 - Image3.Height \ 2
'/��Player�e����l

For f = 0 To 1
    Shape1(f).Top = Line1.Y1 - Shape1(f).Height
    Label4(f).Top = Shape1(f).Top - Label4(f).Height
Next
End Sub
Private Sub Form_Activate()
For f = 0 To 0 + player_2
    Call one(s(f), (f))
Next
End Sub
Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer) 'Player�e��
Select Case KeyCode
    Case 38 '��
        If dptr(1) > 0 And dptr(1) <> 1 Then dptr(1) = dptr(1) - 1: Image3.Top = Image3.Top - Image3.Height - Label5(0).Height
    Case 40 '��
        If dptr(1) > 0 And dptr(1) <> 4 Then dptr(1) = dptr(1) + 1: Image3.Top = Image3.Top + Image3.Height + Label5(0).Height
    Case 13, keya(0), keya(1), keyd(0), keyd(1) 'Enter
        If dptr(1) > 0 And dptr(1) < 4 Then
            Picture1.Visible = False
            Form2.KeyPreview = True
            Timer3.Enabled = False
            label2(1).Visible = True
            Label3(1).Visible = True
        End If
        Select Case dptr(1)
            Case 0 '���UGAME START
                s(1) = 1
                Timer4.Enabled = False
                Label5(4).Visible = False
                dptr(1) = dptr(1) + 1
                For f = 0 To 3
                    Label5(f).Visible = True
                Next
                Image3.Visible = True
                Timer3.Enabled = True
            Case 1 '1P
                player_2 = 0
                Shape1(0).Visible = True
                Shape1(1).Visible = False
                label2(1).Visible = False
                Label3(1).Visible = False
                Label4(1).Visible = False
                pk_mod = 0
                Call imamov(0, 0)
            Case 2, 3 '2P
                If dptr(1) = 3 Then pk_mod = 1 Else pk_mod = 0
                player_2 = 1
                For f = 0 To 1
                    Label4(f).Visible = True
                    Shape1(f).Visible = True
                    Call imamov((f), (f))
                Next
            Case 4 '�ާ@�]�w
                Call game_Click
        End Select
End Select
End Sub
Private Sub Timer3_Timer() '�ӷ��ʵe
If k(4) = 7 Then k(4) = 0 Else k(4) = k(4) + 1
Image3.Picture = Image2(k(4)).Picture
End Sub
Private Sub game_Click() '�\���C���C���]�w
Form3.Show 1
End Sub
Private Sub one(a As Integer, b As Integer)
Call sta_else(b)

Label3(b).Caption = nam(a)
label2(b).Caption = ""
If player_2 = 1 Then label2(b).Caption = b + 1 & "P" & vbCrLf
For i = 0 To 5
    label2(b).Caption = label2(b).Caption & m(i)(a) & vbCrLf
Next
End Sub
Private Sub sta_else(a As Integer)
Select Case keya(a)
    Case 188
        b = ","
    Case 190
        b = "."
    Case 191
        b = "/"
    Case 32
        b = "�ť���"
    Case Else
        b = Chr(keya(a))
End Select
nam = Array("����", "����", "�p��", "����") '����W��
m(0) = Array("���ۢ��G�������" & vbCrLf & "���ӡG" & smp(0) * akiz & "MP", "���ۢ��G���t���W" & vbCrLf & "���ӡG" & smp(1) * akiz & "MP", "���ۢ��G�����ɹp" & vbCrLf & "���ӡG" & smp(2) * akiz & "MP", "���ۢ��G�y�P���u�B" & vbCrLf & "���ӡG" & smp(3) * akiz & "MP")   '����1�W�� ��������
m(2) = Array("���k�G" & b, "���k�G" & b, "���k�G" & b, "���k�G" & b) '����1���k
m(3) = Array("���ۢ��G�ܱ𪺷N��" & vbCrLf & "���ӡG" & smp(4) * akiz & "MP", "���ۢ��G�Y��" & vbCrLf & "���ӡG" & smp(5) * akiz & "MP", "���ۢ��G�ݤ뺧��" & vbCrLf & "���ӡG" & smp(6) * akiz & "MP", "���ۢ��G�W����" & vbCrLf & "���ӡG" & smp(7) * akiz & "MP") '����2�W��
m(5) = Array("���k�G��" & b, "���k�G��" & b, "���k�G��" & b, "���k�G��" & b) '����2���k
If player_2 = 0 Then
    m(1) = Array("²���G�t�ѨϪ��b�Ƥ����A�����|�b�ƼĤH�C", "²���G�������c�������W�A�y���Q�q�s���ˮ`�C", "²���G�y���T�q�s���ˮ`�A�����|�·��C", "²���G�o�g�s�򪺴��u�A�óy�����h�ĪG�C") '����1²��
    m(4) = Array("²���G�۳�ܱ�A�N͢���N�ӤƬ��O�q�C", "²���G���e���d��ˮ`�A�����|�ۤơC", "²���G�Ѥ�ް_�����i�A�y�����q�ˮ`�C", "²���G�H���O�ϯ�q�z���A�y���h�q�ˮ`�C") ''����2²��
Else
    m(1) = Array("²���G�t�ѨϪ��b�Ƥ����A" & vbCrLf & "�@�@�@�����|�b�ƼĤH�C", "²���G�������c�������W�A" & vbCrLf & "�@�@�@�y���Q�q�s���ˮ`�C", "²���G�y���T�q�s���ˮ`�A" & vbCrLf & "�@�@�@�����|�·��C", "²���G�o�g�s�򪺴��u�A" & vbCrLf & "�@�@�@�æ����h�ĪG�C") '����1²��
    m(4) = Array("²���G�۳�ܱ�A�N͢��" & vbCrLf & "�@�@�@�N�ӤƬ��O�q�C", "²���G���e���d��ˮ`�A" & vbCrLf & "�@�@�@�����|�ۤơC", "²���G�Ѥ�ް_�����i�A" & vbCrLf & "�@�@�@�y�����q�ˮ`�C", "²���G�H���O�ϯ�q�z���A" & vbCrLf & "�@�@�@�y���h�q�ˮ`�C") ''����2²��
End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer) '����
For f = 0 To 0 + player_2
    Select Case KeyCode
        Case keyleft(f) '��
            If dptr(2 + f) = 0 Then
                If s(f) > 0 Then
                    If player_2 = 0 Then
                        Call imamov(s(0) - 1, 0)
                    Else
                        If s(f) - 1 = s(1 - f) Then
                            If s(f) - 2 >= 0 Then
                                Call imamov(s(f) - 2, (f))
                            End If
                        Else
                            Call imamov(s(f) - 1, (f))
                        End If
                    End If
                End If
            End If
        Case keyright(f) '�k
            If dptr(2 + f) = 0 Then
                If s(f) < 3 Then
                    If player_2 = 0 Then
                        Call imamov(s(0) + 1, 0)
                    Else
                        If s(f) + 1 = s(1 - f) Then
                            If s(f) + 2 <= 3 Then
                                Call imamov(s(f) + 2, (f)) '��2�ӥ�
                            End If
                        Else
                            Call imamov(s(f) + 1, (f)) '��1�ӥ�
                        End If
                    End If
                End If
            End If
        Case keya(f), keyd(f)
            kikyou(f) = s(f)
            If player_2 = 1 Then
                dptr(2 + f) = 1
                Shape1(f).Visible = True
                If dptr(3 - f) = 1 Then Unload Me: gz = 0: Form1.Show
            Else
                Unload Me: gz = 0: Form1.Show
            End If
        Case 13 'Enter
            If player_2 = 0 Then
                kikyou(1) = -1
                kikyou(0) = s(0): Unload Me: gz = 0: Form1.Show
            End If
        Case keys(f)
            If KeyCode = keys(f) And dptr(2 + f) = 0 Then
                Me.KeyPreview = False: Picture1.Visible = True: Shape1(1).Visible = False
                Timer3.Enabled = True
                
                Label4(f).Visible = False
                Shape1(f).Visible = False
                s(f) = f
                Call imamov((f), (f))
                For ff = 0 To 1
                    dptr(ff + 2) = 0
                Next
            End If
            If KeyCode = keys(f) And dptr(2 + f) = 1 Then dptr(2 + f) = 0
    End Select
Next
End Sub
Private Sub imamov(a As Integer, b As Integer)
If a >= s(b) Then x(b) = 1 Else x(b) = -1 '�M�w�H���W�ٲ��ʪ���V
s(b) = a
Call one(s(b), (b)) '�I�s���⻡�������A
Timer2(b).Enabled = True '����W�ٲ���
For i = 0 To 3
    If i <> s(b) Then Timer1(i).Enabled = False Else Timer1(s(b)).Enabled = True
Next
End Sub
Private Sub Timer1_Timer(Index As Integer)
For f = 0 To 0 + player_2
    If dptr(2) = 0 And f = 0 Or dptr(3) = 0 And f = 1 Then
        k(s(f)) = k(s(f)) + 1
        
        p(f) = p(f) + 1
        If p(f) Mod 2 = 0 Then Shape1(f).Visible = True: p(f) = 0 Else Shape1(f).Visible = False
        
        Select Case s(f)
            Case 0
                If k(s(f)) = 6 Then k(s(f)) = 0
                Image1(s(f)).Picture = Image27(k(s(f))).Picture
            Case 1
                If k(s(f)) = 5 Then k(s(f)) = 0
                Image1(s(f)).Picture = Image26(k(s(f))).Picture
            Case 2
                If k(s(f)) = 4 Then k(s(f)) = 0
                Image1(s(f)).Picture = Image24(k(s(f))).Picture
            Case 3
                If k(s(f)) = 6 Then k(s(f)) = 0
                Image1(s(f)).Picture = image23(k(s(f))).Picture
        End Select
    End If
Next
End Sub
Private Sub Timer2_Timer(Index As Integer)

Shape1(Index).Left = Shape1(Index).Left + 250 * x(Index)

If x(Index) = 1 Then
    If Shape1(Index).Left >= Image1(s(Index)).Left Then Shape1(Index).Left = Image1(s(Index)).Left: Timer2(Index).Enabled = False
Else
    If Shape1(Index).Left <= Image1(s(Index)).Left Then Shape1(Index).Left = Image1(s(Index)).Left: Timer2(Index).Enabled = False
End If
Label4(Index).Left = Shape1(Index).Left + Shape1(Index).Width \ 2 - Label4(Index).Width \ 2
Label3(Index).Left = Shape1(Index).Left + Shape1(Index).Width \ 2 - Label3(Index).Width \ 2

DoEvents
End Sub

Private Sub Timer4_Timer() 'GAME START�{�{
dptr(0) = dptr(0) + 1
If dptr(0) Mod 2 = 0 Then Label5(4).Visible = True: dptr(0) = 0 Else Label5(4).Visible = False
End Sub
