VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "���~���I (�ʧ@�C�� v 3.1)"
   ClientHeight    =   8190
   ClientLeft      =   1770
   ClientTop       =   1410
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   Begin VB.Timer Timer2 
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
      Height          =   405
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
      Height          =   405
      Left            =   960
      TabIndex        =   0
      Top             =   3720
      Width           =   210
   End
   Begin VB.Image Image27 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   5
      Left            =   3120
      Picture         =   "�D���.frx":0000
      Top             =   9120
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image27 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   4
      Left            =   3240
      Picture         =   "�D���.frx":02F9
      Top             =   9000
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image27 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   3
      Left            =   3360
      Picture         =   "�D���.frx":05F2
      Top             =   8880
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image27 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   2
      Left            =   3480
      Picture         =   "�D���.frx":08D1
      Top             =   8760
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image27 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   1
      Left            =   3600
      Picture         =   "�D���.frx":0BCA
      Top             =   8640
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image27 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   0
      Left            =   3720
      Picture         =   "�D���.frx":0EC3
      Top             =   8520
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image26 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   4
      Left            =   4920
      Picture         =   "�D���.frx":11B9
      Top             =   9360
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image26 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   3
      Left            =   5040
      Picture         =   "�D���.frx":1398
      Top             =   9240
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image26 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   2
      Left            =   5160
      Picture         =   "�D���.frx":1592
      Top             =   9120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image26 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   1
      Left            =   5280
      Picture         =   "�D���.frx":1790
      Top             =   9000
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image26 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   0
      Left            =   5400
      Picture         =   "�D���.frx":1994
      Top             =   8880
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image23 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   5
      Left            =   8640
      Picture         =   "�D���.frx":1BC6
      Top             =   9240
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image23 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   4
      Left            =   8760
      Picture         =   "�D���.frx":1DF4
      Top             =   9120
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image23 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   3
      Left            =   8880
      Picture         =   "�D���.frx":2023
      Top             =   9000
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image23 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   2
      Left            =   9000
      Picture         =   "�D���.frx":2253
      Top             =   8880
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image23 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   1
      Left            =   9120
      Picture         =   "�D���.frx":2482
      Top             =   8760
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image23 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   0
      Left            =   9240
      Picture         =   "�D���.frx":26B0
      Top             =   8640
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image24 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   0
      Left            =   6720
      Picture         =   "�D���.frx":28D7
      Top             =   9120
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image24 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   1
      Left            =   6840
      Picture         =   "�D���.frx":2AB6
      Top             =   9000
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image24 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   2
      Left            =   6960
      Picture         =   "�D���.frx":2CEF
      Top             =   8880
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image24 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   3
      Left            =   7080
      Picture         =   "�D���.frx":2ECE
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
      Picture         =   "�D���.frx":3183
      Top             =   1560
      Width           =   1200
   End
   Begin VB.Image Image1 
      Height          =   960
      Index           =   1
      Left            =   4200
      Picture         =   "�D���.frx":3479
      Top             =   1800
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   1200
      Index           =   2
      Left            =   6000
      Picture         =   "�D���.frx":36AB
      Top             =   1680
      Width           =   1200
   End
   Begin VB.Image Image1 
      Height          =   1200
      Index           =   3
      Left            =   7920
      Picture         =   "�D���.frx":3960
      Top             =   1680
      Width           =   1200
   End
   Begin VB.Menu �]�w 
      Caption         =   "�C������"
      Begin VB.Menu game 
         Caption         =   "�C���]�w"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As Integer '���@�Ө���
Dim k(3) As Integer '�ʵe
Dim x As Integer '���ʥ��k�]��
Private Sub game_Click() '�\���C���C���]�w
Form3.Show 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer) '����
Select Case KeyCode
    Case 37 '��
        If s <> 0 Then
            Call Image1_Click(s - 1)
        End If
    Case 39 '�k
        If s <> 3 Then
            Call Image1_Click(s + 1)
        End If
    Case 13 'Enter
        kikyou = s
        Unload Me
        gz = 0
        Form1.Show
End Select
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
        Write #1, 38, 40, 37, 39, 65, 83, 68
    Close #1
    
    Form3.Show 1
Else
    Open "iv" For Input As #1
        If Not EOF(1) Then Input #1, te
    Close #1
    
    Open "keycode" For Input As #1 'Ū����L���s��
        If Not EOF(1) Then Input #1, keyup, keydown, keyleft, keyright, keya, keys, keyd
    Close #1
End If


If reset(0) = 0 Then
    reset(0) = 1
    '����u��Ū�@�����{���X
    
    Form2.Left = 1845
    Form2.Top = 660
    Form2.Width = 12000 '(800*600)
    Form2.Height = 9000
End If

Call Image1_Click(0)
End Sub
Private Sub Image1_Click(Index As Integer)
If Index >= s Then x = 1 Else x = -1 '�M�w�H���W�ٲ��ʪ���V
s = Index
Timer1(s).Enabled = True
Call one(s) '�I�s���⻡�������A
Timer2.Enabled = True '����W�ٲ���
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
        Label3.Caption = "����"
        Label2.Caption = "���ۢ��G�������" & vbCrLf & "²���G�H�ѨϪ����u����ĤH�A" & "�óy�����q�s���ˮ`�C" & vbCrLf & "���k�GA" & vbCrLf & "���ۢ��G�ܱ𪺷N��" & vbCrLf & "²���G�۳�ǻ������ܱ�A�ñN͢���N�ӤƬ��O�q�C" & vbCrLf & "���k�G��A" '��������
    Case 1
        Label3.Caption = "����"
        Label2.Caption = "���ۢ��G���t���W" & vbCrLf & "²���G�V�d�򤺪��ĤH���X�Q�q�s���ˮ`�C" & vbCrLf & "���k�GA" & vbCrLf & "���ۢ��G�Y��" & vbCrLf & "²���G���e���d��ˮ`�A�����|�ۤơC" & vbCrLf & "���k�G��A"
    Case 2
        Label3.Caption = "�p��"
        Label2.Caption = "���ۢ��G�����ɹp" & vbCrLf & "²���G�V�d�򤺪��ĤH���X�T�q�s���ˮ`�A�����|�·��C" & vbCrLf & "���k�GA" & vbCrLf & "���ۢ��G�ݤ뺧��" & vbCrLf & "²���G�Ѥ�Ҥް_���i���A�y���d��ʪ��h�q�ˮ`�C" & vbCrLf & "���k�G��A"
    Case 3
        Label3.Caption = "�_����"
        Label2.Caption = "���ۢ��G�y�P���u�B" & vbCrLf & "²���G�o�g�s�򪺴��u�����ĤH�C" & vbCrLf & "���k�GA" & vbCrLf & "���ۢ��G�W����" & vbCrLf & "²���G�H���O�ϯ�q�z���A�y���d��ʪ��h�q�ˮ`�C" & vbCrLf & "���k�G��A"
End Select
End Sub
Private Sub Timer2_Timer()

Label3.Left = Label3.Left + 250 * x

If x = 1 Then
    If Label3.Left >= Image1(s).Left + Image1(s).Width \ 2 - Label3.Width \ 2 Then Label3.Left = Image1(s).Left + Image1(s).Width \ 2 - Label3.Width \ 2: Timer2.Enabled = False
Else
    If Label3.Left <= Image1(s).Left + Image1(s).Width \ 2 - Label3.Width \ 2 Then Label3.Left = Image1(s).Left + Image1(s).Width \ 2 - Label3.Width \ 2: Timer2.Enabled = False
End If

DoEvents

End Sub

