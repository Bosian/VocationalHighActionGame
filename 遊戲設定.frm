VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�C���]�w"
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
      Caption         =   "�w�]��"
      Height          =   495
      Index           =   0
      Left            =   5640
      TabIndex        =   12
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "����(ESC)"
      Height          =   495
      Index           =   1
      Left            =   5640
      TabIndex        =   11
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�T�w(Enter)"
      Default         =   -1  'True
      Height          =   495
      Index           =   2
      Left            =   5640
      TabIndex        =   10
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "�v���B�n��"
      Height          =   4575
      Left            =   2640
      TabIndex        =   8
      Top             =   120
      Width           =   2655
      Begin VB.CheckBox Check1 
         Caption         =   "���ˮ`�ĪG"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         Value           =   1  '�֨�
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "��L�]�w"
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
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   19
         Top             =   3840
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "S"
         Height          =   495
         Index           =   5
         Left            =   1080
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   18
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "A"
         Height          =   495
         Index           =   4
         Left            =   1080
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   17
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "��"
         Height          =   495
         Index           =   3
         Left            =   1080
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   16
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "��"
         Height          =   495
         Index           =   2
         Left            =   1080
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   15
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "��"
         Height          =   495
         Index           =   1
         Left            =   1080
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   14
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "��"
         Height          =   495
         Index           =   0
         Left            =   1080
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "���D"
         BeginProperty Font 
            Name            =   "�з���"
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
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�з���"
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
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�з���"
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
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�k"
         BeginProperty Font 
            Name            =   "�з���"
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
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�з���"
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
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�U"
         BeginProperty Font 
            Name            =   "�з���"
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
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�W"
         BeginProperty Font 
            Name            =   "�з���"
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
      BackStyle       =   0  '�z��
      BeginProperty Font 
         Name            =   "�з���"
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
Dim s As Integer '��L�]�w�����@��
Dim m(6) As Integer '�w�]��

Private Sub Check1_Click()
If Check1.Value = 1 Then te = 1 Else te = 0
End Sub
Private Sub Command1_Click(Index As Integer) '0�w�]�ȡB1�����B0�T�w
Select Case Index
    Case 0
        Check1.Value = 1: te = 1 ' ���ˮ`�ĪG
        
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
Private Sub Command2_Click(Index As Integer) '��L�]�w
If s <> Index Then Command2(s).BackColor = &H8000000F
s = Index
Command2(Index).BackColor = RGB(128, 255, 250)
End Sub
Private Sub Command2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    Case 0 '�W
        keyup = KeyCode
    Case 1 '�U
        keydown = KeyCode
    Case 2 '��
        keyleft = KeyCode
    Case 3 '�k
        keyright = KeyCode
    Case 4 '����
        keya = KeyCode
    Case 5 '����
        keys = KeyCode
    Case 6 '���D
        keyd = KeyCode
End Select
Command2(Index).Caption = Chr(KeyCode) '�j�g��UCase �p�g��LCase Chr���ন�^��

If KeyCode = 37 Then Command2(Index).Caption = "��"
If KeyCode = 38 Then Command2(Index).Caption = "��"
If KeyCode = 39 Then Command2(Index).Caption = "��"
If KeyCode = 40 Then Command2(Index).Caption = "��"

Command2(Index).BackColor = &H8000000F
End Sub
Private Sub Form_Load()
Call start

Label8.Caption = " �p�n�զ� " & vbCrLf & " �������� " & vbCrLf & "�Ы��w�]��"

Call system(1)
Call imvo

Open "first" For Output As #1
    Write #1, 1
Close #1
End Sub
Public Sub system(a As Integer) '�۰��٭�
If a = 0 Then
    For f = 0 To 6
        Call Command2_KeyDown((f), m(f), 0)
    Next
Else
    Open "keycode" For Append As #1 '�O�@����
    Close #1 '�O�@����
    Open "keycode" For Input As #1 'Ū����L���s��
        For f = 0 To 6
            If Not EOF(1) Then Input #1, m(f)
            Call Command2_KeyDown((f), m(f), 0)
        Next
    Close #1
End If
End Sub
Public Sub imvo() '�v���B�n���]�wŪ���@��
Open "iv" For Append As #1
Close #1
Open "iv" For Input As #1
    If Not EOF(1) Then Input #1, te
    If te = 1 Then Check1.Value = 1 Else Check1.Value = 0
Close #1
End Sub
Public Sub start() '�sŪ�ɤ���
m(0) = 38 '��l��
m(1) = 40
m(2) = 37
m(3) = 39
m(4) = 65
m(5) = 83
m(6) = 68
End Sub
Private Sub Form_Unload(Cancel As Integer) '��X��
Call Command1_Click(1)
End Sub
