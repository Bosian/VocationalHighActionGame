VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  '����
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  '��u�T�w
   Caption         =   "���~���I v 5.0.5 �p��s��"
   ClientHeight    =   8520
   ClientLeft      =   -1455
   ClientTop       =   -450
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "�ʧ@�C��.frx":0000
   ScaleHeight     =   568
   ScaleMode       =   3  '����
   ScaleWidth      =   794
   Begin VB.Timer Timer46 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   8040
      Top             =   6240
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '����
      BackColor       =   &H00000000&
      BorderStyle     =   0  '�S���ؽu
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8160
      ScaleHeight     =   375
      ScaleWidth      =   255
      TabIndex        =   27
      Top             =   6480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer Timer45 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   20
      Left            =   4920
      Top             =   5040
   End
   Begin VB.Timer Timer44 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   20
      Left            =   10920
      Top             =   6600
   End
   Begin VB.Timer Timer43 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   30
      Left            =   9840
      Top             =   9120
   End
   Begin VB.Timer Timer42 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   12240
      Top             =   1680
   End
   Begin VB.Timer Timer41 
      Enabled         =   0   'False
      Index           =   4
      Interval        =   100
      Left            =   14160
      Top             =   1200
   End
   Begin VB.Timer Timer41 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   100
      Left            =   13680
      Top             =   1200
   End
   Begin VB.Timer Timer41 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   100
      Left            =   13200
      Top             =   1200
   End
   Begin VB.Timer Timer41 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   100
      Left            =   12720
      Top             =   1200
   End
   Begin VB.Timer Timer41 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   100
      Left            =   12240
      Top             =   1200
   End
   Begin VB.Timer Timer40 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   10
      Left            =   8520
      Top             =   600
   End
   Begin VB.Timer Timer40 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   2520
      Top             =   600
   End
   Begin VB.Timer Timer39 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   1080
      Top             =   6480
   End
   Begin VB.Timer Timer38 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   500
      Left            =   6120
      Top             =   600
   End
   Begin VB.Timer Timer38 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   200
      Left            =   240
      Top             =   600
   End
   Begin VB.Timer Timer31 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   10
      Left            =   11280
      Top             =   120
   End
   Begin VB.Timer Timer31 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   5400
      Top             =   120
   End
   Begin VB.Timer Timer30 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   30
      Left            =   10800
      Top             =   120
   End
   Begin VB.Timer Timer30 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   30
      Left            =   4920
      Top             =   120
   End
   Begin VB.Timer Timer36 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   1500
      Left            =   6240
      Top             =   1920
   End
   Begin VB.Timer Timer37 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   100
      Left            =   8520
      Top             =   1440
   End
   Begin VB.Timer Timer25 
      Index           =   1
      Interval        =   500
      Left            =   6120
      Top             =   1080
   End
   Begin VB.Timer Timer20 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   200
      Left            =   10920
      Top             =   1080
   End
   Begin VB.Timer Timer21 
      Index           =   3
      Interval        =   100
      Left            =   11400
      Top             =   1080
   End
   Begin VB.Timer Timer21 
      Index           =   2
      Interval        =   100
      Left            =   11400
      Top             =   600
   End
   Begin VB.Timer Timer20 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   200
      Left            =   10920
      Top             =   600
   End
   Begin VB.Timer Timer24 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   5400
      Top             =   8040
   End
   Begin VB.Timer Timer28 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   4200
      Top             =   7320
   End
   Begin VB.Timer Timer37 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   100
      Left            =   2640
      Top             =   1440
   End
   Begin VB.Timer Timer36 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1500
      Left            =   360
      Top             =   1920
   End
   Begin VB.Timer Timer35 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   6000
      Top             =   7560
   End
   Begin VB.Timer Timer18 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   6600
      Top             =   7200
   End
   Begin VB.Timer Timer34 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   120
      Top             =   6480
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
      Left            =   6480
      Top             =   6720
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Index           =   6
      Interval        =   200
      Left            =   6120
      Top             =   6720
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Index           =   5
      Interval        =   10
      Left            =   5760
      Top             =   6720
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Index           =   4
      Interval        =   10
      Left            =   5400
      Top             =   6720
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   100
      Left            =   4920
      Top             =   6720
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   30
      Left            =   4560
      Top             =   6720
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   100
      Left            =   4200
      Top             =   6720
   End
   Begin VB.Timer Timer20 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   200
      Left            =   5040
      Top             =   1080
   End
   Begin VB.Timer Timer21 
      Index           =   1
      Interval        =   100
      Left            =   5520
      Top             =   1080
   End
   Begin VB.Timer Timer29 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   7200
      Top             =   7440
   End
   Begin VB.Timer Timer27 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   30
      Left            =   3600
      Top             =   7440
   End
   Begin VB.Timer Timer26 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   3240
      Top             =   0
   End
   Begin VB.Timer Timer25 
      Index           =   0
      Interval        =   200
      Left            =   240
      Top             =   1080
   End
   Begin VB.Timer Timer23 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   5160
      Top             =   7920
   End
   Begin VB.Timer Timer22 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   100
      Left            =   10800
      Top             =   8040
   End
   Begin VB.Timer Timer21 
      Index           =   0
      Interval        =   100
      Left            =   5520
      Top             =   600
   End
   Begin VB.Timer Timer20 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   200
      Left            =   5040
      Top             =   600
   End
   Begin VB.Timer Timer19 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   9360
      Top             =   6600
   End
   Begin VB.Timer Timer13 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   4080
      Top             =   7200
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   80
      Left            =   9840
      Top             =   7080
   End
   Begin VB.Timer Timer17 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   500
      Left            =   10680
      Top             =   6000
   End
   Begin VB.Timer Timer16 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   50
      Left            =   10200
      Top             =   6000
   End
   Begin VB.Timer Timer15 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   11160
      Top             =   2880
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   3840
      Top             =   6720
   End
   Begin VB.Timer Timer12 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   30
      Left            =   11400
      Top             =   7080
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
      Index           =   0
      Interval        =   25
      Left            =   6360
      Top             =   5040
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   100
      Left            =   6480
      Top             =   5640
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   100
      Left            =   6000
      Top             =   5640
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   5640
      Top             =   4440
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
      Interval        =   10
      Left            =   11160
      Top             =   2520
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
      Index           =   0
      Interval        =   30
      Left            =   6480
      Top             =   6120
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   30
      Left            =   6000
      Top             =   6120
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   30
      Left            =   5520
      Top             =   6120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   30
      Left            =   5040
      Top             =   6120
   End
   Begin VB.Image Image28 
      BorderStyle     =   1  '��u�T�w
      Height          =   330
      Index           =   1
      Left            =   600
      Picture         =   "�ʧ@�C��.frx":240042
      Top             =   8640
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Image Image28 
      BorderStyle     =   1  '��u�T�w
      Height          =   330
      Index           =   0
      Left            =   720
      Picture         =   "�ʧ@�C��.frx":240735
      Top             =   8520
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "X 0"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   9120
      TabIndex        =   26
      Top             =   1920
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "X 0"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   3240
      TabIndex        =   25
      Top             =   1920
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "X 3"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   9120
      TabIndex        =   24
      Top             =   600
      Width           =   540
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "X 3"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   3240
      TabIndex        =   23
      Top             =   600
      Width           =   540
   End
   Begin VB.Line Line7 
      BorderColor     =   &H0000FFFF&
      Index           =   0
      Visible         =   0   'False
      X1              =   680
      X2              =   672
      Y1              =   584
      Y2              =   600
   End
   Begin VB.Image Image26 
      BorderStyle     =   1  '��u�T�w
      Height          =   375
      Index           =   1
      Left            =   13920
      Picture         =   "�ʧ@�C��.frx":240E35
      Top             =   600
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image Image26 
      BorderStyle     =   1  '��u�T�w
      Height          =   375
      Index           =   0
      Left            =   13080
      Picture         =   "�ʧ@�C��.frx":24113F
      Top             =   600
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image Image25 
      Height          =   315
      Index           =   0
      Left            =   12360
      Picture         =   "�ʧ@�C��.frx":24144B
      Top             =   600
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image16 
      BorderStyle     =   1  '��u�T�w
      Height          =   3705
      Index           =   1
      Left            =   12360
      Picture         =   "�ʧ@�C��.frx":241757
      Top             =   -3360
      Visible         =   0   'False
      Width           =   2520
   End
   Begin VB.Image Image16 
      BorderStyle     =   1  '��u�T�w
      Height          =   3705
      Index           =   0
      Left            =   12480
      Picture         =   "�ʧ@�C��.frx":244F53
      Top             =   -3480
      Visible         =   0   'False
      Width           =   2520
   End
   Begin VB.Image Image12 
      Height          =   330
      Index           =   0
      Left            =   1080
      Picture         =   "�ʧ@�C��.frx":248730
      Top             =   6120
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Shape Shape13 
      BorderWidth     =   3
      Height          =   375
      Index           =   0
      Left            =   2280
      Shape           =   3  '���
      Top             =   7800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   0
      Visible         =   0   'False
      X1              =   152
      X2              =   168
      Y1              =   512
      Y2              =   528
   End
   Begin VB.Shape Shape22 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   1215
      Index           =   0
      Left            =   3480
      Shape           =   2  '����
      Top             =   7200
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Shape Shape21 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   1215
      Index           =   0
      Left            =   3360
      Shape           =   2  '����
      Top             =   7320
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Shape Shape20 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   1215
      Index           =   0
      Left            =   3360
      Shape           =   2  '����
      Top             =   7200
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "0 �s��"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   360
      Index           =   1
      Left            =   10080
      TabIndex        =   20
      Top             =   1920
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   6000
      TabIndex        =   21
      Top             =   1560
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label5 
      Alignment       =   2  '�m�����
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "1000/1000"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   6720
      TabIndex        =   22
      Top             =   1920
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Shape xhp 
      BorderColor     =   &H80000005&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   1
      Left            =   6000
      Top             =   1920
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   1
      Left            =   6000
      Top             =   1920
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label label1 
      Alignment       =   1  '�a�k���
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "1000/1000"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   6720
      TabIndex        =   19
      Top             =   600
      Width           =   1620
   End
   Begin VB.Shape hp 
      BorderColor     =   &H80000005&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   1
      Left            =   6000
      Top             =   600
      Width           =   3000
   End
   Begin VB.Label label2 
      Alignment       =   1  '�a�k���
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "1000/1000"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   1
      Left            =   6720
      TabIndex        =   18
      Top             =   1080
      Width           =   1620
   End
   Begin VB.Shape mp 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   1
      Left            =   6000
      Top             =   1080
      Width           =   3000
   End
   Begin VB.Shape shape5 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FF8080&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   1
      Left            =   6000
      Top             =   1080
      Width           =   3000
   End
   Begin VB.Shape shape4 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   1
      Left            =   6000
      Top             =   600
      Width           =   3000
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "��A"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   3
      Left            =   9960
      TabIndex        =   17
      Top             =   1080
      Width           =   510
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   2
      Left            =   9960
      TabIndex        =   16
      Top             =   600
      Width           =   180
   End
   Begin VB.Label label9 
      Alignment       =   1  '�a�k���
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "0"
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
      Index           =   1
      Left            =   8745
      TabIndex        =   15
      Top             =   120
      Width           =   210
   End
   Begin VB.Label label10 
      Alignment       =   1  '�a�k���
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "0 X"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   360
      Index           =   1
      Left            =   10695
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image image23 
      Height          =   645
      Index           =   1
      Left            =   11160
      Picture         =   "�ʧ@�C��.frx":248BB6
      Top             =   1320
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label label12 
      Alignment       =   2  '�m�����
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "2P"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   6180
      TabIndex        =   13
      Top             =   120
      Width           =   360
   End
   Begin VB.Label label12 
      Alignment       =   2  '�m�����
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "1P"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   300
      TabIndex        =   12
      Top             =   120
      Width           =   360
   End
   Begin VB.Shape Shape19 
      BorderStyle     =   0  '�z��
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  '���
      Height          =   135
      Index           =   0
      Left            =   7200
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape17 
      BorderStyle     =   0  '�z��
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  '���
      Height          =   255
      Index           =   0
      Left            =   5280
      Shape           =   3  '���
      Top             =   7200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "0 �s��"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   360
      Index           =   0
      Left            =   4200
      TabIndex        =   11
      Top             =   1920
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Shape Shape16 
      BorderWidth     =   7
      Height          =   255
      Index           =   0
      Left            =   6120
      Shape           =   2  '����
      Top             =   7680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image24 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   1
      Left            =   11760
      Picture         =   "�ʧ@�C��.frx":249104
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image24 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   0
      Left            =   10560
      Picture         =   "�ʧ@�C��.frx":2493F4
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image image23 
      Height          =   645
      Index           =   0
      Left            =   5310
      Picture         =   "�ʧ@�C��.frx":2496E2
      Top             =   1320
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label label10 
      Alignment       =   1  '�a�k���
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "0 X"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   360
      Index           =   0
      Left            =   4815
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image22 
      Height          =   330
      Index           =   0
      Left            =   120
      Picture         =   "�ʧ@�C��.frx":249C30
      Top             =   6120
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label label9 
      Alignment       =   1  '�a�k���
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "0"
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
      Index           =   0
      Left            =   2880
      TabIndex        =   9
      Top             =   120
      Width           =   210
   End
   Begin VB.Image Image21 
      BorderStyle     =   1  '��u�T�w
      Height          =   705
      Index           =   7
      Left            =   14520
      Picture         =   "�ʧ@�C��.frx":24A0B6
      Top             =   7080
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Image21 
      BorderStyle     =   1  '��u�T�w
      Height          =   705
      Index           =   6
      Left            =   15240
      Picture         =   "�ʧ@�C��.frx":24A609
      Top             =   7200
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Image21 
      BorderStyle     =   1  '��u�T�w
      Height          =   705
      Index           =   5
      Left            =   15960
      Picture         =   "�ʧ@�C��.frx":24AB37
      Top             =   7080
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Image21 
      BorderStyle     =   1  '��u�T�w
      Height          =   705
      Index           =   4
      Left            =   16680
      Picture         =   "�ʧ@�C��.frx":24B0B2
      Top             =   7080
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Image21 
      BorderStyle     =   1  '��u�T�w
      Height          =   705
      Index           =   3
      Left            =   14400
      Picture         =   "�ʧ@�C��.frx":24B5FF
      Top             =   6240
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Image21 
      BorderStyle     =   1  '��u�T�w
      Height          =   705
      Index           =   2
      Left            =   15120
      Picture         =   "�ʧ@�C��.frx":24BB65
      Top             =   6360
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Image21 
      BorderStyle     =   1  '��u�T�w
      Height          =   705
      Index           =   1
      Left            =   15840
      Picture         =   "�ʧ@�C��.frx":24C0A4
      Top             =   6240
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Image21 
      BorderStyle     =   1  '��u�T�w
      Height          =   705
      Index           =   0
      Left            =   16680
      Picture         =   "�ʧ@�C��.frx":24C5F2
      Top             =   6240
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Image20 
      Height          =   645
      Index           =   0
      Left            =   12120
      Picture         =   "�ʧ@�C��.frx":24CB42
      Top             =   7440
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Image Image19 
      Height          =   1005
      Left            =   1680
      Picture         =   "�ʧ@�C��.frx":24D092
      Top             =   6480
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image Image18 
      Height          =   270
      Index           =   0
      Left            =   240
      Picture         =   "�ʧ@�C��.frx":24E641
      Top             =   6840
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image Image17 
      Height          =   2775
      Left            =   0
      Picture         =   "�ʧ@�C��.frx":24ED41
      Top             =   5880
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "��A"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   4080
      TabIndex        =   8
      Top             =   1080
      Width           =   510
   End
   Begin VB.Shape delay 
      BorderColor     =   &H80000005&
      FillColor       =   &H0080FF80&
      FillStyle       =   0  '���
      Height          =   315
      Index           =   1
      Left            =   4080
      Top             =   1080
      Width           =   240
   End
   Begin VB.Image Image14 
      Height          =   315
      Index           =   1
      Left            =   3960
      Picture         =   "�ʧ@�C��.frx":2511DE
      Top             =   10080
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image14 
      Height          =   315
      Index           =   0
      Left            =   3960
      Picture         =   "�ʧ@�C��.frx":251329
      Top             =   10440
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image13 
      Height          =   315
      Index           =   0
      Left            =   7200
      Picture         =   "�ʧ@�C��.frx":251475
      Top             =   6960
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   6
      Index           =   0
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   472
      Y2              =   496
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   7
      Top             =   14400
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      Index           =   0
      Visible         =   0   'False
      X1              =   264
      X2              =   248
      Y1              =   536
      Y2              =   560
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   255
      Left            =   3600
      Shape           =   2  '����
      Top             =   8040
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   4080
      TabIndex        =   5
      Top             =   600
      Width           =   180
   End
   Begin VB.Label label2 
      Alignment       =   1  '�a�k���
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "1000/1000"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   0
      Left            =   810
      TabIndex        =   1
      Top             =   1080
      Width           =   1620
   End
   Begin VB.Shape delay 
      BorderColor     =   &H80000005&
      FillColor       =   &H0080FF80&
      FillStyle       =   0  '���
      Height          =   315
      Index           =   0
      Left            =   4080
      Top             =   600
      Width           =   240
   End
   Begin VB.Shape mp 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   0
      Left            =   120
      Top             =   1080
      Width           =   3000
   End
   Begin VB.Shape Shape8 
      FillStyle       =   0  '���
      Height          =   135
      Index           =   0
      Left            =   9360
      Shape           =   2  '����
      Top             =   7680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  '���
      Height          =   255
      Index           =   0
      Left            =   9360
      Shape           =   3  '���
      Top             =   7200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   1200
      Index           =   0
      Left            =   5280
      Picture         =   "�ʧ@�C��.frx":2515C1
      Top             =   4560
      Width           =   1200
   End
   Begin VB.Label Label4 
      Alignment       =   1  '�a�k���
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Left            =   10770
      TabIndex        =   3
      Top             =   6480
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label5 
      Alignment       =   2  '�m�����
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "1000/1000"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   840
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "G O ��"
      BeginProperty Font 
         Name            =   "�з���"
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
      Top             =   2880
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label label1 
      Alignment       =   1  '�a�k���
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "1000/1000"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   810
      TabIndex        =   0
      Top             =   600
      Width           =   1620
   End
   Begin VB.Shape shape5 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FF8080&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   0
      Left            =   120
      Top             =   1080
      Width           =   3000
   End
   Begin VB.Shape hp 
      BorderColor     =   &H80000005&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   0
      Left            =   120
      Top             =   600
      Width           =   3000
   End
   Begin VB.Shape xhp 
      BorderColor     =   &H80000005&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   0
      Left            =   120
      Top             =   1920
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   0
      Left            =   120
      Top             =   1920
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Shape shape4 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '���
      Height          =   375
      Index           =   0
      Left            =   120
      Top             =   600
      Width           =   3000
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   40
      X2              =   728
      Y1              =   160
      Y2              =   160
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   7
      Left            =   12000
      Picture         =   "�ʧ@�C��.frx":2517A0
      Top             =   10440
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   6
      Left            =   12120
      Picture         =   "�ʧ@�C��.frx":251ABE
      Top             =   10320
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   5
      Left            =   12240
      Picture         =   "�ʧ@�C��.frx":251DA2
      Top             =   10200
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   4
      Left            =   12360
      Picture         =   "�ʧ@�C��.frx":252070
      Top             =   10080
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   3
      Left            =   12480
      Picture         =   "�ʧ@�C��.frx":252331
      Top             =   9960
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   2
      Left            =   12600
      Picture         =   "�ʧ@�C��.frx":2525FF
      Top             =   9840
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   1
      Left            =   12720
      Picture         =   "�ʧ@�C��.frx":2528E8
      Top             =   9720
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   0
      Left            =   12840
      Picture         =   "�ʧ@�C��.frx":252C06
      Top             =   9600
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image6 
      Height          =   960
      Index           =   0
      Left            =   10320
      Picture         =   "�ʧ@�C��.frx":252F06
      Top             =   6720
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   1
      Left            =   12960
      Picture         =   "�ʧ@�C��.frx":253203
      Top             =   11640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   2
      Left            =   12840
      Picture         =   "�ʧ@�C��.frx":253520
      Top             =   11760
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   0
      Left            =   13080
      Picture         =   "�ʧ@�C��.frx":253800
      Top             =   11520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   3
      Left            =   12720
      Picture         =   "�ʧ@�C��.frx":253AFD
      Top             =   11880
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   4
      Left            =   12600
      Picture         =   "�ʧ@�C��.frx":253DCF
      Top             =   12000
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   6
      Left            =   12360
      Picture         =   "�ʧ@�C��.frx":254091
      Top             =   12240
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   7
      Left            =   12240
      Picture         =   "�ʧ@�C��.frx":25436D
      Top             =   12360
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   5
      Left            =   12480
      Picture         =   "�ʧ@�C��.frx":25468A
      Top             =   12120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   5
      Left            =   9840
      Picture         =   "�ʧ@�C��.frx":25495C
      Top             =   10320
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   4
      Left            =   9960
      Picture         =   "�ʧ@�C��.frx":254B83
      Top             =   10200
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   3
      Left            =   10080
      Picture         =   "�ʧ@�C��.frx":254DBC
      Top             =   10080
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   2
      Left            =   10200
      Picture         =   "�ʧ@�C��.frx":254FEC
      Top             =   9960
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   1
      Left            =   10440
      Picture         =   "�ʧ@�C��.frx":255225
      Top             =   9720
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image9 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   3
      Left            =   8160
      Picture         =   "�ʧ@�C��.frx":25544C
      Top             =   12120
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image9 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   2
      Left            =   8280
      Picture         =   "�ʧ@�C��.frx":255631
      Top             =   12000
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image9 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   1
      Left            =   8400
      Picture         =   "�ʧ@�C��.frx":25586C
      Top             =   11880
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   4
      Left            =   6480
      Picture         =   "�ʧ@�C��.frx":255A51
      Top             =   10320
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   3
      Left            =   6600
      Picture         =   "�ʧ@�C��.frx":255C34
      Top             =   10200
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   2
      Left            =   6720
      Picture         =   "�ʧ@�C��.frx":255E2E
      Top             =   10080
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   1
      Left            =   6840
      Picture         =   "�ʧ@�C��.frx":256028
      Top             =   9960
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   5
      Left            =   4560
      Picture         =   "�ʧ@�C��.frx":25622E
      Top             =   9960
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   4
      Left            =   4680
      Picture         =   "�ʧ@�C��.frx":256524
      Top             =   9840
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   3
      Left            =   4800
      Picture         =   "�ʧ@�C��.frx":25681C
      Top             =   9720
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   2
      Left            =   4920
      Picture         =   "�ʧ@�C��.frx":256AF9
      Top             =   9600
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   1
      Left            =   5040
      Picture         =   "�ʧ@�C��.frx":256DF1
      Top             =   9480
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   0
      Left            =   10560
      Picture         =   "�ʧ@�C��.frx":2570E7
      Top             =   9600
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image9 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   0
      Left            =   8520
      Picture         =   "�ʧ@�C��.frx":257315
      Top             =   11760
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   0
      Left            =   6960
      Picture         =   "�ʧ@�C��.frx":2575D2
      Top             =   9840
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   0
      Left            =   5160
      Picture         =   "�ʧ@�C��.frx":257806
      Top             =   9360
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   3
      Left            =   8160
      Picture         =   "�ʧ@�C��.frx":257AFE
      Top             =   10080
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   2
      Left            =   8280
      Picture         =   "�ʧ@�C��.frx":257DB3
      Top             =   9960
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   1
      Left            =   8400
      Picture         =   "�ʧ@�C��.frx":257F92
      Top             =   9840
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   0
      Left            =   8520
      Picture         =   "�ʧ@�C��.frx":2581CB
      Top             =   9720
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   0
      Left            =   10320
      Picture         =   "�ʧ@�C��.frx":2583AA
      Top             =   11640
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   1
      Left            =   10200
      Picture         =   "�ʧ@�C��.frx":2585D1
      Top             =   11760
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   2
      Left            =   10080
      Picture         =   "�ʧ@�C��.frx":2587FF
      Top             =   11880
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   3
      Left            =   9960
      Picture         =   "�ʧ@�C��.frx":258A2E
      Top             =   12000
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   4
      Left            =   9840
      Picture         =   "�ʧ@�C��.frx":258C5E
      Top             =   12120
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   5
      Left            =   9720
      Picture         =   "�ʧ@�C��.frx":258E8D
      Top             =   12240
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   0
      Left            =   6480
      Picture         =   "�ʧ@�C��.frx":2590BB
      Top             =   11880
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   1
      Left            =   6360
      Picture         =   "�ʧ@�C��.frx":2592ED
      Top             =   12000
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   2
      Left            =   6240
      Picture         =   "�ʧ@�C��.frx":2594F1
      Top             =   12120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   3
      Left            =   6120
      Picture         =   "�ʧ@�C��.frx":2596EF
      Top             =   12240
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   4
      Left            =   6000
      Picture         =   "�ʧ@�C��.frx":2598E9
      Top             =   12360
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   0
      Left            =   4800
      Picture         =   "�ʧ@�C��.frx":259AC8
      Top             =   11520
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   1
      Left            =   4680
      Picture         =   "�ʧ@�C��.frx":259DBE
      Top             =   11640
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   2
      Left            =   4560
      Picture         =   "�ʧ@�C��.frx":25A0B7
      Top             =   11760
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   3
      Left            =   4440
      Picture         =   "�ʧ@�C��.frx":25A3B0
      Top             =   11880
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   4
      Left            =   4320
      Picture         =   "�ʧ@�C��.frx":25A68F
      Top             =   12000
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   5
      Left            =   4200
      Picture         =   "�ʧ@�C��.frx":25A988
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
      Left            =   3360
      Shape           =   2  '����
      Top             =   7320
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  '���
      Height          =   255
      Index           =   0
      Left            =   10680
      Shape           =   2  '����
      Top             =   7680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  '���
      Height          =   255
      Index           =   0
      Left            =   5640
      Shape           =   2  '����
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      Height          =   255
      Index           =   0
      Left            =   3360
      Shape           =   2  '����
      Top             =   7440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      Visible         =   0   'False
      X1              =   248
      X2              =   128
      Y1              =   752
      Y2              =   808
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      Visible         =   0   'False
      X1              =   8
      X2              =   128
      Y1              =   752
      Y2              =   808
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   5
      Visible         =   0   'False
      X1              =   128
      X2              =   248
      Y1              =   728
      Y2              =   784
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      Visible         =   0   'False
      X1              =   8
      X2              =   128
      Y1              =   784
      Y2              =   728
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      Visible         =   0   'False
      X1              =   6
      X2              =   254
      Y1              =   784
      Y2              =   784
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      Visible         =   0   'False
      X1              =   8
      X2              =   248
      Y1              =   752
      Y2              =   752
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   -120
      Shape           =   2  '����
      Top             =   10920
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Shape Shape15 
      FillStyle       =   0  '���
      Height          =   135
      Index           =   0
      Left            =   12240
      Shape           =   2  '����
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   0
      Visible         =   0   'False
      X1              =   400
      X2              =   416
      Y1              =   480
      Y2              =   488
   End
   Begin VB.Shape Shape18 
      BorderColor     =   &H0080FFFF&
      BorderStyle     =   0  '�z��
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  '���
      Height          =   255
      Index           =   0
      Left            =   5040
      Top             =   7200
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Shape Shape12 
      FillColor       =   &H00400040&
      FillStyle       =   0  '���
      Height          =   1215
      Left            =   7200
      Shape           =   2  '����
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape delay 
      BorderColor     =   &H80000005&
      FillColor       =   &H0080FF80&
      FillStyle       =   0  '���
      Height          =   315
      Index           =   3
      Left            =   9960
      Top             =   1080
      Width           =   240
   End
   Begin VB.Shape delay 
      BorderColor     =   &H80000005&
      FillColor       =   &H0080FF80&
      FillStyle       =   0  '���
      Height          =   315
      Index           =   2
      Left            =   9960
      Top             =   600
      Width           =   240
   End
   Begin VB.Image Image24 
      BorderStyle     =   1  '��u�T�w
      Height          =   3705
      Index           =   3
      Left            =   15000
      Picture         =   "�ʧ@�C��.frx":25AC81
      Top             =   0
      Visible         =   0   'False
      Width           =   2520
   End
   Begin VB.Image Image24 
      BorderStyle     =   1  '��u�T�w
      Height          =   3705
      Index           =   2
      Left            =   14880
      Picture         =   "�ʧ@�C��.frx":25E1F9
      Top             =   120
      Visible         =   0   'False
      Width           =   2520
   End
   Begin VB.Image Image27 
      BorderStyle     =   1  '��u�T�w
      Height          =   2835
      Index           =   1
      Left            =   -1800
      Picture         =   "�ʧ@�C��.frx":261747
      Top             =   8160
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Image Image27 
      BorderStyle     =   1  '��u�T�w
      Height          =   2835
      Index           =   0
      Left            =   -1680
      Picture         =   "�ʧ@�C��.frx":263BE1
      Top             =   8040
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Image Image29 
      BorderStyle     =   1  '��u�T�w
      Height          =   1065
      Index           =   1
      Left            =   1560
      Picture         =   "�ʧ@�C��.frx":26607E
      Top             =   8280
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Image Image29 
      BorderStyle     =   1  '��u�T�w
      Height          =   1065
      Index           =   0
      Left            =   1680
      Picture         =   "�ʧ@�C��.frx":267636
      Top             =   8160
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ���۳Ъ��Ƶ{�����{�������N�����Ƶ{�����p�ɾ�----�p��^^v

Dim s(48) As Integer '0)1P���D�p�� 1)�a�ϭp�� 2)1P����^���p�� 3)���W�d�� 4)GO���{�{ 5)��ܶˮ` 6)�� 7)�O�_����������� 8)�l�u���ͼ� 9)��������p�� 10)�y�P���u�B���ͭp�� 11)�}�]�s�]�����P�_ 12)�}�]�s�]�����ɶ� 13)�|��ɭp�� 14)�}�]�s�]���p�� 15) �y�P���u�B�L�k����ɤ��p�� 16)���d�P�_1-X 17)�ݤ뺧�� 18)�ݤ뺧������� 19)�W���ʤ��_�i���� 20)������𵲧��p�� 21)�s���ư{�t 22)�W���ʤ��}�]�s�]�D�P�B�o�g 23)���]���Ӳ����p�� 24)�ݤ몺�ˮ`�P�_ 25)���W���𵲧��P�_ 26)���W�����P�_ 27)�ɹp�����p�� 28)���W����ˮ` 29)�ɹp����D�P�B�X�{ 30)�ɹp�����P�_ 31)�ɹp���𲾰��p�� 32)�W���ʫD���𲾰��p�� 33)��������@�� 34)�ɹp���ͭp�� 35)2P���D�p�� 36)2P����^���p�� 37)���𲾰��p�� 38)�l�u����=0�ɪ��ѨM���  39)�Ps(38)  40)�Ps(11)  41)�Ps(12) 42)STAGE�X�{��m 43)�]�����y�P���u�B�Ұʭp�� 44)�]�����y�P�u�B�����p�� 45)���ۤw�Χ��i�~��ϥ� 46)�W�t���Ȩϥάy�P���u�B���a�ϲ��� 47)�¹� 48)�·����k�_��
Dim zxcv(13) As Integer '0)1P���D�ɤ��i�W�U���� 1)1P����𤣥i���� 2)'1P�y�P���u�B�����k�ʵe���� 3)1P�P�_���⩹���Ω��k 4)����u��Ū�@�������{�� 5)1P���Ȿ���ɤ������L 6)1P���D���� 7)��GO��S(1)�ܹL����A�p�� 8)'2P�y�P���u�B�����k�ʵe���� 9)2P�P�_���⩹���Ω��k 10)2P���Ȿ���ɤ������L 11)2P���𤣯���� 12)2P���D�ɤ��i�W�U���� 13)2P���D����

Dim ppm() As Integer '0)�� 1)�w���} 2)���{�{ 4)���ʵe 5)����P�_ 6)1P�p��ɶ� �ɶ���hcall�l�u���ͺt��k 7)�s����8)�s�����k0 PK�� 9)1P���� 10)2P���� 11)2P�p��ɶ� �ɶ���hcall�l�u���ͺt��k 12)�l��AI���s��s�ؼФ����� 13)�l��Ai���W�ߧP�_ 14)�O���Ǫ������� 15)�Ǫ��������ɶ� 16)�Ǫ��������ɶ�2(��) 17)�ǥ��X�U����]�� 18)����H���᪺�k�]delay 19)���`���A�x�s 20)�Ǫ����R
Dim hp_GPR() As Single '�Ȧs�t���p�ƪ��Ǫ���q

Dim stat() As Integer '�y�P���u�B���W�U�P�_�}�C
Dim vnn() As Integer '����������������w
Dim hall() As Integer '�}�]�s�]���k�P�_0)���k 1)�W�U
Dim stong() As Integer '�z�𪬺A(�C�ӵ��۳��W�ߡA���|�]���Y�ӵ��۰����z��Ө䥦��۰���)
Dim sup_stong() As Integer '�ϥΥ|���(����)
Dim powstong() As Integer '�W�z��
Dim xy() As Integer '����V�]���k�^
Dim xy2() As Integer '����V�]�W�U�^

Dim fireup(1, 10) As Integer '���W���𤧤ϼu�P�_ 0)�W�U 1)���k
Dim firr(4, 10) As Integer ' 0)���W���𤧬O�_�����k���� 1)�P�_�O�_�����ؼ� 2)�P�_�O���ӥؼ� 3)�P�_���Ǥ���w����(�B�⦡) 4)�P�_���Ǥ��]�w����
Dim ball(1, 300) As Integer '0�l�u��V�x�s 1����w����
Dim slow(5, 0) As Integer 'delay���}�C 0~3 delay 4)1P HP�{�t�p�� 5) 2P HP�{�{�p��
Dim play(2, 1) As Integer '�᭱0)1P 1)2P �e�� 0)�����u����@�� 1)����G�� 2)������j������p��
Dim down(1) As Integer '0)1P 1)2P
Dim p_which(1, 10, 10) As Integer '���@�Ӳɤl�� ( 0)�ڤ� 1)�Ĥ�,�ĴX�ӤH��,�ɤl��)

Dim aop(0, 1) As Single '0)���⪺�R
Dim khp_GPR(1) As Single '�Ȧs�t���p�ƪ��H��Hp
Dim kmp_GPR(1) As Single '�Ȧs�t���p�ƪ��H��Mp
Dim P_Combo(1) As Integer '�s����

Dim sic(10) As Integer '�|��ɮ���
Dim X(1) As Integer '���D��V
Dim k(3) As Integer '�ʵe1�㢲)�H���ʵe
Dim sk(10) '�|������
Dim sh(1) As Integer '�v�l�Z��
Dim ax(3) As Integer '�]�B�P�_
Dim run(3) As Integer '�]�B�W�q��
Dim dzxc(1) As Integer '�y�P���u�B���W�U�P�_
Dim ell(0) As Integer '0) �y�P���u�B(���t)�M�w���k�Υ�����
Dim up(1) As Integer 'Ĳ�o�ĤG���ۤ��@����
Dim ven(15) As Integer '1~15 �� ���~�P�����I �Ĥ@���H�������ۤG
Dim cir(10) As Integer '�W���ʤ���j�t��
Dim ys(8) As Integer '�]��ˮ`�@�P��
Dim metor(1) As Integer '�y�P���u�B�����ʭ���A���ಾ��
Dim met_already(16) As Integer 'PK_mod �y�P���u�B�w����
Dim kb As Integer '���h��
Dim boss01(16) As Integer '�]��1���y�P���u�B���W�U�P�_

Dim kik_ell As Integer '�ܱ𪺷N�Ӥ����k����
Dim yst As Integer '������������O�Ȧs
Dim ok As Integer 'ok�����h�a�ϲ��ʬ����h�H������
Dim ma As Integer '�ƶq
Dim pps As Integer '�Q���������@��
Dim n As Integer '���}��
Dim Y As Integer '�ˮ`��
Dim iss As Integer '�ʵe����
Dim ddd As Integer '�����ƭ�
Dim dx As Integer '�I�������ʶq
Dim gdown As Integer '�Y�����T�w��m�Ȧsleft
Dim gup As Integer '�Y�����T�w��m�Ȧstop
Dim gbl As Integer '�Y���{���y�{����
Dim appshe1 As Integer, appshe2 As Integer '�o��Ӭ��W���ʪ��s�]��m���w
Dim agains As Integer '��ܶi��U�@���C��(��Ū���)
Dim supper As Integer '�}�]�s�]���o�g�q
Dim boss As Integer '�ĴX���]��
Dim boss_power As Integer '�]���i�ε��۪�����
Dim old_xy As Integer '�]�����y�P���u�B���ª���V
Dim ma_speed As Integer '�Ǫ������t��
Dim F8_second As Integer '�K��(����ĤH)
Dim ma_aop As Integer '�Ǫ����R

Const ppmn As Integer = 20 'ppm()���ŧi�ƶq
Const wmy As Integer = 3 '�ݤ뺧�����𤧫ᥴ�X���ƶq
Const holy As Integer = 10 '��������`�ƶq
Const fire_cylinder As Integer = 10 '���W���ƶq(��3 5 10)

Const p_figure As Integer = 100 '�ɤl�ƶq
Dim p(1, 10, 10, p_figure) As boom '�ɤl (�Ĥ�Χڤ�,�ĴX�ӤH��,�ĴX�Ӳɤl��,�ĴX�Ӳɤl)
Private Type boom '�ɤl�ݩ�
    px As Single '��m
    py As Single '��m
    old_px As Single
    old_py As Single
    old_color As Long
    pxs As Single '�[�t��
    pys As Single '�[�t��
    p_aop As Integer '�R
End Type

Private Sub Form_Unload(Cancel As Integer)
'Call sndPlaySound("Data\�L�n.wav", 1)
Call mciSendString("close Data\�I��.mid", vbNullString, 0, 0)
End Sub
Private Sub Timer26_Timer() '��������
Timer26.Enabled = False
DoEvents
Call desd(1)
End Sub
Private Sub ex(a) '�����{�ǡ�
If zxcv(4) = 1 Then Exit Sub
zxcv(4) = 1
If a = 1 Then '������
    Label8.Left = Me.ScaleWidth \ 2 - Label8.Width \ 2
    Label8.Top = Me.ScaleHeight \ 2 - Label8.Height \ 2
    Label8.Caption = "Game Over..."
    Label8.Visible = True: Shape1(0).Visible = False: Timer26.Enabled = True '�Ұʩ���
Else '��F2��
    Call desd(1)
End If
End Sub
Public Sub desd(a) '�����@��
On Error Resume Next

If pk_mod = 0 Then '�p�G���Opk�Ҧ��h
    For i = 1 To s(8) '�����l�u�t��
        If ball(1, i) <> 1 Then '�p�G���Ӫ���s�b�h
            Unload Shape7(i)
            Unload Shape8(i)
            Unload Timer19(i)
        End If
    Next
    For f = 1 To ma '���������t��
        If ppm(1, f) <> 1 Then '�p�G���Ӫ���s�b�h
            Unload Timer12(f)
            Unload Timer11(f)
            Unload Shape2(f)
            Unload Image6(f)
        End If
        Unload Timer22(f)
        Unload Timer43(f)
        Unload Timer44(f)
    Next
    For f = 1 To ppmn
        For ff = 1 To ma
            ppm(f, ff) = 0
        Next
    Next
End If
iss = 0 '����Ū���ɥ��`
s(45) = 0 '�M���]�����b�ϥΪ����۫���
n = 0 '�L���P�_
s(13) = 0
If a = 1 Then
    supper = 0: agains = 0: gbl = 0: dx = 0: gz = 1: music = 0: boss = 0: F8_second = 0: ma_aop = 0: kb = 0
    Erase fireup, firr, ys, hall, vnn, cir, slow, ven, ball, s, zxcv, xy, k, ax, run, ppm, s, stong, up, sup_stong, play, down, powstong, aop, hp_GPR, khp_GPR, kmp_GPR
    Unload Me
    Form2.Show
Else
    agains = 1
    Call Form_Load
End If
End Sub
Private Sub Form_Load() '�����J��
'����ʧ@�C���u��Ū"�@��"��Load�{���X
If reset(1) = 0 Then
    reset(1) = 1
    Form1.Width = 12000 '(800*600)
    Form1.Height = 9000
    
    Me.Caption = Program_Name
End If

If agains <> 1 Then
    ReDim vnn(holy * 2, 1) As Integer
    ReDim stong(7) As Integer
    ReDim sup_stong(7) As Integer
    ReDim powstong(7) As Integer
    ReDim hall(1, 1, 20) As Integer '�Ĥ@�ӤT���}�C 1,2,3
    ok = 0
    ma_aop = 1 '��l�R
    
    Form1.Left = Screen.Width \ 2 - Form1.Width \ 2
    Form1.Top = Screen.Height \ 2 - Form1.Height \ 2
    
    If player_2 = 1 Then Call p2 Else Call p2_hex '�p�G��2P�h �Ϥ� 2P�ݩ���
    Call P12(0) 'Ū���H�����ݩ���
End If


If s(1) > 0 And s(1) Mod 5 = 0 Then '���q���d�άO�]����
    'ma_aop = 5 '�R
    'boss_power = 3 '�]�����ۥi�ϥΪ�����
    kb = 100
    ma_speed = 10 '���ʳt��
    ma = 1: akiz(1) = 50: amax(1) = 10000
    For f = 0 To 1
        Label7(f).Caption = "�W�t����"
        xhp(f).Width = 200
        xhp(f).FillColor = &HFF80FF '�����ܵ���
        Shape6(f).FillColor = &HFFFF& '�����ܶ���
        Shape6(f).Width = xhp(f).Width
        Label5(f).Left = Shape6(f).Left + Shape6(f).Width \ 2 - Label5(f).Width \ 2
        Label5(f).Caption = amax(1) & "/" & amax(1)
        Label14(f).Caption = "X " & ma_aop '�R
        Label14(f).Left = Shape6(f).Left + Shape6(f).Width + 10 '�R����m
    Next
Else
    'ma_aop = (s(1) + 1) Mod 5 '�R
    kb = 30
    ma_speed = 8 '���ʳt��
    ma = 4 * (player_2 + 1): akiz(1) = 10: amax(1) = 2000 '�Ǫ��Ӽ�
    For f = 0 To 1
        Label7(f).Caption = "����"
        xhp(f).Width = 200
        Shape6(f).Width = xhp(f).Width
        Label5(f).Left = Shape6(f).Left + Shape6(f).Width \ 2 - Label5(f).Width \ 2
        Label14(f).Caption = "X " & ma_aop
        Label14(f).Left = Shape6(f).Left + Shape6(f).Width + 10
    Next
End If

ReDim xy(ma), xy2(ma), ppm(ppmn, ma), hp_GPR(ma)
For f = 0 To 0 + player_2
    If hp(f).Visible = False Then hp(f).Width = 100: khp_GPR(f) = hp(f).Width
    hp(f).Visible = True '�_��
    Image1(f).Visible = True '�_��
    Shape1(f).Visible = True
    zxcv(5 + 5 * f) = 0 '�_��
    metor(f) = 1
    
    hp(f).FillColor = &HFFFF&
    label1(f).Caption = hp(f).Width * akiz(0) & "/" & amax(0)
    label2(f).Caption = mp(f).Width * akiz(0) & "/" & amax(0)
    label1(f).Left = shape4(f).Left + shape4(f).Width \ 2 - label1(f).Width \ 2
    label2(f).Left = shape4(f).Left + shape4(f).Width \ 2 - label2(f).Width \ 2
    
    aop(0, f) = aop(0, f) + 1
    Label13(f).Caption = "X " & aop(0, f)
    label10(ff).Visible = True: image23(ff).Visible = True: label10(f).Caption = Val(label10(f).Caption) + 1 & " X"
Next

If pk_mod = 0 Then
    Call nnee '���ͺt��k
Else
    For f = 1 To 2
        Load Timer22(f)
        Load Timer43(f)
    Next
End If

If music = 0 Then
    music = 1
    'Call sndPlaySound("Data\�I��.wav", 9)
    Call mciSendString("play Data\�I��.mid", vbNullString, 0, 0)
    
    '�_������
        'Set dsz = dxa.DirectSoundCreate("")
        'dsz.SetCooperativeLevel Me.hWnd, DSSCL_NORMAL
        'bu.lFlags = DSBCAPS_STATIC
        'Set dsbz = dsz.CreateSoundBufferFromFile(("�I��.mid"), bu, wf)
        'dsbz.play DSBPLAY_LOOPING '''''''''''''''''''''''�w�]����(�榸).
    '�_������
End If
End Sub
Public Sub nnee() '���ͺt��k��
For f = 1 To ma
    Randomize
    
    Load Image6(f)
    Load Timer44(f) '�ɤl
    Image6(f).Top = Int(Rnd * (Form1.ScaleHeight - Image6(f).Height - Line1.Y1 - 10)) + Line1.Y1
    Image6(f).Left = Int(Rnd * (Form1.ScaleWidth * 2 - Form1.ScaleWidth)) + Form1.ScaleWidth
    
    Load Shape2(f)
    ppm(20, f) = ma_aop '�R
    If ma = 1 Then '�u���]���h
        Image6(1).Picture = Image16(0).Picture '����
        Shape2(f).Width = Image6(f).Width - Image6(f).Width \ 2 '�v�l�j�p
        Shape2(f).Height = Image6(f).Height \ 6
        Shape2(f).Top = Image6(f).Top + Image6(f).Height
        Shape2(f).Left = Image6(f).Left
    Else
        Shape2(f).Top = Image6(f).Top + Image6(f).Height + 1
        Shape2(f).Left = Image6(f).Left + Image6(f).Width \ 2
    End If
    
    Image6(f).Visible = True
    Shape2(f).Visible = True
    Shape2(f).ZOrder 1
    
    Load Timer12(f)
    Timer12(f).Enabled = True '����
Next
Call isss
End Sub
Public Sub isss()
For f = 1 To ma
    If f Mod 8 = 0 Then iss = iss + 1
    ppm(4, f) = f - 8 * iss
    
    xy(f) = 1
    
    a = Int(Rnd * 2) + 1
    If a = 2 Then xy2(f) = 1 Else xy2(f) = -1
    
    Load Timer22(f)
    Load Timer43(f)
    Load Timer11(f)
    Timer11(f).Enabled = True '���ʵe
Next
End Sub
Private Sub p2() '2P�ݩ���
Load Image1(1)
Load Shape1(1)
Image1(1).Left = 0
Image1(1).Top = Form1.ScaleHeight \ 2 - Image1(1).Height \ 2
Image1(1).Visible = True
Shape1(1).Visible = True
Image1(1).Left = Form1.ScaleWidth \ 2 - Image1(1).Width \ 2

Load Timer1(1): Load Timer2(1): Load Timer3(1): Load Timer4(1)
Load Timer9(2): Load Timer9(3)
Load Timer8(1)
Load Timer10(1)
Load Timer45(1) '�ɤl�p�ɾ�
End Sub
Private Sub p2_hex() '2P�ݩ�����ܩ�����
'2P���ݩ���
If Label13(1).Visible = True Then Label13(1).Visible = False Else Label13(1).Visible = True
If label1(1).Visible = True Then label1(1).Visible = False Else label1(1).Visible = True
If label2(1).Visible = True Then label2(1).Visible = False Else label2(1).Visible = True
If hp(1).Visible = True Then hp(1).Visible = False Else hp(1).Visible = True
If mp(1).Visible = True Then mp(1).Visible = False Else mp(1).Visible = True
If shape4(1).Visible = True Then shape4(1).Visible = False Else shape4(1).Visible = True
If shape5(1).Visible = True Then shape5(1).Visible = False Else shape5(1).Visible = True
If label12(1).Visible = True Then label12(1).Visible = False Else label12(1).Visible = True
If label9(1).Visible = True Then label9(1).Visible = False Else label9(1).Visible = True
For f = 2 To 3
    If delay(f).Visible = True Then delay(f).Visible = False Else delay(f).Visible = True
    If Label6(f).Visible = True Then Label6(f).Visible = False Else Label6(f).Visible = True
    If Timer21(f).Enabled = True Then Timer21(f).Enabled = False Else Timer21(f).Enabled = True
Next
'/2P���ݩ���
End Sub
Private Sub P12(ByVal a As Integer) '�H�����ݩ���
For f = 0 + a To 0 + player_2
    hp(f).Width = 200
    mp(f).Width = 200
    khp_GPR(f) = hp(f).Width
    kmp_GPR(f) = mp(f).Width
    aop(0, f) = 4
    Timer5(kikyou(f)).Enabled = True
    Select Case kikyou(f)
        Case 0
            Image1(f).Picture = Image2(0).Picture
            sh(f) = 70
            Label6(0 + 2 * f).Caption = "�������"
            Label6(1 + 2 * f).Caption = "�ܱ𪺷N��"
            delay(0 + 2 * f).Width = Label6(0 + 2 * f).Width
            delay(1 + 2 * f).Width = Label6(1 + 2 * f).Width
            Timer20(0 + 2 * f).Interval = 250 - 100
            Timer20(1 + 2 * f).Interval = 250 - 100
        Case 1
            Image1(f).Picture = Image3(0).Picture
            sh(f) = 53
            Label6(0 + 2 * f).Caption = "���t���W"
            Label6(1 + 2 * f).Caption = "�Y��"
            
            Timer20(0 + 2 * f).Interval = 250 - 100
            Timer20(1 + 2 * f).Interval = 250 - 100
        Case 2
            Image1(f).Picture = Image4(0).Picture
            sh(f) = 70
            Label6(0 + 2 * f).Caption = "�����ɹp"
            Label6(1 + 2 * f).Caption = "�ݤ뺧��"
            Timer20(0 + 2 * f).Interval = 250 - 100
            Timer20(1 + 2 * f).Interval = 250 - 100
        Case 3
            Image1(f).Picture = Image5(0).Picture
            sh(f) = 60
            Label6(0 + 2 * f).Caption = "�y�P���u�B"
            Label6(1 + 2 * f).Caption = "�W����"
            Timer20(0 + 2 * f).Interval = 250 - 100
            Timer20(1 + 2 * f).Interval = 250 - 100
    End Select
    
    delay(0 + 2 * f).Width = Label6(0 + 2 * f).Width
    delay(1 + 2 * f).Width = Label6(1 + 2 * f).Width
    
    Shape1(f).Top = Image1(f).Top + sh(f)
    Shape1(f).Left = Image1(f).Left + Image1(f).Width \ 2 - 20
    
    label10(f).Visible = True: image23(f).Visible = True
    label10(f).Caption = Val(label10(f).Caption) + 2 & " X"
    label2(f).Caption = mp(f).Width * akiz(0) & "/" & amax(0)
    
    X(f) = 1
    metor(f) = 1
    run(0 + 2 * f) = 1
    run(1 + 2 * f) = 1
Next
End Sub
Public Sub spdd(ByVal a As Integer) '�l�u���ͺt��k
If s(8) >= 300 Then s(8) = 0
s(8) = s(8) + 1
Load Timer19(s(8)) '���ͤl�u����
Load Shape7(s(8)) '���ͤl�u
Load Shape8(s(8)) '���ͤl�u�v�l

ball(0, s(8)) = xy(a) '�s�J�o�g�l�u������������V

If ball(0, s(8)) = 1 Then Shape7(s(8)).Left = Image6(a).Left - Shape7(s(8)).Width \ 2 + Image6(a).Width \ 2 Else Shape7(s(8)).Left = Image6(a).Left + Image6(a).Width \ 2 - Shape7(s(8)).Width \ 2 '���w�l�u��l��m
Shape7(s(8)).Top = Image6(a).Top + Image6(a).Height \ 2 '���w�l�u��l��m
Shape8(s(8)).Left = Shape7(s(8)).Left '���w�v�l��l��m
Shape8(s(8)).Top = Shape7(s(8)).Top + 32 '���w�v�l��l��m
Shape7(s(8)).Visible = True
Shape8(s(8)).Visible = True

Timer19(s(8)).Enabled = True '�Ұʤl�u����
End Sub
Private Sub Timer19_Timer(Index As Integer) '�ĤH�l�u����
    If Timer6.Enabled = True Then Shape7(Index).Left = Shape7(Index).Left - 5 '�a�ϲ���
    
    Shape7(Index).Left = Shape7(Index).Left - 10 * ball(0, Index) '����
    Shape8(Index).Left = Shape7(Index).Left '�v�l
    
    On Error Resume Next
    
    For f = 0 To 0 + player_2
        If Timer8(f).Enabled = True Then '�D����
            If jump(Shape7(Index), (f)) Then
                If coll2(Shape7(Index), Image1(f)) Then
                    If coll(Shape7(Index), Image1(f)) Then
                        Call kizzs((f))
                        If hp(f).Visible = True Then ball(1, Index) = 1: Unload Timer19(Index): Unload Shape7(Index): Unload Shape8(Index): Exit Sub '�p�G�����h����l��
                    End If
                End If
            End If
        Else
            If coll2(Shape7(Index), Image1(f)) Then
                If coll(Shape7(Index), Image1(f)) Then
                    Call kizzs((f))
                    
                    If hp(f).Visible = True Then ball(1, Index) = 1: Unload Timer19(Index): Unload Shape7(Index): Unload Shape8(Index): Exit Sub '�p�G�����h����l��
                End If
            End If
        End If
    Next
    
    If gz = 1 Then Exit Sub: Timer19(Index).Enabled = False
    If con1(Shape7(Index)) Or con2(Shape7(Index)) Then '�p�G�W�X��ɫh����
        ball(1, Index) = 1
        Unload Timer19(Index)
        Unload Shape7(Index)
        Unload Shape8(Index)
    End If
    DoEvents
End Sub
Private Sub Timer42_Timer(Index As Integer) '�]�����y�P���u�B
Dim a As Integer

Image25(Index).Left = Image25(Index).Left - 16 * old_xy
Image25(Index).Top = Image25(Index).Top - 2 * boss01(Index)

For f = 0 To player_2
    'If coll(Image25(Index), Image1(f)) Or coll(Image1(f), Image25(Index)) Then
    If coll(Image25(Index), Image1(f)) Then
        Call timer42_L2(Index)
        Call kizzs((f))
        Exit Sub
    End If
Next

If Image25(Index).Left + Image25(Index).Width < 0 Or Image25(Index).Left > Form1.ScaleWidth Then Call timer42_L2(Index)

DoEvents
End Sub
Private Sub timer42_L2(Index)
Unload Image25(Index): Unload Timer42(Index)
s(44) = s(44) + 1
If s(44) = 16 Then s(44) = 0: s(45) = 0: s(46) = 0: Timer6.Enabled = False
End Sub
Private Sub Timer41_Timer(Index As Integer) '�]������
Select Case Index
    Case 0 '�y�P���u�B
        On Error Resume Next
        
        s(43) = s(43) + 1
        Image25(s(43)).Left = Image6(1).Left + Image25(s(43)).Width * ((old_xy - 2) Mod 3) * -1
        Image25(s(43)).Top = Int(Rnd * (Image6(1).Height + 1)) + Image6(1).Top
        
        Randomize
        For f = 0 To player_2
            If Image1(f).Top + Image1(f).Height \ 2 <= Image25(s(43)).Top + Image25(s(43)).Height \ 2 Then
                boss01(s(43)) = 1
            Else
                boss01(s(43)) = -1
            End If
            'If Int(Rnd * 2) = 1 Then boss01(s(43)) = 1 Else boss01(s(43)) = -1
        Next
        Image25(s(43)).Visible = True
        Timer42(s(43)).Enabled = True
        If s(43) = 16 Then Timer41(Index).Enabled = False: s(43) = 0
    Case 1 '�u�����\
        
    Case 2 '������
        
    Case 3 '�B�۶üY
        
    Case 4 '�g��ŧ
        
    Case 5 '��j��
        
    Case 6 '�B�ʲy
        
    Case 7 '��������

End Select
DoEvents
End Sub
Private Sub mpss2(a As Integer) '�]����MP�l�Ӻt��k
Select Case a
    Case 0 '�y�P���u�B
        For f = 1 To 16
            Load Image25(f)
            Image25(f).Picture = Image26((old_xy + 2) Mod 3).Picture
            Load Timer42(f)
        Next
        Timer41(a).Enabled = True
    Case 1 '�u�����\
        
    Case 2 '������
        
    Case 3 '�B�۶üY
        
    Case 4 '�g��ŧ
        
    Case 5 '��j��
        
    Case 6 '�B�ʲy
        
    Case 7 '��������

End Select
DoEvents
End Sub
Private Sub Timer12_Timer(Index As Integer) '���ʡ�
Dim a As Integer, b As Integer, c(8) As Integer

    If Timer6.Enabled = True And s(45) = 0 Then Image6(Index).Left = Image6(Index).Left - 10 * run(1): Shape2(Index).Left = Shape2(Index).Left - 10 * run(1)
    
    If ppm(5, Index) = 0 Then '����
    
        If s(45) = 1 And s(1) Mod 5 = 0 Then '�p�G�ΤF�y�P���u�B�h
            Image6(Index).Left = Image6(Index).Left + 1 * old_xy
            Shape2(Index).Left = Shape2(Index).Left + 1 * old_xy
            If Image6(Index).Left < 0 Then Image6(Index).Left = 0: Shape2(Index).Left = Image6(Index).Left + Image6(Index).Width - Shape2(Index).Width
            
            'If Image6(index).Left + Image6(index).Width * 1.5 > Form1.ScaleWidth Then
            '    If Image6(index).Left + Image6(index).Width > Form1.ScaleWidth Then Image6(index).Left = Form1.ScaleWidth - Image6(index).Width
            '    s(46) = 1
            '    Timer6.Enabled = True
            'Else
            '    Timer6.Enabled = False
            'End If
            
            If Image6(Index).Left + Image6(Index).Width > Form1.ScaleWidth - Image6(Index).Width \ 3 Then
                If Image6(Index).Left + Image6(Index).Width >= Form1.ScaleWidth Then Image6(Index).Left = Form1.ScaleWidth - Image6(Index).Width: Shape2(Index).Left = Image6(Index).Left
                If Image1(0).Left <= 10 Or Image1(player_2).Left <= 10 Then
                    Timer6.Enabled = False
                Else
                    For f = 0 To player_2
                        Image1(f).Left = Image1(f).Left - 1
                        Shape1(f).Left = Shape1(f).Left - 1
                        If Image1(f).Left <= 10 Then Image1(f).Left = 10
                    Next
                    s(46) = 1: Timer6.Enabled = True
                End If
            Else
                Timer6.Enabled = False
            End If
            Exit Sub
        End If
        
        ppm(12, Index) = ppm(12, Index) + 1
        If ppm(12, Index) = 50 Then
            ppm(12, Index) = 0
            
            ppm(13, Index) = Int(Rnd * (player_2 + 1)) '�H���P�_
            If hp(ppm(13, Index)).Visible = False Then ppm(13, Index) = ppm(13, Index) * -1 + 1
            
            '��̪�
            'c = ((Image6(Index).Left - Image1(0).Left) ^ 2 + (Image6(Index).Top - Image1(0).Top) ^ 2) ^ 0.5
            'd = ((Image6(Index).Left - Image1(player_2).Left) ^ 2 + ((Image6(Index).Top - Image1(player_2).Top) ^ 2)) ^ 0.5
            'If c > d Then a = 1 Else a = 0
        End If
        If hp(ppm(13, Index)).Visible = False Then Exit Sub
        
        ppm(15, Index) = ppm(15, Index) + 1
        If ppm(15, Index) = 10 Then
            ppm(15, Index) = 0
            ppm(14, Index) = Artificial_Intelligence(Image6(Index), Image1(ppm(13, Index)))
            If ppm(17, Index) >= 1 Then '�p�G�w����H���h
                ppm(18, Index) = ppm(18, Index) + 1 'delay
                If ppm(18, Index) >= 5 Then ppm(18, Index) = 0: ppm(17, Index) = 0 '�l�H
            End If
        End If
        
        Select Case ppm(14, Index)
            Case 1 '��
                a = -1: b = 0
            Case 2 '�k
                a = 1: b = 0
            Case 4 '�W
                a = 0: b = -1
            Case 5 '���W
                a = -1: b = -1
            Case 6 '�k�W
                a = 1: b = -1
            Case 7 '�U
                a = 0: b = 1
            Case 8 '���U
                a = -1: b = 1
            Case 9 '�k�U
                a = 1: b = 1
        End Select
        If a <> 0 Then s(38) = a
        If b <> 0 Then s(39) = b
        If ppm(17, Index) >= 1 Then a = -a: b = -b
        
        Image6(Index).Left = Image6(Index).Left + ma_speed * a: xy(Index) = -1 * s(38) '�k
        Image6(Index).Top = Image6(Index).Top + ma_speed * b: xy2(Index) = -1 * s(39): Shape2(Index).Top = Shape2(Index).Top + ma_speed * b '�U
        
        
        'If coll(Image6(Index), Image1(ppm(13, Index))) = False And coll(Image1(ppm(13, Index)), Image6(Index)) = False Then '�l�u
        If coll(Image6(Index), Image1(ppm(13, Index))) = False Then '�l�u
            For f = 0 To 0 + player_2
                If hp(f).Visible = True Then
                    If ppm(5, Index) = 0 Then '�S���Q�ɹp����ɰ����l�u�o�g����
                        If coll2(Image6(Index), Image1(f)) Or coll2(Image1(f), Image6(Index)) Then
                            If ma = 1 And s(45) = 0 Then '�p�G�O�]�����h
                                e = Int(Rnd * 10) + 1 '�@�w���v�ε���
                                If e = 5 And Image6(1).Left + Image6(1).Width < Form1.ScaleWidth And xhp(0).Width <= 200 \ 2 Then
                                    'If boss_power > 0 Then '�p�G�]���������٨S�Χ��h
                                        s(45) = 1
                                        old_xy = xy(Index)
                                        'boss_power = boss_power - 1 '�i�ϥΪ����۴�֤@��
                                        Call mpss2(0) '�I�s�]��MP
                                    'End If
                                End If
                            End If
                            ppm(6 + 5 * f, Index) = ppm(6 + 5 * f, Index) + 1
                            If ppm(6 + 5 * f, Index) Mod 15 = 0 Then
                                ppm(6 + 5 * f, Index) = 0: Call spdd(Index) '�I�s���ͤl�u�{��
                            End If
                        Else
                            ppm(6 + 5 * f, Index) = 0
                        End If
                    End If
                End If
                DoEvents
            Next
        Else '�񨭥\��
            'Randomize '�����t��k
            'd = Int(Rnd * 3) + 1
            'If d = 1 Then
                ppm(16, Index) = ppm(16, Index) + 1
                If ppm(16, Index) >= 10 Or ma = 1 Then
                    ppm(16, Index) = 0
                    c(Index) = Image6(Index).Left
                    For f = c(Index) To c(Index) - Image6(Index).Width * xy(Index) Step -5 * xy(Index)
                        If ppm(1, Index) = 1 Then Exit Sub
                        Image6(Index).Left = f
                        DoEvents
                    Next
                    Call kizzs(ppm(13, Index))
                    Image6(Index).Left = c(Index)
                    ppm(17, Index) = ppm(17, Index) + 1 '�k
                End If
            'End If
        End If '/�����t��k
    End If
    
    On Error Resume Next
    If ma = 1 Then '�u���]���h
        If xy(Index) = 1 Then
            Shape2(Index).Left = Image6(Index).Left
        Else
            Shape2(Index).Left = Image6(Index).Left + Image6(Index).Width \ 2
        End If
    Else
        If xy(Index) = 1 Then
            Shape2(Index).Left = Image6(Index).Left + Image6(Index).Width \ 2
        Else
            Shape2(Index).Left = (Image6(Index).Left + Image6(Index).Width \ 2) - 20
        End If
    End If
    
    DoEvents
End Sub
Private Sub Timer11_Timer(Index As Integer) '���ʵe��
If s(1) = 0 Or s(1) Mod 5 <> 0 Then
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
Else
    If xy(Index) = 1 Then
        Image6(Index).Picture = Image16(0).Picture
    Else
        Image6(Index).Picture = Image16(1).Picture
    End If
End If
End Sub
Private Sub Timer40_Timer(Index As Integer) 'HP�{�{
Timer40(Index).Interval = 300
slow(4 + Index, 0) = slow(4 + Index, 0) + 1
If slow(4 + Index, 0) Mod 2 = 0 Then hp(Index).FillColor = &HFF& Else hp(Index).FillColor = &HFFFF&
DoEvents
End Sub
Private Sub Timer7_Timer(Index As Integer) '�������ʵe��
ppm(2, Index) = ppm(2, Index) + 1
If ppm(2, Index) Mod 2 = 0 Then
    Image6(Index).Visible = True
    If ppm(2, Index) >= 12 Then
    
        '�|��ɱ���
        Call sico(Index)
        
        Unload Shape2(Index)
        Unload Image6(Index)
        Unload Timer11(Index)
        Unload Timer12(Index)
        Unload Timer7(Index)
        ppm(2, Index) = 0
        n = n + 1
        If n = ma Then ok = 0: Timer15.Enabled = True '���GO�U�@��
        Exit Sub
    End If
Else
    Image6(Index).Visible = False
End If

If Timer6.Enabled = True Then '�p�G�I�����ʫh�Q���}�����]�@�_����
    Image6(Index).Left = Image6(Index).Left - 20 * run(1)
    If xy(Index) = 1 Then
        Shape2(Index).Left = Image6(Index).Left + Image6(Index).Width \ 2
    Else
        Shape2(Index).Left = (Image6(Index).Left + Image6(Index).Width \ 2) - 20
    End If
End If
DoEvents
End Sub
Private Sub Timer46_Timer() '�¹�
s(47) = (s(47) + 1) Mod 2
If s(47) = 1 Then
    Picture1.Picture = Me.Picture
    Me.Picture = LoadPicture()
Else
    Me.Picture = Picture1.Picture
    Call BackPicture_Move '�٭�ثeImage�ݩʭI��
    Timer46.Enabled = False
End If
End Sub
Private Sub p_start(a As Integer, b As Integer, c, e As Object) '�ɤl��l�� (�H�ΩǪ�,���@��,����)

Timer46.Enabled = True '�Ұʶ¹�

Randomize

For f = 1 To p_figure
    p(a, b, c, f).px = e(b).Left + e(b).Width \ 2
    p(a, b, c, f).py = e(b).Top + e(b).Height \ 2
    p(a, b, c, f).pxs = Rnd * 50 - 50 / 2
    p(a, b, c, f).pys = Rnd * 50 - 50 / 2
    p(a, b, c, f).p_aop = 20
Next

If a = 1 Then Timer44(b).Enabled = True Else Timer45(b).Enabled = True

End Sub
Private Sub Timer44_Timer(Index As Integer) '�Ǫ��ɤl����
Call messenger(1, Index)
DoEvents
End Sub
Private Sub Timer45_Timer(Index As Integer) '�H���ɤl����
Call messenger(0, Index)
DoEvents
End Sub
Private Sub messenger(a As Integer, Index As Integer) '�ɤl�֤�
Dim d As Integer
Form1.AutoRedraw = False
For i = 1 To 10
    If p_which(a, Index, i) = 1 Then
        For f = 1 To p_figure '�ɤl��
            Form1.PSet (p(a, Index, i, f).old_px, p(a, Index, i, f).old_py), p(a, Index, i, f).old_color '�M���W�Ӥw�e���ɤl
            p(a, Index, i, f).old_color = Point(p(a, Index, i, f).px, p(a, Index, i, f).py) '�b�e�U�h���e�����o�n�e�������I���C���
            Form1.PSet (p(a, Index, i, f).px, p(a, Index, i, f).py), RGB(255, 255, 255) '�e�X�ɤl
            p(a, Index, i, f).old_px = p(a, Index, i, f).px '���o�w�e��X����
            p(a, Index, i, f).old_py = p(a, Index, i, f).py '���o�w�e��Y����
            p(a, Index, i, f).px = p(a, Index, i, f).px + p(a, Index, i, f).pxs
            p(a, Index, i, f).py = p(a, Index, i, f).py + p(a, Index, i, f).pys
            p(a, Index, i, f).p_aop = p(a, Index, i, f).p_aop - 1
            If p(a, Index, i, f).p_aop <= 0 Then
                Form1.PSet (p(a, Index, i, f).old_px, p(a, Index, i, f).old_py), p(a, Index, i, f).old_color '�M���W�Ӥw�e���ɤl
                p_which(a, Index, i) = 0
            End If
        Next
    Else
        d = d + 1
        If d = 10 Then
            If a = 0 Then Timer45(Index).Enabled = False Else Timer44(Index).Enabled = False
        End If
    End If
Next
Form1.AutoRedraw = True
DoEvents
End Sub
Private Sub Timer22_Timer(Index As Integer) '������h
If pk_mod = 0 Then
    If ppm(1, Index) <> 1 Then
        ppm(5, Index) = 0
        If ppm(19, Index) = 1 Then ppm(19, Index) = 0: Timer43(Index).Enabled = False '���`���A�Ѱ�
        Timer11(Index).Enabled = True '�ҰʳQ����Ӱ���S��
    End If
Else
    metor(Index - 1) = 1
    Timer5(kikyou(Index - 1)).Enabled = True
End If
Timer22(Index).Enabled = False '����ۤv
End Sub
Private Sub sup_st(f As Integer, Index As Integer) '�|��ɦ�m�@��
Select Case f
    Case 1 + 4 * Index
        Line6(1 + 4 * Index).X1 = Image1(Index).Left - Image1(Index).Width
        Line6(1 + 4 * Index).X2 = Image1(Index).Left - Image1(Index).Width \ 2
        Line6(1 + 4 * Index).Y1 = Image1(Index).Top - Image1(Index).Height
        Line6(1 + 4 * Index).Y2 = Image1(Index).Top - Image1(Index).Height \ 2
    Case 2 + 4 * Index
        Line6(2 + 4 * Index).X1 = Image1(Index).Left + Image1(Index).Width * 2
        Line6(2 + 4 * Index).X2 = Image1(Index).Left + Image1(Index).Width * 1.5
        Line6(2 + 4 * Index).Y1 = Image1(Index).Top - Image1(Index).Height
        Line6(2 + 4 * Index).Y2 = Image1(Index).Top - Image1(Index).Height \ 2
    Case 3 + 4 * Index
        Line6(3 + 4 * Index).X1 = Image1(Index).Left - Image1(Index).Width
        Line6(3 + 4 * Index).X2 = Image1(Index).Left - Image1(Index).Width \ 2
        Line6(3 + 4 * Index).Y1 = Image1(Index).Top + Image1(Index).Height * 2
        Line6(3 + 4 * Index).Y2 = Image1(Index).Top + Image1(Index).Height * 1.5
    Case 4 + 4 * Index
        Line6(4 + 4 * Index).X1 = Image1(Index).Left + Image1(Index).Width * 2 '�P2
        Line6(4 + 4 * Index).X2 = Image1(Index).Left + Image1(Index).Width * 1.5 '�P2
        Line6(4 + 4 * Index).Y1 = Image1(Index).Top + Image1(Index).Height * 2 '�P3
        Line6(4 + 4 * Index).Y2 = Image1(Index).Top + Image1(Index).Height * 1.5 '�P3
End Select

End Sub
Private Sub Timer30_Timer(Index As Integer) '�ϥΥ|���(���۶���)
Select Case Index
    Case 0, 1
        Load Timer30(Index + 2)
        For f = 1 + 4 * Index To 4 + 4 * Index
            Load Line6(f)
            Call sup_st((f), Index)
            Line6(f).Visible = True
        Next
        Timer30(Index).Enabled = False
        Timer30(Index + 2).Enabled = True
    Case 2, 3
        For f = 1 + 4 * (Index - 2) To 4 + 4 * (Index - 2)
            Select Case f
                Case 1 + 4 * (Index - 2)
                    Line6(1 + 4 * (Index - 2)).X1 = Line6(1 + 4 * (Index - 2)).X1 + 10
                    Line6(1 + 4 * (Index - 2)).X2 = Line6(1 + 4 * (Index - 2)).X2 + 10
                    Line6(1 + 4 * (Index - 2)).Y1 = Line6(1 + 4 * (Index - 2)).Y1 + 10
                    Line6(1 + 4 * (Index - 2)).Y2 = Line6(1 + 4 * (Index - 2)).Y2 + 10
                Case 2 + 4 * (Index - 2)
                    Line6(2 + 4 * (Index - 2)).X1 = Line6(2 + 4 * (Index - 2)).X1 - 10
                    Line6(2 + 4 * (Index - 2)).X2 = Line6(2 + 4 * (Index - 2)).X2 - 10
                    Line6(2 + 4 * (Index - 2)).Y1 = Line6(2 + 4 * (Index - 2)).Y1 + 10
                    Line6(2 + 4 * (Index - 2)).Y2 = Line6(2 + 4 * (Index - 2)).Y2 + 10
                Case 3 + 4 * (Index - 2)
                    Line6(3 + 4 * (Index - 2)).X1 = Line6(3 + 4 * (Index - 2)).X1 + 10
                    Line6(3 + 4 * (Index - 2)).X2 = Line6(3 + 4 * (Index - 2)).X2 + 10
                    Line6(3 + 4 * (Index - 2)).Y1 = Line6(3 + 4 * (Index - 2)).Y1 - 10
                    Line6(3 + 4 * (Index - 2)).Y2 = Line6(3 + 4 * (Index - 2)).Y2 - 10
                Case 4 + 4 * (Index - 2)
                    Line6(4 + 4 * (Index - 2)).X1 = Line6(4 + 4 * (Index - 2)).X1 - 10
                    Line6(4 + 4 * (Index - 2)).X2 = Line6(4 + 4 * (Index - 2)).X2 - 10
                    Line6(4 + 4 * (Index - 2)).Y1 = Line6(4 + 4 * (Index - 2)).Y1 - 10
                    Line6(4 + 4 * (Index - 2)).Y2 = Line6(4 + 4 * (Index - 2)).Y2 - 10
            End Select
        Next
        If Line6(1 + 4 * (Index - 2)).X2 >= Image1(Index - 2).Left + Image1(Index - 2).Width \ 2 Or Line6(2 + 4 * (Index - 2)).X2 <= Image1(Index - 2).Left + Image1(Index - 2).Width \ 2 Then
            For f = 1 + 4 * (Index - 2) To 4 + 4 * (Index - 2)
                Call sup_st((f), (Index - 2))
            Next
            play(1, Index - 2) = play(1, Index - 2) + 1
            
            If play(1, Index - 2) = 2 Then
                play(1, Index - 2) = 0
                Unload Timer30(Index)
                For f = 1 + 4 * (Index - 2) To 4 + 4 * (Index - 2)
                    Unload Line6(f)
                Next
                
                If play(0, Index - 2) = Index - 1 Then
                    Load Shape13(Index - 1)
                    Call sup_shape13(Index)
                    Shape13(Index - 1).Visible = True
                    Timer31(Index - 2).Enabled = True
                End If
            End If
        End If
End Select
End Sub
Private Sub sup_shape13(Index As Integer)
Shape13(Index - 1).Width = 25
Shape13(Index - 1).Height = 25
Shape13(Index - 1).Left = Image1(Index - 2).Left + Image1(Index - 2).Width \ 2 - Shape13(Index - 1).Width \ 2
Shape13(Index - 1).Top = Image1(Index - 2).Top + Image1(Index - 2).Height \ 2 - Shape13(Index - 1).Height \ 2
End Sub
Private Sub Timer31_Timer(Index As Integer) '�����
Shape13(Index + 1).Left = Shape13(Index + 1).Left - 10
Shape13(Index + 1).Top = Shape13(Index + 1).Top - 10
Shape13(Index + 1).Width = Shape13(Index + 1).Width + 20
Shape13(Index + 1).Height = Shape13(Index + 1).Height + 20
play(2, Index) = play(2, Index) + 1
If play(2, Index) = 10 Then
    'Call sup_shape13(Index + 2)
    'play(3, Index) = play(3, Index) + 1
    'play(2, Index) = 0
    'If play(3, Index) = 3 Then
        'play(3, Index) = 0
        Timer31(Index).Enabled = False
        Unload Shape13(Index + 1)
        play(0, Index) = 0
        play(2, Index) = 0
        
        label10(Index).Visible = True: image23(Index).Visible = True '�|��ɦ���
        label10(Index).Caption = Val(label10(Index).Caption) - 1 & " X"
        If Val(label10(Index).Caption) = 0 Then label10(Index).Visible = False: image23(Index).Visible = False
        label2(Index).Caption = mp(Index).Width * akiz(0) & "/" & amax(0) '/�|��ɦ���
        
        delay(Index * 2).FillColor = &HFF0000
        delay(Index * 2 + 1).FillColor = &HFF0000
        If sup_stong(kikyou(Index)) < 2 Then sup_stong(kikyou(Index)) = sup_stong(kikyou(Index)) + 1 '����1�i����
        If sup_stong(kikyou(Index) + 4) < 2 Then sup_stong(kikyou(Index) + 4) = sup_stong(kikyou(Index) + 4) + 1 '����2�i����
    'End If
End If
End Sub
Private Sub Timer27_Timer(Index As Integer) '�������
Dim a As Integer '1P or 2P
For f = 0 To 0 + player_2
    If kikyou(f) = 0 Then a = f
Next

If vnn(Index, 0) = (Index + 1) * 15 Then '�p�G�w�g���槹���h
    If vnn(Index, 1) Mod 15 <= 12 And vnn(Index, 1) Mod 3 = 0 Then
        If pk_mod = 0 Then
            For f = 1 To ma
                If ppm(1, f) <> 1 Then
                    'If coll(Shape20(Index), Image6(f)) Or coll(Image1(a), Image6(f)) Or coll(Image6(f), Shape20(Index)) Then
                    If coll(Shape20(Index), Image6(f)) Or coll(Image1(a), Image6(f)) Then
                        For i = 1 To 1 + 2 * powstong(0)
                            pps = Image6(f).Index: Call power(0, a)
                            DoEvents
                        Next
                    End If
                End If
            Next
        Else
            Call pks(Index, a * -1 + 1, Shape20(), Image1(), 0)
        End If
    End If
    Unload Shape11(vnn(Index, 1)) '�C�C����
    Unload Line3(vnn(Index, 1)) '�C�C����
    vnn(Index, 1) = vnn(Index, 1) + 1 '���w����������
    
    If vnn(Index, 1) = (Index + 1) * 15 + 1 Then '����
        If Index <> 0 Then '�p�G���O���L�p�ɾ��h
            s(20) = s(20) + 1
            Unload Timer27(Index) '���������p�ɾ�
            Unload Shape20(Index) '���������j����
        Else
            Timer27(Index).Enabled = False
            Shape20(Index).Visible = False
            If stong(0) = 0 Then Call sg(0, a) '�I�s���۲����@��
        End If
        If s(20) = holy * stong(0) Then s(7) = 0: s(20) = 0: Timer14(0).Interval = 10: Call sg(0, a) '�I�s���۲����@��
    End If
Else '��}�l����
    vnn(Index, 0) = vnn(Index, 0) + 1
    
    Shape11(vnn(Index, 0)).Left = Shape11(vnn(Index, 0) - 1).Left '�U�@�ӥ����m
    Shape11(vnn(Index, 0)).Top = Shape11(vnn(Index, 0) - 1).Top + Shape11(vnn(Index, 0) - 1).Height '�U�@�ӥ����m
    Shape11(vnn(Index, 0)).Visible = True
    
    '�]�w��m
        Shape20(Index).Left = Shape11(vnn(Index, 0)).Left - Shape20(Index).Width \ 2 + Shape11(vnn(Index, 0)).Width \ 2 '�d��j�p
        Shape20(Index).Top = Shape11(vnn(Index, 0)).Top - Shape20(Index).Height \ 2 + Shape11(vnn(Index, 0)).Height \ 2 '�d��j�p
        Shape20(Index).Visible = True
    '/�]�w��m
    
    Line3(vnn(Index, 0)).X1 = Shape11(vnn(Index, 0)).Left + Shape11(vnn(Index, 0)).Width \ 2 '�������W�����W����m
    Line3(vnn(Index, 0)).Y1 = Shape11(vnn(Index, 0)).Top '���W�����W��
    Line3(vnn(Index, 0)).X2 = Shape11(vnn(Index, 0)).Left + Shape11(vnn(Index, 0)).Width \ 2 '���W���k�U��
    Line3(vnn(Index, 0)).Y2 = Shape11(vnn(Index, 0)).Top + Shape11(vnn(Index, 0)).Height '���W���k�U
    Line3(vnn(Index, 0)).Visible = True
    
    DoEvents
End If
End Sub
Private Sub Timer34_Timer(Index As Integer) '�}�]�s�]
Call msdn(Index, Image22(), Timer34(), 20, 0, 0, s(12), s(11))
End Sub
Private Sub Timer39_Timer(Index As Integer) '�}�]�s�](�W����)
Call msdn(Index, Image12(), Timer39(), 12, 3, 1, s(41), s(40))
End Sub
Private Sub msdn(ByVal Index As Integer, b As Object, c As Object, ByVal d As Integer, ByVal e As Integer, ByVal g As Integer, h As Integer, j As Integer) '�ĴX���B���Ӫ���(�])�B���ӭp�ɾ��}�C�B�X���]�B���Ө���B�ĴX�Ӷ���(�T��)
Dim a As Integer '(a)1P or 2P
For f = 0 To 0 + player_2
    If kikyou(f) = e Then a = f
Next

b(Index).Left = b(Index).Left - 10 * hall(g, 0, Index)
b(Index).Top = b(Index).Top - 10 * hall(g, 1, Index)

If pk_mod = 0 Then
    For f = 1 To ma
        If ppm(1, f) <> 1 Then If coll(b(Index), Image6(f)) Then pps = Image6(f).Index: Call power(8, (a))
    Next
Else
    Call pks((Index), (a * -1 + 1), b, Image1(), 8)
End If

h = h + 1
If h >= 100 Then
    Unload b(Index): Unload c(Index)
    h = 0
    j = j + 1
    If j = d Then j = 0: Call sg(kikyou(a) + 4, a)
    Exit Sub
End If
If b(Index).Left < 0 Or b(Index).Left + b(Index).Width > Form1.ScaleWidth Then hall(g, 0, Index) = hall(g, 0, Index) * -1
If b(Index).Top < 0 Or b(Index).Top + b(Index).Height > Form1.ScaleHeight Then hall(g, 1, Index) = hall(g, 1, Index) * -1
DoEvents
End Sub
Private Sub Timer13_Timer(Index As Integer) '���W�ʵe
Dim b As Integer, c As Integer 'b=ttop�Ȧs, c=1P or 2P
For f = 0 To 0 + player_2
    If kikyou(f) = 1 Then c = f
Next
If Timer8(c).Enabled = True Then b = ttop(c) Else b = Image1(c).Top

If Shape3(Index).Top <= b - Image1(c).Height * 2 Then '2
    s(26) = s(26) + 1
    If s(26) < 10 Then '3
        Unload Shape3(Index): Unload Timer13(Index) '�����{��
    Else '4
        Shape3(Index).Visible = False
        If Shape18(0).Height >= 11 Then '5
            Shape18(0).Height = Shape18(0).Height - 10
            Shape18(0).Left = Image1(c).Left + Image1(c).Width \ 2 - Shape18(0).Width \ 2
            Shape18(0).Top = b - Image1(c).Height * 2
        Else '6
            Shape18(0).Visible = False
            Unload Shape3(Index): Unload Timer13(Index) '�����{��
            If stong(1) = 0 Then s(26) = 0: s(3) = 0: Call sg(1, (c))
        End If
    End If
Else '1�}�l
    Shape3(Index).Top = Shape3(Index).Top - 8 '���䲾��
    Shape18(0).Left = Image1(c).Left + Image1(c).Width \ 2 - Shape18(0).Width \ 2
    Shape18(0).Top = b - Image1(c).Height * 2
    
    If Timer1(c).Enabled = True Then Shape3(Index).Top = Shape3(Index).Top - 6.5: Shape3(0).Top = Shape3(0).Top - 2.5   '�W�T�w:�ˮ`�T�w
    If Timer2(c).Enabled = True Then Shape3(0).Top = Shape3(0).Top + 2.5 ' �ˮ`�T�w
    If Timer3(c).Enabled = True Or Timer4(c).Enabled = True Then Shape3(Index).Left = Image1(c).Left - Shape3(Index).Width \ 2 + Image1(c).Width \ 2: Shape3(0).Left = Image1(c).Left - Shape3(0).Width \ 2 + Image1(c).Width \ 2 '���k�T�w:�ˮ`�T�w
End If
DoEvents
End Sub

Public Sub supfire(a As Object, ByVal Index As Integer) 'a=���� b=���۪��� ���W���𤧲��ͺt��k
firr(1, Index) = 1
Shape18(Index).Left = a.Left + a.Width \ 2 - Shape18(Index).Width \ 2 '�X�{�b���쪺�������W
Shape18(Index).Top = a.Top + a.Height - Shape18(Index).Height '�X�{�b���쪺�������W
Shape18(Index).Visible = True
firr(2, Index) = a.Top

For ff = 1 To 3 '�X��
    Load Shape3(Index + ff * 10)
    Shape3(Index + ff * 10).Left = a.Left + a.Width \ 2 - Shape3(Index + ff * 10).Width \ 2 '�]�w���W��ئ�m
    Shape3(Index + ff * 10).Top = a.Top + a.Height \ 2 '�]�w���W��ئ�m
Next
Shape3(Index + 10).Visible = True
End Sub
Private Sub Timer28_Timer(Index As Integer) '���W����ʵe
Dim a As Integer '1P or 2P
For f = 0 To 0 + player_2
    If kikyou(f) = 1 Then a = f
Next

If firr(1, Index) = 0 Then '�Ĥ@���q�o�g���]
    '�@�w�ɶ��������]
        s(23) = s(23) + 1
        If s(23) >= 300 Then
            s(23) = 0
            Unload Shape18(Index)
            Unload Shape17(Index)
            Unload Timer28(Index)
            fireup(0, Index) = 0
            s(25) = s(25) + 1
            If s(25) = fire_cylinder Then s(23) = 0: Erase firr, fireup: s(3) = 0: s(25) = 0: s(28) = 0: Call sg(1, (a)) '6����...
            Exit Sub
        End If
    '/�@�w�ɶ��������]
    
    Shape17(Index).Top = Shape17(Index).Top - 10 * fireup(0, Index)
    Shape17(Index).Left = Shape17(Index).Left - 10 * fireup(1, Index) * firr(0, Index)
    If Shape17(Index).Left < 0 Or Shape17(Index).Left + Shape17(Index).Width > Form1.ScaleWidth Then fireup(1, Index) = fireup(1, Index) * -1 '���k�ϼu
    If Shape17(Index).Top < 0 Or Shape17(Index).Top + Shape17(Index).Height > Form1.ScaleHeight Then fireup(0, Index) = fireup(0, Index) * -1: firr(0, Index) = 1
    If firr(0, Index) = 1 Then
        If pk_mod = 0 Then
            For f = 1 To ma
                If ppm(1, f) <> 1 Then
                    If coll(Shape17(Index), Image6(f)) Then
                        For i = 1 To 1 + 2 * powstong(1)
                            pps = Image6(f).Index
                            DoEvents
                        Next
                        Unload Shape17(Index)
                        Call power(9, (a))
                        Call supfire(Image6(f), (Index)) '���W���𤧲��ͺt��k
                        Exit For
                    End If
                End If
            Next
        Else
            If hp(a * -1 + 1).Visible = True Then
                Call pks(Index, a * -1 + 1, Shape17(), Image1(), 1)
                If coll(Shape17(Index), Image1(a * -1 + 1)) Then Unload Shape17(Index): Call supfire(Image1(a * -1 + 1), (Index)) '���W���𤧲��ͺt��k
            End If
        End If
    End If
Else '�ĤG���q���ͤ��W(���]�I��ĤH����)
    If pk_mod = 0 Then
        Call supfire2(Image6(0), (Index), (a)) '���]�������ʵe
    Else
        Call supfire2(Image1(a * -1 + 1), (Index), (a)) '���]�������ʵe
    End If
End If
DoEvents
End Sub
Private Sub supfire2(a As Object, Index, ByVal c As Integer) 'a=�Q�o�ʪ� index=���۪��� c=�o�ʪ� '���]�������ʵe
If Shape18(Index).Top <= firr(2, Index) - a.Height * 2 Then
    Shape3(Index + 10).Visible = False
    Shape3(Index + 10).Top = firr(2, Index) + a.Height \ 2
    If Shape18(Index).Height >= 12 Then Shape18(Index).Height = Shape18(Index).Height - 20 '�C�C��������
    
    
    For f = 2 + firr(3, Index) To 3 '�X��
        
        If Shape3(Index + f * 10).Top <= firr(2, Index) - a.Height * 2 Then '3
            firr(3, Index) = firr(3, Index) + 1 '�w��������(�B�⦡�ݭn)
            Unload Shape3(Index + f * 10)
            If f = 3 Then '4'�X��
                Unload Shape18(Index): Unload Timer28(Index): s(25) = s(25) + 1
                Unload Shape3(10 + Index)
                If s(25) = fire_cylinder Then '5'�X�W(��K�H����)
                    s(23) = 0: Erase firr, fireup: s(3) = 0: s(25) = 0: s(28) = 0: Call sg(1, (c)) '6����...
                End If
            End If
            Exit Sub
            
        Else '2
            Shape3(Index + f * 10).Top = Shape3(Index + f * 10).Top - 8
        End If
    Next
Else '1�ĤG���q�}�l
    Shape18(Index).Top = Shape18(Index).Top - 10 '����
    Shape18(Index).Height = Shape18(Index).Height + 10 '����
    
    Shape3(Index + 10).Top = Shape3(Index + 10).Top - 8 '�Ĥ@�q����
    For f = 1 To 2 '�̫e�@�Ӥ���W�ɪ��{�רӱҰʳo�@�Ӥ��W�W��  '�X��
        If Shape3(Index + f * 10).Top <= firr(2, Index) Then Shape3(Index + (f * 10 + 10)).Visible = True: Shape3(Index + (f * 10 + 10)).Top = Shape3(Index + f * 10 + 10).Top - 8 '�ĤG�q����
    Next
End If
s(28) = s(28) + 1
If s(28) Mod 5 = 0 Then
    If pk_mod = 0 Then
        For ff = 1 To ma '�ˮ`
            If ppm(1, ff) <> 1 Then
                If coll(Image6(ff), Shape3(Index + 30)) Then
                    For i = 1 To 1 + 2 * powstong(1)
                        pps = Image6(ff).Index: Call power(1, (c)) '�X��
                        DoEvents
                    Next
                End If
            End If
            DoEvents
        Next
    Else
        For i = 1 To 1 + 2 * powstong(1)
            Call pks(Index + 30, c * -1 + 1, Shape3(), Image1(), 1)
            DoEvents
        Next
    End If
End If
End Sub
Private Sub Timer23_Timer(Index As Integer) '�ɹp
For f = 0 To 0 + player_2
    If kikyou(f) = 2 Then a = f
Next
If Line2(Index).Y2 >= Shape9.Top + Shape9.Height * 17 Then '2
    s(27) = s(27) + 1: Unload Line2(Index): Unload Timer23(Index)
    If s(27) = 15 Then '3
        Shape9.Visible = False
        If stong(2) = 0 Then s(27) = 0: s(34) = 0: Call sg(2, (a)) Else s(34) = 20  '4�ɹp�w��
    End If
Else '1�}�l
    Line2(Index).Y2 = Line2(Index).Y2 + 9
End If
DoEvents
End Sub
Private Sub Timer24_Timer(Index As Integer) '�ɹp����
Dim a As Integer '1P or 2P
For f = 0 To 0 + player_2
    If kikyou(f) = 2 Then a = f
Next

For i = 1 To 10
    t = Line2((Index - 1) * 10 + i).X1
    Line2((Index - 1) * 10 + i).X1 = Line2((Index - 1) * 10 + i).X2
    Line2((Index - 1) * 10 + i).X2 = t
    DoEvents
Next
Call lightnings

If pk_mod = 0 Then
    For f = 1 To ma
        If ppm(1, f) <> 1 Then
            If fscoll(Image6(f)) Then
                For i = 1 To 1 + 2 * powstong(2)
                    pps = Image6(f).Index: Call power(2, (a))
                    DoEvents
                Next
            End If
        End If
    Next
Else
    If fscoll(Image1(a * -1 + 1)) Then
        For i = 1 To 1 + 2 * powstong(2)
            Call pksk(2, a * -1 + 1)
            DoEvents
        Next
    End If
End If

s(30) = s(30) + 1
If s(30) >= 1 Then
    Unload Timer24(Index)
    s(31) = s(31) + 1
    If s(31) = 5 Then
        For i = 1 To 10 * 5
            Unload Line2(i)
        Next
        s(31) = 0: s(30) = 0: s(27) = 0: s(34) = 0: Call sg(2, a)
    End If
End If
End Sub
Private Sub Timer18_Timer(Index As Integer) '�ݤ뺧��
For f = 0 To 0 + player_2
    If kikyou(f) = 2 Then a = f
Next
Shape21(Index).Visible = True '���i��{
Shape21(Index).Height = Shape21(Index).Height + 4 '���i�d���ܤj
Shape21(Index).Width = Shape21(Index).Width + 16 '���i�d���ܤj
Shape21(Index).Left = Shape21(Index).Left - 8 '���i�d���ܤj
Shape21(Index).Top = Shape21(Index).Top - 2 '���i�d���ܤj
s(24) = s(24) + 1
If s(24) Mod (20 - 12 * stong(6)) = 0 Then
    If pk_mod = 0 Then
        For f = 1 To ma
            If ppm(1, f) <> 1 Then
                'If coll(Shape21(Index), Image6(f)) Or coll(Image6(f), Shape21(Index)) Or coll(Image1(a), Image6(f)) Then
                If coll(Shape21(Index), Image6(f)) Or coll(Image1(a), Image6(f)) Then
                    For i = 1 To 1 + 2 * powstong(6)
                        pps = Image6(f).Index: Call power(6, (a))
                        DoEvents
                    Next
                End If
            End If
        Next
    Else
        For i = 1 To 1 + 2 * powstong(6)
            Call pks(Index, a * -1 + 1, Shape21(), Image1(), 6)
        Next
    End If
End If
DoEvents
If Shape21(Index).Width >= 500 Then
    s(18) = s(18) + 1: Unload Shape21(Index): Unload Timer18(Index)
    If s(18) = 4 + stong(6) * wmy Then s(17) = 0: s(18) = 0: s(24) = 0: Call sg(6, (a))
End If
End Sub
Private Sub Timer29_Timer(Index As Integer) '�y�P���u�B
Dim a As Integer, b As Integer '1P or 2P
For f = 0 To 0 + player_2
    If kikyou(f) = 3 Then a = f
Next

If stong(3) >= 1 Then b = 1
Image13(Index).Left = Image13(Index).Left + 10 * ell(0)
Image13(Index).Top = Image13(Index).Top + 3 * stat(Index)
Shape19(Index * b).Left = Shape19(Index * b).Left + 10 * ell(0) 'Image13(Index).Left - Shape19(Index * b).Width
Shape19(Index * b).Top = Shape19(Index * b).Top + 3 * stat(Index)
If pk_mod = 0 Then
    For f = 1 To ma
        If ppm(1, f) <> 1 Then
            If coll(Image13(Index), Image6(f)) Then
                For i = 1 To 1 + 2 * powstong(3)
                    pps = Image6(f).Index: Call power(3, (a))
                    DoEvents
                Next
                If stong(3) = 0 Then
                    Unload Image13(Index): Unload Timer29(Index): s(15) = s(15) + 1
                    If s(15) >= 16 Then s(10) = 0: s(15) = 0: Call sg(3, (a)): Shape12.Visible = False: metor(a) = 1: Exit Sub
                    Exit Sub
                End If
            End If
        End If
        DoEvents
    Next
Else
    For i = 1 To 1 + 2 * powstong(3)
        Call pks(Index, a * -1 + 1, Image13(), Image1(), 3)
        DoEvents
    Next
    If stong(3) = 0 Then
        If s(15) >= 16 Then Call met_remove((a)): Exit Sub
        If met_already(Index) = 0 Then If Image13(Index).Left > Image1(a).Left + Image1(a).Width * 6 Or Image13(Index).Left < Image1(a).Left - Image1(a).Width * 5 Or con1(Image13(Index)) Or con2(Image13(Index)) Then Unload Image13(Index): Unload Timer29(Index): Call met_remove((a)): Erase met_already
        Exit Sub
    End If
End If
If Image13(Index).Left > Image1(a).Left + Image1(a).Width * 6 Or Image13(Index).Left < Image1(a).Left - Image1(a).Width * 5 Or con1(Image13(Index)) Or con2(Image13(Index)) Then
    Unload Image13(Index): Unload Timer29(Index)
    If stong(3) >= 1 Then
        Unload Shape19(Index)
    End If
    s(15) = s(15) + 1
    If s(15) >= 16 Then Call met_remove((a))
End If
DoEvents
End Sub
Private Sub met_remove(ByVal a As Integer)
s(10) = 0: s(15) = 0: Shape12.Visible = False: metor(a) = 1: Call sg(3, (a)) '�P�_�y�P���u�B�o�g����
End Sub
Private Sub Timer35_Timer(Index As Integer) '�W���ʤ��_�i
For f = 0 To 0 + player_2
    If kikyou(f) = 3 Then a = f
Next

Shape16(Index).Left = Shape16(Index).Left - 4
Shape16(Index).Width = Shape16(Index).Width + 8
Shape16(Index).Top = Shape16(Index).Top - 2
Shape16(Index).Height = Shape16(Index).Height + 4
cir(3) = cir(3) + 1


If cir(3) Mod 5 = 0 Then
    If Shape16(Index).BorderWidth > 1 Then
        Shape16(Index).BorderWidth = Shape16(Index).BorderWidth - 1
        
        If Shape16(Index).BorderWidth Mod 2 = 0 Then
            If pk_mod = 0 Then
                For f = 1 To ma
                    If ppm(1, f) <> 1 Then
                        'If coll(Shape16(Index), Image6(f)) Or coll(Image6(f), Shape16(Index)) Or coll(Image1(a), Image6(f)) Then
                        If coll(Shape16(Index), Image6(f)) Or coll(Image1(a), Image6(f)) Then
                            For i = 1 To 1 + 2 * powstong(7)
                                pps = Image6(f).Index: Call power(7, (a))
                                DoEvents
                            Next
                        End If
                    End If
                    DoEvents
                Next
            Else
                For i = 1 To 1 + powstong(7)
                    Call pks(Index, a * -1 + 1, Shape16(), Image1(), 7)
                    DoEvents
                Next
            End If
        End If
    Else
        Unload Shape16(Index)
        Unload Timer35(Index)
        If stong(7) = 0 Then
            s(32) = s(32) + 1
            If s(32) = 4 Then s(32) = 0: s(19) = 0: Call sg(7, (a)) '�D�����o����
        End If
    End If
End If
DoEvents
End Sub
Private Sub Timer14_Timer(Index As Integer) '���ۡ�
Dim b As Integer, c As Integer, d As Integer, e As Integer 'b=ttop�Ȧs c=��ޯ઺�H(1P or 2P) e=�ܱ𪺷N�Ӥ��^�ǭȥ�
For f = 0 To 0 + player_2
    If Index = kikyou(f) Then c = f
    If Index = kikyou(f) + 4 Then c = f
Next
        
Select Case Index
    Case 0 '�������
        If s(7) = 0 Then
            For f = 0 To holy * stong(0) '�p�G�����𪬺A�h����1~4���p�ɾ�
                If f <> 0 Then
                    Load Shape20(f)
                    Load Timer27(f)
                End If
                s(9) = 15 * (f + 1) '�ƶq
                vnn(f, 0) = 15 * f + 1
                vnn(f, 1) = 15 * f + 1
            Next
            For f = 1 To s(9) '���ͨ�������һݪ�����
                Load Shape11(f) '���ͥ���
                Load Line3(f)
            Next
            For f = 1 To holy * stong(0) '�üƨM�w2~5(4��)�����������m
                Randomize
                Shape11(15 * f + 1).Top = Int(Rnd * (Form1.ScaleHeight - Shape20(0).Height * 2.2)) - Shape20(0).Height * 1.2
                Shape11(15 * f + 1).Left = Int(Rnd * (Form1.ScaleWidth - Shape20(0).Width \ 2)) + Shape20(0).Width \ 4
            Next
        End If
        If stong(0) = 0 Then
            Timer27(0).Interval = 30: Timer27(0).Enabled = True: Timer14(Index).Enabled = False
        Else
            Timer14(Index).Interval = 300
            Timer27(s(7)).Interval = 20
            Timer27(s(7)).Enabled = True
            If s(7) = holy * stong(0) Then Timer14(Index).Enabled = False
            s(7) = s(7) + 1
        End If
    Case 1 '���W
        If Timer8(c).Enabled = True Then b = ttop(c) Else b = Image1(c).Top
        
        Shape18(0).Top = b - Image1(c).Height * 2
        Shape18(0).Height = 196
        
        s(3) = s(3) + 1
        Load Shape3(s(3)) '���ͤ��W(���)
        Load Timer13(s(3)) '���ͤ��W�p�ɾ�
        Shape3(s(3)).Left = Image1(c).Left - Shape3(s(3)).Width \ 2 + Image1(c).Width \ 2 '�]�w���W����m
        Shape3(s(3)).Top = b + Image1(c).Height \ 2 '�]�w���W����m
        Shape18(0).Left = Image1(c).Left + Image1(c).Width \ 2 - Shape18(0).Width \ 2 '�]�w���ߪ���m
        Shape18(0).Top = b + Image1(c).Height - Shape18(0).Height '�]�w���ߪ���m
        Shape18(0).Visible = True
        Shape3(s(3)).Visible = True
        Timer13(s(3)).Enabled = True '�Ұʤ��W�p�ɾ�
        
        If pk_mod = 0 Then
            For f = 1 To ma
                If ppm(1, f) <> 1 Then
                    If ma = 1 Then '�p�G�O�]���h
                        'If coll(Image6(f), Shape3(0)) Or coll(Shape3(0), Image6(f)) Or coll(Image1(c), Image6(f)) Then
                        If coll(Image6(f), Shape3(0)) Or coll(Image1(c), Image6(f)) Then
                            For i = 1 To 1 + powstong(1)
                                pps = Image6(f).Index: Call power(1, c)
                                DoEvents
                            Next
                        End If
                    Else
                        If coll2(Shape2(f), Shape3(0)) Then
                            If coll(Image6(f), Shape3(0)) Then
                                For i = 1 To 1 + powstong(1)
                                    pps = Image6(f).Index: Call power(1, c)
                                    DoEvents
                                Next
                            End If
                        End If
                    End If
                End If
            Next
        Else
            Call pks(0, c * -1 + 1, Shape3(), Image1(), 1)
        End If
        
        If stong(1) >= 1 Then
            Select Case fire_cylinder
                Case Is <= 3
                    d = 1
                Case 4 To 9
                    d = 2
            Case Else
                d = 1
            End Select
            If s(3) Mod d = 0 Then
                Load Shape18(s(3) \ d) '���ͤ��W(����)
                Shape18(s(3) \ d).Height = 1
                
                Shape17(s(3) \ d).Left = Image1(c).Left + Image1(c).Width \ d - Shape17(s(3) \ d).Width \ d '���]
                Shape17(s(3) \ d).Top = b
                Shape17(s(3) \ d).Visible = True
                Timer28(s(3) \ d).Enabled = True '���W�������
            End If
        End If
        If s(3) = 10 Then Timer14(Index).Enabled = False
    Case 2 '�ɹp
        If s(34) < 15 Then
            Timer14(Index).Interval = 30
            s(34) = s(34) + 1
            Load Line2(s(34)) '���͹p�q
            
            Line2(s(34)).X1 = Int(Rnd * Shape9.Width) + Shape9.Left
            Line2(s(34)).X2 = Int(Rnd * Shape9.Width) + Shape9.Left
            Line2(s(34)).Y1 = Int(Rnd * Shape9.Height) + Shape9.Top
            Line2(s(34)).Y2 = gup + sh(c) + 6 '��ޯ઺��m+�v�l+�ǷL�վ�
            
            Line2(s(34)).Visible = True
            Load Timer23(s(34)) '���ͼɹp�p�ɾ�
            Timer23(s(34)).Enabled = True
            
            If s(34) Mod (3 - 2 * stong(2)) = 0 Then '�T�s�q
                If pk_mod = 0 Then
                    For f = 1 To ma
                        If ppm(1, f) <> 1 Then
                            If ma = 1 Then '�p�G�O�]���h
                                'If coll(Image6(f), Shape22(0)) Or coll(Shape22(0), Image6(f)) Or coll(Image1(c), Image6(f)) Then
                                If coll(Image6(f), Shape22(0)) Or coll(Image1(c), Image6(f)) Then
                                    For i = 1 To 1 + powstong(2)
                                        pps = Image6(f).Index: Call power(2, c)
                                        DoEvents
                                    Next
                                End If
                            Else
                                If coll2(Shape2(f), Shape22(0)) Or coll(Image1(c), Image6(f)) Then
                                    If coll(Image6(f), Shape22(0)) Then
                                        For i = 1 To 1 + powstong(2)
                                            pps = Image6(f).Index: Call power(2, c)
                                            DoEvents
                                        Next
                                    End If
                                End If
                            End If
                        End If
                    Next
                Else
                    Call pks(0, c * -1 + 1, Shape22(), Image1(), 2)
                End If
            End If
        Else
            If stong(2) = 0 Then
                Timer14(Index).Enabled = False
            Else
                If s(34) = 20 Then '�P�_��l�ɹp�w��
                    If s(29) = 0 Then '�H�U���u����@�����{���X
                        Timer14(Index).Interval = 100
                        Randomize
                        For f = 1 To 10 * 5 '�X�D
                            Load Line2(f)
                            Line2(f).BorderColor = RGB(255, 255, 0)
                            DoEvents
                        Next
                        s(29) = 1
                    Else
                        Call lightnings
                        Load Timer24(s(29))
                            
                        Timer24(s(29)).Enabled = True
                        If s(29) = 5 Then Timer14(Index).Enabled = False: s(29) = 0: Exit Sub '�X�D
                        s(29) = s(29) + 1
                    End If
                End If
            End If
        End If
    Case 3 '�y�P���u�B
        Randomize
        s(10) = s(10) + 1
        Load Image13(s(10)) '���u
        Load Timer29(s(10))
        Image13(s(10)).Left = Int(Rnd * (Shape12.Width - Image13(s(10)).Width)) + Shape12.Left
        Image13(s(10)).Top = Int(Rnd * (Shape12.Height - Image13(s(10)).Height)) + Shape12.Top
        Image13(s(10)).Visible = True
        stat(s(10)) = dzxc(c)
        Image13(s(10)).ZOrder 0
        If stong(3) >= 1 Then
            Load Shape19(s(10)) '�����
            Shape19(s(10)).Top = Image13(s(10)).Top + Image13(s(10)).Height \ 2 - Shape19(s(10)).Height \ 2
            If ell(0) = -1 Then
                Shape19(s(10)).Left = Image13(s(10)).Left + Image13(s(10)).Width
            Else
                Shape19(s(10)).Left = Image13(s(10)).Left - Shape19(s(10)).Width
            End If
            Shape19(s(10)).Visible = True '/�����
        End If
        Timer29(s(10)).Enabled = True
        If s(10) >= 16 Then s(10) = 0: Timer14(Index).Enabled = False
    Case 4 '�ܱ𪺷N��
        
    '�Ĥ@���D�u....
    
        If ven(14) = 0 Then '�e���~�P
            If ven(0) = 0 Then
                Line4(0).X1 = Line4(0).X1 - (8 + 8 * stong(4))
                If Line4(0).X1 <= Shape10.Left + 16 Then Line4(0).X1 = Shape10.Left + 16: ven(0) = 1: Line4(2).X1 = Line4(0).X1: Line4(2).X2 = Line4(0).X1
            End If
            If ven(1) = 0 Then
                Line4(0).Y1 = Line4(0).Y1 - (4 + 4 * stong(4))
                If Line4(0).Y1 <= Shape10.Top + Shape10.Height \ 4 Then Line4(0).Y1 = Shape10.Top + Shape10.Height \ 4: ven(1) = 1: Line4(2).Y1 = Line4(0).Y1: Line4(2).Y2 = Line4(2).Y1: ven(5) = 1
            End If
            
        '�Ĥ@�����u
            If ven(5) = 1 Then
                If ven(4) = 0 Then
                    Line4(2).X2 = Line4(2).X2 + (8 + 8 * stong(4))
                    If Line4(2).X2 >= Shape10.Left + Shape10.Width - 17 Then Line4(2).X2 = Shape10.Left + Shape10.Width - 17: ven(4) = 1: ven(8) = 1: Line4(1).X1 = Line4(2).X2: Line4(1).Y1 = Line4(2).Y2: Line4(1).Y2 = Line4(1).Y1: Line4(1).X2 = Line4(1).X1
                End If
            End If
        '�ĤG�����u
            If ven(8) = 1 Then
                If ven(9) = 0 Then
                    Line4(1).X2 = Line4(1).X2 - (8 + 8 * stong(4))
                    If Line4(1).X2 <= Shape10.Left + Shape10.Width \ 2 Then Line4(1).X2 = Shape10.Left + Shape10.Width \ 2: ven(9) = 1
                End If
                If ven(10) = 0 Then
                    Line4(1).Y2 = Line4(1).Y2 + (4 + 4 * stong(4))
                    If Line4(1).Y2 >= Shape10.Top + Shape10.Height Then Line4(1).Y2 = Shape10.Top + Shape10.Height: ven(10) = 1: ven(14) = 1 '����
                End If
            End If
            
    '�ĤG���D�u....
        
            If ven(2) = 0 Then
                Line4(5).X2 = Line4(5).X2 + (8 + 8 * stong(4))
                If Line4(5).X2 >= Shape10.Left + Shape10.Width - 17 Then Line4(5).X2 = Shape10.Left + Shape10.Width - 17: ven(2) = 1: Line4(4).X2 = Line4(5).X2: Line4(4).X1 = Line4(4).X2
            End If
            If ven(3) = 0 Then
                Line4(5).Y2 = Line4(5).Y2 + (4 + 4 * stong(4))
                If Line4(5).Y2 >= Shape10.Top + (Shape10.Height \ 4) * 3 Then Line4(5).Y2 = Shape10.Top + (Shape10.Height \ 4) * 3: ven(3) = 1: Line4(4).Y2 = Line4(5).Y2: Line4(4).Y1 = Line4(4).Y2: ven(7) = 1
            End If
        
        '�Ĥ@�����u
            If ven(7) = 1 Then
                If ven(6) = 0 Then
                    Line4(4).X1 = Line4(4).X1 - (8 + 8 * stong(4))
                    If Line4(4).X1 <= Shape10.Left + 16 Then Line4(4).X1 = Shape10.Left + 16: ven(6) = 1: ven(11) = 1: Line4(3).X2 = Line4(4).X1: Line4(3).X1 = Line4(3).X2: Line4(3).Y2 = Line4(4).Y1: Line4(3).Y1 = Line4(3).Y2
                End If
            End If
        '�ĤG�����u
            If ven(11) = 1 Then
                If ven(12) = 0 Then
                    Line4(3).X1 = Line4(3).X1 + (8 + 8 * stong(4))
                    If Line4(3).X1 >= Shape10.Left + Shape10.Width \ 2 Then Line4(3).X1 = Shape10.Left + Shape10.Width \ 2: ven(12) = 1
                End If
                
                If ven(13) = 0 Then
                    Line4(3).Y1 = Line4(3).Y1 - (4 + 4 * stong(4))
                    If Line4(3).Y1 <= Shape10.Top Then Line4(3).Y1 = Shape10.Top: ven(13) = 1
                End If
            End If
                
        '�j�O�@��
        Else
            metor(c) = 1
            If ven(15) = 0 Then
                Image17.Left = Image1(c).Left - Image17.Width \ 2 + Image1(c).Width \ 2 '�ܱ��m
                Image17.Top = Image1(c).Top - Image17.Height \ 2 + Image1(c).Height \ 2
                Image17.Visible = True
                
                Image18(0).Top = Image17.Top + 64 '�b�ڦ�m
                Image18(0).Left = Image17.Left + (Image17.Width * (kik_ell + 2) Mod 3) - (Image18(0).Width * (kik_ell + 2) Mod 3) + (16 * kik_ell)
                Image18(0).Visible = True
                
                Image19.Top = Image18(0).Top - Image19.Height \ 2 + Image18(0).Height \ 2 '�}�]����
                
                Randomize
                If stong(4) = 2 Then rrf = Int(Rnd * 3) - 1 Else rrf = 1
                For f = 0 To 20 '�}�]�s�]��l����V���w
                    hall(0, 0, f) = rrf
                    If f Mod 2 = 0 Then hall(0, 1, f) = -1 Else hall(0, 1, f) = 1
                Next
                
                ven(15) = 1
            Else
                Image17.Left = Image17.Left - 10 * kik_ell
                If Image17.Left >= 1 And Image17.Left <= Form1.ScaleWidth - 1 Then
                    Image18(0).Left = Image17.Left + (Image17.Width * (kik_ell + 2) Mod 3) - (Image18(0).Width * (kik_ell + 2) Mod 3) + (16 * kik_ell) '�b����m
                    Image18(0).Top = Image17.Top + 64
                Else '��ܱ��F��ɫh
                
                    '����P�_
                    If pk_mod = 0 Then
                        For f = 1 To ma
                            If ppm(1, f) <> 1 Then
                                'If coll(Image6(f), Image18(0)) Or coll(Image18(0), Image6(f)) Then
                                If coll(Image6(f), Image18(0)) Then
                                    For i = 1 To 1 + 2 * powstong(4)
                                        pps = Image6(f).Index: Call power(4, c)
                                        DoEvents
                                    Next
                                End If
                            End If
                        Next
                    Else
                        Call pks(0, c * -1 + 1, Image18(), Image1(), 4)
                    End If
                    
                    Select Case stong(4) '��ܶ��𵥯�
                        Case 0
                            Call kiker(c) '�I�s�ܱ𪺷N�Ӧ@��
                            If Image18(0).Left > Form1.ScaleWidth Or Image18(0).Left < 0 - Image18(0).Width Then Call sg(Index, (c)) '���۩�
                        Case 1
                            If (Image19.Left + Image19.Width > Form1.ScaleWidth And kik_ell = 1) Or (Image19.Left < 0 And kik_ell = -1) Then
                        
                                Image18(0).Left = (Form1.ScaleWidth - Image18(0).Width) * (((kik_ell + 2) Mod 3) * -1 + 1)
                                Image19.Top = Image18(0).Top - Image19.Height \ 2 + Image18(0).Height \ 2
                                Call kik_i(Index, 0, c) '�I�s���ͯ}�]�s�]�t��k
                            Else
                                Call kiker(c) '�I�s�ܱ𪺷N�Ӧ@��
                            End If
                        Case 2
                            e = 1
                            Call kik_i(Index, e, c) '�I�s���ͯ}�]�s�]�t��k
                            If e = 2 Then Exit Sub Else Call kiker(c) '�I�s�ܱ𪺷N�Ӧ@��
                    End Select
                End If
            End If
        End If
        '........
        '........
    Case 5 '�Y��
        If gbl = 0 Then '�}�l
            For f = 1 To 3
                b = Line5(f).X2
                Line5(f).X1 = Line5(f).X1 - 20 '����
                Line5(f).X2 = Line5(f).X2 + 20 '����
                If f > 1 Then
                    Line5(f).Y1 = Line5(f).Y1 - 10 * (f * 2 - 5) 'f=2,f=3�@�u�ܥ��t��
                    Line5(f).Y2 = Line5(f).Y2 + 10 * (f * 2 - 5) 'f=2,f=3�@�u�ܥ��t��
                End If
                If Line5(f).X2 >= gdown + Image1(c).Width * 5 Then gbl = 1: Exit For
                DoEvents
            Next
        Else '����
            If stong(5) >= 1 Then
                For f = 1 To 3
                    Form1.Left = Form1.Left - 100
                    Form1.Left = Form1.Left + 100
                    Form1.Left = Form1.Left + 100
                    Form1.Left = Form1.Left - 100
                Next
            End If
            For ff = 1 To 5 + stong(5) * 5 '�s���ˮ`
                If pk_mod = 0 Then
                    For f = 1 To ma
                        If ppm(1, f) <> 1 Then
                            If fscoll(Image6(f)) Then
                                For i = 1 To 1 + 2 * powstong(5)
                                    pps = Image6(f).Index: Call power(5, c)
                                    DoEvents
                                Next
                            End If
                        End If
                    Next
                Else
                    If fscoll(Image1(c * -1 + 1)) Then
                        For i = 1 To 1 + 2 * powstong(5)
                            Call pksk(5, c * -1 + 1)
                            DoEvents
                        Next
                    End If
                End If
            Next
            For f = 1 To 3
                Unload Line5(f)
            Next
            gbl = 0
            Timer14(Index).Enabled = False
            Call sg(Index, (c))
        End If
    Case 6 '�ݤ뺧��
        s(17) = s(17) + 1
        If stong(6) = 0 Then
            Timer14(6).Interval = 200
            Shape21(s(17)).BorderColor = &HFF0000 '���i�C��
            Shape21(s(17)).Left = Image1(c).Left - Shape21(s(17)).Width \ 2 + Image1(c).Width \ 2 '���i��l��m
            Shape21(s(17)).Top = Image1(c).Top + Image1(c).Height - 15 '���i��l��m
        Else
            Timer14(6).Interval = 160
            Randomize
            Shape21(s(17)).BorderColor = &HFF0000 '���i�C��
            Shape21(s(17)).Left = Int(Rnd * (Form1.ScaleWidth - Shape21(s(17)).Width * 3.5)) + Shape21(s(17)).Width * 1.5 '���i��l��m
            Shape21(s(17)).Top = Int(Rnd * (Form1.ScaleHeight - Shape21(s(17)).Height * 2 - Line1.Y1)) + Line1.Y1 '���i��l��m
        End If
        Timer18(s(17)).Enabled = True
        If s(17) = 4 + stong(6) * wmy Then Timer14(Index).Enabled = False
    Case 7 '�W����
        Form1.FillStyle = 0
        Form1.AutoRedraw = False
        Form1.Refresh
        
        If cir(1) >= 150 Then
            If cir(0) >= 50 + 50 * stong(7) Then
                If cir(2) >= 150 Then
                    If s(19) >= 4 Then '5���I
                        If stong(7) >= 1 Then
                            If s(22) < 1 Then Timer14(7).Interval = 100
                            s(22) = s(22) + 1
                            For f = 0 To 1
                                Select Case s(22) Mod 4
                                    Case 1
                                        hall(1, f, s(22)) = 1 * (f * 2 - 1)
                                    Case 2
                                        hall(1, f, s(22)) = -1
                                    Case 3
                                        hall(1, f, s(22)) = 1
                                    Case 0
                                        hall(1, f, s(22)) = -1 * (f * 2 - 1)
                                End Select
                            Next
                            Load Timer39(s(22))
                            Load Image12(s(22))
                            Image12(s(22)).Left = appshe1 - Image12(s(22)).Width \ 2 + Shape16(0).Width \ 2
                            Image12(s(22)).Top = appshe2 + Image12(s(22)).Height
                            Image12(s(22)).Visible = True
                            Timer39(s(22)).Enabled = True
                            If s(22) = 12 Then s(22) = 0: Timer14(Index).Enabled = False
                        Else
                            Timer14(Index).Enabled = False
                        End If
                    Else '4���;_�i
                        Timer14(7).Interval = 130
                        s(19) = s(19) + 1
                        
                        If s(19) = 1 Then '���_�i����I�B�A���|��ۥD�H�]
                            For f = 1 To 4
                                Shape16(f).Left = Image1(c).Left + Image1(c).Width \ 2 - Shape16(f).Width \ 2 '�_�i��m�]�w
                                Shape16(f).Top = Image1(c).Top + Image1(c).Height '/�_�i��m�]�w
                            Next
                            appshe1 = Shape16(1).Left
                            appshe2 = Shape16(1).Top
                        End If
                        
                        Shape16(s(19)).Visible = True
                        Timer35(s(19)).Enabled = True
                    End If
                Else '3���ܭp�ɾ��t�� & �V�U��
                    cir(2) = cir(2) + 20
                    Timer14(7).Interval = 10
                End If
            Else '2�ܤj
                cir(0) = cir(0) + 5
            End If
        Else '1�V�W��(�_�l�I)
            cir(1) = cir(1) + 15
        End If
        If cir(2) < 150 Then Form1.Circle (Image1(c).Left + Image1(c).Width \ 2, ((Image1(c).Top) - cir(1)) + cir(2)), cir(0)
End Select
DoEvents
End Sub
Private Sub kik_i(Index As Integer, a As Integer, b As Integer) '�ܱ𪺷N�Ӳ��ͦ@�� b=1P or 2P
If s(14) = 20 Then
    Image18(0).Visible = False
    Image19.Visible = False
    s(14) = 0
    Timer14(Index).Enabled = False '���۩�
    If a = 1 Then a = 2: Exit Sub
Else
    If a = 0 Then Image19.Left = (Form1.ScaleWidth - Image19.Width) * (((kik_ell + 2) Mod 3) * -1 + 1)

    '�}�]�s�]
    s(14) = s(14) + 1
    Load Timer34(s(14))
    Load Image22(s(14))

    Randomize
    Image22(s(14)).Left = Image19.Left + (Image19.Width - Image22(s(14)).Width) * (((kik_ell + 2) Mod 3) * -1 + 1) '�]�w��m
    Image22(s(14)).Top = Int(Rnd * Image19.Height) + Image19.Top '�]�w��m

    Image22(s(14)).Visible = True
    Timer34(s(14)).Enabled = True '�Ұʯ}�]�s�]����
End If
End Sub
Private Sub lightnings() '�ɹp�@��
For f = 0 To 4
    '�X�D--�]�w�C�D�p���Ĥ@�Ӧ�m
    Line2(10 * f + 1).X1 = Int(Rnd * Form1.ScaleWidth)
    Line2(10 * f + 1).X2 = Line2(10 * f + 1).X1 + 30
    Line2(10 * f + 1).Y1 = Int(Rnd * (Form1.ScaleHeight)) - Form1.ScaleHeight \ 2
    Line2(10 * f + 1).Y2 = Line2(10 * f + 1).Y1 + 40
    Line2(10 * f + 1).Visible = True
    
    For i = 2 To 10 '�]�w�C�D�p���ĤG��̫�Ӧ�m
        Line2(f * 10 + i).X1 = Line2(f * 10 + i - 1).X2 '�p���S��
        Line2(f * 10 + i).Y1 = Line2(f * 10 + i - 1).Y2 '�p���S��
        Line2(f * 10 + i).X2 = Line2(f * 10 + i - 1).X1 '�p���S��
        Line2(f * 10 + i).Y2 = Line2(f * 10 + i - 1).Y2 + 40 '�p���S��
        Line2(f * 10 + i).Visible = True
        DoEvents
    Next
Next
End Sub
Private Sub tai(a As Integer, b As Integer) '���`���A�t��k a=���ӵ���,b=���ӨϥΪ�
Dim c As Integer, d As Integer
ys(a) = yst
If a = 0 Then
    Randomize
    c = Int(Rnd * 3) + 1
    If c = 2 Then ys(a) = ys(a) * 2 '�s���ˮ`
Else
    If pk_mod = 1 Then pps = b * -1 + 2 '�p�G�Opk�Ҧ��hpps=�Q�����ϥΪ�
    Randomize
    c = Int(Rnd * 3) + 1
    If c = 2 Then
        If ppm(19, pps) = 0 Then ppm(19, pps) = 1: Timer22(pps).Interval = 3000
        If pk_mod = 0 Then
            If a = 5 Then '�ۤ�
                If ma = 1 Then
                    Image6(pps).Picture = Image24((xy(pps) + 2) Mod 3 + 2).Picture
                Else
                    Image6(pps).Picture = Image24((xy(pps) + 2) Mod 3).Picture
                End If
                Exit Sub
            Else
                Timer43(pps).Enabled = True
                Exit Sub '�·�
            End If
        End If
    End If
End If
End Sub
Private Sub Timer43_Timer(Index As Integer) '�·��ʵe
Dim a As Integer, b As Integer
If ppm(1, Index) <> 1 Then
    b = 5
    s(48) = s(48) + 1
    If s(48) Mod 2 = 0 Then s(48) = 0: a = 1 Else a = -1
    Image6(Index).Left = Image6(Index).Left + b * a
End If
DoEvents
End Sub
Public Sub mpss(a As Integer, e As Integer) 'MP�l�Ӻt��k�� a=���ۺ��� e=1P or 2P
If kmp_GPR(e) >= smp(a) + smp(a) * sup_stong(a) Then
    Dim d As Integer
    If Timer8(e).Enabled = True Then d = ttop(e) Else d = Image1(e).Top
    
    stong(a) = sup_stong(a)
    sup_stong(a) = 0
    If khp_GPR(e) <= (amax(0) / akiz(0)) / 3 Then delay(a \ 4 + 2 * e).FillColor = &HFF&: powstong(a) = 1 Else delay(a \ 4 + 2 * e).FillColor = &H80FF80: powstong(a) = 0
    
    gdown = Image1(e).Left '�D����m Left
    gup = Image1(e).Top '�D����mTop
    
    Shape3(0).Left = Image1(e).Left - Shape3(0).Width \ 2 + Image1(e).Width \ 2 '���W��
    Shape3(0).Top = d + Image1(e).Height \ 2
    Shape20(0).Left = Image1(e).Left - Shape20(0).Width \ 2 + Image1(e).Width \ 2  '���W��
    Shape20(0).Top = d + Image1(e).Height \ 2
    Shape21(0).Left = Image1(e).Left - Shape21(0).Width \ 2 + Image1(e).Width \ 2  '���W��
    Shape21(0).Top = d + Image1(e).Height \ 2
    Shape22(0).Left = Image1(e).Left - Shape22(0).Width \ 2 + Image1(e).Width \ 2  '���W��
    Shape22(0).Top = d + Image1(e).Height \ 2

    Timer25(e).Enabled = False 'mp�W�[����
    If a = 3 Or a = 4 Then metor(e) = 0: zxcv(2 + 6 * e) = zxcv(3 + 6 * e) '��������
    If a = kikyou(e) + 4 Then
        Timer21(1 + 2 * e).Enabled = False '����A delay�{�{(���ۤG)
        delay(1 + 2 * e).Visible = False: delay(1 + 2 * e).Width = 1
    Else
        Timer21(0 + 2 * e).Enabled = False '����A delay�{�{
        delay(0 + 2 * e).Visible = False: delay(0 + 2 * e).Width = 1
    End If
    kmp_GPR(e) = kmp_GPR(e) - smp(a) - smp(a) * stong(a)
    mp(e).Width = kmp_GPR(e) \ 1
    label2(e).Caption = (kmp_GPR(e) * akiz(0)) \ 1 & "/" & amax(0)
    Select Case a
        Case 0 '�������
            Shape20(0).BorderColor = RGB(255, 255, 0) '���w�d���C���
        
            Shape11(0).Left = Image1(e).Left + Image1(e).Width \ 2 - Shape11(0).Width \ 2
            Shape11(0).Top = d - 190
            
            'Shape11(0).Left = image1(e).Left - 175
            'Shape11(0).Top = image1(e).Top - 100
            Line3(0).X1 = Shape11(0).Left + Shape11(0).Width \ 4  '�������W�����W����m
            Line3(0).Y1 = Shape11(0).Top '���W�����W��
            Line3(0).X2 = Shape11(0).Left + Shape11(0).Width \ 2  '���W���k�U��
            Line3(0).Y2 = Shape11(0).Top + Shape11(0).Height \ 2  '���W���k�U��
            
        Case 1 '���W
            If stong(1) >= 1 Then
                For f = 1 To fire_cylinder '�X�W
                    Load Shape17(f)
                    Load Timer28(f)
                    fireup(0, f) = 1
                    If f Mod 2 = 0 Then fireup(1, f) = 1 Else fireup(1, f) = -1
                Next
            End If
            Shape18(0).Height = 1
            Shape18(0).Left = Image1(e).Left + Image1(e).Width \ 2 - Shape18(0).Width \ 2
            Shape18(0).Top = d + Image1(e).Height - Shape18(0).Height
            Shape18(0).Visible = True
          
            '------�_������
            Set ds = dxa.DirectSoundCreate("")
            ds.SetCooperativeLevel Me.hWnd, DSSCL_NORMAL
            bu.lFlags = DSBCAPS_STATIC
            Set dsb = ds.CreateSoundBufferFromFile("Data\���W.wav", bu, wf)
            dsb.play DSBPLAY_DEFAULT '''''''''''''''''''''''�w�]����(�榸).
            '------�_������
            
            If stong(a) = 1 Then c = 0: b = 0 Else c = 0: b = 0
        Case 2 '�ɹp
            Shape9.Left = Image1(e).Left - 80 '�W�����¬}
            Shape9.Top = d - 180 '�W�����¬}
            Shape9.Visible = True
        
            Set dscc = dxa.DirectSoundCreate("")
            dscc.SetCooperativeLevel Me.hWnd, DSSCL_NORMAL
            bu.lFlags = DSBCAPS_STATIC
            Set dsbcc = dscc.CreateSoundBufferFromFile(("Data\�ɹp.wav"), bu, wf)
            dsbcc.play DSBPLAY_DEFAULT '''''''''''''''''''''''�w�]����(�榸).
            
        Case 3 '�y�P���u�B
            ReDim stat(20)
            dzxc(e) = 0
            If zxcv(3 + 6 * e) = 1 Then Image13(0).Picture = Image14(1).Picture: ell(0) = -1: Shape12.Left = Image1(e).Left + Image1(e).Width - Shape12.Width Else Image13(0).Picture = Image14(0).Picture: ell(0) = 1: Shape12.Left = Image1(e).Left '0���k1����
            Shape12.Top = d + Image1(e).Height \ 2 - Shape12.Height \ 2
            Shape12.Visible = True
            
        Case 4 '�ܱ𪺷N��
            Image17.Picture = Image27(zxcv(3 + 6 * e)).Picture '�ܱ𥪥k�� 0)�k 1)��
            Image18(0).Picture = Image28(zxcv(3 + 6 * e)).Picture
            Image19.Picture = Image29(zxcv(3 + 6 * e)).Picture
            kik_ell = zxcv(3 + 6 * e) * -2 + 1
        
            Shape10.Left = Image1(e).Left - Shape10.Width \ 2 + Image1(e).Width \ 2
            Shape10.Top = d + Image1(e).Height \ 2
            Shape10.Visible = True
            
            Line4(5).X1 = Shape10.Left + Shape10.Width \ 2 '�]�w��l�Ƴ��I
            Line4(5).Y1 = Shape10.Top
            Line4(0).X2 = Shape10.Left + Shape10.Width \ 2 '�]�w��l�Ƴ��I
            Line4(0).Y2 = Shape10.Top + Shape10.Height
            
            Line4(0).X1 = Line4(0).X2
            Line4(0).Y1 = Line4(0).Y2
            
            Line4(5).X2 = Line4(5).X1
            Line4(5).Y2 = Line4(5).Y1
            
            For f = 0 To 5
                Line4(f).Visible = True
            Next
            
        Case 5 '�Y��
            For f = 1 To 3
                Load Line5(f) '�Y���u��
                Line5(f).X1 = Image1(e).Left + Image1(e).Width \ 2
                Line5(f).Y1 = d + Image1(e).Height \ 2
                Line5(f).X2 = Line5(f).X1
                Line5(f).Y2 = Line5(f).Y1
                Line5(f).Visible = True
            Next
            
        '-----�_������
            Set dscd = dxa.DirectSoundCreate("")
            dscd.SetCooperativeLevel Me.hWnd, DSSCL_NORMAL
            bu.lFlags = DSBCAPS_STATIC
            Set dsbd = dscd.CreateSoundBufferFromFile(("Data\�Y��.wav"), bu, wf)
            dsbd.play DSBPLAY_DEFAULT '''''''''''''''''''''''�w�]����(�榸).
        '-----�_������
        
        Case 6 '�ݤ뺧��
            For f = 1 To 4 + stong(6) * wmy
                Load Shape21(f)
                Shape21(f).Width = 60 '���i��l�j�p
                Shape21(f).Height = 30 '���i��l�j�p
                
                Load Timer18(f)
            Next
            
        Case 7 '�W����
            cir(0) = 5
            Timer14(7).Interval = 50
            For f = 1 To 4
                Load Shape16(f)
                Load Timer35(f)
            Next
            
    End Select
    
    Select Case stong(a)
        Case 2
            c = 300: b = 100
    End Select
    
    ys(kikyou(e)) = Int(Rnd * (100 + c - 10 - b + 1)) + 10 + b '�s���ˮ`
    ys(kikyou(e) + 4) = Int(Rnd * (100 + c - 10 - b + 1)) + 10 + b '/�s���ˮ`
    
    yst = ys(kikyou(e))
    
    Timer14(a).Enabled = True
End If
End Sub
Private Sub kiker(a As Integer) '�ܱ𪺷N�Ӧ@��

Image18(0).Left = Image18(0).Left + 30 * kik_ell '�b�ڮg�X
Image19.Left = Image18(0).Left + (Image18(0).Width * (((kik_ell + 2) Mod 3) * -1 + 1)) - Image19.Width \ 2 '�}�]����
Image19.Visible = True

If Image17.Left + Image17.Width < 0 Or Image17.Left > Form1.ScaleWidth Then '��ܱ�W�X��ɫh
    Image17.Visible = False
End If

End Sub
Public Sub sg(a As Integer, b As Integer) '���ۦ@��
Timer14(a).Enabled = False

If a <= 3 Then delay(0 + 2 * b).Visible = True: Timer20(0 + 2 * b).Enabled = True Else delay(1 + 2 * b).Visible = True: Timer20(1 + 2 * b).Enabled = True '����delay

Select Case a
    Case 0 '�������
        Erase vnn
        ReDim vnn(holy * 2, 1) As Integer
    Case 3 '�y�P���u�B
        metor(b) = 1
    Case 4 '�ܱ𪺷N��
        Shape10.Visible = False
        Image18(0).Visible = False
        Image19.Visible = False
        Image19.Left = 1
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
    Case 7 '�W����
        Erase cir: s(19) = 0: s(22) = 0
End Select
stong(a) = 0
End Sub
Private Sub Timer20_Timer(Index As Integer) '����delay
delay(Index).Width = delay(Index).Width + Label6(Index).Width \ 20
If delay(Index).Width >= Label6(Index).Width Then Timer20(Index).Enabled = False: Timer21(Index).Enabled = True
DoEvents
End Sub
Private Sub Timer21_Timer(Index As Integer) 'A delay�{�{
slow(Index, 0) = slow(Index, 0) + 1
If slow(Index, 0) Mod 2 = 0 Then slow(Index, 0) = 0: delay(Index).Visible = True Else delay(Index).Visible = False
DoEvents
End Sub
Private Sub Timer10_Timer(Index As Integer) '����
Dim b(1) As Integer
If (Timer3(Index).Enabled = True Or Timer4(Index).Enabled = True) And Timer8(Index).Enabled = True Then   '���D����
    If Timer6.Enabled = True Then s(2 + 34 * Index) = 0: Timer10(Index).Enabled = False: Exit Sub
    If Timer3(Index).Enabled = True Then b(Index) = 1 Else b(Index) = -1 '��
    Image1(Index).Left = Image1(Index).Left - 20 * b(Index)
    Shape1(Index).Left = Shape1(Index).Left - 20 * b(Index)
    
    s(2 + 34 * Index) = s(2 + 34 * Index) + 1
    If pk_mod = 0 Then Call sppr(10, Index) Else Call pks(Index, Index * -1 + 1, Image1(), Image1(), 10) '�P�_�O�_����
    
    If s(2 + 34 * Index) >= 5 Then
        s(2 + 34 * Index) = 0: Timer10(Index).Enabled = False
        For f = 1 To ma
            ppm(9 + 1 * Index, f) = 0
        Next
    End If
Else
    Image1(Index).Left = Image1(Index).Left - 10 * (zxcv(3 + 6 * Index) * 2 - 1)
    Shape1(Index).Left = Shape1(Index).Left - 10 * (zxcv(3 + 6 * Index) * 2 - 1)
    s(2 + 34 * Index) = s(2 + 34 * Index) + 1
    If s(2 + 34 * Index) >= 5 Then
        If pk_mod = 0 Then Call sppr(9, Index) Else Call pks(Index, Index * -1 + 1, Image1(), Image1(), 9) '�P�_�O�_����
        
        Image1(Index).Left = Image1(Index).Left + 50 * (zxcv(3 + 6 * Index) * 2 - 1)
        Shape1(Index).Left = Shape1(Index).Left + 50 * (zxcv(3 + 6 * Index) * 2 - 1)
        s(2 + 34 * Index) = 0
        Timer10(Index).Enabled = False
    End If
End If
DoEvents
End Sub
Private Sub pks(ByVal a As Integer, ByVal b As Integer, c As Object, d As Object, ByVal e As Integer) 'pk�l��t��k a=�o�ʪ� b= �Q�o�ʪ� c=a���� d=b���� e=���ۺ���
If hp(b).Visible = True And hp(b * -1 + 1).Visible = True Then
    'If coll(c(a), d(b)) Or coll(d(b), c(a)) Then
    If coll(c(a), d(b)) Then
        If stong(3) = 0 And e = 3 Then s(15) = s(15) + 1: Unload c(a): Unload Timer29(a): met_already(a) = 1
        Call pksk((e), (b))
    End If
End If
End Sub
Private Sub pksk(ByVal a As Integer, ByVal b As Integer) 'pk�l��t��k-2 a=���ۺ��� b=�Q�o����
Dim Burst As Byte

If hp(b).Visible = True And hp(b * -1 + 1).Visible = True Then
    Timer36(b * -1 + 1).Enabled = False
    
    If a = 0 Or a = 2 Or a = 5 Then Call tai(a, b * -1 + 1)
    
    metor(b) = 0
    Timer5(kikyou(b)).Enabled = False
    
    If a < 8 And a <> 3 Then '�]��
        Y = ys(a)
        Burst = 1 '���F��bug(�]�����ۨS���z���]�w)
    Else '����
        Randomize
        Y = Int(Rnd * (100 - 10 + 1)) + 10
        If Int(Rnd * 5) + 1 = 5 Then
            For i = 1 To 10
                If p_which(0, b, i) = 0 Then
                    p_which(0, b, i) = 1
                    Burst = 3
                    Y = Y * 10 '�[�`�ˮ`
                    Call p_start(0, b, i, Image1())
                    Exit For
                End If
            Next
        Else
            Burst = 1 '�M�w���ˮ`���C���(�z���v)
        End If
    End If
    
    If khp_GPR(b) - Y / akiz(0) >= 1 Then
        khp_GPR(b) = khp_GPR(b) - Y / akiz(0)
        hp(b).Width = khp_GPR(b) \ 1
        label1(b).Caption = (khp_GPR(b) * akiz(0)) \ 1 & "/" & amax(0)
    Else
        aop(0, b) = aop(0, b) - 1
        If aop(0, b) >= 0 Then
            hp(b).Width = 200
            khp_GPR(b) = hp(b).Width
            Label13(b).Caption = "X " & aop(0, b)
            label1(b).Caption = (khp_GPR(b) * akiz(0)) \ 1 & amax(0)
        Else
            hp(b).Visible = False
            Image1(b).Visible = False
            Shape1(b).Visible = False
            metor(b) = 0
            zxcv(5 + 5 * b) = 1
            label1(b).Caption = "0/" & amax(0)
        End If
    End If
    ppm(7, b * -1 + 1) = ppm(7, b * -1 + 1) + 1
    
    Label11(b * -1 + 1).Caption = ppm(7, b * -1 + 1) & "�s��"
    Timer37(b * -1 + 1).Enabled = True
    Timer36(b * -1 + 1).Enabled = True
    Timer22(b + 1).Interval = 200
    Timer22(b + 1).Enabled = True '�Ȯɰ��y
    Call voice
    If te = 1 Then Call news(Image1(b), Y, Burst)  '���ˮ`
End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer) '������U��

Select Case KeyCode
    Case 113 'F2
        If Timer20(0 + 2 * f).Enabled = True And Timer20(1 + 2 * f).Enabled = True Or _
           Timer20(0 + 2 * f).Enabled = True And Timer21(1 + 2 * f).Enabled = True Or _
           Timer21(0 + 2 * f).Enabled = True And Timer20(1 + 2 * f).Enabled = True Or _
           Timer21(0 + 2 * f).Enabled = True And Timer21(1 + 2 * f).Enabled = True Then Call ex(0)
    Case 118 'F7
        For f = 0 To player_2
            aop(0, f) = aop(0, f) + 1
            Label13(f).Caption = "X " & aop(0, f)
            hp(f).Width = 200
            khp_GPR(f) = hp(f).Width
            label1(f).Caption = amax(0) & "/" & amax(0)
            Timer40(f).Enabled = False
            hp(f).FillColor = &HFFFF&
            label10(ff).Visible = True: image23(ff).Visible = True: label10(f).Caption = Val(label10(f).Caption) + 1 & " X"
        Next
    Case 119 'F8
        If F8_second = 99 Then F8_second = 0 Else F8_second = 99
End Select

If player_2 = 0 And (KeyCode = keya(1) Or KeyCode = keys(1) Or KeyCode = keyd(1)) Then
    player_2 = 1
    If kikyou(0) = 3 Then kikyou(1) = 0 Else kikyou(1) = kikyou(0) + 1
    Call p2
    Call p2_hex
    Call P12(1)
End If
For f = 0 To 0 + player_2
    If zxcv(5 + 5 * f) <> 1 Then
        If (zxcv(0 + 12 * f) = 1 And KeyCode = keyup(f)) Or (zxcv(0 + 12 * f) = 1 And KeyCode = keydown(f)) Then
        Else
            Select Case KeyCode
                Case keyleft(f) '��37
                    If ax(0 + 2 * f) = 1 Then run(0 + 2 * f) = 2
                    
                    Timer3(f).Enabled = True '����
                    zxcv(3 + 6 * f) = 1 '�ʵe���k�P�_
                Case keyup(f) '�W38
                    up(f) = 1
                    
                    dzxc(f) = -1 '�y�P���u�B�W�U�P�_
                    
                    If zxcv(0 + 12 * f) = 0 Then Timer1(f).Enabled = True '�D���ɲ���
                Case keyright(f) '�k39
                    If ax(1 + 2 * f) = 1 Then run(1 + 2 * f) = 2
                    
                    Timer4(f).Enabled = True '����
                    zxcv(3 + 6 * f) = 0 '�ʵe���k�P�_
                Case keydown(f) '�U40
                    down(f) = 1
                
                    dzxc(f) = 1 '�y�P�W�U�P�_
                
                    If zxcv(0 + 12 * f) = 0 Then Timer2(f).Enabled = True '�D���ɲ���
                Case keya(f) 'A65
                    If down(f) = 1 And play(0, f) = 0 Then
                        If Val(label10(f).Caption) > 0 Then play(0, f) = f + 1: Timer30(f).Enabled = True '�p�G���|��ɫh���۶���
                    Else
                        If up(f) = 1 Then
                            If Timer21(1 + 2 * f).Enabled = True Then Call mpss(kikyou(f) + 4, (f)) 'Ū���l��MP�t��k(���ۤG)
                        Else
                            If Timer21(0 + 2 * f).Enabled = True Then Call mpss(kikyou(f), (f)) 'Ū���l��MP�t��k
                        End If
                    End If
                Case keys(f) 'S ����83
                    zxcv(1 + 10 * f) = zxcv(1 + 10 * f) + 1
                    If zxcv(1 + 10 * f) = 1 Then Timer10(f).Enabled = True
                Case keyd(f) 'D ���D68
                    If zxcv(0 + 12 * f) = 0 Then
                        zxcv(0 + 12 * f) = 1
                        ttop(f) = Image1(f).Top
                        Timer8(f).Enabled = True
                    End If
            End Select
        End If
    End If
Next
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer) '����u�_��
For f = 0 To 0 + player_2
    If KeyCode = keyleft(f) Or KeyCode = keyup(f) Or KeyCode = keyright(f) Or KeyCode = keydown(f) Or KeyCode = keys(f) Or KeyCode = keya(f) Then a = f
    Select Case KeyCode
        Case keyleft(f) '��
            ax(0 + 2 * f) = 1
            Timer9(0 + 2 * f).Enabled = True
            ax(1 + 2 * f) = 0 '�����k�]
            Timer9(1 + 2 * f).Enabled = False '�����k�]
            
            run(0 + 2 * f) = 1 '�������A
            
            Timer3(f).Enabled = False
        Case keyup(f) '�W
            up(f) = 0
            Timer1(f).Enabled = False
            dzxc(f) = 0
        Case keyright(f) '�k
            ax(1 + 2 * f) = 1
            Timer9(1 + 2 * f).Enabled = True
            ax(0 + 2 * f) = 0 '�������]
            Timer9(0 + 2 * f).Enabled = False '�������]
            
            run(1 + 2 * f) = 1 '�������A
            
            Timer6.Enabled = False '�a�ϲ�������
            Timer4(f).Enabled = False
        Case keydown(f) '�U
            down(f) = 0
            Timer2(f).Enabled = False
            dzxc(f) = 0
        Case keys(f) 's����
            zxcv(1 + 10 * f) = 0
        Case keya(f) 'A
            Timer25(f).Enabled = True 'MP�W�[
    End Select
Next
End Sub
Private Sub Timer9_Timer(Index As Integer) '�]�B�P�_��
For f = 0 To 0 + player_2
    If Index = 0 + 2 * f Then ax(0 + 2 * f) = 0 Else ax(1 + 2 * f) = 0
Next
Timer9(Index).Enabled = False
End Sub
Private Sub Timer8_Timer(Index As Integer) '���D��
If X(Index) = 1 Then s(0 + 35 * Index) = s(0 + 35 * Index) + 1 Else s(0 + 35 * Index) = s(0 + 35 * Index) - 1
Image1(Index).Top = Image1(Index).Top - (12 - s(0 + 35 * Index)) * X(Index)
'shape1(index).Left = shape1(index).Left + 1 * x
'shape1(index).Width = shape1(index).Width - 2 * x
If s(0 + 35 * Index) = 11 Then
    If zxcv(6 + 7 * Index) = 0 Then
        s(0 + 35 * Index) = 12: X(Index) = -1
        zxcv(6 + 7 * Index) = zxcv(6 + 7 * Index) + 1
    Else
        zxcv(6 + 7 * Index) = zxcv(6 + 7 * Index) - 1
    End If
End If
If Image1(Index).Top >= ttop(Index) Then X(Index) = 1: zxcv(0 + 12 * Index) = 0: s(0 + 35 * Index) = 0: zxcv(6 + 7 * Index) = 0: Image1(Index).Top = ttop(Index): Timer8(Index).Enabled = False
DoEvents
End Sub
Private Sub Timer1_Timer(Index As Integer) '�W��
If metor(Index) = 0 Or Timer8(Index).Enabled = True Then Exit Sub

Image1(Index).Top = Image1(Index).Top - 10
Shape1(Index).Top = Image1(Index).Top + sh(Index)
If Image1(Index).Top < 0 + Line1.Y1 Then Image1(Index).Top = 0 + Line1.Y1: Shape1(Index).Top = Image1(Index).Top + sh(Index)
End Sub
Private Sub Timer2_Timer(Index As Integer) '�U��
If metor(Index) = 0 Or Timer8(Index).Enabled = True Then Exit Sub

Image1(Index).Top = Image1(Index).Top + 10
Shape1(Index).Top = Image1(Index).Top + sh(Index)
If Image1(Index).Top + Image1(Index).Height > Form1.ScaleHeight - 10 Then Image1(Index).Top = Form1.ScaleHeight - Image1(Index).Height - 10: Shape1(Index).Top = Image1(Index).Top + sh(Index)
End Sub
Private Sub Timer3_Timer(Index As Integer) '����
If metor(Index) = 0 Then Exit Sub

Image1(Index).Left = Image1(Index).Left - 10 * run(0 + 2 * Index)
Shape1(Index).Left = Image1(Index).Left + Image1(Index).Width \ 2
Timer5(kikyou(Index)).Enabled = True '�Ұʰʵe
If Image1(Index).Left <= 0 + 10 Then Image1(Index).Left = 0 + 10
End Sub
Private Sub Timer4_Timer(Index As Integer) '�k��
If metor(Index) = 0 Then Exit Sub

If Image1(Index).Left + Image1(Index).Width \ 2 >= (Form1.ScaleWidth \ 5) * 4 Then '�a�ϬO�_����
    If ok = 0 Then
        If hp(0).Visible = True And hp(1).Visible = True And player_2 = 1 Then
            If Image1(Index * -1 + 1).Left > 10 Then
                Timer6.Enabled = True
                Shape1(Index).Left = Image1(Index).Left + Image1(Index).Width \ 2 - 20
                If Timer6.Enabled = True Then Image1(Index * -1 + 1).Left = Image1(Index * -1 + 1).Left - 10 * run(1 + 2 * Index): Shape1(Index * -1 + 1).Left = Shape1(Index * -1 + 1).Left - 10 * run(1 + 2 * Index)
            Else
                Timer6.Enabled = False
                Call s_timer4((Index))
            End If
            Exit Sub
        Else
            Timer6.Enabled = True
            Shape1(Index).Left = Image1(Index).Left + Image1(Index).Width \ 2 - 20
            Exit Sub
        End If
    End If
End If
Call s_timer4((Index))
End Sub
Private Sub s_timer4(ByVal Index As Integer)
Image1(Index).Left = Image1(Index).Left + 10 * run(1 + 2 * Index)
Shape1(Index).Left = Image1(Index).Left + Image1(Index).Width \ 2 - 20
Timer5(kikyou(Index)).Enabled = True '�Ұʰʵe
If Image1(Index).Left + Image1(Index).Width > Form1.ScaleWidth - 10 Then Image1(Index).Left = Form1.ScaleWidth - Image1(Index).Width - 10
End Sub
Private Sub Timer5_Timer(Index As Integer) '����ʵe��
For f = 0 To 0 + player_2
    If kikyou(f) = Index Then a = f
Next
k(kikyou(a)) = k(kikyou(a)) + 1
Select Case kikyou(a)
    Case 0
        If k(kikyou(a)) = 6 Then k(kikyou(a)) = 0
        If metor(a) = 0 Then zxcv(3 + 6 * a) = zxcv(2 + 6 * a) '�ܱ𪺷N�Ӫ����k�ʵe����
        Image1(a).Picture = IIf(zxcv(3 + 6 * a) = 0, Image2(k(kikyou(a))).Picture, Image7(k(kikyou(a))).Picture)
    Case 1
        If k(kikyou(a)) = 5 Then k(kikyou(a)) = 0
        Image1(a).Picture = IIf(zxcv(3 + 6 * a) = 0, Image3(k(kikyou(a))).Picture, Image8(k(kikyou(a))).Picture)
    Case 2
        If k(kikyou(a)) = 4 Then k(kikyou(a)) = 0
        Image1(a).Picture = IIf(zxcv(3 + 6 * a) = 0, Image4(k(kikyou(a))).Picture, Image9(k(kikyou(a))).Picture)
    Case 3
        If k(kikyou(a)) = 6 Then k(kikyou(a)) = 0
        If metor(a) = 0 Then zxcv(3 + 6 * a) = zxcv(2 + 6 * a) '�y�P���u�B�����k�ʵe����
        Image1(a).Picture = IIf(zxcv(3 + 6 * a) = 0, Image5(k(kikyou(a))).Picture, Image10(k(kikyou(a))).Picture)
End Select
DoEvents
End Sub
Private Sub sico(a As Integer) '�|��ɱ����t��k
Randomize
If Int(Rnd * 3) + 1 = 2 Then
    If s(13) = 10 Then
        s(13) = 0
    Else
        s(13) = s(13) + 1
    End If
    Load Image20(s(13)) '�|��ɥ���
    Load Shape15(s(13)) '�|��ɼv�l
    Load Timer32(s(13)) '�|�����ʰʵe
    Load Timer33(s(13)) '�|��ɮ���
    
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
Private Sub Timer33_Timer(Index As Integer) '�|��ɮ���
sic(Index) = sic(Index) + 1
If sic(Index) >= 10 Then
    If sic(Index) Mod 2 = 0 Then Image20(Index).Visible = True Else Image20(Index).Visible = False
    
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
Private Sub Timer32_Timer(Index As Integer) '�|�����ʰʵe

If sk(Index) = 7 Then
    sk(Index) = 0
Else
    sk(Index) = sk(Index) + 1
End If

Image20(Index).Picture = Image21(sk(Index)).Picture

For f = 0 To 0 + player_2
    If Timer6.Enabled = True Then Image20(Index).Left = Image20(Index).Left - 20 * run(1 + 2 * f): Shape15(Index).Left = Shape15(Index).Left - 20 * run(1 + 2 * f) '�a�ϲ���

    If coll(Image20(Index), Image1(f)) Then
    
        For ff = 0 To player_2
            label10(ff).Visible = True: image23(ff).Visible = True: label10(ff).Caption = Val(label10(ff).Caption) + 1 & " X": Image20(Index).Left = -1111: Shape15(Index).Left = -1111
        Next
        
        If khp_GPR(f) >= 4000 / akiz(0) Then '�ɦ�
            hp(f).Width = 200
            khp_GPR(f) = hp(f).Width
        Else
            khp_GPR(f) = khp_GPR(f) + 1000 / akiz(0)
            hp(f).Width = khp_GPR(f) \ 1
        End If
        label1(f).Caption = (khp_GPR(f) * akiz(0)) \ 1 & "/" & amax(0)
        
        If khp_GPR(f) >= (amax(0) / akiz(0)) / 3 Then Timer40(f).Enabled = False: hp(f).FillColor = &HFFFF&
        
        If kmp_GPR(f) >= 4000 / akiz(0) Then '���]
            mp(f).Width = 200
            kmp_GPR(f) = mp(f).Width
        Else
            kmp_GPR(f) = kmp_GPR(f) + 1000 / akiz(0)
            mp(f).Width = kmp_GPR(f) \ 1
        End If
        label2(f).Caption = (kmp_GPR(f) * akiz(0)) \ 1 & "/" & amax(0)
    End If
Next
End Sub
Private Sub Timer38_Timer(Index As Integer) 'Hp�W�[�t��k
hp(Index).Width = hp(Index).Width + 1
If hp(Index).Width >= 200 Then hp(Index).Width = 200
label1(Index).Caption = hp(Index).Width * akiz(0) & "/" & amax(0)
End Sub
Private Sub Timer25_Timer(Index As Integer) 'MP�W�[�t��k
Dim a As Single

If khp_GPR(Index) <= (amax(0) / akiz(0)) / 3 Then Timer25(Index).Interval = 100 Else Timer25(Index).Interval = 200

kmp_GPR(Index) = kmp_GPR(Index) + 10 / akiz(0)
mp(Index).Width = kmp_GPR(Index) \ 1
If mp(Index).Width >= 200 Then mp(Index).Width = 200: kmp_GPR(Index) = mp(Index).Width
label2(Index).Caption = (kmp_GPR(Index) * akiz(0)) \ 1 & "/" & amax(0)
End Sub
Public Sub kizzs(a As Integer) '����l��t��k��
Dim Burst As Byte, ym As Integer

If s(10) = 0 And hp(a).Visible = True Then

    Call voice
    
    Randomize
    ym = Int(Rnd * 100) + 1
    If ma = 1 Then ym = ym * 3
    
    If Int(Rnd * 5) + 1 = 5 Then
        For i = 1 To 10 '�M��Ŷ������ɤl
            If p_which(0, a, i) = 0 Then
                Burst = 3
                ym = ym * 10
                p_which(0, a, i) = 1
                Call p_start(0, a, i, Image1())
                Exit For
            End If
        Next
    Else
        Burst = 1
    End If
    
    If Image1(a).Left > kb And Image1(a).Left + Image1(a).Width + kb < Me.ScaleWidth Then Call unadd(Image1(a), Shape1(a), 0) '���h
    
    If khp_GPR(a) - ym / akiz(0) <= 0 Then '�p�G�S��h
        Timer40(a).Enabled = False '�{�{����
        hp(a).FillColor = &HFFFF& '�C�⬰����
        aop(0, a) = aop(0, a) - 1
        If aop(0, a) >= 0 Then
            Label13(a).Caption = "X " & aop(0, a)
            hp(a).Width = 200
            khp_GPR(a) = hp(a).Width
            label1(a).Caption = hp(a).Width * akiz(0) & amax(0)
        Else
            hp(a).Visible = False
            Image1(a).Visible = False
            Shape1(a).Visible = False
            label1(a).Caption = "0/" & amax(0)
            metor(a) = 0
            zxcv(5 + 5 * a) = 1
            If hp(0).Visible = False And hp(player_2).Visible = False Then Call ex(1): Exit Sub
        End If
    Else '����
        khp_GPR(a) = khp_GPR(a) - ym / akiz(0)
        hp(a).Width = khp_GPR(a) \ 1
        label1(a).Caption = (khp_GPR(a) * akiz(0)) \ 1 & "/" & amax(0)
        If khp_GPR(a) <= (amax(0) / akiz(0)) / 3 Then Timer40(a).Interval = 10: Timer40(a).Enabled = True
    End If
    
    If te = 1 Then Call news(Image1(a), ym, Burst) '�O�_���ˮ`
    
    DoEvents
End If
End Sub
Private Sub voice() '�I���n��
'-----�_������
    Set dsx = dxa.DirectSoundCreate("")
    dsx.SetCooperativeLevel Me.hWnd, DSSCL_NORMAL
    bu.lFlags = DSBCAPS_STATIC
    Set dsbx = dsx.CreateSoundBufferFromFile(("Data\�I��.wav"), bu, wf)
    dsbx.play DSBPLAY_DEFAULT '''''''''''''''''''''''�w�]����(�榸).
'-----�_������
End Sub
Public Sub sppr(a As Integer, b As Integer) '�P�_�O�_������ a=�� or ����, b=1P or 2P
For f = 1 To ma
    If ppm(1, f) <> 1 Then
        If ppm(9 + 1 * b, f) <> 1 Then
            If Timer8(b).Enabled = True Then '����
                If jump(Image6(f), (b)) Then
                    If coll4(Image1(b), Image6(f), (b)) Then pps = Image6(f).Index: Call power(a, b): If a = 10 Then ppm(9 + 1 * b, f) = 1
                End If
            Else '9����
                'If coll(Image1(b), Image6(f)) Or coll(Image6(f), Image1(b)) Then pps = Image6(f).Index: Call power(a, b)
                If coll(Image1(b), Image6(f)) Then pps = Image6(f).Index: Call power(a, b)
            End If
        End If
    End If
    DoEvents
Next
End Sub
Private Function jump(a As Object, b As Integer) As Boolean '�D�����D�ɡA�����I��
If ttop(b) > a.Top And ttop(b) < a.Top + a.Height Or ttop(b) + Image1(b).Height > a.Top And ttop(b) + Image1(b).Height < a.Top + a.Height Then jump = True
End Function
Public Sub power(a As Integer, c As Integer) '�ĤH�l��t��k�� a=���@�ا���,c= 1P or 2P
'---------------------------------------------------------------------------------------
Dim b As Integer, d As Integer, e As Integer, Burst As Byte 'd=���@�ӵ��ۤ��Ȧs,b,e=�̤j�̤p�ˮ`��

Label7(c).Visible = True: Label14(c).Visible = True: Label5(c).Visible = True: xhp(c).Visible = True: Shape6(c).Visible = True: Label11(c).Visible = True '�Ǫ���q���
Timer37(c).Enabled = False '�s���Ƥ��{
Timer36(c).Enabled = False '�@�w�ɶ���q

ppm(5, pps) = 1 '����ɪ��S��
Timer11(pps).Enabled = False '���ʵe����

If a = 0 Or a = 2 Or a = 5 Then Call tai(a, c)

d = a '�P�_�O���@�ӵ���

If hp_GPR(pps) = 0 Then
    xhp(c).Width = Shape6(c).Width
    xhp(c).Visible = True
    hp_GPR(pps) = xhp(c).Width
Else
    xhp(c).Visible = True
    xhp(c).Width = hp_GPR(pps) \ 1
    Label14(c).Caption = "X " & ppm(20, pps)
End If

Call voice '�n��

'�ˮ`�t��k
If d < 8 And d <> 3 Then '�]��
    Y = ys(d) + F8_second
Else '����
    Select Case d
        Case 3 '�y�P���u�B
            If stong(3) = 2 Then e = 200: b = 50
        Case 8 '�}�]�s��
            If stong(kikyou(c) + 4) = 2 Then e = 200: b = 50
        Case 9 '����
            e = 200: b = 100
        Case 10 '����
            e = 100: b = 50
    End Select
    
    Randomize
    Y = Int(Rnd * (100 + e - 10 - b + 1)) + 10 + b + F8_second
    
    If Int(Rnd * 5) + 1 = 5 Then '�z���v
        For i = 1 To 10
            If p_which(1, pps, i) = 0 Then
                Y = Y * 10
                Burst = 2
                p_which(1, pps, i) = 1
                Call p_start(1, pps, i, Image6())
                Exit For
            End If
        Next
    End If
    
End If
'/�ˮ`�t��k

If hp_GPR(pps) - Y / akiz(1) <= 0 Then
    ppm(20, pps) = ppm(20, pps) - 1
    If ppm(20, pps) >= 0 Then
        Label14(c).Caption = "X " & ppm(20, pps)
        xhp(c).Width = amax(1) / akiz(1)
        hp_GPR(pps) = xhp(c).Width
        Label5(c).Caption = (hp_GPR(pps) * akiz(1)) \ 1 & "/" & amax(1)
    Else
        hp_GPR(pps) = 0
        xhp(c).Visible = False
        Label5(c).Caption = hp_GPR(pps) & "/" & amax(1)
        On Error Resume Next
        Load Timer7(pps)
        ppm(1, pps) = 1
        Image6(pps).Top = Image6(pps).Top + 15 '�����U
        Timer7(pps).Enabled = True '�Ұʳ������ʵe
        'Call p_start(pps, 0, Image6()) '���Ͳɤl
    End If
    
    If ma = 1 And ppm(20, pps) = 0 Then '�]���Q����@����
        xhp(c).FillColor = &HFFFF&
        Shape6(c).FillColor = &HFF&
        boss_power = 3
    End If
Else
    hp_GPR(pps) = hp_GPR(pps) - Y / akiz(1)
    xhp(c).Width = hp_GPR(pps) \ 1
    
    Label5(c).Caption = (hp_GPR(pps) * akiz(1)) \ 1 & "/" & amax(1)
End If

If te = 1 Then Call news(Image6(pps), Y, Burst) '�O�_���ˮ`

If P_Combo(c) < P_Combo(c * -1 + 1) Then P_Combo(c) = P_Combo(c * -1 + 1)
For i = 0 To player_2 '�P�B��s
    P_Combo(i) = P_Combo(i) + 1 '�s����
    For j = 1 To P_Combo(i) \ 10
        f = f & "!"
    Next
    Label11(i).Caption = P_Combo(i) & " �s��" & f
Next


If d > 2 And d <> 5 And d <> 7 Then
    Call unadd(Image6(pps), Shape2(pps), 1) '���h
End If

'����
label9(c).Caption = Val(label9(c).Caption) + 50
If Val(label9(c).Caption) >= 1000000 Then label9(c).Caption = 1000000

If ppm(19, pps) = 0 Then Timer22(pps).Interval = 200
Timer22(pps).Enabled = True

Timer36(c).Enabled = True '�@�w�ɶ���q
Timer37(c).Enabled = True '�s���ư{�t
DoEvents
End Sub
Private Sub unadd(a As Object, b As Object, c As Byte) '���h c = 0,1 �M�w1P��2P
If zxcv(3 + 6 * c) = -c + 1 Then
    a.Left = a.Left + kb
    b.Left = a.Left + a.Width \ 2
Else
    a.Left = a.Left - kb
    b.Left = (a.Left + a.Width \ 2) - 20
End If
End Sub
Public Sub news(a As Object, b As Integer, ByVal d As Integer) '��ܶˮ`�t��k��
Select Case d
    Case 0
        c = RGB(255, 0, 0)
    Case 1
        c = RGB(255, 255, 0)
    Case 2, 3
        c = RGB(255, 0, 255)
        d = d - 2 '���L��akiz(��j���v���T)
End Select

'���ͪ���
s(5) = s(5) + 1
If s(5) = 30000 Then s(5) = 0
Load Label4(s(5))
'���t��m
Label4(s(5)).Top = a.Top - Label4(s(5)).Height
Label4(s(5)).Left = a.Left + a.Width \ 2
'�]�w�ݩ�
Label4(s(5)).ForeColor = c '�C��
Label4(s(5)).Caption = b
Label4(s(5)).Visible = True
Load Timer16(s(5)) '���ͤ@�ӭӧO����    �p�ɾ�
Load Timer17(s(5)) '���ͤ@�ӭӧO���ʲ����p�ɾ�
Timer16(s(5)).Enabled = True '�}�Ҳ��ʭp�ɾ�
Timer17(s(5)).Enabled = True '�}�Ҳ����p�ɾ�
DoEvents
End Sub
Private Sub Timer16_Timer(Index As Integer) '��ܶˮ`�����ʡ�
Label4(Index).Top = Label4(Index).Top - 10
DoEvents
End Sub
Private Sub Timer17_Timer(Index As Integer) '�����ˮ`��ܡ�
Unload Timer16(Index)
Unload Label4(Index)
Unload Timer17(Index)
DoEvents
End Sub
Private Sub Timer36_Timer(Index As Integer) '�@�w�ɶ����Ǫ�����q����
Label14(Index).Visible = False
Label7(Index).Visible = False
Label5(Index).Visible = False
xhp(Index).Visible = False
Shape6(Index).Visible = False
Timer37(Index).Enabled = False
Label11(Index).Visible = False
P_Combo(Index) = 0
If pk_mod = 1 Then
    ppm(7, Index) = 0
    Label11(b * -1 + 1).Caption = ppm(7, b * -1 + 1) & "�s��"
End If
End Sub
Private Sub Timer37_Timer(Index As Integer) '�s���ư{�t
s(21) = (s(21) + 1) Mod 2
If s(21) = 0 Then Label11(Index).Visible = True Else Label11(Index).Visible = False
End Sub
Private Sub Timer6_Timer() '�D�I�����ʡ�'���{�Ǫ������{���X���C�Y��U�I�I
Form1.AutoRedraw = True
If Image1(0).Left >= Image1(player_2).Left Then a = 0 Else a = 1

dx = dx + 20 * run(1 + 2 * a)

If s(46) = 0 Then
    If s(42) = 0 Then
        s(42) = 1
        s(1) = s(1) + 1
        c = s(1) \ 5 + 1
        If s(1) Mod 5 = 0 Then b = 5: c = c - 1: boss = c Else b = s(1) Mod 5
        Me.Caption = Program_Name & "STAGE " & c & "-" & b
    End If
    
    If dx >= 1024 Then
        dx = 0
        ok = 1
        s(42) = 0
        
        Timer6.Enabled = False
        If s(1) Mod 5 = 0 Then zxcv(7) = 1
        'If s(1) = 5 Then s(1) = 0 '����
    End If
End If

If s(47) = 0 Then Call BackPicture_Move '�p�G�S���¹��hø��
DoEvents
End Sub
Private Sub BackPicture_Move() '�I������ø��
If dx >= 1024 Then dx = 0

If dx <= 100 Then '=====�j��564��N�W�ϤF=====
    Form1.PaintPicture Form1.Picture, 0, 0, 924, 708, dx, 0, 924, 708, vbSrcCopy
Else '�W��form1�I�� �ƻs��          0,0,��e,��,dx,0,��e,��
    Form1.PaintPicture Form1.Picture, 0, 0, 1024 - dx, 708, dx, 0, 1024 - dx, 708, vbSrcCopy
    '�W��form1�I�� �ƻs��           0,0,�ϼe-dx,��,dx,0,�ϼe-dx,��
    Form1.PaintPicture Form1.Picture, 1024 - dx, 0, dx - 100, 708, 0, 0, dx - 100, 708, vbSrcCopy
    '�W��form1�I�� �ƻs��           �ϼe-dx,0,��e+(dx-�ϼe),��,0,0,��e+(dx+�ϼe),��
End If
End Sub
Private Sub Timer15_Timer() 'GO��ܡ�
s(4) = s(4) + 1
If s(4) Mod 2 = 0 Then Label3.Visible = False Else Label3.Visible = True

If s(16) = 0 And s(1) > 1 And s(1) Mod 5 <> 0 And s(42) = 1 Then s(16) = 1: ma_aop = s(1): Call desd(0)
If ok = 1 Then
    If s(1) <> 1 Then
        s(16) = 0
        If n = ma And s(1) Mod 5 <> 0 Then ok = 0: Exit Sub
        Timer15.Enabled = False: Label3.Visible = False
        If s(1) Mod 5 = 0 And zxcv(7) = 1 Then zxcv(7) = 0: ma_aop = s(1): Call desd(0)
    Else
        ok = 0
    End If
End If
DoEvents
End Sub
