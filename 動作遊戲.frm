VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '��u�T�w
   Caption         =   "���~���I (�ʧ@�C�� v 3.3) �p��s��"
   ClientHeight    =   8520
   ClientLeft      =   4635
   ClientTop       =   1755
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "�ʧ@�C��.frx":0000
   ScaleHeight     =   568
   ScaleMode       =   3  '����
   ScaleWidth      =   794
   Begin VB.Timer Timer24 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   6480
      Top             =   2640
   End
   Begin VB.Timer Timer28 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   5400
      Top             =   4680
   End
   Begin VB.Timer Timer37 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4680
      Top             =   1080
   End
   Begin VB.Timer Timer36 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   7080
      Top             =   360
   End
   Begin VB.Timer Timer35 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   2760
      Top             =   3120
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
      Top             =   2040
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Index           =   6
      Interval        =   200
      Left            =   2400
      Top             =   2040
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Index           =   5
      Interval        =   10
      Left            =   2040
      Top             =   2040
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Index           =   4
      Interval        =   10
      Left            =   1680
      Top             =   2040
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   100
      Left            =   1200
      Top             =   2040
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   30
      Left            =   840
      Top             =   2040
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   100
      Left            =   480
      Top             =   2040
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
   Begin VB.Timer Timer27 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   30
      Left            =   3720
      Top             =   3120
   End
   Begin VB.Timer Timer26 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   240
      Top             =   480
   End
   Begin VB.Timer Timer25 
      Interval        =   1000
      Left            =   2760
      Top             =   1080
   End
   Begin VB.Timer Timer23 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   6000
      Top             =   2520
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
      Top             =   2040
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
      Top             =   3840
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
   Begin VB.Shape Shape19 
      BorderStyle     =   0  '�z��
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  '���
      Height          =   135
      Index           =   0
      Left            =   3600
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape17 
      BorderStyle     =   0  '�z��
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  '���
      Height          =   255
      Index           =   0
      Left            =   6360
      Shape           =   3  '���
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label11 
      Alignment       =   1  '�a�k���
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "0 �s��"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   405
      Left            =   2040
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Shape Shape16 
      BorderWidth     =   7
      Height          =   255
      Index           =   0
      Left            =   2760
      Shape           =   2  '����
      Top             =   2880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image24 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   1
      Left            =   12960
      Picture         =   "�ʧ@�C��.frx":240042
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image24 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   0
      Left            =   11520
      Picture         =   "�ʧ@�C��.frx":240303
      Top             =   8640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image23 
      Height          =   645
      Left            =   4080
      Picture         =   "�ʧ@�C��.frx":2405C5
      Top             =   960
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label10 
      Alignment       =   1  '�a�k���
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "0 X"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Left            =   3480
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Image Image22 
      Height          =   330
      Index           =   0
      Left            =   1080
      Picture         =   "�ʧ@�C��.frx":240B13
      Top             =   5160
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label9 
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
      Left            =   2985
      TabIndex        =   9
      Top             =   120
      Width           =   210
   End
   Begin VB.Image Image21 
      BorderStyle     =   1  '��u�T�w
      Height          =   705
      Index           =   7
      Left            =   14520
      Picture         =   "�ʧ@�C��.frx":240F99
      Top             =   7080
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Image21 
      BorderStyle     =   1  '��u�T�w
      Height          =   705
      Index           =   6
      Left            =   15240
      Picture         =   "�ʧ@�C��.frx":2414EC
      Top             =   7200
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Image21 
      BorderStyle     =   1  '��u�T�w
      Height          =   705
      Index           =   5
      Left            =   15960
      Picture         =   "�ʧ@�C��.frx":241A1A
      Top             =   7080
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Image21 
      BorderStyle     =   1  '��u�T�w
      Height          =   705
      Index           =   4
      Left            =   16680
      Picture         =   "�ʧ@�C��.frx":241F95
      Top             =   7080
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Image21 
      BorderStyle     =   1  '��u�T�w
      Height          =   705
      Index           =   3
      Left            =   14400
      Picture         =   "�ʧ@�C��.frx":2424E2
      Top             =   6240
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Image21 
      BorderStyle     =   1  '��u�T�w
      Height          =   705
      Index           =   2
      Left            =   15120
      Picture         =   "�ʧ@�C��.frx":242A48
      Top             =   6360
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Image21 
      BorderStyle     =   1  '��u�T�w
      Height          =   705
      Index           =   1
      Left            =   15840
      Picture         =   "�ʧ@�C��.frx":242F87
      Top             =   6240
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Image21 
      BorderStyle     =   1  '��u�T�w
      Height          =   705
      Index           =   0
      Left            =   16680
      Picture         =   "�ʧ@�C��.frx":2434D5
      Top             =   6240
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Image20 
      Height          =   645
      Index           =   0
      Left            =   12120
      Picture         =   "�ʧ@�C��.frx":243A25
      Top             =   7440
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Image Image19 
      Height          =   1005
      Left            =   2160
      Picture         =   "�ʧ@�C��.frx":243F75
      Top             =   6480
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image Image18 
      Height          =   270
      Left            =   720
      Picture         =   "�ʧ@�C��.frx":245524
      Top             =   6840
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image Image17 
      Height          =   2775
      Left            =   480
      Picture         =   "�ʧ@�C��.frx":245C24
      Top             =   5880
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Shape Shape14 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      Height          =   375
      Left            =   480
      Shape           =   3  '���
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
      Alignment       =   2  '�m�����
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "��A"
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
      Left            =   8730
      TabIndex        =   8
      Top             =   0
      Width           =   540
   End
   Begin VB.Shape delay 
      BorderColor     =   &H80000005&
      FillColor       =   &H0080FF80&
      FillStyle       =   0  '���
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
      Picture         =   "�ʧ@�C��.frx":2480C1
      Top             =   10080
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image14 
      Height          =   315
      Index           =   0
      Left            =   3960
      Picture         =   "�ʧ@�C��.frx":24820C
      Top             =   10440
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image13 
      Height          =   315
      Index           =   0
      Left            =   3960
      Picture         =   "�ʧ@�C��.frx":248358
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
      BackStyle       =   0  '�z��
      Caption         =   "Game Over..."
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
      Left            =   5160
      TabIndex        =   7
      Top             =   4200
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label Label7 
      Alignment       =   2  '�m�����
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
      Left            =   5700
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
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
      Shape           =   2  '����
      Top             =   2640
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label6 
      Alignment       =   2  '�m�����
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "A"
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
      Left            =   10800
      TabIndex        =   5
      Top             =   0
      Width           =   180
   End
   Begin VB.Label Label2 
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
      Left            =   960
      TabIndex        =   1
      Top             =   1080
      Width           =   1620
   End
   Begin VB.Shape delay 
      BorderColor     =   &H80000005&
      FillColor       =   &H0080FF80&
      FillStyle       =   0  '���
      Height          =   855
      Index           =   0
      Left            =   10800
      Top             =   360
      Width           =   240
   End
   Begin VB.Shape mp 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '���
      Height          =   375
      Left            =   240
      Top             =   1080
      Width           =   3000
   End
   Begin VB.Shape Shape8 
      FillStyle       =   0  '���
      Height          =   135
      Index           =   0
      Left            =   9120
      Shape           =   2  '����
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  '���
      Height          =   255
      Index           =   0
      Left            =   9120
      Shape           =   3  '���
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   1200
      Left            =   6000
      Picture         =   "�ʧ@�C��.frx":2484A4
      Top             =   4200
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
      Left            =   10530
      TabIndex        =   3
      Top             =   4680
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label5 
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
      Left            =   5280
      TabIndex        =   4
      Top             =   360
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
      Top             =   2520
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Label1 
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
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   1620
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FF8080&
      FillStyle       =   0  '���
      Height          =   375
      Left            =   240
      Top             =   1080
      Width           =   3000
   End
   Begin VB.Shape hp 
      BorderColor     =   &H80000005&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '���
      Height          =   375
      Left            =   240
      Top             =   600
      Width           =   3000
   End
   Begin VB.Shape xhp 
      BorderColor     =   &H80000005&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '���
      Height          =   375
      Left            =   4560
      Top             =   360
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '���
      Height          =   375
      Left            =   4560
      Top             =   360
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '���
      Height          =   375
      Left            =   240
      Top             =   600
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
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   7
      Left            =   12000
      Picture         =   "�ʧ@�C��.frx":248683
      Top             =   10440
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   6
      Left            =   12120
      Picture         =   "�ʧ@�C��.frx":2489A1
      Top             =   10320
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   5
      Left            =   12240
      Picture         =   "�ʧ@�C��.frx":248C85
      Top             =   10200
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   4
      Left            =   12360
      Picture         =   "�ʧ@�C��.frx":248F53
      Top             =   10080
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   3
      Left            =   12480
      Picture         =   "�ʧ@�C��.frx":249214
      Top             =   9960
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   2
      Left            =   12600
      Picture         =   "�ʧ@�C��.frx":2494E2
      Top             =   9840
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   1
      Left            =   12720
      Picture         =   "�ʧ@�C��.frx":2497CB
      Top             =   9720
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   0
      Left            =   12840
      Picture         =   "�ʧ@�C��.frx":249AE9
      Top             =   9600
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image6 
      Height          =   960
      Index           =   0
      Left            =   10080
      Picture         =   "�ʧ@�C��.frx":249DE9
      Top             =   4920
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   1
      Left            =   12960
      Picture         =   "�ʧ@�C��.frx":24A0E6
      Top             =   11640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   2
      Left            =   12840
      Picture         =   "�ʧ@�C��.frx":24A403
      Top             =   11760
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   0
      Left            =   13080
      Picture         =   "�ʧ@�C��.frx":24A6E3
      Top             =   11520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   3
      Left            =   12720
      Picture         =   "�ʧ@�C��.frx":24A9E0
      Top             =   11880
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   4
      Left            =   12600
      Picture         =   "�ʧ@�C��.frx":24ACB2
      Top             =   12000
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   6
      Left            =   12360
      Picture         =   "�ʧ@�C��.frx":24AF74
      Top             =   12240
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   7
      Left            =   12240
      Picture         =   "�ʧ@�C��.frx":24B250
      Top             =   12360
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   5
      Left            =   12480
      Picture         =   "�ʧ@�C��.frx":24B56D
      Top             =   12120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   5
      Left            =   9840
      Picture         =   "�ʧ@�C��.frx":24B83F
      Top             =   10320
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   4
      Left            =   9960
      Picture         =   "�ʧ@�C��.frx":24BA66
      Top             =   10200
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   3
      Left            =   10080
      Picture         =   "�ʧ@�C��.frx":24BC9F
      Top             =   10080
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   2
      Left            =   10200
      Picture         =   "�ʧ@�C��.frx":24BECF
      Top             =   9960
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   1
      Left            =   10440
      Picture         =   "�ʧ@�C��.frx":24C108
      Top             =   9720
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image9 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   3
      Left            =   8160
      Picture         =   "�ʧ@�C��.frx":24C32F
      Top             =   12120
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image9 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   2
      Left            =   8280
      Picture         =   "�ʧ@�C��.frx":24C514
      Top             =   12000
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image9 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   1
      Left            =   8400
      Picture         =   "�ʧ@�C��.frx":24C74F
      Top             =   11880
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   4
      Left            =   6480
      Picture         =   "�ʧ@�C��.frx":24C934
      Top             =   10320
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   3
      Left            =   6600
      Picture         =   "�ʧ@�C��.frx":24CB17
      Top             =   10200
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   2
      Left            =   6720
      Picture         =   "�ʧ@�C��.frx":24CD11
      Top             =   10080
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   1
      Left            =   6840
      Picture         =   "�ʧ@�C��.frx":24CF0B
      Top             =   9960
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   5
      Left            =   4560
      Picture         =   "�ʧ@�C��.frx":24D111
      Top             =   9960
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   4
      Left            =   4680
      Picture         =   "�ʧ@�C��.frx":24D407
      Top             =   9840
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   3
      Left            =   4800
      Picture         =   "�ʧ@�C��.frx":24D6FF
      Top             =   9720
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   2
      Left            =   4920
      Picture         =   "�ʧ@�C��.frx":24D9DC
      Top             =   9600
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   1
      Left            =   5040
      Picture         =   "�ʧ@�C��.frx":24DCD4
      Top             =   9480
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   0
      Left            =   10560
      Picture         =   "�ʧ@�C��.frx":24DFCA
      Top             =   9600
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image9 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   0
      Left            =   8520
      Picture         =   "�ʧ@�C��.frx":24E1F8
      Top             =   11760
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   0
      Left            =   6960
      Picture         =   "�ʧ@�C��.frx":24E4B5
      Top             =   9840
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   0
      Left            =   5160
      Picture         =   "�ʧ@�C��.frx":24E6E9
      Top             =   9360
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   3
      Left            =   8160
      Picture         =   "�ʧ@�C��.frx":24E9E1
      Top             =   10080
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   2
      Left            =   8280
      Picture         =   "�ʧ@�C��.frx":24EC96
      Top             =   9960
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   1
      Left            =   8400
      Picture         =   "�ʧ@�C��.frx":24EE75
      Top             =   9840
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   0
      Left            =   8520
      Picture         =   "�ʧ@�C��.frx":24F0AE
      Top             =   9720
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   0
      Left            =   10320
      Picture         =   "�ʧ@�C��.frx":24F28D
      Top             =   11640
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   1
      Left            =   10200
      Picture         =   "�ʧ@�C��.frx":24F4B4
      Top             =   11760
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   2
      Left            =   10080
      Picture         =   "�ʧ@�C��.frx":24F6E2
      Top             =   11880
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   3
      Left            =   9960
      Picture         =   "�ʧ@�C��.frx":24F911
      Top             =   12000
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   4
      Left            =   9840
      Picture         =   "�ʧ@�C��.frx":24FB41
      Top             =   12120
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   5
      Left            =   9720
      Picture         =   "�ʧ@�C��.frx":24FD70
      Top             =   12240
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   0
      Left            =   6480
      Picture         =   "�ʧ@�C��.frx":24FF9E
      Top             =   11880
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   1
      Left            =   6360
      Picture         =   "�ʧ@�C��.frx":2501D0
      Top             =   12000
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   2
      Left            =   6240
      Picture         =   "�ʧ@�C��.frx":2503D4
      Top             =   12120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   3
      Left            =   6120
      Picture         =   "�ʧ@�C��.frx":2505D2
      Top             =   12240
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  '��u�T�w
      Height          =   1020
      Index           =   4
      Left            =   6000
      Picture         =   "�ʧ@�C��.frx":2507CC
      Top             =   12360
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   0
      Left            =   4800
      Picture         =   "�ʧ@�C��.frx":2509AB
      Top             =   11520
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   1
      Left            =   4680
      Picture         =   "�ʧ@�C��.frx":250CA1
      Top             =   11640
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   2
      Left            =   4560
      Picture         =   "�ʧ@�C��.frx":250F9A
      Top             =   11760
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   3
      Left            =   4440
      Picture         =   "�ʧ@�C��.frx":251293
      Top             =   11880
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   4
      Left            =   4320
      Picture         =   "�ʧ@�C��.frx":251572
      Top             =   12000
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  '��u�T�w
      Height          =   1260
      Index           =   5
      Left            =   4200
      Picture         =   "�ʧ@�C��.frx":25186B
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
      Shape           =   2  '����
      Top             =   4560
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  '���
      Height          =   255
      Index           =   0
      Left            =   10440
      Shape           =   2  '����
      Top             =   5880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  '���
      Height          =   255
      Left            =   6360
      Shape           =   2  '����
      Top             =   5280
      Width           =   495
   End
   Begin VB.Image Image12 
      Height          =   13365
      Left            =   15360
      Picture         =   "�ʧ@�C��.frx":251B64
      Top             =   -960
      Visible         =   0   'False
      Width           =   18270
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      Height          =   255
      Index           =   0
      Left            =   3960
      Shape           =   2  '����
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Shape12 
      FillColor       =   &H00400040&
      FillStyle       =   0  '���
      Height          =   1215
      Left            =   4080
      Shape           =   2  '����
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape13 
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  '���
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
      Y1              =   760
      Y2              =   816
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      Visible         =   0   'False
      X1              =   16
      X2              =   136
      Y1              =   760
      Y2              =   816
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   5
      Visible         =   0   'False
      X1              =   136
      X2              =   256
      Y1              =   736
      Y2              =   792
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      Visible         =   0   'False
      X1              =   16
      X2              =   136
      Y1              =   792
      Y2              =   736
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      Visible         =   0   'False
      X1              =   14
      X2              =   262
      Y1              =   792
      Y2              =   792
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      Visible         =   0   'False
      X1              =   16
      X2              =   256
      Y1              =   760
      Y2              =   760
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   0
      Shape           =   2  '����
      Top             =   11040
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
      X1              =   136
      X2              =   152
      Y1              =   168
      Y2              =   176
   End
   Begin VB.Shape Shape18 
      BorderColor     =   &H0080FFFF&
      BorderStyle     =   0  '�z��
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  '���
      Height          =   255
      Index           =   0
      Left            =   6840
      Top             =   5880
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ���۳Ъ��ƥZ�����{�������N�����Ƶ{�����p�ɾ�----�p��^^v

Dim s(32) As Integer '0)���D�p�� 1)�a�ϭp�� 2)����^���p�� 3)���W�d�� 4)GO���{�{ 5)��ܶˮ` 6)A delay�{�{ 7)�O�_����������� 8)�l�u���ͼ� 9)��������p�� 10)�y�P���u�B���ͭp�� 11)���۶��� 12)�W����(����)���p������ 13)�|��ɭp�� 14)�}�]�s�]���p�� 15) �y�P���u�B�L�k����ɤ��p�� 16)�}�]�s�] 17)�ݤ뺧�� 18)�ݤ뺧������� 19)�W���ʤ��_�i���� 20)������𵲧��p�� 21)�s���ư{�t 22)�W���ʤ��}�]�s�]�D�P�B�o�g 23)�}�]�s�]�T�{�o�g���� 24)�ݤ몺�ˮ`�P�_ 25)���W���𵲧��P�_ 26)���W�����P�_ 27)�ɹp�����p�� 28)���W����ˮ` 29)�ɹp����D�P�B�X�{ 30)�ɹp�����P�_ 31)�ɹp���𲾰��p�� 32)�W���ʫD���𲾰��p��
Dim sic(10) As Integer '�|��ɮ���
Dim zxcv(7) As Integer '0)������D�ɤ��i�W�U���� 1)����𤣥i���� 2)'�y�P���u�B�����k�ʵe���� 3)�P�_���⩹���Ω��k 4)����u��Ū�@�������{�� 5)������L�����(�Ω󵴩۶��� 6)���D���� 7)�a�ϲ��ʭ��������
Dim x As Integer '���D��V
Dim xy() As Integer '����V
Dim ttop '�Ȧstop
Dim k(3) As Integer '�ʵe1�㢲)�H���ʵe
Dim sk(10) '�|������
Dim dx As Integer '�I�������ʶq
Dim sh As Integer '�v�l�Z��
Dim ax(1) As Integer '�]�B�P�_
Dim run(1) As Integer '�]�B�W�Ⱦ�
Dim ok As Integer 'ok�����h�a�ϲ��ʬ����h�H������
Dim ma As Integer '�ƶq
Dim pps As Integer '74�Q���������@��
Dim ppm() As Integer '0)�� 1)�w���} 2)���{�{ 4)���ʵe 5)����P�_ 6)�p��ɶ� �ɶ���hcall�l�u���ͺt��k 7)�s���� 8)�@�w�ɶ��S���h�s�����k�s
Dim n As Integer '���}��
Dim y As Integer '�ˮ`��
Dim akiz As Integer '�u��e�ש�j�v
Dim amax As Integer '�u��e�ש�j�v(�̤j��)
Dim iss As Integer '�ʵe����
Dim ball(1, 30) As Integer '0�l�u��V�x�s 1����w����
Dim ddd As Integer '�����ƭ�
Dim dzxc As Integer '�y�P���u�B���W�U�P�_
Dim stat() As Integer '�y�P���u�B���W�U�P�_�}�C
Dim ell(0) As Integer '0) �y�P���u�B(���t)�M�w���k�Υ�����
Dim up As Integer 'Ĳ�o�ĤG���ۤ��@����
Dim slow(1, 0) As Integer 'delay���}�C
Dim ven(15) As Integer '1~15 �� ���~�P�����I �Ĥ@���H�������ۤG
Dim hall() As Integer '�}�]�s�]���k�P�_0)���k 1)�W�U
Dim stong() As Integer '�z�𪬺A(�C�ӵ��۳��W�ߡA���|�]���Y�ӵ��۰����z��Ө䥦��۰���)
Dim vnn() As Integer '����������������w
Dim cir(10) As Integer '�W���ʤ���j�t��
Dim ys(8) As Integer '�]��ˮ`�@�P��
Dim gdown As Integer '�Y�����T�w��m�Ȧsleft
Dim gup As Integer '�Y�����T�w��m�Ȧstop
Dim gbl As Integer '�Y���{���y�{����
Dim metor As Integer '�y�P���u�B�����S����
Dim appshe1 As Integer, appshe2 As Integer '�o��Ӭ��W���ʪ��s�]��m���w
Dim agains As Integer '��ܶi��U�@���C��(��Ū���)
Dim supper As Integer '�}�]�s�]���o�g�q
Dim fireup(1, 10) As Integer '���W���𤧤ϼu�P�_ 0)�W�U 1)���k
Dim firr(3, 10) As Integer ' 0)���W���𤧬O�_�����k���� 1)�P�_�O�_�����ؼ� 2)�P�_�O���ӥؼ� 3)�P�_���Ǥ���w����(�B�⦡)

Const ppmn = 8 'ppm()���ŧi�ƶq
Const wmy = 3 '�ݤ뺧�����𤧫ᥴ�X���ƶq
Const holy = 8 '��������`�ƶq

Private Sub Timer26_Timer() '��������
Timer26.Enabled = False
DoEvents
Call desd(1)
End Sub
Private Sub ex(a) '�����{�ǡ�
If zxcv(4) = 1 Then Exit Sub
zxcv(4) = 1
If a = 1 Then '������
    Label8.Visible = True: Image1.Visible = False: Shape1.Visible = False: Timer26.Enabled = True '�Ұʩ���
Else '��F2��
    Call desd(1)
End If
End Sub
Public Sub desd(a) '�����@��
For i = 1 To s(8) '�����l�u�t��
    If ball(1, i) <> 1 Then '�p�G���Ӫ���s�b�h
        Unload Timer19(i)
        Unload Shape7(i)
        Unload Shape8(i)
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
Next
For f = 1 To ppmn
    For ff = 1 To ma
        ppm(f, ff) = 0
    Next
Next
iss = 0 '����Ū���ɥ��`
n = 0 '�L���P�_
s(13) = 0
If a = 1 Then
    supper = 0
    up = 0
    agains = 0
    aikz = 0
    amax = 0
    gdown = 0
    gup = 0
    gbl = 0
    y = 0
    dx = 0
    gz = 1
    Erase fireup, firr, ys, hall, vnn, cir, slow, ven, ball, s, zxcv, xy, k, ax, run, ppm, s, stong
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
    
    Form1.Left = 1845
    Form1.Top = 660
    Form1.Width = 12000 '(800*600)
    Form1.Height = 9000
    
End If

ok = 1
metor = 1
ma = 10

If agains <> 1 Then
    ReDim vnn(holy, 1) As Integer
    ReDim stong(7) As Integer
    
    Timer5(kikyou).Enabled = True
    Select Case kikyou
        Case 0
            Image1.Picture = Image2(0).Picture
            sh = 70
            Label6(0).Caption = "�������"
            Label6(1).Caption = "�ܱ𪺷N��"
            Timer20(0).Interval = 200
            Timer20(1).Interval = 400
        Case 1
            Image1.Picture = Image3(0).Picture
            sh = 53
            Label6(0).Caption = "���t���W"
            Label6(1).Caption = "�Y��"
            Timer20(0).Interval = 300
            Timer20(1).Interval = 600
        Case 2
            Image1.Picture = Image4(0).Picture
            sh = 70
            Label6(0).Caption = "�����ɹp"
            Label6(1).Caption = "�ݤ뺧��"
            Timer20(0).Interval = 200
            Timer20(1).Interval = 200
        Case 3
            Image1.Picture = Image5(0).Picture
            sh = 60
            Label6(0).Caption = "�y�P���u�B"
            Label6(1).Caption = "�W����"
            Timer20(0).Interval = 250
            Timer20(1).Interval = 500
    End Select
    Shape1.Top = Image1.Top + sh
    Shape1.Left = Image1.Left + Image1.Width \ 2 - 20
End If

'For f = 0 To 7
'    stong(f) = 1
'Next

ReDim xy(ma), ppm(ppmn, ma)
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

Call nnee '���ͺt��k


For f = 1 To ma
    If f Mod 8 = 0 Then iss = iss + 1
    ppm(4, f) = f - 8 * iss
    
    xy(f) = 1
    Load Timer22(f)
    Load Timer11(f)
    Timer11(f).Enabled = True '���ʵe
Next

If music = 0 Then
    music = 1
    sndPlaySound "�I��.wav", 9
    '�_������
    '    Set dsz = dxa.DirectSoundCreate("")
    '    dsz.SetCooperativeLevel Me.hWnd, DSSCL_NORMAL
    '    bu.lFlags = DSBCAPS_STATIC
    '    Set dsbz = dsz.CreateSoundBufferFromFile(("�I��.wav"), bu, wf)
    '    dsbz.Play DSBPLAY_LOOPING '''''''''''''''''''''''�w�]����(�榸).
    '�_������
End If
End Sub
Public Sub nnee() '���ͺt��k��
For f = 1 To 10
    Randomize
    
    Load Image6(f)
    Image6(f).Top = Int(Rnd * (Form1.ScaleHeight - Image6(f).Height - Line1.Y1 - 10)) + Line1.Y1
    Image6(f).Left = Int(Rnd * (Form1.ScaleWidth * 2 - Form1.ScaleWidth)) + Form1.ScaleWidth
    Image6(f).Visible = True
    
    Load Shape2(f)
    Shape2(f).Top = Image6(f).Top + 65
    Shape2(f).Left = Image6(f).Left + Image6(f).Width \ 2
    Shape2(f).Visible = True
    Shape2(f).ZOrder 1
    
    Load Timer12(f)
    Timer12(f).Enabled = True '����
Next
End Sub
Public Sub spdd(a) '�l�u���ͺt��k
If s(8) >= 30 Then s(8) = 0
s(8) = s(8) + 1
Load Timer19(s(8)) '���ͤl�u����
Load Shape7(s(8)) '���ͤl�u
Load Shape8(s(8)) '���ͤl�u�v�l

ball(0, s(8)) = xy(a) '�s�J�o�g�l�u������������V

If ball(0, s(8)) = 1 Then Shape7(s(8)).Left = Image6(a).Left - Shape7(s(8)).Width Else Shape7(s(8)).Left = Image6(a).Left + Image6(a).Width '���w�l�u��l��m
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
    Shape8(Index).Left = Shape7(Index).Left
    
    If coll2(Shape8(Index), Shape1) Then '�P�_���L��������
        If coll(Shape7(Index), Image1) Then
            Call kizzs
            If hp.Visible = True Then ball(1, Index) = 1: Unload Timer19(Index): Unload Shape7(Index): Unload Shape8(Index): Exit Sub '�p�G�����h����l��
        End If
    End If
    
    If gz = 1 Then Exit Sub: Timer19(Index).Enabled = False
    If con1(Shape7(Index)) Or con2(Shape7(Index)) Then '�p�G�W�X��ɫh����
        ball(1, Index) = 1
        Unload Timer19(Index)
        Unload Shape7(Index)
        Unload Shape8(Index)
    End If
    DoEvents
End Sub
Private Sub Timer12_Timer(Index As Integer) '���ʡ�

    ppm(8, Index) = ppm(8, Index) + 1 '�s�����k�s�t��k
    If ppm(8, Index) >= 14 Then
        ppm(7, Index) = 0
        ppm(8, Index) = 0
    End If '/�s�����k�s�t��k

    If Timer6.Enabled = True Then Image6(Index).Left = Image6(Index).Left - 10 * run(1)
    
    If ppm(5, Index) = 0 Then Image6(Index).Left = Image6(Index).Left - 10 * xy(Index) '���\����(�S���Q�ɹp����)
    
    If xy(Index) = 1 Then
        Shape2(Index).Left = Image6(Index).Left + Image6(Index).Width \ 2
    Else
        Shape2(Index).Left = (Image6(Index).Left + Image6(Index).Width \ 2) - 20
    End If
    
    If ppm(5, Index) = 0 Then '�S���Q�ɹp����ɰ����l�u�o�g����
        If xy(Index) = 1 And Image1.Left < Image6(Index).Left Or xy(Index) = -1 And Image1.Left + Image1.Width > Image6(Index).Left + Image6(Index).Width Then '�o�g����(���ݨ쨤��)
            If coll2(Shape2(Index), Shape1) Or coll2(Shape1, Shape2(Index)) Then
                ppm(6, Index) = ppm(6, Index) + 1
                If ppm(6, Index) Mod 10 = 0 Then ppm(6, Index) = 0: Call spdd(Index) '�I�s���ͤl�u�{��
            Else
                ppm(6, Index) = 0
            End If
        End If
    End If
    
    If Timer6.Enabled = True Then Exit Sub '�a�ϲ��ʮɡA���P�_�Ǫ�����ɤϼu
    
    If Image6(Index).Left < 0 Then xy(Index) = -1: Shape2(Index).Left = Image6(Index).Left + Image6(Index).Width \ 2 - 20 '�ϦV�P�_
    If Image6(Index).Left + Image6(Index).Width > Form1.ScaleWidth Then xy(Index) = 1
    DoEvents
End Sub
Private Sub Timer11_Timer(Index As Integer) '���ʵe��
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
Private Sub Timer7_Timer(Index As Integer) '�������ʵe��
ppm(2, Index) = ppm(2, Index) + 1
If ppm(2, Index) Mod 2 = 0 Then
    Image6(Index).Visible = True
    If ppm(2, Index) >= 12 Then
    
        '�|��ɱ���
        Call sico(Index)
        
        Unload Shape2(Index)
        Unload Image6(Index)
        Unload Timer7(Index)
        Unload Timer11(Index)
        Unload Timer12(Index)
        ppm(2, Index) = 0
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
Private Sub Timer22_Timer(Index As Integer) '������h
If ppm(1, Index) <> 1 Then ppm(5, Index) = 0: Timer11(Index).Enabled = True '�ҰʳQ����Ӱ���S��
Timer22(Index).Interval = 200 '����ɶ�
Timer22(Index).Enabled = False '����ۤv
DoEvents
End Sub
Private Sub timer30_Timer(Index As Integer) '���۶���
Image16.Picture = Form1.Picture
Select Case Index
    Case 0
        zxcv(5) = 1 '��L��w
        For f = 1 To 9
            Load Timer30(f) '���J����ʵe
            If f >= 2 Then Timer30(f).Enabled = False: Timer30(f).Interval = 50
        Next
        Timer30(1).Interval = 10
        Timer30(1).Enabled = True
        Timer30(0).Enabled = False
        Exit Sub
    Case 1 '����ʵe
        For f = 1 To 8 'Ū8��
            s(11) = s(11) + 1
            Load Line10(s(11))
            Load Shape13(s(11))
            Select Case s(11)
                Case 1 To 3 '���W
                    Shape13(s(11)).Left = Image1.Left - Image1.Width
                    Shape13(s(11)).Top = Image1.Top + (-Image1.Height * -(f - 2)) + Image1.Height \ 2
                    'Shape13(s(11)).Visible = True
                    
                    Line10(s(11)).X1 = Shape13(s(11)).Left
                    If s(11) = 1 Then '���G�ڹ�b�Q���X�ӥL�̪��@�q�I�A�u�n�p��...(�Q��ֵo��.....= ="")
                        
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
            Timer30(f).Enabled = True '����
        Next
        Unload Timer30(1)
        Exit Sub
    Case 2 '�H�U������
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
Private Sub superx() '�W����(���𵴩�)
Form1.Picture = LoadPicture()
Shape14.Width = 25 '��l��
Shape14.Height = 25 '��l��
Shape14.Left = Image1.Left + Image1.Width \ 2 '�]�w��m
Shape14.Top = Image1.Top + Image1.Height \ 2 '�]�w��m
Shape14.Visible = True
Timer31.Enabled = True '�ҰʮĪG
End Sub
Private Sub Timer31_Timer() '�W�������S��
s(12) = s(12) + 1
If s(12) = 25 Then
    Timer31.Enabled = False
    s(12) = 0
    Shape14.Visible = False
    Form1.Picture = Image16.Picture
    dx = 0
    zxcv(5) = 0
    
    stong(kikyou) = 1 '�N�|��ɪ��O�q�ǵ�����= =
    stong(kikyou + 4) = 1 '�N�|��ɪ��O�q�ǵ�����= =
    
    Label10.Caption = Val(Label10.Caption) - 1 & " X"
    If Val(Label10.Caption) = 0 Then Label10.Visible = False: Image23.Visible = False
    For f = 0 To 1
        delay(f).FillColor = &H80FFFF '�z�o���A'�C�����
    Next
    Exit Sub
End If
Shape14.Left = Shape14.Left - 20
Shape14.Width = Shape14.Width + 40
Shape14.Top = Shape14.Top - 20
Shape14.Height = Shape14.Height + 40
DoEvents
End Sub
Private Sub Timer27_Timer(Index As Integer) '�������
If vnn(Index, 0) = (Index + 1) * 15 Then '�p�G�w�g���槹���h
    If vnn(Index, 1) Mod 15 <= 12 And vnn(Index, 1) Mod 3 = 0 Then
        For f = 1 To ma
            If ppm(1, f) <> 1 Then
                If coll(Image6(f), Shape3(Index)) Then pps = Image6(f).Index: Call power(0)
            End If
        Next
    End If
    Unload Shape11(vnn(Index, 1)) '�C�C����
    Unload Line3(vnn(Index, 1)) '�C�C����
    vnn(Index, 1) = vnn(Index, 1) + 1 '���w����������
    
    If vnn(Index, 1) = (Index + 1) * 15 + 1 Then '����
        If Index <> 0 Then '�p�G���O���L�p�ɾ��h
            s(20) = s(20) + 1
            Unload Timer27(Index) '���������p�ɾ�
            Unload Shape3(Index) '���������j����
        Else
            Timer27(Index).Enabled = False
            Shape3(Index).Visible = False
            If stong(0) = 0 Then Call sg(0) '�I�s���۲����@��
        End If
        If s(20) = holy Then Call sg(0): s(7) = 0: s(20) = 0: Timer14(0).Interval = 10 '�I�s���۲����@��
    End If
Else '��}�l����
    vnn(Index, 0) = vnn(Index, 0) + 1
    
    Shape11(vnn(Index, 0)).Left = Shape11(vnn(Index, 0) - 1).Left '�U�@�ӥ����m
    Shape11(vnn(Index, 0)).Top = Shape11(vnn(Index, 0) - 1).Top + Shape11(vnn(Index, 0) - 1).Height '�U�@�ӥ����m
    Shape11(vnn(Index, 0)).Visible = True
    
    '�]�w��m
        Shape3(Index).Left = Shape11(vnn(Index, 0)).Left - Shape3(Index).Width \ 2 + Shape11(vnn(Index, 0)).Width \ 2 '�d��j�p
        Shape3(Index).Top = Shape11(vnn(Index, 0)).Top - Shape3(Index).Height \ 2 + Shape11(vnn(Index, 0)).Height \ 2 '�d��j�p
        Shape3(Index).Visible = True
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
Image22(Index).Left = Image22(Index).Left - 10 * hall(0, Index)
Image22(Index).Top = Image22(Index).Top - 10 * hall(1, Index)

For f = 1 To ma
    If ppm(1, f) <> 1 Then
        If coll(Image22(Index), Image6(f)) Then pps = Image6(f).Index: Call power(8)
    End If
Next
s(16) = s(16) + 1
If s(16) >= 100 Then
    Unload Image22(Index): Unload Timer34(Index)
    s(23) = s(23) + 1
    s(16) = 0
    If s(23) = supper Then Call sg(kikyou + 4): s(23) = 0
    Exit Sub
End If
If Image22(Index).Left < 0 Or Image22(Index).Left + Image22(Index).Width > Form1.ScaleWidth Then hall(0, Index) = hall(0, Index) * -1
If Image22(Index).Top < 0 Or Image22(Index).Top + Image22(Index).Height > Form1.ScaleHeight Then hall(1, Index) = hall(1, Index) * -1
DoEvents
End Sub
Private Sub Timer13_Timer(Index As Integer) '���W�ʵe
If Timer8.Enabled = True Then b = ttop Else b = Image1.Top

If Shape3(Index).Top <= b - Image1.Height * 2 Then '2
    s(26) = s(26) + 1
    If s(26) < 10 Then '3
        Unload Shape3(Index): Unload Timer13(Index) '�����{��
    Else '4
        Shape3(Index).Visible = False
        If Shape18(0).Height >= 11 Then '5
            Shape18(0).Height = Shape18(0).Height - 10
            Shape18(0).Left = Image1.Left + Image1.Width \ 2 - Shape18(0).Width \ 2
            Shape18(0).Top = b - Image1.Height * 2
        Else '6
            Shape18(0).Visible = False
            Unload Shape3(Index): Unload Timer13(Index) '�����{��
            If stong(1) = 0 Then s(26) = 0: s(3) = 0: Call sg(1)
        End If
    End If
Else '1�}�l
    Shape3(Index).Top = Shape3(Index).Top - 8 '���䲾��
    Shape18(0).Left = Image1.Left + Image1.Width \ 2 - Shape18(0).Width \ 2
    Shape18(0).Top = b - Image1.Height * 2
    
    If Timer1.Enabled = True Then Shape3(Index).Top = Shape3(Index).Top - 6.5: Shape3(0).Top = Shape3(0).Top - 2.5   '�W�T�w:�ˮ`�T�w
    If Timer2.Enabled = True Then Shape3(0).Top = Shape3(0).Top + 2.5 ' �ˮ`�T�w
    If Timer3.Enabled = True Or Timer4.Enabled = True Then Shape3(Index).Left = Image1.Left - Shape3(Index).Width \ 2 + Image1.Width \ 2: Shape3(0).Left = Image1.Left - Shape3(0).Width \ 2 + Image1.Width \ 2 '���k�T�w:�ˮ`�T�w
End If
DoEvents
End Sub
Private Sub Timer28_Timer(Index As Integer) '���W����ʵe
If firr(1, Index) = 0 Then '�Ĥ@���q�o�g���]
    Shape17(Index).Top = Shape17(Index).Top - 10 * fireup(0, Index)
    Shape17(Index).Left = Shape17(Index).Left - 10 * fireup(1, Index) * firr(0, Index)
    If Shape17(Index).Left < 0 Or Shape17(Index).Left + Shape17(Index).Width > Form1.ScaleWidth Then fireup(1, Index) = fireup(1, Index) * -1 '���k�ϼu
    If Shape17(Index).Top < 0 Or Shape17(Index).Top + Shape17(Index).Height > Form1.ScaleHeight Then fireup(0, Index) = fireup(0, Index) * -1: firr(0, Index) = 1
    If firr(0, Index) = 1 Then
        For f = 1 To ma
            If ppm(1, f) <> 1 Then
                If coll(Shape17(Index), Image6(f)) Then
                    pps = Image6(f).Index
                    Unload Shape17(Index)
                    Call power(9)
                                        
                    firr(1, Index) = 1
                    Shape18(Index).Left = Image6(f).Left + Image6(f).Width \ 2 - Shape18(Index).Width \ 2 '�X�{�b���쪺�������W
                    Shape18(Index).Top = Image6(f).Top + Image6(f).Height - Shape18(Index).Height '�X�{�b���쪺�������W
                    Shape18(Index).Visible = True
                    firr(2, Index) = Image6(f).Top
                    
                    For ff = 1 To 3 '�X��
                        Load Shape3(Index + ff * 10)
                        Shape3(Index + ff * 10).Left = Image6(f).Left + Image6(f).Width \ 2 - Shape3(Index + ff * 10).Width \ 2 '�]�w���W��ئ�m
                        Shape3(Index + ff * 10).Top = Image6(f).Top + Image6(f).Height \ 2 '�]�w���W��ئ�m
                    Next
                    Shape3(Index + 10).Visible = True
                    Exit For
                End If
            End If
        Next
    End If
Else '�ĤG���q���ͤ��W(���]�I��ĤH����)
    If Shape18(Index).Top <= firr(2, Index) - Image6(0).Height * 2 Then
        Shape3(Index + 10).Visible = False
        Shape3(Index + 10).Top = firr(2, Index) + Image6(0).Height \ 2
        If Shape18(Index).Height >= 12 Then Shape18(Index).Height = Shape18(Index).Height - 20 '�C�C��������
        
        
        For f = 2 + firr(3, Index) To 3 '�X��
            
            If Shape3(Index + f * 10).Top <= firr(2, Index) - Image6(0).Height * 2 Then '3
                firr(3, Index) = firr(3, Index) + 1 '�w��������(�B�⦡�ݭn)
                Unload Shape3(Index + f * 10)
                If f = 3 Then '4'�X��
                    Unload Shape18(Index): Unload Timer28(Index): s(25) = s(25) + 1
                    If s(25) = 5 Then '5'�X�W(��K�H����)
                        For ff = 1 To 5 '�X�W
                            Unload Shape3(10 + ff)
                        Next
                        Erase firr, fireup: s(3) = 0: s(25) = 0: s(28) = 0: Call sg(1) '6����...
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
        For ff = 1 To ma '�ˮ`
            If ppm(1, ff) <> 1 Then
                If coll(Image6(ff), Shape3(Index + 30)) Then pps = Image6(ff).Index: Call power(1) '�X��
            End If
            DoEvents
        Next
    End If
    
End If
DoEvents
End Sub
Private Sub Timer23_Timer(Index As Integer) '�ɹp
If Line2(Index).Y2 >= Shape9.Top + Shape9.Height * 17 Then '2
    s(27) = s(27) + 1: Unload Line2(Index): Unload Timer23(Index)
    If s(27) = 15 Then '3
        Shape9.Visible = False
        If stong(2) = 0 Then s(27) = 0: s(3) = 0: Call sg(2) Else s(3) = 20   '4�ɹp�w��
    End If
Else '1�}�l
    Line2(Index).Y2 = Line2(Index).Y2 + 9
End If
DoEvents
End Sub
Private Sub Timer24_Timer(Index As Integer) '�ɹp����
For i = 1 To 10
    t = Line2((Index - 1) * 10 + i).X1
    Line2((Index - 1) * 10 + i).X1 = Line2((Index - 1) * 10 + i).X2
    Line2((Index - 1) * 10 + i).X2 = t
    DoEvents
Next
Call lightnings
For f = 1 To ma
    If ppm(1, f) <> 1 Then
        If fscoll(Image6(f)) Then pps = Image6(f).Index: Call power(2)
    End If
Next
s(30) = s(30) + 1
If s(30) >= 3 Then
    Unload Timer24(Index)
    s(31) = s(31) + 1
    If s(31) = 5 Then
        For i = 1 To 10 * 5
            Unload Line2(i)
        Next
        s(31) = 0: s(30) = 0: s(27) = 0: s(3) = 0: Call sg(2)
    End If
End If
End Sub
Private Sub Timer18_Timer(Index As Integer) '�ݤ뺧��
Shape3(Index).Visible = True '���i��{
Shape3(Index).Height = Shape3(Index).Height + 4 '���i�d���ܤj
Shape3(Index).Width = Shape3(Index).Width + 16 '���i�d���ܤj
Shape3(Index).Left = Shape3(Index).Left - 8 '���i�d���ܤj
Shape3(Index).Top = Shape3(Index).Top - 2 '���i�d���ܤj
s(24) = s(24) + 1
If s(24) Mod (20 - 12 * stong(6)) = 0 Then
    For f = 1 To ma
        If ppm(1, f) <> 1 Then
            If coll(Shape3(Index), Image6(f)) Or coll(Image6(f), Shape3(Index)) Then pps = Image6(f).Index: Call power(6)
        End If
    Next
End If
DoEvents
If Shape3(Index).Width >= 500 Then
    s(18) = s(18) + 1: Unload Shape3(Index): Unload Timer18(Index)
    If s(18) = 4 + stong(6) * wmy Then s(17) = 0: s(18) = 0: s(24) = 0: Call sg(6)
End If
End Sub
Private Sub Timer29_Timer(Index As Integer) '�y�P���u�B
Image13(Index).Left = Image13(Index).Left + 10 * ell(0)
Image13(Index).Top = Image13(Index).Top + 3 * stat(Index)
Shape19(Index * stong(3)).Left = Shape19(Index * stong(3)).Left + 10 * ell(0) 'Image13(Index).Left - Shape19(Index * stong(3)).Width
Shape19(Index * stong(3)).Top = Shape19(Index * stong(3)).Top + 3 * stat(Index)
For f = 1 To ma
    If ppm(1, f) <> 1 Then
        If coll(Image13(Index), Image6(f)) Then
            pps = Image6(f).Index: Call power(3)
            If stong(3) = 0 Then
                Unload Image13(Index): Unload Timer29(Index): s(15) = s(15) + 1
                If s(15) >= 16 Then s(10) = 0: s(15) = 0: Call sg(3): Shape12.Visible = False: metor = 1
                Exit Sub
            End If
        End If
    End If
    DoEvents
Next
If Image13(Index).Left > Image1.Left + Image1.Width * 6 Or Image13(Index).Left < Image1.Left - Image1.Width * 5 Or con1(Image13(Index)) Or con2(Image13(Index)) Then
    Unload Image13(Index): Unload Timer29(Index)
    If stong(3) = 1 Then Unload Shape19(Index)
    s(15) = s(15) + 1
    If s(15) >= 16 Then s(10) = 0: s(15) = 0: Call sg(3): Shape12.Visible = False: metor = 1 '�P�_���𪬺A���y�P���u�B�o�g����
End If
DoEvents
End Sub
Private Sub Timer35_Timer(Index As Integer) '�W���ʤ��_�i

Shape16(Index).Left = Shape16(Index).Left - 4
Shape16(Index).Width = Shape16(Index).Width + 8
Shape16(Index).Top = Shape16(Index).Top - 2
Shape16(Index).Height = Shape16(Index).Height + 4
cir(3) = cir(3) + 1


If cir(3) Mod 5 = 0 Then
    If Shape16(Index).BorderWidth > 1 Then
        Shape16(Index).BorderWidth = Shape16(Index).BorderWidth - 1
        
        If Shape16(Index).BorderWidth Mod 2 = 0 Then
            For f = 1 To ma
                If ppm(1, f) <> 1 Then
                    If coll(Shape16(Index), Image6(f)) Or coll(Image6(f), Shape16(Index)) Then
                        pps = Image6(f).Index: Call power(7)
                    End If
                End If
                DoEvents
            Next
        End If
    Else
        Unload Shape16(Index)
        Unload Timer35(Index)
        If stong(7) = 0 Then
            s(32) = s(32) + 1
            If s(32) = 4 Then s(32) = 0: Call sg(7): s(19) = 0  '�D�����o����
        End If
    End If
End If
DoEvents
End Sub
Private Sub Timer14_Timer(Index As Integer) '���ۡ�
Select Case Index
    Case 0 '�������
        If s(7) = 0 Then
            For f = 0 To holy * stong(0) '�p�G�����𪬺A�h����1~4���p�ɾ�
                If f <> 0 Then
                    Load Shape3(f)
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
                Shape11(15 * f + 1).Top = Int(Rnd * (Form1.ScaleHeight - Shape3(0).Height * 2.2)) - Shape3(0).Height * 1.2
                Shape11(15 * f + 1).Left = Int(Rnd * (Form1.ScaleWidth - Shape3(0).Width \ 2)) + Shape3(0).Width \ 4
            Next
        End If
        If stong(0) = 0 Then
            Timer27(0).Interval = 30: Timer27(0).Enabled = True: Timer14(Index).Enabled = False
        Else
            Timer14(Index).Interval = 300
            Timer27(s(7)).Interval = 20
            Timer27(s(7)).Enabled = True
            If s(7) = holy Then Timer14(Index).Enabled = False
            s(7) = s(7) + 1
        End If
    Case 1 '���W
        Dim b As Integer
        If Timer8.Enabled = True Then b = ttop Else b = Image1.Top
        
        Shape18(0).Top = b - Image1.Height * 2
        Shape18(0).Height = 196
        
        s(3) = s(3) + 1
        Load Shape3(s(3)) '���ͤ��W(���)
        Load Timer13(s(3)) '���ͤ��W�p�ɾ�
        Shape3(s(3)).Left = Image1.Left - Shape3(s(3)).Width \ 2 + Image1.Width \ 2 '�]�w���W����m
        Shape3(s(3)).Top = b + Image1.Height \ 2 '�]�w���W����m
        Shape18(0).Left = Image1.Left + Image1.Width \ 2 - Shape18(0).Width \ 2 '�]�w���ߪ���m
        Shape18(0).Top = b + Image1.Height - Shape18(0).Height '�]�w���ߪ���m
        Shape18(0).Visible = True
        Shape3(s(3)).Visible = True
        Timer13(s(3)).Enabled = True '�Ұʤ��W�p�ɾ�
        
        For f = 1 To ma
            If ppm(1, f) <> 1 Then
                If coll2(Shape2(f), Shape3(0)) Then
                    If coll(Image6(f), Shape3(0)) Then pps = Image6(f).Index: Call power(1)
                End If
            End If
        Next
        
        If stong(1) = 1 Then
            If s(3) Mod 2 = 0 Then
                Load Shape18(s(3) \ 2) '���ͤ��W(����)
                Shape18(s(3) \ 2).Height = 1
                
                Shape17(s(3) \ 2).Left = Image1.Left + Image1.Width \ 2 - Shape17(s(3) \ 2).Width \ 2 '���]
                Shape17(s(3) \ 2).Top = b
                Shape17(s(3) \ 2).Visible = True
                Timer28(s(3) \ 2).Enabled = True '���W�������
            End If
        End If
        If s(3) = 10 Then Timer14(Index).Enabled = False
    Case 2 '�ɹp
        If s(3) < 15 Then
            Timer14(Index).Interval = 30
            s(3) = s(3) + 1
            Load Line2(s(3)) '���͹p�q
            
            Line2(s(3)).X1 = Int(Rnd * Shape9.Width) + Shape9.Left
            Line2(s(3)).X2 = Int(Rnd * Shape9.Width) + Shape9.Left
            Line2(s(3)).Y1 = Int(Rnd * Shape9.Height) + Shape9.Top
            Line2(s(3)).Y2 = gup + sh + 6 '��ޯ઺��m+�v�l+�ǷL�վ�
            
            Line2(s(3)).Visible = True
            Load Timer23(s(3)) '���ͼɹp�p�ɾ�
            Timer23(s(3)).Enabled = True
            
            If s(3) Mod (5 - 2 * stong(2)) = 0 Then '�T�s�q
                For f = 1 To ma
                    If ppm(1, f) <> 1 Then
                        If coll2(Shape2(f), Shape3(0)) Then
                            If coll(Image6(f), Shape3(0)) Then pps = Image6(f).Index: Call power(2)
                        End If
                    End If
                Next
            End If
        Else
            If stong(2) = 0 Then
                Timer14(Index).Enabled = False
            Else
                If s(3) = 20 Then '�P�_��l�ɹp�w��
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
        stat(s(10)) = dzxc
        Image13(s(10)).ZOrder 0
        If stong(3) = 1 Then
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
                Line4(0).X1 = Line4(0).X1 - (4 + 4 * stong(4))
                If Line4(0).X1 <= Shape10.Left + 16 Then Line4(0).X1 = Shape10.Left + 16: ven(0) = 1: Line4(2).X1 = Line4(0).X1: Line4(2).X2 = Line4(0).X1
            End If
            If ven(1) = 0 Then
                Line4(0).Y1 = Line4(0).Y1 - (2 + 2 * stong(4))
                If Line4(0).Y1 <= Shape10.Top + Shape10.Height \ 4 Then Line4(0).Y1 = Shape10.Top + Shape10.Height \ 4: ven(1) = 1: Line4(2).Y1 = Line4(0).Y1: Line4(2).Y2 = Line4(2).Y1: ven(5) = 1
            End If
            
        '�Ĥ@�����u
            If ven(5) = 1 Then
                If ven(4) = 0 Then
                    Line4(2).X2 = Line4(2).X2 + (4 + 4 * stong(4))
                    If Line4(2).X2 >= Shape10.Left + Shape10.Width - 17 Then Line4(2).X2 = Shape10.Left + Shape10.Width - 17: ven(4) = 1: ven(8) = 1: Line4(1).X1 = Line4(2).X2: Line4(1).Y1 = Line4(2).Y2: Line4(1).Y2 = Line4(1).Y1: Line4(1).X2 = Line4(1).X1
                End If
            End If
        '�ĤG�����u
            If ven(8) = 1 Then
                If ven(9) = 0 Then
                    Line4(1).X2 = Line4(1).X2 - (4 + 4 * stong(4))
                    If Line4(1).X2 <= Shape10.Left + Shape10.Width \ 2 Then Line4(1).X2 = Shape10.Left + Shape10.Width \ 2: ven(9) = 1
                End If
                If ven(10) = 0 Then
                    Line4(1).Y2 = Line4(1).Y2 + (2 + 2 * stong(4))
                    If Line4(1).Y2 >= Shape10.Top + Shape10.Height Then Line4(1).Y2 = Shape10.Top + Shape10.Height: ven(10) = 1: ven(14) = 1 '����
                End If
            End If
            
    '�ĤG���D�u....
        
            If ven(2) = 0 Then
                Line4(5).X2 = Line4(5).X2 + (4 + 4 * stong(4))
                If Line4(5).X2 >= Shape10.Left + Shape10.Width - 17 Then Line4(5).X2 = Shape10.Left + Shape10.Width - 17: ven(2) = 1: Line4(4).X2 = Line4(5).X2: Line4(4).X1 = Line4(4).X2
            End If
            If ven(3) = 0 Then
                Line4(5).Y2 = Line4(5).Y2 + (2 + 2 * stong(4))
                If Line4(5).Y2 >= Shape10.Top + (Shape10.Height \ 4) * 3 Then Line4(5).Y2 = Shape10.Top + (Shape10.Height \ 4) * 3: ven(3) = 1: Line4(4).Y2 = Line4(5).Y2: Line4(4).Y1 = Line4(4).Y2: ven(7) = 1
            End If
        
        '�Ĥ@�����u
            If ven(7) = 1 Then
                If ven(6) = 0 Then
                    Line4(4).X1 = Line4(4).X1 - (4 + 4 * stong(4))
                    If Line4(4).X1 <= Shape10.Left + 16 Then Line4(4).X1 = Shape10.Left + 16: ven(6) = 1: ven(11) = 1: Line4(3).X2 = Line4(4).X1: Line4(3).X1 = Line4(3).X2: Line4(3).Y2 = Line4(4).Y1: Line4(3).Y1 = Line4(3).Y2
                End If
            End If
        '�ĤG�����u
            If ven(11) = 1 Then
                If ven(12) = 0 Then
                    Line4(3).X1 = Line4(3).X1 + (4 + 4 * stong(4))
                    If Line4(3).X1 >= Shape10.Left + Shape10.Width \ 2 Then Line4(3).X1 = Shape10.Left + Shape10.Width \ 2: ven(12) = 1
                End If
                
                If ven(13) = 0 Then
                    Line4(3).Y1 = Line4(3).Y1 - (2 + 2 * stong(4))
                    If Line4(3).Y1 <= Shape10.Top Then Line4(3).Y1 = Shape10.Top: ven(13) = 1
                End If
            End If
                
        '�j�O�@��
        Else
            metor = 1
            If ven(15) = 0 Then
                Image17.Left = Image1.Left - Image17.Width \ 2 + Image1.Width \ 2 '�ܱ��m
                Image17.Top = Image1.Top - Image17.Height \ 2 + Image1.Height \ 2
                Image17.Visible = True
                
                Image18.Top = Image17.Top + 64 '�b�ڦ�m
                Image18.Left = Image17.Left + 16
                Image18.Visible = True
                
                Image19.Top = Image18.Top - Image19.Height \ 2 + Image18.Height \ 2 '�}�]����
                
                Call ham99(0) '�}�]�s�]��V���w
                
                ven(15) = 1
            Else
                Image17.Left = Image17.Left - 10
                If Image17.Left >= 1 Then
                    Image18.Left = Image17.Left + 16
                    Image18.Top = Image17.Top + 64
                Else '��ܱ��F��ɫh
                
                    '����P�_
                    For f = 1 To ma
                        If ppm(1, f) <> 1 Then
                            If coll(Image6(f), Image18) Or coll(Image18, Image6(f)) Then pps = Image6(f).Index: Call power(4)
                        End If
                    Next
                    
                    If stong(4) = 1 Then '����P�_
                        If Image19.Left + Image19.Width > Form1.ScaleWidth Then
                            Image18.Left = Form1.ScaleWidth - Image18.Width
                            Image19.Top = Image18.Top - Image19.Height \ 2 + Image18.Height \ 2
                            If s(14) = 20 Then
                                Image18.Visible = False
                                Image19.Visible = False
                                
                                s(14) = 0
                                Timer14(Index).Enabled = False '���۩�
                            Else
                                Image19.Left = Form1.ScaleWidth - Image19.Width
                                
                                '�}�]�s�]
                                s(14) = s(14) + 1
                                Load Timer34(s(14))
                                Load Image22(s(14))
                                
                                Randomize
                                Image22(s(14)).Left = Form1.ScaleWidth - Image22(s(14)).Width '�]�w��m
                                Image22(s(14)).Top = Int(Rnd * Image19.Height) + Image19.Top '�]�w��m
                                
                                Image22(s(14)).Visible = True
                                Timer34(s(14)).Enabled = True '�Ұʯ}�]�s�]����
                            End If
                        Else
                            Call kiker '�I�s�ܱ𪺷N�Ӧ@��
                        End If
                    Else
                        Call kiker '�I�s�ܱ𪺷N�Ӧ@��
                        If Image18.Left > Form1.ScaleWidth Then Call sg(Index) '���۩�
                    End If
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
                If Line5(f).X2 >= gdown + Image1.Width * 5 Then gbl = 1: Exit For
                DoEvents
            Next
        Else '����
            If stong(5) = 1 Then
                For f = 1 To 3
                    Form1.Left = Form1.Left - 100
                    Form1.Left = Form1.Left + 100
                    Form1.Left = Form1.Left + 100
                    Form1.Left = Form1.Left - 100
                Next
            End If
            For ff = 1 To 3 + stong(5) * 7 '�s���ˮ`
                For f = 1 To ma
                    If ppm(1, f) <> 1 Then
                        If fscoll(Image6(f)) Then pps = Image6(f).Index: Call power(5)
                    End If
                Next
            Next
            For f = 1 To 3
                Unload Line5(f)
            Next
            gbl = 0
            Timer14(Index).Enabled = False
            Call sg(Index)
        End If
    Case 6 '�ݤ뺧��
        s(17) = s(17) + 1
        If stong(6) = 0 Then
            Timer14(6).Interval = 200
            Shape3(s(17)).BorderColor = &HFF0000 '���i�C��
            Shape3(s(17)).Left = Image1.Left - Shape3(s(17)).Width \ 2 + Image1.Width \ 2 '���i��l��m
            Shape3(s(17)).Top = Image1.Top + Image1.Height - 15 '���i��l��m
        Else
            Timer14(6).Interval = 160
            Randomize
            Shape3(s(17)).BorderColor = &HFF0000 '���i�C��
            Shape3(s(17)).Left = Int(Rnd * (Form1.ScaleWidth - Shape3(s(17)).Width * 3.5)) + Shape3(s(17)).Width * 1.5 '���i��l��m
            Shape3(s(17)).Top = Int(Rnd * (Form1.ScaleHeight - Shape3(s(17)).Height * 2 - Line1.Y1)) + Line1.Y1 '���i��l��m
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
                        If stong(7) = 1 Then
                            If s(22) < 1 Then Timer14(7).Interval = 100: Call ham99(1) '�}�]�s�]��V���w
                            s(22) = s(22) + 1
                            Image22(s(22)).Left = appshe1 - Image22(s(22)).Width \ 2 + Shape16(0).Width \ 2
                            Image22(s(22)).Top = appshe2 + Image22(s(22)).Height
                            Image22(s(22)).Visible = True
                            Timer34(s(22)).Enabled = True
                            If s(22) = 12 Then Timer14(Index).Enabled = False
                        Else
                            Timer14(Index).Enabled = False
                        End If
                    Else '4���;_�i
                        Timer14(7).Interval = 130
                        s(19) = s(19) + 1
                        
                        If s(19) = 1 Then '���_�i����I�B�A���|��ۥD�H�]
                            For f = 1 To 4
                                Shape16(f).Left = Image1.Left + Image1.Width \ 2 - Shape16(f).Width \ 2 '�_�i��m�]�w
                                Shape16(f).Top = Image1.Top + Image1.Height '/�_�i��m�]�w
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
        If cir(2) < 150 Then Form1.Circle (Image1.Left + Image1.Width \ 2, ((Image1.Top) - cir(1)) + cir(2)), cir(0)
End Select
DoEvents
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
Private Sub tai(a As Integer) '���`���A�t��k
Randomize
c = Int(Rnd * 5) + 1
If c = 2 Then Timer22(pps).Interval = 5000 Else Timer22(pps).Interval = 200 '�ɹp���S��ĪG
End Sub
Public Sub ham99(a As Integer) '�}�]�s�]��V
Select Case a
    Case 0 '�ܱ𪺷N��
        supper = 20
        ReDim hall(1, supper)
        For f = 0 To 1 '�}�]�s�]��l����V���w
            For ff = 0 To 20
                If ff Mod 2 = 0 Then hall(f, ff) = -1 Else hall(f, ff) = 1
            Next
        Next
    Case 1 '�W����
        supper = 12
        ReDim hall(1, supper)
        For f = 0 To 1 '�}�]�s�]��l����V���w
            For ff = 0 To 12
                Select Case ff Mod 4
                    Case 1
                        hall(f, ff) = 1 * (f * 2 - 1)
                    Case 2
                        hall(f, ff) = -1
                    Case 3
                        hall(f, ff) = 1
                    Case 0
                        hall(f, ff) = -1 * (f * 2 - 1)
                End Select
            Next
        Next
End Select
End Sub
Public Sub mpss(a As Integer) '�l��MP�t��k��
If mp.Width >= 50 Then
    Dim d As Integer
    If Timer8.Enabled = True Then d = ttop Else d = Image1.Top
    
    gdown = Image1.Left '�D����m Left
    gup = Image1.Top '�D����mTop
    
    Shape3(0).Left = Image1.Left - Shape3(0).Width \ 2 + Image1.Width \ 2 '���W��
    Shape3(0).Top = d + Image1.Height \ 2
    

    Timer25.Enabled = False 'mp�W�[����
    If a = 3 Or a = 4 Then metor = 0: zxcv(2) = zxcv(3) '��������
    If a = kikyou + 4 Then
        Timer21(1).Enabled = False '����A delay�{�{(���ۤG)
        delay(1).Visible = False: delay(1).Height = 1: delay(1).Top = 81
    Else
        Timer21(0).Enabled = False '����A delay�{�{
        delay(0).Visible = False: delay(0).Height = 1: delay(0).Top = 81
    End If
    mp.Width = mp.Width - 40
    Label2.Caption = mp.Width * akiz & "/" & amax
    Image12.Visible = True: Image12.ZOrder 1
    Select Case a
        Case 0 '�������
            Shape3(0).BorderColor = RGB(255, 255, 0) '���w�d���C���
        
            Shape11(0).Left = Image1.Left + Image1.Width \ 2 - Shape11(0).Width \ 2
            Shape11(0).Top = d - 190
            
            'Shape11(0).Left = Image1.Left - 175
            'Shape11(0).Top = Image1.Top - 100
            Line3(0).X1 = Shape11(0).Left + Shape11(0).Width \ 4  '�������W�����W����m
            Line3(0).Y1 = Shape11(0).Top '���W�����W��
            Line3(0).X2 = Shape11(0).Left + Shape11(0).Width \ 2  '���W���k�U��
            Line3(0).Y2 = Shape11(0).Top + Shape11(0).Height \ 2  '���W���k�U��
            
            If stong(a) = 1 Then c = 2: b = 2 Else c = 2: b = 1
        Case 1 '���W
            If stong(1) = 1 Then
                For f = 1 To 5 '�X�W
                    Load Shape17(f)
                    Load Timer28(f)
                    fireup(0, f) = 1
                    If f Mod 2 = 0 Then fireup(1, f) = 1 Else fireup(1, f) = -1
                Next
            End If
            Shape18(0).Height = 1
            Shape18(0).Left = Image1.Left + Image1.Width \ 2 - Shape18(0).Width \ 2
            Shape18(0).Top = d + Image1.Height - Shape18(0).Height
            Shape18(0).Visible = True
            
            '------�_������
            Set ds = dxa.DirectSoundCreate("")
            ds.SetCooperativeLevel Me.hWnd, DSSCL_NORMAL
            bu.lFlags = DSBCAPS_STATIC
            Set dsb = ds.CreateSoundBufferFromFile(("���W.wav"), bu, wf)
            dsb.Play DSBPLAY_DEFAULT '''''''''''''''''''''''�w�]����(�榸).
            '------�_������
            
            If stong(a) = 1 Then c = 2: b = 2 Else c = 2: b = 1
        Case 2 '�ɹp
            Shape9.Left = Image1.Left - 80 '�W�����¬}
            Shape9.Top = d - 180 '�W�����¬}
            Shape9.Visible = True
        
            Set dscc = dxa.DirectSoundCreate("")
            dscc.SetCooperativeLevel Me.hWnd, DSSCL_NORMAL
            bu.lFlags = DSBCAPS_STATIC
            Set dsbcc = dscc.CreateSoundBufferFromFile(("�ɹp.wav"), bu, wf)
            dsbcc.Play DSBPLAY_DEFAULT '''''''''''''''''''''''�w�]����(�榸).
            
            If stong(a) = 1 Then c = 5: b = 2 Else c = 5: b = 3
        Case 3 '�y�P���u�B
            ReDim stat(20)
            dzxc = 0
            If zxcv(3) = 1 Then Image13(0).Picture = Image14(1).Picture: ell(0) = -1: Shape12.Left = Image1.Left + Image1.Width - Shape12.Width Else Image13(0).Picture = Image14(0).Picture: ell(0) = 1: Shape12.Left = Image1.Left '0���k1����
            Shape12.Top = d + Image1.Height \ 2 - Shape12.Height \ 2
            Shape12.Visible = True
            
        Case 4 '�ܱ𪺷N��
            Shape10.Left = Image1.Left - Shape10.Width \ 2 + Image1.Width \ 2
            Shape10.Top = d + Image1.Height \ 2
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
            
            If stong(a) = 1 Then c = 2: b = 2 Else c = 2: b = 1
        Case 5 '�Y��
            For f = 1 To 3
                Load Line5(f) '�Y���u��
                Line5(f).X1 = Image1.Left + Image1.Width \ 2
                Line5(f).Y1 = d + Image1.Height \ 2
                Line5(f).X2 = Line5(f).X1
                Line5(f).Y2 = Line5(f).Y1
                Line5(f).Visible = True
            Next
            
        '-----�_������
            Set dscd = dxa.DirectSoundCreate("")
            dscd.SetCooperativeLevel Me.hWnd, DSSCL_NORMAL
            bu.lFlags = DSBCAPS_STATIC
            Set dsbd = dscd.CreateSoundBufferFromFile(("�Y��.wav"), bu, wf)
            dsbd.Play DSBPLAY_DEFAULT '''''''''''''''''''''''�w�]����(�榸).
        '-----�_������
        
            If stong(a) = 1 Then c = 2: b = 2 Else c = 2: b = 1
        Case 6 '�ݤ뺧��
            For f = 1 To 4 + stong(6) * wmy
                Load Shape3(f)
                Shape3(f).Width = 60 '���i��l�j�p
                Shape3(f).Height = 30 '���i��l�j�p
                
                Load Timer18(f)
            Next
            
            If stong(a) = 1 Then c = 2: b = 2 Else c = 2: b = 1
        Case 7 '�W����
            cir(0) = 5
            Timer14(7).Interval = 50
            For f = 1 To 4
                Load Shape16(f)
                Load Timer35(f)
            Next
            If stong(7) = 1 Then
                For f = 1 To 12
                    Load Image22(f)
                    Load Timer34(f)
                Next
            End If
            
            If stong(a) = 1 Then c = 2: b = 2 Else c = 2: b = 1
    End Select
    
    ys(kikyou) = Int(Rnd * 10 + c) + 2 + b�s���ˮ`
    ys(kikyou + 4) = Int(Rnd * 10 + c) + 2 + b '/�s���ˮ`
    
    Timer14(a).Enabled = True
End If
End Sub
Private Sub kiker() '�ܱ𪺷N�Ӧ@��

Image18.Left = Image18.Left + 30 '�b�ڮg�X
Image19.Left = Image18.Left + Image18.Width - Image19.Width \ 2 '�}�]����
Image19.Visible = True

If Image17.Left + Image1.Width < 0 Then '��ܱ�W�X��ɫh
    Image17.Visible = False
End If

End Sub
Public Sub sg(a As Integer) '���ۦ@��
Timer14(a).Enabled = False

If a <= 3 Then delay(0).Visible = True: Timer20(0).Enabled = True Else delay(1).Visible = True: Timer20(1).Enabled = True '����delay

Select Case a
    Case 0 '�������
        Erase vnn
        ReDim vnn(holy, 1) As Integer
    Case 3 '�y�P���u�B
        metor = 1
    Case 4 '�ܱ𪺷N��
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
    Case 7 '�W����
        Erase cir: s(19) = 0: s(22) = 0
End Select

delay(a \ 4).FillColor = &H80FF80
stong(a) = 0

End Sub
Private Sub Timer20_Timer(Index As Integer) '����delay
delay(Index).Height = delay(Index).Height + 5
delay(Index).Top = delay(Index).Top - 5
If delay(Index).Height >= 57 Then delay(Index).Height = 57: delay(Index).Top = 24:: Timer20(Index).Enabled = False: Timer21(Index).Enabled = True
DoEvents
End Sub
Private Sub Timer21_Timer(Index As Integer) 'A delay�{�{
slow(Index, s(6)) = slow(Index, s(6)) + 1
If slow(Index, s(6)) Mod 2 = 0 Then
    delay(Index).Visible = True
Else
    delay(Index).Visible = False
End If
DoEvents
End Sub
Private Sub Timer10_Timer() '����
If (Timer3.Enabled = True And Timer8.Enabled = True) Or (Timer4.Enabled = True And Timer8.Enabled = True) Then '���D����
    If Timer3.Enabled = True Then '��
        Image1.Left = Image1.Left - 20
        Shape1.Left = Shape1.Left - 20
    Else '�k
        Image1.Left = Image1.Left + 20
        Shape1.Left = Shape1.Left + 20
    End If
    s(2) = s(2) + 1
    Call sppr(10) '�P�_�O�_����
    
    If s(2) >= 5 Then s(2) = 0: Timer10.Enabled = False
Else
    If zxcv(3) = 0 Then '�k
        Image1.Left = Image1.Left + 10
        Shape1.Left = Shape1.Left + 10
        s(2) = s(2) + 1
        If s(2) >= 5 Then
            Call sppr(9) '�P�_�O�_����
            
            Image1.Left = Image1.Left - 50
            Shape1.Left = Shape1.Left - 50
            s(2) = 0
            Timer10.Enabled = False
        End If
    Else '��
        Image1.Left = Image1.Left - 10
        Shape1.Left = Shape1.Left - 10
        s(2) = s(2) + 1
        If s(2) >= 5 Then
            Call sppr(9) '�P�_�O�_����
            
            Image1.Left = Image1.Left + 50
            Shape1.Left = Shape1.Left + 50
            s(2) = 0
            Timer10.Enabled = False
        End If
    End If
End If
DoEvents
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer) '������U��
If zxcv(5) = 1 Then Exit Sub
If KeyCode = keyup Then dzxc = -1
If KeyCode = keydown Then dzxc = 1 '�y�P�W�U�P�_
    If (zxcv(0) = 1 And KeyCode = keyup) Or (zxcv(0) = 1 And KeyCode = keydown) Then Exit Sub
    Select Case KeyCode
        Case keyleft '��37
            If ax(0) = 1 Then run(0) = 2
            
            Timer3.Enabled = True '����
            zxcv(3) = 1 '�ʵe���k�P�_
        Case keyup '�W38
            up = 1
            If zxcv(0) = 0 Then
                Timer1.Enabled = True '����
            End If
        Case keyright '�k39
            If ax(1) = 1 Then run(1) = 2
            
            Timer4.Enabled = True '����
            zxcv(3) = 0 '�ʵe���k�P�_
        Case keydown '�U40
            If zxcv(0) = 0 Then
                Timer2.Enabled = True '����
            End If
        Case keya 'A65
            If Val(Label10.Caption) > 0 Then Timer30(0).Enabled = True '�p�G���|��ɫh���۶���
            If up = 1 Then
                If Timer21(1).Enabled = True Then Call mpss(kikyou + 4) 'Ū���l��MP�t��k(���ۤG)
            Else
                If Timer21(0).Enabled = True Then Call mpss(kikyou)  'Ū���l��MP�t��k
            End If
        Case keys 'S ����83
            zxcv(1) = zxcv(1) + 1
            If zxcv(1) = 1 Then Timer10.Enabled = True
        Case keyd 'D ���D68
            If zxcv(0) = 0 Then
                zxcv(0) = 1
                ttop = Image1.Top
                Timer8.Enabled = True
            End If
        Case 113 'F2
            If Timer20(0).Enabled = True And Timer20(1).Enabled = True Or _
               Timer20(0).Enabled = True And Timer21(1).Enabled = True Or _
               Timer21(0).Enabled = True And Timer20(1).Enabled = True Or _
               Timer21(0).Enabled = True And Timer21(1).Enabled = True Then Call ex(0)
    End Select
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer) '����u�_��
Select Case KeyCode
    Case keyleft '��
        ax(0) = 1
        Timer9(0).Enabled = True
        ax(1) = 0 '�����k�]
        Timer9(1).Enabled = False '�����k�]
        
        run(0) = 1 '�������A
        
        Timer3.Enabled = False
    Case keyup '�W
        up = 0
        Timer1.Enabled = False
        dzxc = 0
    Case keyright '�k
        ax(1) = 1
        Timer9(1).Enabled = True
        ax(0) = 0 '�������]
        Timer9(0).Enabled = False '�������]
        
        run(1) = 1 '�������A
        
        Timer6.Enabled = False '�a�ϲ�������
        Timer4.Enabled = False
    Case keydown '�U
        Timer2.Enabled = False
        dzxc = 0
    Case keys 's����
        zxcv(1) = 0
    Case keya 'A
        Timer30(0).Enabled = False '���۰����
        Timer25.Enabled = True 'MP�W�[
End Select
End Sub
Private Sub Timer9_Timer(Index As Integer) '�]�B�P�_��
If Index = 0 Then
    ax(0) = 0
Else
    ax(1) = 0
End If
Timer9(Index).Enabled = False
End Sub
Private Sub Timer8_Timer() '���D��
If x = 1 Then s(0) = s(0) + 1 Else s(0) = s(0) - 1
Image1.Top = Image1.Top - (12 - s(0)) * x
'Shape1.Left = Shape1.Left + 1 * x
'Shape1.Width = Shape1.Width - 2 * x
If s(0) = 11 Then
    If zxcv(6) = 0 Then
        s(0) = 12: x = -1
        zxcv(6) = zxcv(6) + 1
    Else
        zxcv(6) = zxcv(6) - 1
    End If
End If
If Image1.Top >= ttop Then x = 1: zxcv(0) = 0: s(0) = 0: zxcv(6) = 0: Image1.Top = ttop: Timer8.Enabled = False
DoEvents
End Sub
Private Sub Timer1_Timer() '�W��
If metor = 0 Or Timer8.Enabled = True Then Exit Sub

Image1.Top = Image1.Top - 10
Shape1.Top = Image1.Top + sh
If Image1.Top < 0 + Line1.Y1 Then Image1.Top = 0 + Line1.Y1: Shape1.Top = Image1.Top + sh
End Sub
Private Sub Timer2_Timer() '�U��
If metor = 0 Or Timer8.Enabled = True Then Exit Sub

Image1.Top = Image1.Top + 10
Shape1.Top = Image1.Top + sh
If Image1.Top + Image1.Height > Form1.ScaleHeight - 10 Then Image1.Top = Form1.ScaleHeight - Image1.Height - 10: Shape1.Top = Image1.Top + sh
End Sub
Private Sub Timer3_Timer() '����
If metor = 0 Then Exit Sub

Image1.Left = Image1.Left - 10 * run(0)
Shape1.Left = Image1.Left + Image1.Width \ 2
Timer5(kikyou).Enabled = True '�Ұʰʵe
If Image1.Left < 0 + 10 Then Image1.Left = 0 + 10
End Sub
Private Sub Timer4_Timer() '�k��
If metor = 0 Then Exit Sub

If Image1.Left + Image1.Width \ 2 >= Form1.ScaleWidth \ 2 Then '�a�ϬO�_����
    If ok = 0 Then
        zxcv(7) = 1
        Timer6.Enabled = True
        Shape1.Left = Image1.Left + Image1.Width \ 2 - 20
        Exit Sub
    Else
        If Image1.Left + Image1.Width > Form1.ScaleWidth - 10 Then Image1.Left = Form1.ScaleWidth - Image1.Width - 10
    End If
End If
Image1.Left = Image1.Left + 10 * run(1)
Shape1.Left = Image1.Left + Image1.Width \ 2 - 20
Timer5(kikyou).Enabled = True '�Ұʰʵe
End Sub
Private Sub Timer5_Timer(Index As Integer) '����ʵe��
k(kikyou) = k(kikyou) + 1
Select Case kikyou
    Case 0
        If k(kikyou) = 6 Then k(kikyou) = 0
        If metor = 0 Then zxcv(3) = zxcv(2) '�ܱ𪺷N�Ӫ����k�ʵe����
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
        If metor = 0 Then zxcv(3) = zxcv(2) '�y�P���u�B�����k�ʵe����
        If zxcv(3) = 0 Then
            Image1.Picture = Image5(k(kikyou)).Picture
        Else
            Image1.Picture = Image10(k(kikyou)).Picture
        End If
End Select
End Sub
Private Sub Timer32_Timer(Index As Integer) '�|�����ʰʵe

If sk(Index) = 7 Then
    sk(Index) = 0
Else
    sk(Index) = sk(Index) + 1
End If

Image20(Index).Picture = Image21(sk(Index)).Picture

If Timer6.Enabled = True Then Image20(Index).Left = Image20(Index).Left - 20 * run(1): Shape15(Index).Left = Shape15(Index).Left - 20 * run(1) '�a�ϲ���
If coll(Image20(Index), Image1) Then
    Label10.Visible = True: Image23.Visible = True: Label10.Caption = Val(Label10.Caption) + 1 & " X": Image20(Index).Left = -1111: Shape15(Index).Left = -1111
    If hp.Width >= 150 Then '�ɦ�
        hp.Width = 200
    Else
        hp.Width = hp.Width + 50
    End If
    Label1.Caption = hp.Width * akiz & "/" & amax
    
    If mp.Width >= 150 Then '���]
        mp.Width = 200
    Else
        mp.Width = mp.Width + 50
    End If
    Label2.Caption = mp.Width * akiz & "/" & amax
End If
End Sub
Private Sub Timer15_Timer() 'GO��ܡ�
s(4) = s(4) + 1
If s(4) Mod 2 = 0 Then
    Label3.Visible = False
Else
    Label3.Visible = True
End If
If ok = 1 Then Timer15.Enabled = False: Label3.Visible = False: Call desd(0)
DoEvents
End Sub
Private Sub Timer25_Timer() 'MP�W�[�t��k
mp.Width = mp.Width + 5
If mp.Width >= 200 Then mp.Width = 200
Label2.Caption = mp.Width * akiz & "/" & amax
End Sub
Public Sub kizzs() '����l��t��k��
If s(10) = 0 And hp.Visible = True Then
    Call voice
    Randomize
    ym = Int(Rnd * 7) + 1
    If hp.Width - ym <= 0 Then
        hp.Visible = False
        Label1.Caption = "0/" & amax
        
        metor = 0
        zxcv(5) = 1
        Call ex(1): Exit Sub
    Else
        hp.Width = hp.Width - ym
        Label1.Caption = hp.Width * akiz & "/" & amax
    End If
    DoEvents
End If
End Sub
Private Sub voice() '�I���n��
'-----�_������
    Set dsx = dxa.DirectSoundCreate("")
    dsx.SetCooperativeLevel Me.hWnd, DSSCL_NORMAL
    bu.lFlags = DSBCAPS_STATIC
    Set dsbx = dsx.CreateSoundBufferFromFile(("�I��.wav"), bu, wf)
    dsbx.Play DSBPLAY_DEFAULT '''''''''''''''''''''''�w�]����(�榸).
'-----�_������
End Sub
Public Sub sppr(a As Integer) '�P�_�O�_������
For f = 1 To ma
    If ppm(1, f) <> 1 Then
        If coll2(Shape1, Shape2(f)) Or coll2(Shape2(f), Shape1) Then
            If coll(Image1, Image6(f)) Or coll(Image6(f), Image1) Then pps = Image6(f).Index: Call power(a): Exit Sub
        End If
    End If
    DoEvents
Next
End Sub
Public Sub power(a As Integer) '�ĤH�l��t��k��
'-------------------------------------'0�㢲�d���,4 ����,5 ���D����
Label7.Visible = True: Label5.Visible = True: xhp.Visible = True: Shape6.Visible = True: Label11.Visible = True '�Ǫ���q���
Timer37.Enabled = False '�s���Ƥ��{
Timer36.Enabled = False '�@�w�ɶ���q


ppm(5, pps) = 1 '����ɪ��S��
Timer11(pps).Enabled = False '���ʵe����
If a = 2 Or a = 5 Then Call tai(a)

d = a '�P�_�O���@�ӵ���
'�ˮ`�[��
If a = 10 Then '���D����
    a = 0: b = 2
Else
    a = 5: b = 0
End If

If ppm(0, pps) = 0 Then
    xhp.Width = 200
    xhp.Visible = True
    ppm(0, pps) = xhp.Width
Else
    xhp.Visible = True
    xhp.Width = ppm(0, pps)
End If

Call voice

'�ˮ`�t��k
If d < 8 And d <> 3 Then '�]��
    y = ys(d)
Else '����
    Randomize
    y = Int(Rnd * 10 + a - b) + 2
End If
'/�ˮ`�t��k

If xhp.Width - y <= 0 Then
    ppm(0, pps) = 0
    xhp.Visible = False
    Label5.Caption = ppm(0, pps) & "/" & amax
    Load Timer7(pps)
    ppm(1, pps) = 1
    Image6(pps).Top = Image6(pps).Top + 15 '�����U��
    Timer7(pps).Enabled = True '�Ұʳ������ʵe
    n = n + 1
Else
    xhp.Width = xhp.Width - y
    ppm(0, pps) = xhp.Width
    Label5.Caption = ppm(0, pps) * akiz & "/" & amax
End If

If te = 1 Then Call news '�O�_���ˮ`

ppm(8, pps) = 0
ppm(7, pps) = ppm(7, pps) + 1 '�s����
If ppm(7, pps) >= 2 Then
    f = "�I"
    If ppm(7, pps) \ 10 >= 1 Then
        f = "�I�I"
    End If
End If
Label11.Caption = ppm(7, pps) & " �s��" & f

'���h
If d > 2 And d <> 5 And d <> 7 Then
    If zxcv(3) = 0 Then
        Image6(pps).Left = Image6(pps).Left + 30
        Shape2(pps).Left = Image6(pps).Left + Image6(pps).Width \ 2
    Else
        Image6(pps).Left = Image6(pps).Left - 30
        Shape2(pps).Left = (Image6(pps).Left + Image6(pps).Width \ 2) - 20
    End If
End If

'����
Label9.Caption = Val(Label9.Caption) + 50

Timer22(pps).Enabled = True

Timer36.Enabled = True '�@�w�ɶ���q
Timer37.Enabled = True '�s���ư{�t
DoEvents
End Sub
Public Sub news() '��ܶˮ`�t��k��
'���ͪ���
s(5) = s(5) + 1
If s(5) = 30000 Then s(5) = 0
Load Label4(s(5))
'���t��m
Label4(s(5)).Top = Image6(pps).Top - Label4(s(5)).Height
Label4(s(5)).Left = Image6(pps).Left + Image6(pps).Width \ 2
'�]�w�ݩ�
Label4(s(5)).Caption = y * akiz
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
Private Sub Timer36_Timer() '�@�w�ɶ����Ǫ�����q����
Label7.Visible = False
Label5.Visible = False
xhp.Visible = False
Shape6.Visible = False
Timer37.Enabled = False
Label11.Visible = False
End Sub
Private Sub Timer37_Timer() '�s���ư{�t
s(21) = s(21) + 1
If s(21) Mod 2 = 0 Then
    s(21) = 0
    Label11.Visible = True
Else
    Label11.Visible = False
End If
End Sub
Private Sub Timer6_Timer() '�D�I�����ʡ�
Form1.AutoRedraw = True
dx = dx + 20 * run(1)
If dx >= 1024 Then dx = 0
If dx <= 100 Then '=====�j��564��N�W�ϤF=====
    Form1.PaintPicture Form1.Picture, 0, 0, 924, 708, dx, 0, 924, 708, vbSrcCopy
Else '�W��form1�I�� �ƻs��          0,0,��e,��,dx,0,��e,��
    Form1.PaintPicture Form1.Picture, 0, 0, 1024 - dx, 708, dx, 0, 1024 - dx, 708, vbSrcCopy
    '�W��form1�I�� �ƻs��           0,0,�ϼe-dx,��,dx,0,�ϼe-dx,��
    Form1.PaintPicture Form1.Picture, 1024 - dx, 0, dx - 100, 708, 0, 0, dx - 100, 708, vbSrcCopy
    '�W��form1�I�� �ƻs��           �ϼe-dx,0,��e+(dx-�ϼe),��,0,0,��e+(dx+�ϼe),��
End If

If dx >= 500 Then
    s(1) = s(1) + 1
    If s(1) = 3 Then s(1) = 0: ok = 1: Timer6.Enabled = False: zxcv(7) = 0 '====�ϼe=====
End If

DoEvents
End Sub
Private Sub sico(a As Integer) '�|��ɱ����t��k
Randomize
f = Int(Rnd * 5) + 1
If f = 2 Then
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
