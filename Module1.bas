Attribute VB_Name = "Module1"
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function BitBlt Lib "gdi32" _
(ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'---------�H�U���_������+���䪺
Public dxa As New DirectX7
Public ds As DirectSound '���W
Public dsb As DirectSoundBuffer
Public bu As DSBUFFERDESC
Public wf As WAVEFORMATEX

Public dscc As DirectSound '�ɹp
Public dsbcc As DirectSoundBuffer

Public dsx As DirectSound '�I��
Public dsbx As DirectSoundBuffer

Public dscd As DirectSound '�ɹp
Public dsbd As DirectSoundBuffer

'Public dsz As DirectSound '�I������
'Public dsbz As DirectSoundBuffer
'---------�H�W���_������+���䪺
Public Const Program_Name As String = "���~���I v 5.0.5 �p��s��"
Public kikyou(1) As Integer '0) 1P���� 1) 2P����
Public player_2 As Integer '�O�_��2P
Public pk_mod As Integer '1P VS 2P
Public music
Public gz '������X��
Public ttop(1) As Integer '�Ȧsimage1.top
Public reset(1) As Integer '����ʧ@�C���u��Ū"�@��"��Load�{���X
Public te As Integer '�M�w�O�_���ˮ`�X��
Public smp() As Single '���Ӫ�MP�q
Public akiz(1) As Integer '�u��e�ש�j�v
Public amax(1) As Integer '�u��e�ש�j�v(�̤j��)

Public keya(1) As Integer, keys(1) As Integer, keyd(1) As Integer, keyup(1) As Integer, keydown(1) As Integer, keyleft(1) As Integer, keyright(1) As Integer '0) 1P��L�]�w 1)2P��L�]�w

Public Function coll(a As Object, b As Object) As Boolean '�I���[�ơ�
If a.Left < b.Left + b.Width And b.Left < a.Left + a.Width And _
   a.Top < b.Top + b.Width And b.Top < a.Top + a.Width Then coll = True
   
'If (a.Left >= b.Left And a.Left <= b.Left + b.Width And _
   a.Top >= b.Top And a.Top <= b.Top + b.Height) Or _
   (a.Left >= b.Left And a.Left <= b.Left + b.Width And _
   a.Top + a.Height >= b.Top And a.Top + a.Height <= b.Top + b.Height) Or _
   (a.Left + a.Width >= b.Left And a.Left + a.Width <= b.Left + b.Width And _
   a.Top + a.Height >= b.Top And a.Top + a.Height <= b.Top + b.Height) Or _
   (a.Left + a.Width >= b.Left And a.Left + a.Width <= b.Left + b.Width And _
   a.Top >= b.Top And a.Top <= b.Top + b.Height) Then coll = True
End Function
Public Function coll2(a As Object, b As Object) As Boolean '�I���[��2��("����"�ȤW�U����Top)
If a.Top >= b.Top And a.Top <= b.Top + b.Height Then coll2 = True
End Function
Public Function coll3(a As Object, b As Object) As Boolean '�I���[��3�� Line2
If (a.X1 >= b.Left And a.X1 <= b.Left + b.Width And _
   a.Y1 >= b.Top And a.Y1 <= b.Top + b.Height) Or _
   (a.X1 >= b.Left And a.X1 <= b.Left + b.Width And _
   a.Y1 + a.Y2 >= b.Top And a.Y1 + a.Y2 <= b.Top + b.Height) Or _
   (a.X1 + a.X2 >= b.Left And a.X1 + a.X2 <= b.Left + b.Width And _
   a.Y1 + a.Y2 >= b.Top And a.Y1 + a.Y2 <= b.Top + b.Height) Or _
   (a.X1 + a.X2 >= b.Left And a.X1 + a.X2 <= b.Left + b.Width And _
   a.Y1 >= b.Top And a.Y1 <= b.Top + b.Height) Then coll3 = True
End Function
Public Function coll4(a As Object, b As Object, c As Integer) As Boolean '�I���[�ơ�
If (a.Left >= b.Left And a.Left <= b.Left + b.Width And _
   ttop(c) >= b.Top And ttop(c) <= b.Top + b.Height) Or _
   (a.Left >= b.Left And a.Left <= b.Left + b.Width And _
   ttop(c) + a.Height >= b.Top And ttop(c) + a.Height <= b.Top + b.Height) Or _
   (a.Left + a.Width >= b.Left And a.Left + a.Width <= b.Left + b.Width And _
   ttop(c) + a.Height >= b.Top And ttop(c) + a.Height <= b.Top + b.Height) Or _
   (a.Left + a.Width >= b.Left And a.Left + a.Width <= b.Left + b.Width And _
   ttop(c) >= b.Top And ttop(c) <= b.Top + b.Height) Then coll4 = True
End Function
Public Function con1(a As Object) As Boolean '���󪺤p�������
If a.Left + a.Width < 0 Then con1 = True
End Function
Public Function con2(a As Object) As Boolean '���󪺤j��k�����
If a.Left > Form1.ScaleWidth Then con2 = True
End Function
Function IsLoad(ByVal a) As Boolean
    Dim s
    
    '�����~�~�����
    On Error Resume Next
    '���wS�O�o����Name,���ɭY���󤣦s�b�N�|�o�Ϳ��~
    s = a
    '�H���L���~�ӧP�_������O�_�s�b
    IsLoad = Not CBool(Err)
    Err = 0
End Function
Public Function fscoll(a As Object) As Boolean '���e��(�ù�)�ˮ`�I��
If a.Left > 0 And a.Left < Form1.ScaleWidth And a.Top > 0 And a.Top < Form1.ScaleHeight Or _
   a.Left + a.Width > 0 And a.Left + a.Width < Form1.ScaleWidth And a.Top > 0 And a.Top < Form1.ScaleHeight Or _
   a.Top + a.Height > 0 And a.Top + a.Height < Form1.ScaleHeight And a.Left > 0 And a.Left < Form1.ScaleWidth Or _
   a.Top + a.Height > 0 And a.Top + a.Height < Form1.ScaleHeight And a.Left + a.Width > 0 And a.Left + a.Width < Form1.ScaleWidth Then fscoll = True
End Function
Public Function Artificial_Intelligence(a As Object, b As Object) As Integer '�Ǫ��l��AI
If a.Left + a.Width \ 2 > b.Left + b.Width Then Artificial_Intelligence = Artificial_Intelligence + 1 '�Ǧb���⪺�k��(�����])
If a.Left + a.Width \ 2 < b.Left Then Artificial_Intelligence = Artificial_Intelligence + 2 '�Ǧb���⪺����(���k�])
If a.Top + a.Height \ 2 > b.Top + b.Height Then Artificial_Intelligence = Artificial_Intelligence + 4 '�Ǧb���⪺�U��(���W�])
If a.Top + a.Height \ 2 < b.Top Then Artificial_Intelligence = Artificial_Intelligence + 7 '�Ǧb���⪺�W��(���U�])
End Function
'����
'If Artificial_Intelligence = 1 Then '���]
'If Artificial_Intelligence = 2 Then '�k�]
'If Artificial_Intelligence = 4 Then '�W�]
'If Artificial_Intelligence = 5 Then '���W
'If Artificial_Intelligence = 6 Then '�k�W
'If Artificial_Intelligence = 7 Then '�U�]
'If Artificial_Intelligence = 8 Then '���U
'If Artificial_Intelligence = 9 Then '�k�U


