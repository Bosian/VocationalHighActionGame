Attribute VB_Name = "Module1"
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public kikyou As Integer '����
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
Public music
Public gz '������X��
Public reset(1) As Integer '����ʧ@�C���u��Ū"�@��"��Load�{���X
Public te As Integer '�M�w�O�_���ˮ`�X��
Public keya As Integer, keys As Integer, keyd As Integer, keyup As Integer, keydown As Integer, keyleft As Integer, keyright As Integer '��L�]�w

Public Function coll(a As Object, b As Object) As Boolean '�I���[�ơ�
If (a.Left >= b.Left And a.Left <= b.Left + b.Width And _
   a.Top >= b.Top And a.Top <= b.Top + b.Height) Or _
   (a.Left >= b.Left And a.Left <= b.Left + b.Width And _
   a.Top + a.Height >= b.Top And a.Top + a.Height <= b.Top + b.Height) Or _
   (a.Left + a.Width >= b.Left And a.Left + a.Width <= b.Left + b.Width And _
   a.Top + a.Height >= b.Top And a.Top + a.Height <= b.Top + b.Height) Or _
   (a.Left + a.Width >= b.Left And a.Left + a.Width <= b.Left + b.Width And _
   a.Top >= b.Top And a.Top <= b.Top + b.Height) Then coll = True
End Function
Public Function coll2(a As Object, b As Object) As Boolean '�I���[��2��("����"�ȤW�U����Top)
If a.Top >= b.Top And a.Top <= b.Top + b.Height Or _
   a.Top + a.Height >= b.Top And a.Top + a.Height <= b.Top + b.Height Then coll2 = True
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
