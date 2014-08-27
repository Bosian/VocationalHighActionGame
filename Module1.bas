Attribute VB_Name = "Module1"
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'---------以下為奇摩知識+內找的
Public dxa As New DirectX7
Public ds As DirectSound '火柱
Public dsb As DirectSoundBuffer
Public bu As DSBUFFERDESC
Public wf As WAVEFORMATEX

Public dscc As DirectSound '暴雷
Public dsbcc As DirectSoundBuffer

Public dsx As DirectSound '碰撞
Public dsbx As DirectSoundBuffer

Public dscd As DirectSound '暴雷
Public dsbd As DirectSoundBuffer

'Public dsz As DirectSound '背景音樂
'Public dsbz As DirectSoundBuffer
'---------以上為奇摩知識+內找的
Public kikyou(1) As Integer '0) 1P角色 1) 2P角色
Public player_2 As Integer '是否有2P
Public pk_mod As Integer '1P VS 2P
Public music
Public gz '防止移除出錯
Public ttop(1) As Integer '暫存image1.top
Public reset(1) As Integer '限制動作遊戲只能讀"一次"的Load程式碼
Public te As Integer '決定是否跳傷害出來
Public smp() As Integer '消耗的MP量
Public akiz As Integer '真實寬度放大率
Public amax As Integer '真實寬度放大率(最大值)
Public keya(1) As Integer, keys(1) As Integer, keyd(1) As Integer, keyup(1) As Integer, keydown(1) As Integer, keyleft(1) As Integer, keyright(1) As Integer '0) 1P鍵盤設定 1)2P鍵盤設定

Public Function coll(a As Object, b As Object) As Boolean '碰撞涵數●
If (a.Left >= b.Left And a.Left <= b.Left + b.Width And _
   a.Top >= b.Top And a.Top <= b.Top + b.Height) Or _
   (a.Left >= b.Left And a.Left <= b.Left + b.Width And _
   a.Top + a.Height >= b.Top And a.Top + a.Height <= b.Top + b.Height) Or _
   (a.Left + a.Width >= b.Left And a.Left + a.Width <= b.Left + b.Width And _
   a.Top + a.Height >= b.Top And a.Top + a.Height <= b.Top + b.Height) Or _
   (a.Left + a.Width >= b.Left And a.Left + a.Width <= b.Left + b.Width And _
   a.Top >= b.Top And a.Top <= b.Top + b.Height) Then coll = True
End Function
Public Function coll2(a As Object, b As Object) As Boolean '碰撞涵數2●("絕招"僅上下偵測Top)
If a.Top >= b.Top And a.Top <= b.Top + b.Height Then coll2 = True
End Function
Public Function coll3(a As Object, b As Object) As Boolean '碰撞涵數3● Line2
If (a.X1 >= b.Left And a.X1 <= b.Left + b.Width And _
   a.Y1 >= b.Top And a.Y1 <= b.Top + b.Height) Or _
   (a.X1 >= b.Left And a.X1 <= b.Left + b.Width And _
   a.Y1 + a.Y2 >= b.Top And a.Y1 + a.Y2 <= b.Top + b.Height) Or _
   (a.X1 + a.X2 >= b.Left And a.X1 + a.X2 <= b.Left + b.Width And _
   a.Y1 + a.Y2 >= b.Top And a.Y1 + a.Y2 <= b.Top + b.Height) Or _
   (a.X1 + a.X2 >= b.Left And a.X1 + a.X2 <= b.Left + b.Width And _
   a.Y1 >= b.Top And a.Y1 <= b.Top + b.Height) Then coll3 = True
End Function
Public Function coll4(a As Object, b As Object, c As Integer) As Boolean '碰撞涵數●
If (a.Left >= b.Left And a.Left <= b.Left + b.Width And _
   ttop(c) >= b.Top And ttop(c) <= b.Top + b.Height) Or _
   (a.Left >= b.Left And a.Left <= b.Left + b.Width And _
   ttop(c) + a.Height >= b.Top And ttop(c) + a.Height <= b.Top + b.Height) Or _
   (a.Left + a.Width >= b.Left And a.Left + a.Width <= b.Left + b.Width And _
   ttop(c) + a.Height >= b.Top And ttop(c) + a.Height <= b.Top + b.Height) Or _
   (a.Left + a.Width >= b.Left And a.Left + a.Width <= b.Left + b.Width And _
   ttop(c) >= b.Top And ttop(c) <= b.Top + b.Height) Then coll4 = True
End Function
Public Function con1(a As Object) As Boolean '物件的小於左邊邊界
If a.Left + a.Width < 0 Then con1 = True
End Function
Public Function con2(a As Object) As Boolean '物件的大於右邊邊界
If a.Left > Form1.ScaleWidth Then con2 = True
End Function
Function IsLoad(ByVal a) As Boolean
    Dim s
    
    '有錯誤繼續執行
    On Error Resume Next
    '指定S是這物件的Name,此時若物件不存在就會發生錯誤
    s = a
    '以有無錯誤來判斷此物件是否存在
    IsLoad = Not CBool(Err)
    Err = 0
End Function
Public Function fscoll(a As Object) As Boolean '全畫面(螢幕)傷害碰撞
If a.Left > 0 And a.Left < Form1.ScaleWidth And a.Top > 0 And a.Top < Form1.ScaleHeight Or _
   a.Left + a.Width > 0 And a.Left + a.Width < Form1.ScaleWidth And a.Top > 0 And a.Top < Form1.ScaleHeight Or _
   a.Top + a.Height > 0 And a.Top + a.Height < Form1.ScaleHeight And a.Left > 0 And a.Left < Form1.ScaleWidth Or _
   a.Top + a.Height > 0 And a.Top + a.Height < Form1.ScaleHeight And a.Left + a.Width > 0 And a.Left + a.Width < Form1.ScaleWidth Then fscoll = True
End Function
