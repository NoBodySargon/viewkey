'17/02/2013 ITVolna
'������ ������ � ���� ����� �������� ���� - ��� �� � ��� MAC
'������ �������� ��� �������:macfunc - ����������� MAC � MACWrite - ������ MAC � ����
'
'
set WshShell = CreateObject("WScript.Shell")
Set objNetwork = CreateObject("WScript.Network")
strComputer = objNetwork.ComputerName '��� ��
Set oFSO = CreateObject("Scripting.FileSystemObject")

strArchPath = oFSO.GetParentFolderName(WScript.ScriptFullName) & "\"
strArchName = strArchPath & "key.txt"
If oFSO.FileExists(strArchName) Then
MsgBox "��������� ������ � ������������ ������"
Else
MsgBox "����� ������ ����� ����"
oFSO.CreateTextFile(strArchName)
End If

' ������� ������������ MAC ������ ������� �����, � ������� ������ ������������ ��� ��
Function macfunc(strComputer)
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set IPConfigSet = objWMIService.ExecQuery _
 ("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")
For Each IPConfig In IPConfigSet
'MAC �����
   strMACAddress = IPConfig.MACAddress(i)
   macfunc=strMACAddress & " IP:" & IPConfig.IPAddress(i)'������������ ��������
  
Next
End Function 

Function CreateFile(text)
  Dim fso, tf
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set tf = fso.OpenTextFile(strArchName, 8)
  tf.WriteLine(text) 
  tf.Close
End Function

Function ConvertToKey(regKey)
  Const KeyOffset = 52
  isWin8 = (regKey(66) \ 6) And 1
  regKey(66) = (regKey(66) And &HF7) Or ((isWin8 And 2) * 4)
  j = 24
  Chars = "BCDFGHJKMPQRTVWXY2346789"
  Do
    Cur = 0
    y = 14
    Do 
      Cur = Cur * 256
      Cur = regKey(y + KeyOffset) + Cur
      regKey(y + KeyOffset) = (Cur \ 24)
      Cur = Cur Mod 24
      y = y -1
    Loop While y >= 0
    j = j -1
    winKeyOutput = Mid(Chars, Cur + 1, 1) & winKeyOutput
    Last = Cur
  Loop While j >= 0
  If (isWin8 = 1) Then
    keypart1 = Mid(winKeyOutput, 2, Last)
    insert = "N"
    winKeyOutput = Replace(winKeyOutput, keypart1, _
      keypart1 & insert, 2, 1, 0)
    If Last = 0 Then winKeyOutput = insert & winKeyOutput
  End If
  a = Mid(winKeyOutput, 1, 5)
  b = Mid(winKeyOutput, 6, 5)
  c = Mid(winKeyOutput, 11, 5)
  d = Mid(winKeyOutput, 16, 5)
  e = Mid(winKeyOutput, 21, 5)
  ConvertToKey = a & "-" & b & "-" & c & "-" & d & "-" & e
End Function


'��������� ������� ������ � ���� ��� ������
Function MACWrite()
  Dim fso, f
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set f = fso.OpenTextFile(strArchName,8) '1-������ 2-��� ������ 8 - ��������
  ' �������� ������ � ��������� �� ����� ������.
  f.WriteLine((strComputer)) '& ": "&+macfunc(strComputer)) '����� ��� �� � MAC
  f.WriteLine("MAC: "&+macfunc(strComputer)) '����� ��� �� � MAC
  f.Close
End Function

' ��������� ������� ��������� ���� � ���� � ��� ������ � ������ �� � MAC �������, ���� MAC ������ � ��� �� ��������� �� ��������� �� �����, �� ������� �� �������
' ���� ������ ��� ������ ���, �� ����� � ���� MAC ������ ������
Sub OpenFileScan()
  Dim fso, f, readmac
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set f = fso.OpenTextFile(strArchName,1) '1-������ 2-��� ������ 8 - ��������
If fso.FileExists(strArchName) Then '���� �� ������� ��
set readfile = fso.OpenTextFile(strArchName,1) '����� ����
While Not readfile.AtEndOfStream '������ ���������, ���� �� ����� �����
    readmac=readfile.ReadLine
 s="MAC: "&+macfunc(strComputer)
  If s=readmac Then ' ���� ��� �� � ��������������� ��� ��� ����� � ����� ���� �� ������� �� �������
 MsgBox "���� �������� ��� ����������� � ������"
 WScript.Quit '�����
 End If
Wend
MACWrite() ' �������� ������� ������ � ���� ��� �� � ��� MAC ������
End If
End Sub

OpenFileScan()
regKey = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\"
DigitalProductId = WshShell.RegRead(regKey & "DigitalProductId")
WinProductName = "Windows Product Name: " & _
  WshShell.RegRead(regKey & "ProductName") & vbNewLine
WinProductID = "Windows Product ID: " & _
  WshShell.RegRead(regKey & "ProductID") & vbNewLine
WinProductKey = ConvertToKey(DigitalProductId)
strProductKey ="Windows Key: " & WinProductKey
WinProductID = WinProductName & WinProductID & strProductKey
MsgBox(WinProductID)
CreateFile(WinProductID)