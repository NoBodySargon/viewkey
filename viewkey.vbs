'17/02/2013 ITVolna
'Скрипт записи в файл новых значений пары - имя пк и его MAC
'Скрипт содержит две функции:macfunc - определения MAC и MACWrite - Записи MAC в файл
'
'
set WshShell = CreateObject("WScript.Shell")
Set objNetwork = CreateObject("WScript.Network")
strComputer = objNetwork.ComputerName 'Имя ПК
Set oFSO = CreateObject("Scripting.FileSystemObject")

strArchPath = oFSO.GetParentFolderName(WScript.ScriptFullName) & "\"
strArchName = strArchPath & "key.txt"
If oFSO.FileExists(strArchName) Then
MsgBox "Продолжим работу с существующим файлом"
Else
MsgBox "Будет создан новый файл"
oFSO.CreateTextFile(strArchName)
End If

' функция опеределения MAC адреса сетевой карты, в функцию должно передаваться имя ПК
Function macfunc(strComputer)
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set IPConfigSet = objWMIService.ExecQuery _
 ("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")
For Each IPConfig In IPConfigSet
'MAC адрес
   strMACAddress = IPConfig.MACAddress(i)
   macfunc=strMACAddress & " IP:" & IPConfig.IPAddress(i)'Возвращенное значение
  
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


'Объявляем функцию записи в файл МАК адреса
Function MACWrite()
  Dim fso, f
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set f = fso.OpenTextFile(strArchName,8) '1-чтение 2-для записи 8 - дозапись
  ' Записать строку с переводом на новую строку.
  f.WriteLine((strComputer)) '& ": "&+macfunc(strComputer)) 'Пишем имя ПК и MAC
  f.WriteLine("MAC: "&+macfunc(strComputer)) 'Пишем имя ПК и MAC
  f.Close
End Function

' Процедура которая открывает файл и ищет в нем строку с именем ПК и MAC адресом, если MAC Машины и имя ПК совпадает со значением из файла, то выходим из скрипта
' Если такого МАК адреса нет, то пишем в файл MAC данной машины
Sub OpenFileScan()
  Dim fso, f, readmac
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set f = fso.OpenTextFile(strArchName,1) '1-чтение 2-для записи 8 - дозапись
If fso.FileExists(strArchName) Then 'Файл со списком ПК
set readfile = fso.OpenTextFile(strArchName,1) 'Чтаем файл
While Not readfile.AtEndOfStream 'Читаем построчно, пока не конец файла
    readmac=readfile.ReadLine
 s="MAC: "&+macfunc(strComputer)
  If s=readmac Then ' Если имя ПК и соответствующий ему мак адрес в файле есть то выходим из скрипта
 MsgBox "Этот компютер уже учавствовал в опросе"
 WScript.Quit 'выход
 End If
Wend
MACWrite() ' Вызываем функцию записи в файл Имя ПК и его MAC адреса
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