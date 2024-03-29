```VBA
Private n As Long
'Эта переменная используется для хранения количества символов в исходной строке.
Const ABB As String = "абвгдеёжзийклмнопрстуфхцчшщъыьэюя"
'Эта константа определяет исходный алфавит.
Const HoBABB As String = "рстуфхцчшщъыьэюяабвгдеёжзийклмноп"
'Эта константа определяет новый алфавит.
Private Sub AnotherCipher()
'Этот подпрограмма запускает шифрование и расшифровку.
Dim STR As String, strT As String, strR As String
'Эти переменные используются для хранения исходной строки, зашифрованной строки и расшифрованной строки.
STR = "Простая замена это один из самых древних шифров"
'Эта строка используется в качестве исходной строки для шифрования.
STR = LCase(STR)
'Эта строка преобразует исходную строку в нижний регистр.
Debug.Print STR
'Эта строка выводит исходную строку в окно Immediate.
Decode STR, strR, ABB, HoBABB
'Эта строка зашифровывает исходную строку с использованием исходного и нового алфавитов.
Debug.Print strR
'Эта строка выводит зашифрованную строку в окно Immediate.
Decode strR, strT, HoBABB, ABB
'Эта строка расшифровывает зашифрованную строку с использованием нового и исходного алфавитов.
Debug.Print strT
'Эта строка выводит расшифрованную строку в окно Immediate.
Private Sub Decode(STR, STRS, Oldkey, NewKey)
'Эта подпрограмма выполняет шифрование или расшифровку строки.
Dim n As Long, i As Long
'Эти переменные используются для хранения количества символов в исходной строке и индекса текущего символа в исходной строке.
Dim tmp As String
'Эта переменная используется для хранения текущего символа исходной строки.
n = Len(STR)
'Эта строка определяет количество символов в исходной строке.
STRS = Space(n)
'Эта строка создает строку, состоящую из n пробелов.
For i = 1 To n
'Эта цикл повторяется n раз.
tmp = Mid(STR, i, 1)
'Эта строка выделяет текущий символ исходной строки.
k = InStr(1, Oldkey, tmp)
'Эта строка определяет место текущего символа исходной строки в исходном алфавите.
If k = 0 Then
'Эта ветвь выполняется, если в исходном алфавите нет текущего символа.
Mid(STRS, i, 1) = tmp
'Эта строка копирует текущий символ исходной строки в преобразованную строку.
Else
'Эта ветвь выполняется, если в исходном алфавите есть текущий символ.
Mid(STRS, i, 1) = Mid(NewKey, k, 1)
'Эта строка заменяет текущий символ исходной строки на соответствующий символ нового алфавита.
End If
'Эта строка завершает ветвь.
Next i
'Эта строка завершает цикл.
End Sub
'Эта строка завершает подпрограмму.
```


