```VBA
Private Sub UserForm_Initialize()
'Эта процедура выполняется при создании формы. Она устанавливает заголовки кнопок и переключателей, а также устанавливает начальное состояние переключателя OptionButton1 в значение True.
CommandButton1.Caption = "Выполнить"
'Устанавливает заголовок кнопки CommandButton1 в значение "Выполнить".
CommandButton2.Caption = "Отменить"
'Устанавливает заголовок кнопки CommandButton2 в значение "Отменить".
OptionButton1.Caption = "Зашифровать"
'Устанавливает заголовок переключателя OptionButton1 в значение "Зашифровать".
OptionButton2.Caption = "Расшифровать"
'Устанавливает заголовок переключателя OptionButton2 в значение "Расшифровать".
OptionButton1.Value = True
'Устанавливает значение переключателя OptionButton1 в True.
OptionButton2.Value = False
'Устанавливает значение переключателя OptionButton2 в False.
End Sub
Sub текст(видимость)
'Эта процедура управляет видимостью текстовых полей и кнопок на форме.
Frame2.Visible = видимость
'Устанавливает видимость фрейма Frame2 в значение видимость.
TextBox1.Visible = видимость
'Устанавливает видимость текстового поля TextBox1 в значение видимость.
TextBox2.Visible = видимость
'Устанавливает видимость текстового поля TextBox2 в значение видимость.
TextBox3.Visible = видимость
'Устанавливает видимость текстового поля TextBox3 в значение видимость.
End Sub

Private Sub CommandButton1_Click()
  Dim str As String, strT As String, strR As String
  Dim N As Long, i As Long
  slogan = "компьютер"
  oldkey = "абвгдежзийклмнопрстуфхцчшщъыьэюя"
  newKey = slogan

  ' Удаляем повторения символов из нового ключа
  Dim j As Long, char As String
  For j = 1 To Len(oldkey)
    char = Mid(oldkey, j, 1)
    If InStr(newKey, char) = 0 Then
      newKey = newKey & char
    End If
  Next j
   
  str = TextBox1.text
  N = Len(str)
  strT = Space(N)
   
  oldk = UCase(oldkey)
  newk = UCase(newKey)
   
  If OptionButton1.Value = True Then
    If TextBox1.text = "" Then
      MsgBox ("Введите исходный текст"): GoTo 1
    Else
      'зашифровка по лозунговому шифру
      For i = 1 To N
        'зашифровка по лозунговому шифру
      tmp = Mid(str, i, 1)
'определяем место строки исходного текста в исходном алфавите

k = InStr(1, oldkey, tmp)
If k = 0 Then
k = InStr(1, oldk, tmp)
If k = 0 Then
'если в исходном алфавите такого символа нет
'то переносим его в новую строку не изменяя
Mid(strT, i, 1) = tmp
Else
'заменяем соответствующий пробел на символ нового алфавита
Mid(strT, i, 1) = Mid(newk, k, 1)
End If
Else
Mid(strT, i, 1) = Mid(newKey, k, 1)
End If
Next i
    End If
    End If
TextBox2.text = strT
  If OptionButton2.Value = True Then
  str = TextBox1.text
  N = Len(str)
  strR = Space(N)
    If TextBox1.text = "" Then
      MsgBox ("Введите шифротекст"): GoTo 1
    Else
      'расшифровка по лозунговому шифру
      For i = 1 To N
      tmp = Mid(str, i, 1)
        k = InStr(1, newKey, tmp)
If k = 0 Then
k = InStr(1, newk, tmp)
If k = 0 Then
'если в исходном алфавите такого символа нет
'то переносим его в новую строку не изменяя
Mid(strR, i, 1) = tmp
Else
'заменяем соответствующий пробел на символ нового алфавита
Mid(strR, i, 1) = Mid(oldk, k, 1)
End If
Else
Mid(strR, i, 1) = Mid(oldkey, k, 1)
End If
Next i
      End If
  End If
TextBox3.text = strR
1:
End Sub
Private Sub CommandButton2_Click()
'Эта процедура выполняется при нажатии кнопки CommandButton2. Она закрывает форму.
UserForm2.Hide
'Закрывает форму.
End Sub
Private Sub OptionButton1_Click()
'Эта процедура выполняется при щелчке переключателем OptionButton1. Она устанавливает состояние переключателя OptionButton1 в значение True и состояние переключателя OptionButton2 в значение False.
OptionButton1.Value = True
'Устанавливает состояние переключателя OptionButton1 в значение True.
расшифровать = False
'Устанавливает значение переменной расшифровать в значение False.
TextBox1.text = TextBox3.text
'Записывает текст из текстового поля TextBox3 в текстовое поле TextBox1.
TextBox1.SetFocus
'Устанавливает фокус на текстовое поле TextBox1.
End Sub
Private Sub OptionButton2_Click()
'Эта процедура выполняется при щелчке переключателем OptionButton2. Она устанавливает состояние переключателя OptionButton2 в значение True и состояние переключателя OptionButton1 в значение False.
зашифровать = False
'Устанавливает значение переменной зашифровать в значение False.
OptionButton2.Value = True
'Устанавливает состояние переключателя OptionButton2 в значение True.
TextBox1.text = TextBox2.text
'Записывает текст из текстового поля TextBox2 в текстовое поле TextBox1.
TextBox1.SetFocus
'Устанавливает фокус на текстовое поле TextBox1.
End Sub
```

##Программа
![image](https://github.com/mvchemezov1/sequrityinfo/assets/144443468/b88ac88a-03c1-4521-8c37-f38b93d7cf04)
