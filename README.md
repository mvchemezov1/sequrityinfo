```VBA
Private Sub UserForm_Initialize()
    CommandButton1.Caption = "Выполнить"
    CommandButton2.Caption = "Отменить"
    OptionButton1.Caption = "Зашифровать"
    OptionButton2.Caption = "Расшифровать"
    'первоначальный выбор переключателя
    OptionButton1.Value = True
    OptionButton2.Value = False
End Sub

Sub текст(видимость)
    'процедура управляет видимостью
    Frame2.Visible = видимость
    TextBox1.Visible = видимость
    TextBox2.Visible = видимость
    TextBox3.Visible = видимость
End Sub

Private Sub CommandButton1_Click()
    If OptionButton1 = True Then
        ' Шифрование текста
        TextBox2.text = EncryptText(TextBox1.text)
    ElseIf OptionButton2 = True Then
        ' Расшифрование текста
        TextBox3.text = DecryptText(TextBox2.text)
    End If
End Sub

Private Function EncryptText(inputText As String) As String
    Dim outputText As String
    Dim key As String
    Dim i As Integer
    Dim charCode As Integer

    Randomize

    ' Генерация случайной гаммы из букв русского языка
    For i = 1 To Len(inputText)
        key = key & Chr(Int(32 * Rnd) + 32)
    Next i

    ' Шифрование с учетом регистра
    For i = 1 To Len(inputText)
        charCode = Asc(Mid(inputText, i, 1))
        If charCode >= 192 And charCode <= 255 Then
            key = key & Chr(charCode)
        Else
            key = key & Chr(Asc(" ") Xor Asc(Mid(inputText, i, 1)))
        End If
    Next i

    ' Шифрование текста
    For i = 1 To Len(inputText)
        charCode = Asc(Mid(inputText, i, 1)) Xor Asc(Mid(key, i, 1))
        If charCode >= 192 And charCode <= 255 Then
            outputText = outputText & Chr(charCode)
        Else
            outputText = outputText & Chr(Asc("@") Xor charCode)
        End If
    Next i

    EncryptText = outputText
End Function

Private Function DecryptText(inputText As String) As String
    Dim outputText As String
    Dim key As String
    Dim i As Integer
    Dim charCode As Integer

    ' Получение ключа из зашифрованного текста
    For i = 1 To Len(inputText)
        key = key & Chr(Asc(Mid(inputText, i, 1)) Xor Asc(Mid(TextBox1.text, i, 1)))
    Next i

    ' Расшифрование текста
    For i = 1 To Len(inputText)
        charCode = Asc(Mid(inputText, i, 1)) Xor Asc(Mid(key, i, 1))
        If charCode >= 192 And charCode <= 255 Then
            outputText = outputText & Chr(charCode)
        Else
            outputText = outputText & Chr(Asc("@") Xor charCode)
        End If
    Next i

    DecryptText = outputText
End Function
```

##Программа
![image](https://github.com/mvchemezov1/sequrityinfo/assets/144443468/4acd68ea-fe62-4f4a-9d5a-6ac5d2de906e)

