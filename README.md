##Работа в Excel в строке для формул
![image](https://github.com/mvchemezov1/sequrityinfo/assets/144443468/3acb0bd0-f468-43b6-a43e-832c64651e87)

##Внутренняя часть
```VBA
Private Sub CommandButton1_Click()
'шифровка
m = TextBox1.Value
n = Len(m)
k = TextBox4.Value
If TextBox1.Text = "" Then MsgBox "Введите текст": GoTo 1
If TextBox4.Text = "" Then MsgBox "Введите ключ": GoTo 1
For i = 1 To n
r = Ca(Mid(m, i, 1), k)
c = c & r
Next i
TextBox2.Value = c
1:
End Sub
Private Function Ca(A, Sh)
Dim chrnew As Integer
Select Case Asc(A)
        Case 224 To 255 'строчные буквы
        chrnew = ModMew(Asc(A) + Sh - 224, 32) + 224
        Ca = Chr(chrnew)
        Case 192 To 223 'прописные буквы
        chrnew = ModMew(Asc(A) + Sh - 192, 32) + 192
        Ca = Chr(chrnew)
        Case Else
        Ca = A
        End Select
End Function
Private Function ModMew(A, B)
If A >= 0 Then
ModMew = A Mod B
Else
ModMew = (B + A) Mod B
End If
End Function
Private Sub CommandButton2_Click()
'расшифровка
m = TextBox2.Value
c = Space(n)
k = TextBox4.Value
If TextBox2.Text = "" Then MsgBox "Введите шифротекст": GoTo 1
If TextBox4.Text = "" Then MsgBox "Введите ключ": GoTo 1
n = Len(m)
For i = 1 To n
r = Ca(Mid(m, i, 1), -k)
c = c & r
Next i
TextBox3.Value = c
1:
End Sub

Private Sub TextBox1_Change()
TextBox1.Value = "Чем тоньше лед, тем больше всем хочется узнать, выдержит ли он."
End Sub
```

##Программа
![image](https://github.com/mvchemezov1/sequrityinfo/assets/144443468/5c231f3a-daad-4a0d-9649-77547400eaba)

