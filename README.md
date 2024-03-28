```VBA
Dim S As String
Dim m() As Integer
Dim HO, H, P, Z, K, D As Integer
Private Sub CommandButton1_Click()
K = 0
Randomize
Z = Round(1000 * Rnd())
I = 1
D = 0
While K <> 1
For I = 1 To Z
If (Z \ I) = (Z / I) Then D = D + 1
Next I
If D = 2 Then K = 1: GoTo 1 Else Z = Z - 1: K = 0: D = 0
Wend
1: HO = Z
TextBox2.Text = HO

K = 0
Randomize
Z = Round(1000 * Rnd())
I = 1
D = 0
While K <> 1
For I = 1 To Z
If (Z \ I) = (Z / I) Then D = D + 1
Next I
If D = 2 Then K = 1: GoTo 2 Else Z = Z - 1: K = 0: D = 0
Wend
2: P = Z
TextBox3.Text = P
TextBox4.Text = ""
End Sub

Private Sub CommandButton2_Click()
S = TextBox1.Text
HO = Val(TextBox2.Text)
P = Val(TextBox3.Text)
If S = "" Then
    MsgBox ("Задана пустая строка сообщения")
End If
If HO = 0 Then
    MsgBox ("Не задан начальнай хэш")
End If
If P = 0 Then
    MsgBox ("Не задан хэш модуль")
End If

For I = 1 To Len(S)
ReDim m(I)
m(I) = Val(Mid(S, I, 1))
Next I

For I = 1 To Len(S)
H = ((m(I) + HO) ^ 2) Mod P
HO = H

Next
TextBox4.Text = H
End Sub

Private Sub CommandButton3_Click()
UserForm1.Hide
TextBox1.Text = ""
TextBox2.Text = ""
TextBox3.Text = ""
TextBox4.Text = ""
End Sub
```

##Программа 1
![image](https://github.com/mvchemezov1/sequrityinfo/assets/144443468/5d4a9d5b-978b-4ec0-916a-e3c43066a8c3)

##Программа 2
![image](https://github.com/mvchemezov1/sequrityinfo/assets/144443468/96e19c5d-3330-42d4-9b8e-63f70569a075)

##Программа 3
![image](https://github.com/mvchemezov1/sequrityinfo/assets/144443468/359b54c0-73d2-4108-8e54-39b1de5390c4)


