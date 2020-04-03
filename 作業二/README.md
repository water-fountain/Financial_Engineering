4/1 02:19 未完成操作說明<br>



學習歷程:在算ytm時一直做不出正確答案，後來發現是儲存格的小數點位數不夠<br>
關於spot rate的內容仍不太了解，不確定該做哪種債券的spot RATE<br>
如果是美國公債的話就跟一般債券稍有不同了。<br>

4/3後記:上完課後更加理解老師的意思了，為了更方便操作製作了excel表格式的總計算器<br>
操作說明如下:<br>
https://i.imgur.com/YXDXPzY.jpg
https://i.imgur.com/ULLI8ug.jpg
https://i.imgur.com/AWcJEwl.jpg

程式碼:
``` 
Private Sub CommandButton1_Click()

Dim 合約價格#, 複利期#, d3#
Dim d1#, 利率期#, 年利率#
Dim 售價#, 到期年#, 複利年#

With sheet1
   Range("A1:G10000").Clear
   
    Cells(2, "A") = TextBox1.Text
    Cells(2, "B") = TextBox2.Text
    Cells(2, "C") = TextBox3.Text / 100
    Cells(2, "D") = TextBox4.Text
    Cells(2, "E") = TextBox5.Text
  
售價 = Cells(2, 1)
合約價格 = Cells(2, 2)
年利率 = Cells(2, 3)
到期年 = Cells(2, 4)
複利年 = Cells(2, 5)


Range("F2").Select
ActiveCell.FormulaR1C1 = "=RC[-2]/RC[-1]"
複利期 = Cells(2, 6)

Range("G2").Select
ActiveCell.FormulaR1C1 = "=RC[-4]*RC[-2]"
利率期 = Cells(2, 7)





Cells(1, 1) = "售價"
Cells(1, 2) = "合約價格"
Cells(1, 3) = "利率(年)"
Cells(1, 4) = "多久後到期(年)"
Cells(1, 5) = "多久複利1次(年)"
Cells(1, 6) = "共幾期"
Cells(1, 7) = "利率(期)"

Range("C2").Select
Selection.NumberFormatLocal = "0.00%"
Range("G2").Select
Selection.NumberFormatLocal = "0.00%"

For i = 1 To 複利期

d1 = 合約價格 * 利率期

Cells(6, 1) = 0
Cells(6 + i, 1) = i

Cells(6, 2) = -售價
Cells(6 + i, 2) = d1
Cells(6 + 複利期, 2) = d1 + 合約價格


Next


Range("B4").Select
ActiveCell.FormulaR1C1 = "=IRR(R[2]C:R[10000]C)"
Selection.NumberFormatLocal = "0.00000%"

Range("D4").Select
Selection.NumberFormatLocal = "0.00000%"
ActiveCell.FormulaR1C1 = "=RC[-2]/R[-2]C[1]"

Cells(4, 1) = "IRR(期)"
Cells(4, 3) = "到期值利率(年)"



End With
End Sub

``` 
用bootstrap方法計算spot rate，計算的是benchmark one year annual pay bond
程式碼:
``` 
Private Sub CommandButton1_Click()

Dim 合約價格#, c1#, d3#
Dim D1#, D2#, c2#
Dim YTMONE#, YTMTWO#, YTMTHREE#
Dim E1#, E2#, E3#
Dim F1#, F2#, F3#
Dim G1#, G2#, G3#

With sheet1
   Range("A1:G10000").Clear
   
    Cells(2, "A") = TextBox1.Text
    Cells(2, "B") = TextBox2.Text / 100
  
    Cells(2, "C") = TextBox4.Text / 100
    Cells(2, "D") = TextBox5.Text / 100
  
Range("B2").Select
Selection.NumberFormatLocal = "0.00%"

Range("C2").Select
Selection.NumberFormatLocal = "0.00%"

Range("D2").Select
Selection.NumberFormatLocal = "0.00%"

合約價格 = Cells(2, 1)
YTMONE = Cells(2, 2)
YTMTWO = Cells(2, 3)
YTMTHREE = Cells(2, 4)


'第一年YTM = 第一年SPOTRATE

D2 = (合約價格 * YTMTWO) / (1 + YTMTWO) + (((合約價格 * YTMTWO) + 合約價格) / ((1 + YTMTWO) ^ 2))
D1 = (合約價格 * YTMTWO) / (1 + YTMONE)

Cells(3, 1) = D2
Cells(3, 2) = D1
c2 = Cells(3, 1)
c1 = Cells(3, 2)


d3 = ((合約價格 * YTMTWO) + 合約價格) / (c2 - c1)
Cells(3, 3) = d3

Cells(3, 1).Font.ColorIndex = 2
Cells(3, 2).Font.ColorIndex = 2
Cells(3, 3).Font.ColorIndex = 2

Range("C4").Select
ActiveCell.FormulaR1C1 = "=SQRT(R[-1]C)-1"
Selection.NumberFormatLocal = "0.0000%"

'3

E3 = Cells(4, 3)
 
E1 = ((合約價格 * YTMTHREE) / (1 + YTMTHREE)) + ((合約價格 * YTMTHREE) / ((1 + YTMTHREE) ^ 2)) + (((合約價格 * YTMTHREE) + 合約價格) / ((1 + YTMTHREE) ^ 3))

E2 = ((合約價格 * YTMTHREE) / (1 + YTMONE)) + ((合約價格 * YTMTHREE) / (1 + E3) ^ 2)

Cells(5, 1) = E1
Cells(5, 2) = E2

F1 = Cells(5, 1)
F2 = Cells(5, 2)

F3 = (合約價格 * (1 + YTMTHREE)) / (F1 - F2)
Cells(5, 4) = F3

G3 = F3 ^ (1 / 3)
Cells(5, 5) = G3


Range("D4").Select
ActiveCell.FormulaR1C1 = "=R[1]C[1]-1"
Selection.NumberFormatLocal = "0.0000%"

Cells(5, 1).Font.ColorIndex = 2
Cells(5, 2).Font.ColorIndex = 2
Cells(5, 3).Font.ColorIndex = 2
Cells(5, 4).Font.ColorIndex = 2
Cells(5, 5).Font.ColorIndex = 2

Cells(4, 2) = YTMONE

Cells(4, 1) = "每年SpotRate"
Cells(1, 1) = "合約價格"
Cells(1, 2) = "第1年YTM"
Cells(1, 3) = "第2年YTM"
Cells(1, 4) = "第3年YTM"
Range("B4").Select

Selection.NumberFormatLocal = "0.0000%"

End With
End Sub

``` 
forward rate程式碼:
``` 
Private Sub CommandButton1_Click()

Dim 合約價格#, 複利期#, d3#
Dim d1#, ra#, rb#
Dim 售價#, 到期年#, f1#
Dim YTM#, 到#, 複#

With sheet1

   Range("A1:G10000").Clear
   
    Cells(2, "A") = TextBox1.Text / 100
    Cells(2, "B") = TextBox2.Text / 100
   
  
ra = Cells(2, 1)
rb = Cells(2, 2)


d1 = (1 + ra)
d3 = (1 + rb) ^ 2


f1 = d3 / d1 - 1

Cells(2, 3) = f1



    Range("A2:C2").Select
    Selection.Style = "Percent"






End With

End Sub
``` 
