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


Private Sub Label1_Click()

End Sub

``` 
