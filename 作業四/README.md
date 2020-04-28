以下為bsformula非正式作業 的程式碼:</bs>
```
Private Sub 按鈕1_Click()

Dim d1#, d2#, Nd1#, Nd2#
Dim rf#, Std#, Tt#
Dim callvalue#, implicedvalue#
Dim r%, us#, ls#, s%, k%
With sheet1
k = Cells(2, 1)
rf = Cells(2, 2)

Std = Cells(2, 3)
Tt = Cells(2, 4) / 250
us = Cells(5, 2)
ls = Cells(6, 2)

Range("a8:c300").Clear
Cells(8, 1) = "標的物價格"
Cells(8, 2) = "理論價值"
Cells(8, 3) = "內含價值"
Cells(8, 4) = "Delta值"

Cells(8, 1).Interior.ColorIndex = 35
Cells(8, 2).Interior.ColorIndex = 35
Cells(8, 3).Interior.ColorIndex = 35


For s = ls To us
d1 = CCur(Log(CCur(s) / CCur(k)) + CCur(CCur(rf) + CCur(0.5) * Std ^ 2) * Tt) / CCur(Std * Tt ^ 0.5)
d2 = CCur(Log(CCur(s) / CCur(k)) + CCur(CCur(rf) - CCur(0.5) * Std ^ 2) * Tt) / CCur(Std * Tt ^ 0.5)
Nd1 = Application.WorksheetFunction.NormSDist(d1)
Nd2 = Application.WorksheetFunction.NormSDist(d2)
callvalue = s * (Nd1) - k * Exp(-rf * Tt) * (Nd2)
implicedvalue = Application.Max(0, (s - k))
Cells(s - ls + 9, 1) = s
Cells(s - ls + 9, 2) = callvalue
Cells(s - ls + 9, 3) = implicatedvalue
Cells(s - ls + 9, 4) = Nd1
Next
End With

End Sub
```
