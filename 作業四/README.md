學習歷程:</bs>

![](https://i.imgur.com/XZuKtph.jpg) </bs>
流程圖
![](https://i.im</bs>gur.com/3vz7guI.jpg) </bs>

計算結果:</bs>
![](https://i.imgur.com/L3Moua2.jpg) </bs>

以下為作業四的程式碼:</bs>
```
Private Sub 按鈕1_Click()

Dim d1#, d2#, Nd1#, Nd2#
Dim rf#, Std#, T#
Dim ca#, im#, dd#
Dim r%, sigma#, ls#, S%, K%
Dim fre#, div#, kk#, D#
Dim Ndd1#, Ndd2#
With sheet1

Range("a6:b300").Clear

'S為現價currently price，年化波動度σ為sigma，T為年，r為利率，K為履約價格exercise price，fre一年幾次股利，div一次發多少

S = Cells(2, 1)
r = Cells(2, 2)
sigma = Cells(2, 3)
T = Cells(2, 4)
kk = Cells(2, 5)
fre = Cells(2, 6)
div = Cells(2, 7)



Cells(1, 1) = "現價"
Cells(1, 2) = "無風險利率"
Cells(1, 3) = "年化波動度"
Cells(1, 4) = "距到期幾年"
Cells(1, 5) = "履約價格"
Cells(1, 6) = "一年幾次股利"
Cells(1, 7) = "一次發多少"
Cells(4, 7) = "股利現值"

im = T * fre

ca = 12 / fre

'計算股利現值:

For i = 1 To im

ls = ca * i - 2



Cells(5 + i, 1) = i
Cells(5 + i, 2) = ls



Cells(5, 7).Value = Application.WorksheetFunction.Sum(Range(Cells(6, 3), Cells(6 + im, 3)))

Next

dd = Cells(5, 7)


K = Cells(5, 8)




End With

End Sub



Function BSECall(K, r, sigma, T, S)

d1 = (Log(S / K) + (r + 0.5 * sigma ^ 2) * T) / (sigma * Sqr(T))
d2 = d1 - sigma * Sqr(T)
Nd1 = Application.NormSDist(d1)
Nd2 = Application.NormSDist(d2)
BSECall = S * Nd1 - Exp(-r * T) * K * Nd2


End Function


Function BSEPut(K, r, sigma, T, S)

d1 = (Log(S / K) + (r + 0.5 * sigma ^ 2) * T) / (sigma * Sqr(T))
d2 = d1 - sigma * Sqr(T)
Ndd1 = Application.NormSDist(-d1)
Ndd2 = Application.NormSDist(-d2)
BSEPut = (Exp(-r * T) * K * Ndd2) - S * Ndd1


End Function

```
