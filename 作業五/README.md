使用說明:</bs>

啟用編輯 </bs>

![](https://i.imgur.com/czttS3k.jpg) </bs>

啟用巨集 </bs>

![](https://i.imgur.com/wAyjbDX.jpg) </bs>

按下按鈕後開始計算(需等待) </bs>

黃色格子可更改數值，白色格子較不建議更改，N建議<1000 </bs>

![](https://i.imgur.com/hsrEG49.jpg) </bs>


學習歷程:</bs>

![](https://i.imgur.com/tkFHT2l.jpg) </bs>

公式:</bs>
![](https://i.imgur.com/bNULOjP.png) </bs>
![](https://i.imgur.com/YDn6RvN.jpg) </bs>

流程圖</bs>
![](https://i.imgur.com/gsSIf0z.jpg) </bs>

計算結果:</bs>

沒有考慮股利，Convention 1~3處的值來自InputR工作表的P(P=1)</bs>

SHORT RATE的值是每個PATH的最終值，礙於EXCEL限制沒有畫路線圖</bs>

STOCK PRICE的值也是每個PATH走到最後的值。</bs>

![](https://i.imgur.com/hsrEG49.jpg) </bs>

以下為作業五的主要程式碼:</bs>
```
Public Sub GetValues(ByRef T As Variant, ByRef f As Variant, ByRef Df As Variant, Optional Instn As Integer = 1)
    Dim nPoints As Integer, i As Integer
    
    With Worksheets("InputR")
    nPoints = Application.Count(.Range("E:E"))
    
    ReDim T(0 To nPoints - 1) As Double
    ReDim f(0 To nPoints - 1) As Double
    ReDim Df(0 To nPoints - 1) As Double
    
    For i = 0 To nPoints - 1
        T(i) = .Cells(i + 3, 5)
        f(i) = .Cells(i + 3, 6 + Instn)
        Df(i) = .Cells(i + 3, 10 + Instn)
        
   '此處在後台取得forward rate及df/dt，因hull white公式中需要
   
    Next i
    End With
End Sub

Public Function HW_path(T As Variant, f As Variant, Df As Variant, r0 As Double, a As Double, sigma As Double, N As Long) As Variant
    Dim M As Integer, i As Long, j As Long
    Dim dt As Double
    '以上為定義參數
    M = UBound(T) 
    dt = T(1) - T(0)
    ReDim theta(0 To M) As Double
    ReDim r(0 To M, 1 To N) As Double
    '以上為定義上下界

    For j = 1 To N
        r(0, j) = r0
        For i = 0 To M - 1
            theta(i) = Df(i) + a * f(i) + sigma ^ 2 / 2 / a * (1 - Exp(-2 * a * i * dt))
            r(i + 1, j) = r(i, j) + (theta(i) - a * r(i, j)) * dt + sigma * Sqr(dt) * rGauss()
            
'公式中hull white圖片上的兩條公式

    Range("m11").Select
    ActiveCell(i, 1) = r(i, j)

'
            
        Next i
    Next j
    
    HW_path = r
End Function

Public Function HW_path_1(T As Variant, f As Variant, Df As Variant, r0 As Double, a As Double, sigma As Double, N As Long) As Variant
    Dim M As Integer, i As Long, j As Long
    Dim dt As Double
    M = UBound(T)
    dt = T(1) - T(0)
    ReDim theta(0 To M) As Double
    ReDim r(0 To M, 1 To N) As Double

    For j = 1 To N
        r(0, j) = r0
        For i = 0 To M - 1
            theta(i) = Df(i) + a * f(i) + sigma ^ 2 / 2 / a * (1 - Exp(-2 * a * i * dt))
            r(i + 1, j) = r(i, j) + (theta(i) - a * r(i, j)) * dt + sigma * Sqr(dt) * rGauss()
            If r(i + 1, j) < 0 Then r(i + 1, j) = 0 'Convention for r < 0, take r = 0
        Next i
    Next j
    
    HW_path_1 = r
End Function

Public Function HW_path_2(T As Variant, f As Variant, Df As Variant, r0 As Double, a As Double, sigma As Double, N As Long) As Variant
    Dim M As Integer, i As Long, j As Long
    Dim dt As Double
    M = UBound(T)
    dt = T(1) - T(0)
    ReDim theta(0 To M) As Double
    ReDim r(0 To M, 1 To N) As Double

    For j = 1 To N
        r(0, j) = r0
        For i = 0 To M - 1
            theta(i) = Df(i) + a * f(i) + sigma ^ 2 / 2 / a * (1 - Exp(-2 * a * i * dt))
            r(i + 1, j) = r(i, j) + (theta(i) - a * r(i, j)) * dt + sigma * Sqr(dt) * rGauss()
            If r(i + 1, j) < 0 Then r(i + 1, j) = -r(i + 1, j) 'Convention for r < 0, take r = -r
        Next i
    Next j
    
    HW_path_2 = r
End Function

Function Intr(r As Variant, t0 As Double, t1 As Double, dt As Double) As Variant
    Dim i As Long, j As Long, N As Long
    N = UBound(r, 2)
    ReDim Sum(1 To N)
    For j = 1 To N
    Sum(j) = 0
        For i = Int(t0 / dt) To Int(t1 / dt) - 1
            Sum(j) = Sum(j) + r(i, j) * dt
        Next i
    Next j
    Intr = Sum
End Function

Function Bond(Intr As Variant) As Variant
    Dim N As Long, i As Long
    N = UBound(Intr)
    ReDim Val(1 To N) As Double
    For i = 1 To N
        Val(i) = Exp(-Intr(i))
    Next i
    Bond = Val
End Function
```
```


Sub CommandButton1_Click()

  Calculate
  
End Sub
Private Sub Click()

Dim d1#, d2#, Nd1#, Nd2#
Dim rf#, Std#, T#
Dim ca#, im#, dd#
Dim r%, sigma#, ls#, S%, K%
Dim fre#, div#, kk#, D#
Dim Ndd1#, Ndd2#


End Sub

Function BSECall(K, r, sigma, T, S) 'call price計算function

d1 = (Log(S / K) + (r + 0.5 * sigma ^ 2) * T) / (sigma * Sqr(T))
d2 = d1 - sigma * Sqr(T)
Nd1 = Application.NormSDist(d1)
Nd2 = Application.NormSDist(d2)
BSECall = S * Nd1 - Exp(-r * T) * K * Nd2


End Function


Function BSEPut(K, r, sigma, T, S) 'put price計算function

d1 = (Log(S / K) + (r + 0.5 * sigma ^ 2) * T) / (sigma * Sqr(T))
d2 = d1 - sigma * Sqr(T)
Ndd1 = Application.NormSDist(-d1)
Ndd2 = Application.NormSDist(-d2)
BSEPut = (Exp(-r * T) * K * Ndd2) - S * Ndd1


End Function






Private Function mc(S, u, sg, T, N) '股價計算function

cum = 0
For i = 1 To N
random = Application.WorksheetFunction.NormInv(Rnd(), 0, 1)

st = S * Exp((u - sg ^ 2 / 2) * T + sg * Sqr(T) * random)

cum = cum + st
Next
mc = CDbl(CDbl(cum) / CDbl(N))

End Function
```
