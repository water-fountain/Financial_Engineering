本金平均攤還試算
========
   以下內容皆本人(倪采靖)所製作<br>
      維護修課倫理請勿抄襲<br>
      1版上傳於2020.3.15.00:32<br>

執行說明:
-------

點擊download下載 作業1.xlsm 檔<br>
-----
![下載](https://i.imgur.com/CMth6df.jpg)

開啟作業1檔案後 點擊 啟用編輯 <br>
-----


![啟用編輯](https://i.imgur.com/m1tBU2k.jpg)

點擊 啟用巨集 <br>
-----


![啟用巨集](https://i.imgur.com/5VIhVCL.jpg)

表單會自然跳出 依序輸入表單，按下開始按鈕 <br>
-----

![輸入表單](https://i.imgur.com/JMDmk5p.jpg)


表單不會自然關閉，如需關閉表單請手動按叉 <br>
-----

![按叉](https://i.imgur.com/00UgeAY.jpg)

表單關閉後可見文件全貌，也就是計算結果 <br>
-----

![全貌](https://i.imgur.com/pTGlZ4w.jpg)

如果需要計算其他資料，按下文件上的計算按鈕 <br>
-----
![按鈕](https://i.imgur.com/Qm9OcgN.jpg)


按下按鈕後表單會跳出 依序輸入表單，按下開始按鈕可開始計算 <br>
-----

![按後](https://i.imgur.com/SNkiesw.jpg)



程式碼公開
``` 
Private Sub CommandButton1_Click()
Dim rf#, d4#, d3#
Dim d1#, d2#, 年利率#
Dim 本金%, 期數年%, 期數月%
With sheet1
   Range("A1:G10000").Clear
    Cells(2, "A") = TextBox1.Text
    Cells(2, "B") = TextBox2.Text
    Cells(2, "C") = TextBox3.Text / 100
  
本金 = Cells(2, 1)
期數年 = Cells(2, 2)
年利率 = Cells(2, 3)
Range("D2").Select
ActiveCell.FormulaR1C1 = "=RC[-2]*12"
期數月 = Cells(2, 4)

Cells(1, 1) = "本金(萬元)"
Cells(1, 2) = "期數(年)"
Cells(1, 3) = "利率(年)"
Cells(1, 4) = "期數(月)"
Cells(1, 5) = "平均每月攤還本金"
Cells(1, 6) = "平均每月攤還利息"
Cells(1, 7) = "全部利息"
Cells(5, 1) = "期數(月)"
Cells(5, 2) = "本金(元)"
Cells(5, 3) = "利息(元)"
Cells(2, 6) = "請參考下表"
Cells(5, 4) = "本金利息累計(元)"


Range("C2").Select
Selection.NumberFormatLocal = "0.00%"


For i = 1 To 期數月

d1 = (本金 / 期數月) * 10000
d2 = CCur(IPmt(CCur(年利率) / 12, i, CCur(期數月), CCur(本金) * 10000))

Cells(5 + i, 1) = i
Cells(2, 5) = d1
Cells(5 + i, 2) = d1
Cells(5 + i, 3) = -d2
Cells(5 + i, 4).Value = Application.WorksheetFunction.Sum(Range(Cells(6, 3), Cells(6 + i, 3))) + d1 * i
Cells(2, 7).Value = Application.WorksheetFunction.Sum(Range(Cells(6, 3), Cells(期數月 + 6, 3)))


Next
End With
End Sub
```
``` 
Private Sub Workbook_Open()

UserForm1.Show


End Sub
``` 
``` 
Private Sub 按鈕1_Click()
UserForm1.Show

End Sub
``` 
1版:已公開程式碼，完成詳細說明之編輯<br>

發想過程圖:
![](https://i.imgur.com/5NndJwk.jpg)
