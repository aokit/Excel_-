' -*- coding:shift_jis -*-

'./x2p》ダッシュボード.bas

Sub 名前の定義確認の生成()
   Call 開始時抑制
   Range("B6").Value = "=isref(" & Range("A6").Value & ")"
   Range("B7").Value = "=isref(" & Range("A7").Value & ")"
   Range("B8").Value = "=isref(" & Range("A8").Value & ")"
   Range("B9").Value = "=isref(" & Range("A9").Value & ")"
   Range("B10").Value = "=isref(" & Range("A10").Value & ")"
   Range("B11").Value = "=isref(" & Range("A11").Value & ")"
   Range("B12").Value = "=isref(" & Range("A12").Value & ")"
   Range("B13").Value = "=isref(" & Range("A13").Value & ")"
   Range("B14").Value = "=isref(" & Range("A14").Value & ")"
   Range("B15").Value = "=isref(" & Range("A15").Value & ")"
   Call 集計状況の各範囲の名前定義
   Call 終了時解放
End Sub

Private Sub 集計状況の各範囲の名前定義()
   Call newName2Range(Range("B18"), Range("A18").Value)
   Call newName2Range(Range("B19"), Range("A19").Value)
   Call newName2Range(Range("B20"), Range("A20").Value)
   Call newName2Range(Range("B21"), Range("A21").Value)
   Call newName2Range(Range("B22"), Range("A22").Value)
   Call newName2Range(Range("B23"), Range("A23").Value)
End Sub

' 名前をつけた範囲にあらたな名前をつける
' （まだ名前をつけていなければはじめて名前をつける）
Private Sub newName2Range(rng As Range, strName As String)
   Dim nm As Name
   For Each nm In Names
      If rng.Address = nm.RefersToRange.Address Then
         nm.Delete
      End If
   Next
   rng.Parent.Parent.Names.Add _
      Name := strName, _
      RefersToLocal := "=" & rng.Address(External := True)
End Sub

Sub 集計名_組織_初期化()
   Call 組織略称初期化
End Sub

Private Sub 組織略称初期化()
   Dim cc() As Long
   Dim cs() As String
   Dim 組織略称 As Variant
   Dim 組織 As Range
   Set 組織 = ThisWorkbook.Names("組織").RefersToRange
   ' 名前付きの範囲（Named Range）から配列へ：.RefersToRange
   組織略称 = 組織
   Dim n As Long
   Dim m As Long
   m = UBound(組織略称)
   Debug.Print m
   Debug.Print LBound(組織略称)
   n = 0
   ' ReDim cc(n)
   ' ReDim cs(n)
   Dim i As Long, j As Long, k As Long
   i = 0
   j = m ' 最小値が m ということは無いはずなのでシードにする。
   k = 0
   For Each c In 組織
      ' Debug.Print c.Row + 0 & ":" & セルの固定色(c) & ":" & c.Value
      ' Debug.Print n > c.Row
      i = c.Row
      If i < j Then j = i
      If i > k Then k = i
      If n < i Then
         n = n + m
         ReDim cc(n)
         ReDim cs(n)
      End If
      cc(i) = セルの固定色(c)
      cs(i) = c.Value
   Next c
   Debug.Print j
   Debug.Print k
End Sub

Function カラムの最終行(n As String, k) As Long
   ' n - 範囲に与えた名前（文字列）
   ' k - その範囲の中のカラム番号
   Dim r1 As Long
   Dim r2 As Long
   Dim mr As Long
   Dim s As Variant
   mr = Rows.count
   Set s = Range(n).Columns(k).End(xlDown)
   r1 = s.Row
   If r1 = mr Then
      カラムの最終行 = 0
      Exit Function
   End If
   Do While Not (r1 = mr)
      ' Debug.Print s.Value
      r2 = r1
      Set s = s.End(xlDown)
      r1 = s.Row
   Loop
   カラムの最終行 = r2
End Function

