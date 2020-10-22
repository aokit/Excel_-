' -*- coding:shift_jis -*-

'./x2p》ダッシュボード.bas

' 範囲に対して　.End(xlDown)　であらたな範囲を得る。これは
' 内容があるセルの連続の最後の行のセルである。
' 繰り返し適用したときに、行が Rows.Count（行の仕様最大値）
' になったとき、その適用の前の行を返す。

Function カラムの最終行(n As String, k) As Long
   ' n - 範囲与えた名前（文字列）
   ' k - その範囲の中のカラム番号
   Dim r1 As Long
   Dim r2 As Long
   Dim mr As Long
   Dim s As Variant
   mr = Rows.Count
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

