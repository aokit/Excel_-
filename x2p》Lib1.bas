' -*- coding:shift_jis-dos -*-

'./x2p》Lib1

Function セルの固定色(セル)
     '// セルの色 = セル.Interior.ColorIndex
     '// セルの色 = セル.FormatConditions.Interior.Color
     '// セルの色 = セル.DisplayFormat.Interior.Color
     Dim a
     Dim b
     Dim c
     a = セル.Interior.ColorIndex
     '// b = セル.DisplayFormat.Interior.Color
     '// 条件付き背景色を得たいとしても
     '// ユーザ関数として実行するとここでエラーが起きるので使えない
     If (a = a) Then
        セルの固定色 = a
     Else
        セルの■■色 = a & "変色"
     End If
End Function
