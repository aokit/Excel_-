Attribute VB_Name = "x2p_tabjump"
Option Explicit
' -*- coding:shift_jis-dos -*-

'./x2p_tabjump.bas

Function range_TabWiden_range(R_n As Range, _
                              Optional k As Long = 1) As Range
   '
   ' range_TabBottom_range を改造。指定した範囲 R_n の 列 k の
   ' 空セルでない行　まで　範囲 R_n を拡張して返す。
   ' 関数が返した範囲に再度関数を適用しても範囲は不変となること
   ' は定義から明らかなので、下方向の次の範囲を得るためには、
   ' まず、返された範囲のひとつ下の行の　列 k で次の
   ' 空セルでないセル を探すことが必要となる。これは
   ' range_n_TabWiden_range （range_n_TabBottom_range の別名）
   ' で記述してある。
   ' そのため、 range_TabBottom_range で参照する引数 q は不要
   ' となる。
   ' 列 k の値が R_n の列数（R_n.Rows.Count）を超えていても、
   ' 有効である。その列に対象として処理して、作用して、行を得て
   ' 列については R_n の範囲を返す。
   ' 引数 R_n の 列 k の下方向に空セルしかない場合は、範囲が
   ' 存在しないものとして、引数で指定した範囲をそのまま返す。
   '
   Dim r0 As Long
   Dim r1 As Long
   On Error GoTo RowError
   If (R_n.Cells(1,k).Value <> "") And _
      (R_n.Cells(2,k).Value = "") Then
      Set range_TabWiden_range = R_n
   Else
      r0 = R_n.Cells(1,k).Row
      ' r1 = R_n.Cells(1,k).End(xlDown).Row
      ' ┗１行だけではなくて
      ' ┏複数行の範囲が指定されたときは、指定された範囲の
      ' ┃最後の行から .End(xlDown) をする。これで、
      ' ┃指定された範囲内に 空欄 があっても指定された範囲
      ' ┃をもとに範囲を返すようになる。
      r1 = R_n.Cells(R_n.Rows.Count,k).End(xlDown).Row
      If r1 = Rows.Count Then
         ' ↓・・・空セルしかなかった
         Set range_TabWiden_range = R_n
      Else
         Set range_TabWiden_range = R_n.Resize((r1 - r0 + 1))
      End If
   End If
   Exit Function
RowError:
   Debug.Print("range_TabWiden_rangeでエラー：指定した範囲が最下行に達しているなど")
   Set range_TabWiden_range = R_n
End Function

Function range_n_TabWiden_range(R_n As Range, _
                                Optional k As Long = 1, _
                                Optional n As Long = 1) As Range
   '
   ' 指定した範囲 R_n の k 列目から下の方向に値のあるセルが
   ' 連続している範囲を R_n から数えて、n 個めの範囲を返す。
   ' k と n を記載しない場合はいずれも 1 とする。
   ' 『列の最大行』に相当する。
   ' range_TabWiden_range を n回『繰り返して』呼ぶ
   ' ・・・・・・・・・・繰り返すために・・・・・・・・・・
   ' １回めは、 range_TabWiden_range を実行して範囲を返す。
   ' 　範囲の次（下）のセル（＝空セル）の範囲を Rt で２回め
   ' 　へ引き継ぐ。
   ' ２回め以降は、Rt を非空きセルの先頭へ移動し、
   ' 　その後、 range_TabWiden_range を実行して範囲を返す。
   ' 　範囲の次（下）のセル（＝空セル）の範囲を Rt で次回へ
   ' 　引き継ぐ。
   '
   Dim q As Long
   Dim Rt As Range
   Dim Rs As Range
   Set Rt = R_n
   Set range_n_TabWiden_range = R_n
   On Error GoTo EndOfRange
   For q = 1 To n
      If q > 1 Then
         Set Rt = Rt.Cells(1,k).End(xlDown).Offset(0,-(k-1)).Resize(1,Rt.Columns.Count)
         ' Set Rt = Rt.Cells(1,k).End(xlDown).Offset(0,-(k-1)).Resize(1,R_n.Columns.Count)
      End If
      Set Rs = range_n_TabWiden_range
      Set Rt = range_TabWiden_range(Rt, k)
      If (Rt.Row + Rt.Rows.Count - 1) = Rows.Count Then
         Set range_n_TabWiden_range = Rs
         Debug.Print("範囲が指定された数に足りませんでした。")
         Exit Function
      Else
         Set range_n_TabWiden_range = Rt
         Set Rt = Rt.Offset(1,0).Cells(Rt.Rows.Count,1).Resize(1,Rt.Columns.Count)
         ' Set Rt = Rt.Offset(1,0).Cells(Rt.Rows.Count,1).Resize(1,R_n.Columns.Count)
      End If
   Next q
   On Error GoTo 0
   Exit Function
   '
EndOFRange:
   range_n_TabWiden_range = R_n
   Debug.Print("範囲が指定された数に足りませんでした。")
End Function

' ========================================================================================
' ========================================================================================
' ========================================================================================

Function range_TabBottom_range(R_n As Range, _
                               Optional k As Long = 1, _
                               Optional q As Long = 1) As Range
   '
   ' R_n で与えられた範囲の 第１行の k 列目（デフォルトは１列目）
   ' から下方向に値があるセルをたどって一番下のセルまで含むように
   ' R_n の列数はそのまま、行数だけ拡大して返す。
   ' ※ R_n で与えられた範囲の 第１行の k 列目の下がすでに 空セル
   '    だったら、拡大しない。
   '
   ' R_n が１行のみで、k 列の上のセルも下のセルも空セルの場合、
   ' （２行以上あれば、２回の実行で次の範囲に移動するのだが）
   ' １回の実行で次の範囲に移動してしまう。それを抑制するために
   ' 複数回の呼び出しは、回数を指定する別の関数によって記述して、
   ' この関数自体には、何回目の呼び出しであるかを伝えるようにする。
   ' そのために 引数 q を使う。デフォルト値は 1 である。
   ' 続けて実行するときには、連続何回目であるかを q で与える。
   ' （連続３回実行するときは、q が 3,2,1 と変化する）
   '
   Dim r0 As Long
   Dim r1 As Long
   On Error GoTo RowError
   If (R_n.Cells(1,k).Value <> "") And _
      (R_n.Cells(2,k).Value = "") And _
      ((q mod 2) = 1) Then
      Set range_TabBottom_range = R_n
   Else
      r0 = R_n.Cells(1,k).Row
      r1 = R_n.Cells(1,k).End(xlDown).Row
      If r1 = Rows.Count Then
         Set range_TabBottom_range = R_n
         ' ? range_TabBottom_range(Range("A4:B4"),2).Address
         ' $A$4:$B$1048576
         ' → $A4:$B4
         ' 範囲が存在しないときは、引数で指定した範囲を返す。
      Else
         Set range_TabBottom_range = R_n.Resize((r1 - r0 + 1))
      End If
   End If
   Exit Function
RowError:
   Set range_TabBottom_range = R_n
End Function

' ? range_n_TabBottom_range(Range("D4")).Address
' ? range_n_TabBottom_range(Range("D4"),,5).Address
' ? range_n_TabBottom_range(Range("C4"),2,5).Address
' ? range_n_TabBottom_range(Range("B4"),3,3).Address
' ? range_n_TabBottom_range(Range("B4:D4"),3,3).Address
' $B$30
' → $B$30:$D$30
'
' ? range_TabBottom_range(Range("C4"), 2, 1).Address
' ? range_TabBottom_range(Range("C4"), 2, 1).Offset(1,1).Address
' ? range_TabBottom_range(Range("C4"), 2, 1).Offset(1,0+1).Cells(21,1).Address
' 0+1 = k
' 負の値も使いたいので Offset を充てる。
' 21 = range_TabBottom_range(Range("C4"), 2, 1).Rows.Count
' ? range_TabBottom_range(Range("C4"), 2, 1).Rows.Count
'
Function range_n_TabBottom_range(R_n As Range, _
                                 Optional k As Long = 1, _
                                 Optional n As Long = 1) As Range
   '
   ' 指定した範囲 R_n の k 列目から下の方向に値のあるセルが
   ' 連続している範囲を R_n から数えて、n 個めの範囲を返す。
   ' k と n を記載しない場合はいずれも 1 とする。
   ' 『列の最大行』に相当する。
   ' range_TabBottom_range を n回『繰り返して』呼ぶ
   ' ・・・・・・・・・・繰り返すために・・・・・・・・・・
   ' １回めは、 range_TabBottom_range を実行して範囲を返す。
   ' 　範囲の次（下）のセル（＝空セル）の範囲を Rt で２回め
   ' 　へ引き継ぐ。
   ' ２回め以降は、Rt を非空きセルの先頭へ移動し、
   ' 　その後、 range_TabBottom_range を実行して範囲を返す。
   ' 　範囲の次（下）のセル（＝空セル）の範囲を Rt で次回へ
   ' 　引き継ぐ。
   '
   Dim q As Long
   Dim Rt As Range
   Dim Rs As Range
   Set Rt = R_n
   Set range_n_TabBottom_range = R_n
   On Error GoTo EndOfRange
   For q = 1 To n
      ' If q > 1 Then Set Rt = Rt.Cells(1,k).End(xlDown)
      If q > 1 Then
         ' Set Rt = Rt.Cells(1,k).End(xlDown).Offset(0,-(k-1))
         Set Rt = Rt.Cells(1,k).End(xlDown).Offset(0,-(k-1)).Resize(1,Rt.Columns.Count)
         ' Set Rt = Rt.Cells(1,k).End(xlDown).Offset(0,-(k-1)).Resize(1,R_n.Columns.Count)
         ' Stop
      End If
      Set Rs = range_n_TabBottom_range
      ' Set Rt = range_TabBottom_range(Rt, k, q)
      Set Rt = range_TabBottom_range(Rt, k)
      ' Stop
      If (Rt.Row + Rt.Rows.Count - 1) = Rows.Count Then
         Set range_n_TabBottom_range = Rs
         Debug.Print("範囲が指定された数に足りませんでした。")
         Exit Function
      Else
         Set range_n_TabBottom_range = Rt
         ' Set Rt = Rt.Cells(Rt.Rows.Count + 1, k)
         ' Set Rt = Rt.Cells(Rt.Rows.Count + 1, 1).Resize(1, Rt.Columns.Count)
         ' Set Rt = Rt.Cells(Rt.Rows.Count + 1, k).Offset(0, -(k-1)).Resize(1, Rt.Columns.Count)
         ' Set Rt = Rt.Offset(1,k).Cells(Rt.Rows.Count,1)
         Set Rt = Rt.Offset(1,0).Cells(Rt.Rows.Count,1).Resize(1,Rt.Columns.Count)
         ' Set Rt = Rt.Offset(1,0).Cells(Rt.Rows.Count,1).Resize(1,R_n.Columns.Count)
         ' Stop
         '
' ? range_TabBottom_range(Range("C4"), 2, 1).Offset(1,0+1).Cells(21,1).Address
' range_TabBottom_range(Range("C4"), 2, 1) = Rt
' 0+1 = k
' 負の値も使いたいので Offset を充てる。
' 21 = range_TabBottom_range(Range("C4"), 2, 1).Rows.Count
' ? range_TabBottom_range(Range("C4"), 2, 1).Rows.Count

         '
      End If
   Next q
   On Error GoTo 0
   Exit Function
   '
EndOFRange:
   ' range_n_TabBottom_range = Rs
   range_n_TabBottom_range = R_n
   Debug.Print("範囲が指定された数に足りませんでした。")
End Function
   
' k が機能していないような気がする。範囲を超えて k を機能させることはできるか？
' q で奇遇を分ける必要がないかもしれないがどうか。
' ? Range("C27").Resize(1,2).Cells(1,2).End(xlDown).Offset(0,-1).Resize(1,2).Address
' $C$28:$D$28

' ? Range("C27").Cells(1,2).End(xlDown).Offset(0,-1).Resize(1,2).Address