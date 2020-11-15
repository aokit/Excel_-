Attribute VB_Name = "x2p_Lib1"
' -*- coding:shift_jis-dos -*-

'./x2p_Lib1.bas
' ＜注意＞
' 　このスクリプトでは関数定義しているので標準モジュールに置かないと機能しない。
' 　そのためファイルの先頭に　Attribute VB_Name = "x2p_Lib1"　と記述しておく。
' 　（外部ファイルとして編集しているときのみ見える）

Function セルの固定色(セル)
     Dim a
     '// a = セル.Interior.ColorIndex
     '// a = セル.FormatConditions.Interior.Color ' - これはうまく行かない
     '// 　　条件付き背景色を得たいとしても
     '// 　　ユーザ関数として実行するとエラーが起きるので使えない
     '// a = セル.DisplayFormat.Interior.Color

     a = セル.Interior.ColorIndex
     セルの固定色 = a
End Function

Public Sub 開始時抑制()
    Application.ScreenUpdating = False '画面描画を停止
    Application.Cursor = xlWait 'ウエイトカーソル
    Application.EnableEvents = False 'イベントを抑止
    Application.DisplayAlerts = False '確認メッセージを抑止
    Application.Calculation = xlCalculationManual '計算を手動に
End Sub

Public Sub 終了時解放()
    Application.StatusBar = False 'ステータスバーを消す
    Application.Calculation = xlCalculationAutomatic '計算を自動に
    Application.DisplayAlerts = True '確認メッセージを開始
    Application.EnableEvents = True 'イベントを開始
    Application.Cursor = xlDefault '標準カーソル
    Application.ScreenUpdating = True '画面描画を開始
End Sub

Function 列の最終行(n As String, _
                          Optional k As Long = 1, _
                          Optional q As Long = 0) As Long
   ' n - 開始するセル（範囲でもよい）に名付けた名前（文字列）
   ' k - ＜オプション＞その範囲の中の列番号
   ' q - ＜オプション＞何回目の空白を終わりとみなすか：0だと無制限
   Dim R_n As Range
   Set R_n = ThisWorkbook.Names(n).RefersToRange
   If IsMissing(k) Then
      If IsMissing(q) Then
         列の最終行 = 列の最終行_range(R_n)
      Else
         列の最終行 = 列の最終行_range(R_n, , q)
      End If
   Else
      If IsMissing(q) Then
         列の最終行 = 列の最終行_range(R_n, k)
      Else
         列の最終行 = 列の最終行_range(R_n, k, q)
      End If
   End If
End Function

Function 行の最終列(n As String, _
                          Optional k As Long = 1, _
                          Optional q As Long = 0) As Long
   ' n - 開始するセル（範囲でもよい）に名付けた名前（文字列）
   ' k - ＜オプション＞その範囲の中の行番号
   ' q - ＜オプション＞何回目の空白を終わりとみなすか：0だと無制限
   Dim R_n As Range
   Set R_n = ThisWorkbook.Names(n).RefersToRange
   If IsMissing(k) Then
      If IsMissing(q) Then
         行の最終列 = 行の最終列_range(R_n)
      Else
         行の最終列 = 行の最終列_range(R_n, , q)
      End If
   Else
      If IsMissing(q) Then
         行の最終列 = 行の最終列_range(R_n, k)
      Else
         行の最終列 = 行の最終列_range(R_n, k, q)
      End If
   End If
End Function

Function 列の最終行_range(ByRef R_n As Range, _
                          Optional k As Long = 1, _
                          Optional q As Long = 0) As Long
   ' R_n - 開始するセルを含む範囲-Range-
   ' k - ＜オプション＞その範囲の中の列番号
   ' q - ＜オプション＞何回目の空白を終わりとみなすか：0だと無制限
   Dim R1 As Long
   Dim r2 As Long
   Dim mr As Long
   Dim s As Variant
   mr = Rows.count ' 行の最大値・・・ここで飽和する。
   Set s = R_n.Columns(k)
   r2 = s.Row
   q = q - 1
   Do
      R1 = r2
      Set s = s.End(xlDown)
      r2 = s.Row
      q = q - 1
   Loop While Not ((r2 >= mr) Or (q = 0))
   列の最終行_range = R1
End Function

Function range_列の最終行_range(ByRef R_n As Range, _
                                Optional k As Long = 1, _
                                Optional q As Long = 0) As Range
   ' R_n - 開始するセルを含む範囲-Range-
   ' k - ＜オプション＞その範囲の中の列番号
   ' q - ＜オプション＞何回目の空白を終わりとみなすか：0だと無制限
   Dim r0 As Long
   Dim r1 As Long
   Dim r2 As Long
   Dim mr As Long
   Dim s As Variant
   r0 = R_n.Row
   mr = Rows.count ' 行の最大値・・・ここで飽和する。
   Set s = R_n.Columns(k)
   r2 = s.Row
   q = q - 1
   Do
      r1 = r2
      Set s = s.End(xlDown)
      r2 = s.Row
      q = q - 1
   Loop While Not ((r2 >= mr) Or (q = 0))
   ' 列の最終行_range = R1
   Set range_列の最終行_range = R_n.resize((r1 - r0 + 1))
   '   　　　　　　　　　　　　　　列は省略して行のみ拡張┛
   '┗戻り値が範囲つまり『オブジェクト』なので Set を使う！！
End Function

Function 行の最終列_range(R_n As Range, _
                          Optional k As Long = 1, _
                          Optional q As Long = 0) As Long
   ' R_n - 開始するセルを含む範囲-Range-
   ' k - ＜オプション＞その範囲の中の行番号
   ' q - ＜オプション＞何回目の空白を終わりとみなすか：0だと無制限
   Dim c1 As Long
   Dim c2 As Long
   Dim mc As Long
   Dim s As Variant
   mc = Columns.count ' 列の最大値・・・ここで飽和する。
   Set s = R_n.Rows(k)
   c1 = 0
   c2 = s.Column
   Do
      c1 = c2
      Set s = s.End(xlToRight)
      c2 = s.Column
      q = q - 1
   Loop While Not ((c2 >= mc) Or (q = 0))
   行の最終列_range = c1
End Function

Function 複数行の最終列_range(R_n As Range, _
                              Optional q As Long = 0) As Long
   ' R_n - 開始するセルを含む範囲-Range-
   ' q - ＜オプション＞何回目の空白を終わりとみなすか：0だと無制限
   Dim c1 As Long
   Dim c2 As Long
   Dim mc As Long
   Dim s As Range ' yet Variant
   mc = Columns.count ' 列の最大値・・・ここで飽和する。
   Dim k As Long
   Dim cx As Long
   cx = 0
   For k = 1 To R_n.Rows.count
      Set s = R_n.Rows(k)
      c1 = 0
      ' c2 = 0
      c2 = s.Column
      ' 初期値はここで設定しておかないといけなそう。
      Do
         c1 = c2
         Set s = s.End(xlToRight)
         c2 = s.Column
         q = q - 1
      Loop While Not ((c2 >= mc) Or (q = 0))
      If cx < c1 Then cx = c1
   Next k
   複数行の最終列_range = cx
End Function

Function range_連続列最大行_range(R_n As Range, _
                                    Optional q As Long = 0) As Range
   '
   ' 複数列の最終行を使って、範囲をひろげる
   ' ・（先頭列の行数）×（最長行の列数）を範囲とする
   ' R_n - 開始するセルを含む範囲-Range-
   ' q - ＜オプション＞何回目の空白を終わりとみなすか：0だと無制限
   Dim r0 As Long
   Dim c0 As Long
   r0 = R_n.Row
   c0 = R_n.Column
   Dim nr As Long
   Dim nC As Long
   nr = R_n.Rows.count
   nC = R_n.Columns.count
   If (nr = 1) And (nC = 1) Then
      ' Call ExpandRangeCont(R_n, strName, cx)
      Dim rz As Long
      rz = 列の最終行_range(R_n)
      Set R_n = R_n.Resize((rz - r0 + 1), 1)
      Dim cz As Long
      cz = 複数行の最終列_range(R_n)
      ' cz が 0 になってしまうのはなぜ。
      Set R_n = R_n.Resize((rz - r0 + 1), (cz - c0 + 1))
   End If
   Set range_連続列最大行_range = R_n
End Function

