Attribute VB_Name = "x2p_Lib1"
Option Explicit

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
   列の最終行 = 列の最終行_range(R_n, k, q)
End Function

Function a_列の最終行(n As String, _
                          Optional k As Long = 1, _
                          Optional ByVal q As Long = 0) As Long
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
   '
   Dim R_n As Range
   Set R_n = ThisWorkbook.Names(n).RefersToRange
   行の最終列 = 行の最終列_range(R_n, k, q)
End Function

Function 列の最終行_range(ByRef R_n As Range, _
                          Optional k As Long = 1, _
                          Optional g As Long = 0) As Long
   ' R_n - 開始するセルを含む範囲-Range-
   ' k - ＜オプション＞その範囲の中の列番号
   ' q - ＜オプション＞何回目の空白を終わりとみなすか：0だと無制限
   ' q=1 のとき：
   ' ・下のセルが空白でないとき、値のあるセルの連続の最後のセルに移動
   ' 　（下のセルが空白のセルになるセルに移動）
   ' ・下のセルが空白のとき、値のあるセルの連続の最初のセルに移動
   ' 　（上のセルが空白のセルになるセルに移動）
   ' 　／または、スプレッドシートの論理上の最大行のセルに移動
   Dim q As Long
   q = g
   ' ---
   Dim r1 As Long
   Dim r2 As Long
   Dim mr As Long
   Dim s As Range
   Dim w As Long
   w = 0
   mr = Rows.count ' 行の最大値・・・ここで飽和する。
   Set s = R_n.Columns(k)
   r2 = s.Row
   If (mr = r2) Or (mr = s.Rows.Count) Then
      Debug.Print("列の最終行_rangeに与えられた範囲がシートの最下行に達しています")
      列の最終行_range = mr
      Exit Function
   End If
   '
   Do
      r1 = r2
      ' ---
      ' Set s = s.End(xlDown)
      ' ＜↓変更↓＞
      ' s が複数セル／行の範囲のこともあるので：
      If "" = s.Cells((s.Rows.Count + 1), 1).Value Then
         If w > 0 Then
            w = 0
            Set s = s.End(xlDown)
         Else
            w = 1
         End If
      Else
         w = 0
         Set s = s.End(xlDown)
      End If
      ' ---
      r2 = s.Row
      q = q - 1
      ' ......... ((r2 >= mr) Or (q = 0) Or (r1 = r2))
      ' ＜↓変更↓＞
   Loop While Not ((r2 >= mr) Or (q = 0))
   列の最終行_range = r2
End Function

Function range_列の最終行_range(ByRef R_n As Range, _
                          Optional k As Long = 1, _
                          Optional q As Long = 0) As Range
   ' R_n - 開始するセルを含む範囲-Range-
   ' k - ＜オプション＞その範囲の中の列番号
   ' q - ＜オプション＞何回目の空白を終わりとみなすか：0だと無制限
   Dim r2 As Long
   Dim s As Range
   Set s = R_n.Columns(k)
   r2 = s.Row
   ' Set range_列の最終行_range = s.Offset((列の最終行_range(R_n, k, q) - r2), 0)
   Set range_列の最終行_range = s.Resize((列の最終行_range(R_n, k, q) - r2 + 1), 1)
   '
End Function

Function range_列の最終行_namedrange(strName As String, _
                                     Optional k As Long = 1, _
                                     Optional q As Long = 0) As Range
   Dim R_n As range
   Set R_n = ThisWorkbook.Names(strName).RefersToRange
   Set range_列の最終行_namedrange = _
        range_列の最終行_range(R_n, k, q)
   '
End Function

Function a_range_列の最終行_range(ByRef R_n As Range, _
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

Function 変更前_行の最終列_range(R_n As Range, _
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

Function 行の最終列_range(R_n As Range, _
                          Optional k As Long = 1, _
                          Optional g As Long = 0) As Long
   ' R_n - 開始するセルを含む範囲-Range-
   ' k - ＜オプション＞その範囲の中の行番号
   ' q - ＜オプション＞何回目の空白を終わりとみなすか
   ' ：指定されなければ 1 指定が 0 だと無制限。←これはまずいかも。
   ' もどした。
   '
   ' 引数であるが、ByVal としておかないと、
   ' 呼び出したあと、q の値が変わってしまうので、繰り返しのなかで
   ' q を指定して呼び出すと、予期せぬ結果になる。
   ' 関数なのに、引数についてさえ、デフォルトが ByRef という・・・
   '
   Dim q As Long
   q = g
   ' ---
   Dim c1 As Long
   Dim c2 As Long
   Dim mc As Long
   Dim s As Range
   Dim w As long
   w = 0
   mc = Columns.count ' 列の最大値・・・ここで飽和する。
   Set s = R_n.Rows(k)
   c2 = s.Column
   If (mc = c2) Or (mc = s.Columns.Count) Then
      Debug.Print("行の最終列_rangeに与えられた範囲がシートの最右列に達しています")
      行の最終列_range = mc
      Exit Function
   End If
   '
   Do
      c1 = c2
      ' ---
      ' Set s = s.End(xlToRight)
      ' ＜↓変更↓＞
      ' s が複数セル／列からなる範囲であることもあるので：
      If "" = s.Cells(1, (s.Columns.Count + 1)).Value Then
         If w > 0 Then
            w = 0
            Set s = s.End(xlToRight)
         Else
            w =1
         End If
      Else
         w = 0
         Set s = s.End(xlToRight)
      End If
      ' ---
      c2 = s.Column
      q = q - 1
      ' ......... ((c2 >= mc) Or (q = 0) Or (c1 = c2))
   ' ＜↓変更↓＞
   Loop While Not ((c2 >= mc) Or (q = 0))
   行の最終列_range = c2
End Function

Function range_行の最終列_range(ByRef R_n As Range, _
                                Optional k As Long = 1, _
                                Optional q As Long = 0) As Range
   ' R_n - 開始するセルを含む範囲-Range-
   ' k - ＜オプション＞その範囲の中の行番号
   ' q - ＜オプション＞何回目の空白を終わりとみなすか：0だと無制限
   Dim c2 As Long
   Dim s As Range
   Set s = R_n.Rows(k)
   c2 = s.Column
   ' Set range_行の最終列_range = s.Offset(0, (行の最終列_range(R_n, k, q) - c2))
   Set range_行の最終列_range = s.Resize(1, (行の最終列_range(R_n, k, q) - c2 + 1))
   '
End Function

Function range_行の最終列_namedrange(strName As String, _
                                     Optional k As Long = 1, _
                                     Optional q As Long = 0) As Range
   Dim R_n As range
   Set R_n = ThisWorkbook.Names(strName).RefersToRange
   Set range_行の最終列_namedrange = _
        range_行の最終列_range(R_n, k, q)
   '
End Function

Function a_複数行の最終列_range(R_n As Range, _
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
      qi = q
      Set s = R_n.Rows(k)
      ' c1 = 0
      ' c2 = 0
      c2 = s.Column
      ' 初期値はここで設定しておかないといけなそう。
      Do
         c1 = c2
         Set s = s.End(xlToRight)
         c2 = s.Column
         qi = qi - 1
      Loop While Not ((c2 >= mc) Or (qi = 0))
      ' If cx < c1 Then cx = c1
      If cx < c2 Then cx = c2
   Next k
   複数行の最終列_range = cx
End Function

Function 複数行の最終列_range(R_n As Range, _
                              Optional q As Long = 0) As Long
   ' R_n - 開始するセルを含む範囲-Range-
   ' q - ＜オプション＞何回目の空白を終わりとみなすか：0だと無制限
   '
   Dim cx As Long
   cx = 0
   Dim c2 As Long
   c2 = 0
   Dim r2 As Long
   r2 = R_n.Row
   ' R_n.Rows.Count > 1 であることもあるので：
   If R_n.Rows.Count > 1 then
      Set R_n = R_n.Resize((列の最終行_range(R_n.Cells(1, 1), 1, 1) - r2), 1)
   End If
   ' ┗R_n.Rows.Count = 1 なら R_nはそのまま。
   '
   Dim k As Long
   For k = 1 To R_n.Rows.Count
      c2 = 行の最終列_range(R_n, k, q)
      If cx < c2 Then cx = c2
   Next k
   '
   複数行の最終列_range = cx
End Function

Function range_連続列最大行_range(R_n As Range, _
                                  Optional q As Long = 1) As Range
   '
   ' 複数列の最終行を使って、範囲をひろげる
   ' ・（先頭列の行数）×（最長行の列数）を範囲とする
   ' R_n - 開始するセルを含む範囲-Range-
   ' q - ＜オプション＞何回目の空白を終わりとみなすか：0だと無制限
   '
   ' 連続列最大行を求める場合には、デフォルトを q = 1 に固定してみる。
   ' 
   Dim r0 As Long
   Dim c0 As Long
   r0 = R_n.Row
   c0 = R_n.Column
   Dim nr As Long
   Dim nC As Long
   nr = R_n.Rows.count
   nC = R_n.Columns.count
   Dim qi As long
   If (nr = 1) And (nC = 1) Then
      ' Call ExpandRangeCont(R_n, strName, cx)
      qi = q
      Dim rz As Long
      rz = 列の最終行_range(R_n, , qi)
      Set R_n = R_n.Resize((rz - r0 + 1), 1)
      qi = q
      Dim cz As Long
      cz = 複数行の最終列_range(R_n, qi)
      ' cz が 0 になってしまうのはなぜ。
      Set R_n = R_n.Resize((rz - r0 + 1), (cz - c0 + 1))
   End If
   Set range_連続列最大行_range = R_n
End Function

Function range_連続列最大行_namedrange(strRangeName As String, _
                                    Optional q As Long = 1) As Range
   '
   ' 複数列の最終行を使って、範囲をひろげる
   ' ・（先頭列の行数）×（最長行の列数）を範囲とする
   ' R_n - 開始するセルを含む範囲-Range-
   ' q - ＜オプション＞何回目の空白を終わりとみなすか：0だと無制限
   '
   ' 連続列最大行を求める場合には、デフォルトを q = 1 に固定してみる。
   ' 
   ' 名前をつけた左上セルから、不定列数の連続行の範囲に拡張して範囲を返す
   Dim R_n As Range
   Set R_n = ThisWorkbook.Names(strRangeName).RefersToRange
   ' Set R_n =...
   Set range_連続列最大行_namedrange = range_連続列最大行_range(R_n, q)
End Function
