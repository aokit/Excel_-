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

Function 列の最終行(n As String, Optional k As Long = 1) As Long
   ' n - 開始するセル（範囲でもよい）に名付けた名前（文字列）
   ' k - その範囲の中の列番号（オプション）
   Dim r1 As Long
   Dim r2 As Long
   Dim mr As Long
   Dim s As Variant
   Dim R_n As Range
   Set R_n = ThisWorkbook.Names(n).RefersToRange
   r1 = R_n.Row
   mr = Rows.count ' 行の最大値・・・ここで飽和する。
   Set s = R_n.Columns(k).End(xlDown)
   r2 = s.Row
   If r2 = mr Then
      列の最終行 = r1
      Exit Function
   End If
   Do While Not (r2 = mr)
      ' Debug.Print s.Value
      r1 = r2
      Set s = s.End(xlDown)
      r2 = s.Row
   Loop
   列の最終行 = r1
End Function

Function 行の最終列(n As String, Optional k As Long = 1) As Long
   ' n - 開始するセル（範囲でもよい）に名付けた名前（文字列）
   ' k - その範囲の中の行番号（オプション）
   Dim c1 As Long
   Dim c2 As Long
   Dim mc As Long
   Dim s As Variant
   Dim R_n As Range
   Set R_n = ThisWorkbook.Names(n).RefersToRange
   c1 = R_n.Column
   mc = Columns.count ' 行の最大値・・・ここで飽和する。
   Set s = R_n.Rows(k).End(xlToRight)
   c2 = s.Column
   If c2 = mc Then
      行の最終列 = c1
      Exit Function
   End If
   Do While Not (c2 = mc)
      ' Debug.Print s.Value
      c1 = c2
      Set s = s.End(xlToRight)
      c2 = s.Column
   Loop
   行の最終列 = c1
End Function
