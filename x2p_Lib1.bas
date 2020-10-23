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
