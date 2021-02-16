Attribute VB_Name = "x2p_▼７"
Option Explicit

' -*- coding:shift_jis-dos -*-

'./x2p_▼７.bas

' ┏━━
' ┃▼７
'
' ◆仕向地辞書（集計名／別名⇒行番号）の生成
' ┃　・名付け範囲『仕向地別名』から
' ◆仕向地集計名配列（行番号→集計名）の生成
' ◆仕向地集計
' ◆仕向地別名回数表示（０回と２回以上を名付け範囲『仕向地別名回数』に表示）
' 　　・名付け範囲『仕向地別名』と同じシートに『仕向地別名回数』（２列）を設定
' 　　　仕向地の別名に関して、別名に記載されていないもの（０回）　と
' 　　　別名に複数回（２回以上）記載されているもの　を回数とともに表示している。
'
Sub 仕向地別集計()
   '
   ' Stop
   '
   Dim str承認記録() As String
   Dim dic仕向番号 As New Dictionary
   Dim ary集計名() As String
   Dim ary集計名件数金額() As Variant 'Long
   ' Dim ary未割当仕向() As String
   Dim ary未割当仕向() As Variant
   Dim nClass As Long
   '
   Call 承認記録配列格納("承認記録", str承認記録)
   Call 仕向地集計期間フィルタ(str承認記録, 3)
   '
   Stop
   '
   Call 仕向地辞書生成("仕向地別名", nClass, dic仕向番号)
   ' ┗・・・ここで　Debug.Print(dic仕向番号("イタリヤ")(1))　とすれば
   ' 　　　　dic仕向番号に読み込まれていることがわかる
   Stop
   '
   Call 仕向地集計名配列生成("仕向地別名", nClass, ary集計名件数金額)
   ' ┗・・・このあと、str承認記録(2,3) などアクセスすると落ちる。
   ' 直後ではなくて、１分程度して、Excel自体が落ちる。
   '
   Stop
   '
   Call 仕向地集計(str承認記録, 11, 10, dic仕向番号, _
                   ary集計名件数金額, ary未割当仕向, 3)
   ' Stop
   ' ┗・・・このあと　集計名件数金額　と　未割当仕向　を表示する
   '
   Stop
   '
   ' ┏仕向地集計名件数金額表示
   Call PrintArrayOnNamedRange("仕向地集計", ary集計名件数金額, 3, -4)
   ' ┏未割当別名表示
   ' Stop
   'Call PrintArrayOnNamedRange("未割当別名", ary集計名件数金額, 1)
   '
   ' ソートし、所定の範囲に記入
   ' 
   Dim i As Long
   
   Call 降順ソート(ary集計名件数金額, 2) '件数降順
   Dim ary集計名件数() As Variant
   ReDim ary集計名件数(1 To nClass, 1 To 2)
   For i = 1 To nClass
      ary集計名件数(i, 1) = ary集計名件数金額(i, 1)
      ary集計名件数(i, 2) = ary集計名件数金額(i, 2)
   Next i
   Call PrintArrayOnNamedRange("仕向地件数＿トップ", ary集計名件数, 2, -4)
   
   Call 降順ソート(ary集計名件数金額, 3) '金額降順
   Dim ary集計名金額() As Variant
   ReDim ary集計名金額(1 To nClass, 1 To 2)
   For i = 1 To nClass
      ary集計名金額(i, 1) = ary集計名件数金額(i, 1)
      ary集計名金額(i, 2) = ary集計名件数金額(i, 3) / 1000
   Next i
   Call PrintArrayOnNamedRange("仕向地金額＿トップ", ary集計名金額, 2, -4)

   '
   Call PrintArrayOnNamedRange("未割当別名", ary未割当仕向, 1)
   'Call PrintArrayOnNamedRange("未割当別名", Transpose(ary未割当仕向), 1)
   ' Dim aryT As Variant
   ' Dim Ur As Long
   ' Dim Lr As Long
   ' Ur = UBound(ary未割当仕向, 1)
   ' Lr = LBound(ary未割当仕向, 1)
   ' ReDim aryT(Lr To Ur, 1 To 1)
   ' For r = Lr To Ur
   '    aryT(r, 1) = ary未割当仕向(r)
   ' Next r
   ' Stop
   ' Call PrintArrayOnNamedRange("未割当別名", aryT, 1)
   ' 
   ' Stop
   ' Call 仕向地別名回数表示("仕向地別名回数")
End Sub

Private Sub 承認記録配列格納(strNameD As String, _
                             str承認記録() As String)
   Call NamedRangeSQ2ArrStr(strNameD, str承認記録)
End Sub

Private Sub 仕向地集計期間フィルタ(ByRef str承認記録() As String, _
                                   ByVal iC_pTerm As Long)
   '
   Dim U2 As Long
   U2 = UBound(str承認記録, 1)
   Dim iC_承認日 As Long
   iC_承認日 = 2
   Dim pTerm As Boolean
   ' Dim iC_pTerm As Long
   ' iC_pTerm = 3
   '
   Dim i As Long
   For i = 2 To U2
      pTerm = p有効期間(str承認記録(i, iC_承認日))
      ' これだけで落ちるか？→ Yes
      ' 矢継ぎ早にやると、これだけで落ちる。ゆっくり F8 で２周ほどすれば
      ' そのあと F5 でも大丈夫・・・なんだろこれ。
      ' strx = str承認記録(i, iC_承認日)
      ' これだけで落ちるか？→ No
      str承認記録(i, iC_pTerm) = Str(pTerm)
   Next i
End Sub


Sub 仕向地集計(ByRef str承認記録() As String, _
                       ByVal iC_仕向地 As Long, _
                       ByVal iC_金額 As Long, _
                       ByRef dicDIN As Dictionary, _
                       ByRef aryNTM() As Variant, _
                       ByRef aryYAD() As Variant, _
                       ByVal iC_pTerm)
   '
   Stop
   '
   '...................ByRef aryYAD() As String)
   ' aryYAD　未割当仕向　の扱い
   ' 表示させるときに aryYAD も Variant でないと型が一致せずコンパイルエラー
   ' を起こしてしまうので、なんだか腑に落ちないが Variant にした。
   ' また、表に戻す際に、１次元だと行になってしまうので、２次元の１列配列に
   ' 移し替えている。
   '
   ' 第１引数：strNameD：の名前が与えられた範囲から承認データを読み取る。
   ' 第２引数：iC_仕向地：承認データにおける仕向地のカラム番号
   ' 第３引数：iC_金額：承認データにおける金額のカラム番号
   ' 第４引数：dicDIN：仕向地の別名辞書
   ' 第５引数：aryNTM：（返す値）仕向地名・合計金額・合計回数の表
   ' 第６引数：aryYAD：（返す値）割当解決ができなかった仕向地の表
   ' 
   ' ("承認記録", 11, 10, dic仕向番号, ary集計名件数金額)
   ' 第６引数：未割り当て仕向地を返す（）
   '
   ' ReDim aryYAD(1 To 200)
   ' ２次元にしておく
   ' ReDim aryYAD(1 To 200, 1 To 1)
   Dim TaryYAD(1 To 200) As String
   ' ┗・・・未割当の仕向を格納する（返すときには未使用を削る）
   Dim nYAD As Long
   nYAD = 0
   '
   ' Dim str承認記録() As String
   ' Call 承認記録読み取り(str承認記録)
   ' Call NamedRangeSQ2ArrStr("承認記録", str承認記録)
   ' Call NamedRangeSQ2ArrStr(strNameD, str承認記録)
   '
   ' Stop
   '
   ' ここから　配列　str承認記録　に対して　▼３／▼５を参考にして集計処理を行う。
   ' Dim U2 As Long
   ' 件数と金額を初期化（Emptyではなく0に）
   ' U2 = UBound(aryNTM, 1)
   ' Dim i As Long
   ' For i = 1 To U2
   '    aryNTM(i, 2) = 0
   '    aryNTM(i, 3) = 0
   ' Next i
   ' str承認記録　の全行をスキャン
   U2 = UBound(str承認記録, 1)
   ' Dim iC_承認日 As Long
   ' iC_承認日 = 2
   Dim IX As Long
   Dim str仕向地 As String
   Dim str金額 As String
   Dim strx As String
   Dim pterm As Boolean
   Dim aryIX() As Variant
   ' Dim iC_pTerm As Long
   ' iC_pTerm = 3
   '
   Stop
   ' --- ここからステップ実行で確認
   ' For i = 2 To U2
   '    pTerm = p有効期間(str承認記録(i, iC_承認日))
   '    ' これだけで落ちるか？→ Yes
   '    ' 矢継ぎ早にやると、これだけで落ちる。ゆっくり F8 で２周ほどすれば
   '    ' そのあと F5 でも大丈夫・・・なんだろこれ。
   '    ' strx = str承認記録(i, iC_承認日)
   '    ' これだけで落ちるか？→ No
   '    str承認記録(i, iC_pTerm) = Str(pTerm)
   ' Next i
   '
   ' ここまでですでに問題が起きている・・・
   '
   ' Stop
   '
   For i = 2 To U2
      ' pTerm = p有効期間(str承認記録(i, iC_承認日))
      pTerm = CStr(str承認記録(i, iC_pTerm))
      If pTerm Then
      End If
   Next i
   '
   Stop
   '
   For i = 2 To U2
      pTerm = CStr(str承認記録(i, iC_pTerm))
      If pTerm Then
      ' p有効期限　を呼ばなければ大丈夫か？
      ' よばなければそこでは落ちないが、あとから落ちることが判明。
      ' If True Then
         str仕向地 = str承認記録(i, iC_仕向地)
         str金額 = str承認記録(i, iC_金額)
         '
         Stop
         '
         If dicDIN.Exists(str仕向地) Then
            ' Dim aryIX() As Variant
            aryIX = dicDIN.Item(str仕向地)
            ' ┗・・・仕向地の Identification Number（複数割当もあり得るのでの配列）
            Dim U22 As Long
            u22 = UBound(aryIX, 1)
            Dim j As Long
            For j = 1 To U22
               If j > 1 Then Debug.Print j
               IX = aryIX(j)
               aryNTM(IX, 2) = aryNTM(IX, 2) + 1
               ' aryNTM(IX, 3) = aryNTM(IX, 3) + Val(str金額)
               aryNTM(IX, 3) = aryNTM(IX, 3) + Val(str金額) / U22
               ' ┗・・・別名割当回数で均等分割して積算
            Next j
            ' Erase aryIX
            ' ＜↓変更↓＞
            ''' 不要では？
         Else
            Debug.Print(str仕向地)
            ' ┗集計名にも別名にもない仕向地はひとまず印刷しておく。
            ' 　名付け範囲『別名回数』に０回として表示するために
            ' 　配列を構成しておく予定。
            nYAD = nYAD + 1
            TaryYAD(nYAD) = str仕向地
         End If
      End If
   Next i
   ' --- ここまでの実行になにか問題がある。
   If nYAD = 0 Then
      ' 未割当仕向がまったく無かったとき、配列もなくしてしまうと例外処理
      ' が面倒なので、特別に１要素で空文字列の配列にしておく。
      ReDim aryYAD(1 To 1, 1 To 1)
      aryYAD(1, 1) = ""
   Else
      ReDim aryYAD(1 To nYAD, 1 To 1)
      ' ┗・・・未割当仕向の配列で書き込んでいないところを切り落とす
      ' 　　　　名付け範囲『別名回数』に０回として表示する対象
      For j = 1 To nYAD
         aryYAD(j, 1) = TaryYAD(j)
      Next j
   End If
   ' Stop
   '
End Sub
