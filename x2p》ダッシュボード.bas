' -*- coding:shift_jis -*-

'./x2p》ダッシュボード.bas

' ┏━━
' ┃▼０
'
Sub 名前の定義確認の生成()
   '
   ' 集計処理で参照／表示する範囲を指定するために名前付けが済んでいるかの
   ' チェックリストを生成する。
   ' 集計処理の結果や状況をまとめて表示する名前付け範囲を生成する。
   '
   Call 開始時抑制
   Dim BA As Variant
   Set BA = ActiveSheet.Shapes(Application.Caller)
   '   ┗関数を起動したボタンのあるセル範囲を確保しておく
   Dim c1 As Long
   Dim c2 As Long
   Dim r As Long
   Dim r0 As Long
   Dim rz As Long
   ' チェックリスト生成：
   ' ボタンの左下のセルから名前付けに用意した文字列がセルに格納してあるので
   ' それらの名前について、範囲が割り当てられているか表示するような式を隣の
   ' セルに与える。
   c2 = BA.TopLeftCell.Column
   c1 = c2 - 1
   r0 = BA.TopLeftCell.Row + 1
   ' rz = 列の最終行_range(Cells(r0, c1), , 2) ' 最初の空白行の手前の行
   ' rz = 列の最終行_range(Cells(r0, c1), , 3) ' 最初の空白行の手前の行
   ' ┗・・・なぜかここ、パラメータ（上記の３）を増やさないといけなかった。
   ' 　　　　どうしてかは未解明
   '＜↓変更↓＞
   rz = 列の最終行_range(Cells(r0, c1),, 1) ' 最初の空白行の手前の行
   ' stop
   For r = r0 To rz
      Cells(r, c2).Value = "=isref(" & Cells(r, c1).Value & ")"
   Next r
   ' 表示のための名前付け範囲生成：
   ' その下の空白につづいて状況表示用のセルとその名前を配置する。
   ' r0 = 列の最終行_range(Cells(rz, c1), , 4)
   ' rz = 列の最終行_range(Cells(rz, c1), , 5)
   '＜↓変更↓＞
   r0 = 列の最終行_range(Cells(rz, c1),, 1)
   rz = 列の最終行_range(Cells(r0, c1),, 1) ' 最初の空白行の手前の行
   For r = r0 To rz
      Call newName2Range(Cells(r, c2), Cells(r, c1).Value)
   Next r
   Call 終了時解放
End Sub

' ┏━━
' ┃▼１
'
Sub 組織辞書初期化()
   '
   ' 組織表（名前『組織』で定義した範囲-Range-　いまのところ、
   ' 組織名称・略称・英文呼称のシートにある）の表記内容に従って
   ' 名前『Range_組織辞書』で定義した範囲に組織辞書を展開する。
   ' 名前『Range_組織辞書』で定義した範囲は最初は左上となる１セルだけであるが
   ' 初期化によって、必要な大きさの範囲に書き換えられる。
   ' ＜↓変更↓＞
   ' ▼参照：ブック内の名前『組織』
   ' 　　　　＞３文字の組織名を１列に表記。
   ' 　　　　＞上位組織は色付きのセル・下位組織は色なしのセルで表現。
   ' 　　　　＞シート『組織名称・略称・英文呼称』に置く。
   ' 　　　　＞他の名前、他のシートであっても問題ない。
   ' ▲結果：ブック内の名前『組織辞書┏』
   ' 　　　　＞左上の１セルだけの名前として定義。　組織辞書　として展開。
   ' 　　　　＞組織辞書　は、１列目が上位組織、２列め以降が属する下位組織。
   ' 　　　　＞シート上では、１列以上の、不定長（不定列数）連続行　
   ' 　　　　＞他のすべてのプロシジャでの　組織辞書　の参照は、この名前に
   ' 　　　　　よって行う。名前から　不定長連続行　の範囲を確定する機能を
   ' 　　　　　それぞれのプロシジャが用意する。
   ' 名前『Range_組織辞書』で定義した範囲を組織の辞書として使う
   ' ・組織集計の第１列、集計名を初期化するとき
   ' ・同第３列、件数を集計するとき
   ' ＜↓変更↓＞
   ' 名前『組織辞書┏』の内容は、名前『組織集計』をもつ列（ダッシュボード
   ' にある）を（『組織』にもとづいて）更新する途中で更新される。
   ' また、名前『組織集計』の列を第１列とした第３列（集計）を更新するとき
   ' に、集計処理で参照される（下位組織を上位組織で集計するために必要）
   ' 
   ' なお初期化で生成される組織辞書は初期値であって、書き換えることが想定されて
   ' いる。ただし、直接セルを書き換えることは想定しない。Range_組織辞書（範囲）
   ' の更新、途中に空白がない、多重帰属がない、などの整合性を維持したり、操作の
   ' 利便性（２セルの選択とボタンクリックで登録）のために別途プロシジャを用意す
   ' る予定。
   ' ＜↓変更↓＞
   ' 『組織辞書┏』で左上セルを定義した、不定長連続行の範囲である組織辞書は、
   ' 『▼１』で『組織』から生成される。組織辞書は、『組織集計』の第１列のセルへ
   ' の記載と、第３列のセルへの集計処理に利用される。
   ' 組織辞書では、多重帰属（複数の上位組織に同一の下位組織が所属すること）は
   ' 起こらないものとする。今回は、『組織』を公式の機関から入手したが、組織変更
   ' の際に、あらたに『組織』の内容を入手することができない場合にも、組織辞書で
   ' はなく『組織』を編集（現在の設定をコピー、編集して、『組織』と名前定義）
   ' するものとした。
   '
   ' Call 組織略称クリア
   ' ＜↓変更↓＞
   ' ＝＝＝＝＝＝
   '
   ' １┏・・・背景色識別で作られた組織表の背景色と内容を配列にそれぞれ格納する
   Dim 組織略称CI() As Long
   Dim 組織略称ST() As String
   Call 組織略称読み取り(組織略称CI(), 組織略称ST())
   '
   ' ２┏・・・背景色識別で作られた組織表の配列を行単位の組織表に変換する
   Dim str組織辞書() As Variant
   Call CIonST2Arr(組織略称CI(), 組織略称ST(), str組織辞書())
   '
   Stop
   
   ' ３┏・・・名付けした範囲に配列を書き出し範囲を広げて名付けを更新する
   Dim strName As String
   ' strName = "Range_組織辞書"
   ' Call Arr2ReNamedRange(str組織辞書(), strName)
   ' ＜↓変更↓＞
   strName = "組織辞書┏"
   Call PrintArrayOnNamedRange(strName, str組織辞書)

End Sub

' ┏━━
' ┃▼２
'
Sub 組織集計1列初期化()
   '
   ' 名前『Range_組織辞書』で名付けた範囲に組織辞書にもとづいて、
   ' 名前『組織集計』で名付けた範囲（集計名の左上セルから必要な行数の１列の
   ' 範囲を生成する。
   '
   ' 名前『組織集計』で名付けた範囲をもとに
   ' 名前『Range_組織集計１列』で名付けた範囲を構成する。
   ' まず、既存の領域をクリアするようにした。既存の領域が大きくても機能する。
   '
   Dim strName As String
   ' strName = "Range_組織辞書"
   ' Dim Range_組織辞書 As Range
   ' Set Range_組織辞書 = _
   '     ThisWorkbook.Names(strName).RefersToRange
   ' Dim str組織辞書() As Variant
   ' str組織辞書 = Range_組織辞書.Value
   '┗所定行数×１列の配列
   ' ＜↓変更↓＞
   strName = "組織辞書┏"
   Dim rh As long
   Dim 組織辞書 As Range
   Set 組織辞書 = ThisWorkbook.Names(strName).RefersToRange
   rh = Range_列の最終行_Range(組織辞書,, 1).Rows.Count
   Set 組織辞書 = 組織辞書.Resize(rh,1)
   Dim str組織辞書() As Variant
   str組織辞書 = 組織辞書.Value
   '┗所定行数×１列の配列
   
   ' Dim j As Long
   ' j = UBound(str組織辞書, 1)
   ' strName = "組織集計"
   ' Dim Range_組織集計1列 As Range
   ' Dim r0 As Long
   ' Dim r1 As Long
   ' Set Range_組織集計1列 = _
   '     ThisWorkbook.Names(strName).RefersToRange
   ' ＜↓変更↓＞
   Dim j As Long
   j = UBound(str組織辞書, 1)
   strName = "組織集計"
   Dim Range_組織集計1列 As Range
   Set Range_組織集計1列 = _
       ThisWorkbook.Names(strName).RefersToRange

   ' r0 = Range_組織集計1列.Row
   ' r1 = 列の最終行(strName)
   ' Set Range_組織集計1列 = Range_組織集計1列.Resize((r1 - r0 + 1), 1)
   ' Range_組織集計1列.Clear
   Set Range_組織集計1列 = Range_組織集計1列.Resize(j, 1)
   Call ClearColumnRowEnd(Range_組織集計1列)
   Range_組織集計1列.Font.Name = "BIZ UDゴシック"
   ' Range_組織集計1列.Value = str組織辞書
   Range_組織集計1列 = str組織辞書
   '
End Sub

' ┏━━
' ┃▼３
'
Sub 組織別個別審査件数集計()
   Dim str承認記録() As String
   Call 承認記録読み取り(str承認記録)
   Dim dic組織辞書 As New Dictionary
   Dim U1 As Long
   If False Then 'True then
      ' Unused
      ' 以下の True節 は使われなくなったはず。リファクタリングでバサッと
      ' 消せるように以下のアサートの残す。
      Debug.Print("使われなくなったTrue節が実行されました。")
      Exit Sub
      '
      Dim str組織辞書() As String
      Call 組織辞書読み取り(str組織辞書)
      Call 組織辞書構成(str組織辞書, dic組織辞書)
      U1 = UBound(str組織辞書, 1)
   Else
      ' 以下のに置き換えてみる
      ' Call SingleHomeDict_namedRange("Range_組織辞書", U1, dic組織辞書, 0)
      '＜↓変更↓＞
      Call SingleHomeDict_namedRange("組織辞書┏", U1, dic組織辞書, 0)
   End If
   '
   Dim 組織別個別審査件数() As Long
   ReDim 組織別個別審査件数(1 To U1, 1 To 1)
   '
   Dim 申請者所属 As String
   Dim IX As Long
   Dim U2 As Long
   U2 = UBound(str承認記録, 1)
   For i = 2 To U2
      ' Debug.Print str承認記録(i, 2)
      If p有効期間(str承認記録(i, 2)) Then
         ' 15 - 申請者所属
         ' Debug.Print str承認記録(i, 15)
         申請者所属 = str承認記録(i, 15)
         If (dic組織辞書.Exists(申請者所属)) Then
            IX = dic組織辞書.Item(申請者所属)
            ' Debug.Print IX
            組織別個別審査件数(IX, 1) = 組織別個別審査件数(IX, 1) + 1
         End If
      End If
   Next i
   For i = 1 To U1
      ' Debug.Print 組織別個別審査件数(i, 1)
   Next i
   Call 組織集計個別審査件数更新(組織別個別審査件数)
   '
   ' Stop
   '
   Call 組織集計個別審査件数転記()
End Sub

Sub 組織集計個別審査件数転記()
   Dim R_1 As Range
   Dim R_2 As Range
   Dim R_3 As range
   Set R_1 = range_列の最終行_namedrange("組織集計",, 1)
   Set R_2 = R_1.Offset(0, 2) ' 組織集計の集計値
   Dim v1() As Variant
   Dim v2() As Variant
   Dim v3() As Variant ' - v3 は２次元配列であることを明示すること
   ' 　　┗ここの『()』がないと、コンパイラが配列でないと判断してエラー。
   v1 = R_1.value
   v2 = R_2.value
   Dim LB1 As Long
   Dim UB1 As Long
   LB1 = LBound(v1, 1)
   UB1 = UBound(v1, 1)
   If LB1 <> 1 Then Debug.Print("組織集計に異常があります。＠転記")
   If LB1 <> LBound(v2, 1) Then Debug.Print("組織集計に異常があります。＠転記")
   If UB1 <> UBound(v2, 1) Then Debug.Print("組織集計に異常があります。＠転記")
   ReDim v3(1 To UB1, 1 To 2)
   For i = 1 To UB1
      v3(i, 1) = v1(i, 1)
      v3(i, 2) = v2(i, 1)
   Next i
   ' Set R_3 = range_列の最終行_namederange("別表１").Offset(1,0)
   
   Set R_3 = ThisWorkBook.Names("別表１").RefersToRange.Offset(1, 0).Cells(1, 1)

   ' Call PrintArrayOnRange(R_3, v3, 2)

   ' Call PrintArrayOnRange(R_3, v3, 0)
   Call PrintArrayOnRange(R_3, v3, -1)
   
End Sub

'┗━━━━━━
Function PickHeadWord(ByRef dicSyn As Dictionary, _
                 ByVal word As String) As String
   ' 番号で類別された類義語辞書の見出し語を返す
   '
   PickHeadWord = ""
   If Not dicSyn.Exists(word) Then Exit Function
   Dim ix As Long
   ix = dicSyn(word)
   Dim i  As Long
   For i = 0 To dicSyn.Count - 1
      If dicSyn(dicSyn.Keys(i)) = ix Then
         PickHeadWord = dicSyn.Keys(i)
         Exit For
      End If
   Next i   
End Function
'┏━━━━━━

' ┏━━
' ┃▼４
'
Sub 組織集計_非ゼロ更新()
   Dim 組織集計_非ゼロ() As Variant
   ' ┗・・・この配列は、文字列以外も引き渡す目的で Variant としておく。
   Call 組織集計_非ゼロ抽出(組織集計_非ゼロ)
   Call 組織集計_非ゼロ書出(組織集計_非ゼロ)
End Sub

' ┏━━
' ┃▼５
'
Sub 取引区分別個別審査件数集計()
   '
   Dim str承認記録() As String
   Call 承認記録読み取り(str承認記録)
   '
   Dim dic取引区分辞書 As New Dictionary
   Dim U1 As Long
   Call SingleHomeDict_namedRange("取引集計", U1, dic取引区分辞書, 1)
   ' ┗『取引集計』と名付けたセルの直下１列を dic取引区分辞書 として確保
   '
   Dim 取引区分別個別審査件数() As Long
   ReDim 取引区分別個別審査件数(1 To U1, 1 To 1)
   Dim 取引区分別個別審査金額() As Long
   ReDim 取引区分別個別審査金額(1 To U1, 1 To 1)
   '
   Dim 取引区分 As String ' 表では『取引内容区分』
   Dim 金額 As Long ' 表では『金額（円）』
   Dim IX As Long
   Dim U2 As Long
   U2 = UBound(str承認記録, 1)
   For i = 2 To U2
      If p有効期間(str承認記録(i, 2)) Then
         ' 8 - 表では『取引内容区分』
         ' 10 - 表では『金額（円）』
         ' Debug.Print str承認記録(i, 8)
         ' Debug.Print str承認記録(i, 10)
         取引区分 = str承認記録(i, 8)
         金額 = str承認記録(i, 10)
         If (dic取引区分辞書.Exists(取引区分)) Then
            IX = dic取引区分辞書.Item(取引区分)
            ' Debug.Print IX
            取引区分別個別審査件数(IX, 1) = 取引区分別個別審査件数(IX, 1) + 1
            取引区分別個別審査金額(IX, 1) = 取引区分別個別審査金額(IX, 1) + 金額
         End If
      End If
   Next i
   ' stop
   Call 取引区分集計個別審査件数更新(取引区分別個別審査件数)
   Call 取引区分集計個別審査金額更新(取引区分別個別審査金額)
End Sub

' ┏━━
' ┃▼６
'
Sub 取引集計_非ゼロ更新()
   Dim 取引集計_非ゼロ() As Variant
   ' ┗・・・この配列は、文字列以外も引き渡す目的で Variant としておく。
   Call 取引集計_非ゼロ抽出(取引集計_非ゼロ)
   Call 取引集計_非ゼロ書出(取引集計_非ゼロ)
End Sub

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
   Dim dic仕向番号 As New Dictionary
   Dim ary集計名() As String
   Dim ary集計名件数金額() As Variant
   ' Dim ary未割当仕向() As String
   Dim ary未割当仕向() As Variant
   Dim nClass As Long
   Call 仕向地辞書生成("仕向地別名", nClass, dic仕向番号)
   ' Stop
   ' ┗・・・ここで　Debug.Print(dic仕向番号("イタリヤ")(1))　とすれば
   ' 　　　　dic仕向番号に読み込まれていることがわかる
   Call 仕向地集計名配列生成("仕向地別名", nClass, ary集計名件数金額)
   ' ┗・・・
   Call 仕向地集計("承認記録", 11, 10, dic仕向番号, ary集計名件数金額, ary未割当仕向)
   ' Stop
   ' ┗・・・このあと　集計名件数金額　と　未割当仕向　を表示する
   ' ┏仕向地集計名件数金額表示
   Call PrintArrayOnNamedRange("仕向地集計", ary集計名件数金額, 3, -4)
   ' ┏未割当別名表示
   ' Stop
   'Call PrintArrayOnNamedRange("未割当別名", ary集計名件数金額, 1)
   '
   ' ソートし、所定の範囲に記入
   ' 
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

Private Sub 仕向地集計(strNameD As String, _
                       iC_仕向地 As Long, _
                       iC_金額 As Long, _
                       ByRef dicDIN As Dictionary, _
                       ByRef aryNTM() As Variant, _
                       ByRef aryYAD() As Variant)
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
   Dim TaryYAD(1 To 200)
   ' ┗・・・未割当の仕向を格納する（返すときには未使用を削る）
   Dim nYAD As Long
   nYAD = 0
   Dim str承認記録() As String
   ' Call 承認記録読み取り(str承認記録)
   ' Call NamedRangeSQ2ArrStr("承認記録", str承認記録)
   Call NamedRangeSQ2ArrStr(strNameD, str承認記録)
   ' Stop
   ' ここから　配列　str承認記録　に対して　▼３／▼５を参考にして集計処理を行う。
   Dim U2 As Long
   ' 件数と金額を初期化（Emptyではなく0に）
   U2 = UBound(aryNTM, 1)
   For i = 1 To U2
      aryNTM(i, 2) = 0
      aryNTM(i, 3) = 0
   Next i
   ' str承認記録　の全行をスキャン
   U2 = UBound(str承認記録, 1)
   Dim iC_承認日 As Long
   iC_承認日 = 2
   Dim IX As Long
   Dim str仕向地 As String
   Dim str金額 As String
   For i = 2 To U2
      If p有効期間(str承認記録(i, iC_承認日)) Then
         str仕向地 = str承認記録(i, iC_仕向地)
         str金額 = str承認記録(i, iC_金額)
         If dicDIN.Exists(str仕向地) Then
            Dim aryIX() As Variant
            aryIX = dicDIN.Item(str仕向地)
            ' ┗・・・仕向地の Identification Number（複数割当もあり得るのでの配列）
            Dim U22 As Long
            u22 = UBound(aryIX, 1)
            For j = 1 To U22
               If j > 1 Then Debug.Print j
               IX = aryIX(j)
               aryNTM(IX, 2) = aryNTM(IX, 2) + 1
               ' aryNTM(IX, 3) = aryNTM(IX, 3) + Val(str金額)
               aryNTM(IX, 3) = aryNTM(IX, 3) + Val(str金額) / U22
               ' ┗・・・別名割当回数で均等分割して積算
            Next j
            Erase aryIX
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

' ┏━━
' ┃▼８
'
Sub 許可特例等抽出()
   '
   '
   ' ▼３のコードを流用
   Dim str承認記録() As String
   Call 承認記録読み取り(str承認記録)
   Dim U2 As Long
   U2 = UBound(str承認記録, 1)

   Dim dic組織辞書 As New Dictionary
   Dim U1 As Long

   If False Then
      ' 使われなくなった（範囲の変わる　Range_組織辞書　ではなく
      ' 単一セルの　組織辞書┏　を使うようにした）
      ' Call SingleHomeDict_namedRange("Range_組織辞書", U1, dic組織辞書, 0)
   Else
      Call SingleHomeDict_namedRange("組織辞書┏", U1, dic組織辞書, 0)
   End If
   
   Dim 申請者所属 As String

   Dim A1() As String ' 包括許可
   Dim A2() As String ' 包括許可（貨物）
   Dim A3() As String ' 包括許可（役務）
   Dim A4() As String ' 少額特例
   Dim A5() As String ' 公知特例
   Dim A6() As String ' 該当国内
   Dim A1A() As String
   Dim A4A() As String
   Dim A1V() As Variant
   Dim A4V() As variant
      
   ReDim A1(1 To U2)
   ReDim A2(1 To U2)
   ReDim A3(1 To U2)
   ReDim A4(1 To U2)
   ReDim A5(1 To U2)
   ReDim A6(1 To U2)
   ReDim A1A(1 To U2, 1 To 7)
   ReDim A4A(1 To U2, 1 To 7)
   
   Dim i1 As Long
   Dim i2 As Long
   Dim i3 As Long
   Dim i4 As Long
   Dim i5 As Long
   Dim i6 As Long

   i1 = 1
   i2 = 1
   i3 = 1
   i4 = 1
   i5 = 1
   i6 = 1
   
   Dim c1 As String
   Dim c17 As String
   Dim c18 As String
   Dim c19 As String
   Dim c20 As String
   Dim q As Boolean

   Dim yen As Long
   
   For i = 2 To U2
      q = False
      If p有効期間(str承認記録(i, 2)) Then
         c1 = str承認記録(i, 1)
         c17 = str承認記録(i, 17)
         c18 = str承認記録(i, 18)
         c19 = str承認記録(i, 19)
         c20 = str承認記録(i, 20)
         If (c18 = "包括許可適用") Or (c20 = "包括許可適用") Then
            A1(i1) = c1
            ' ┏＜依存＞承認記録のフィールド構造
            A1A(i1, 1) = str承認記録(i, 1) ' 管理番号
            A1A(i1, 2) = str承認記録(i, 7) ' 件名
            A1A(i1, 3) = str承認記録(i, 11) ' 仕向地
            A1A(i1, 4) = str承認記録(i, 13) ' 顧客・契約先
            A1A(i1, 5) = str承認記録(i, 14) ' 最終需要者
            ' A1A(i1, 6) = str承認記録(i, 15) ' 申請者所属＿本部変換前
            申請者所属 = str承認記録(i, 15) ' 申請者所属＿本部変換前
            yen = str承認記録(i, 10) ' 金額＿円
            A1A(i1, 7) = CStr(CLng(yen) / 1000)
            If (dic組織辞書.Exists(申請者所属)) Then
               A1A(i1, 6) = PickHeadWord(dic組織辞書, 申請者所属)
            Else
               A1A(i1, 6) = "*"
            End If
            i1 = i1 + 1
            q = True
         End If
         If (c18 = "包括許可適用") Then
            A2(i2) = c1
            i2 = i2 + 1
            q = True
         End If
         If (c20 = "包括許可適用") Then
            A3(i3) = c1
            i3 = i3 + 1
            q = True
         End If
         If (Right(c18, 2) = "特例") Then ' 少額特例
            A4(i4) = c1
            ' ┏＜依存＞承認記録のフィールド構造
            A4A(i4, 1) = str承認記録(i, 1) ' 管理番号
            A4A(i4, 2) = str承認記録(i, 7) ' 件名
            A4A(i4, 3) = str承認記録(i, 11) ' 仕向地
            A4A(i4, 4) = str承認記録(i, 13) ' 顧客・契約先
            A4A(i4, 5) = str承認記録(i, 14) ' 最終需要者
            ' A1A(i1, 6) = str承認記録(i, 15) ' 申請者所属＿本部変換前
            申請者所属 = str承認記録(i, 15) ' 申請者所属＿本部変換前
            yen = str承認記録(i, 10) ' 金額＿円
            A4A(i4, 7) = CStr(CLng(yen) / 1000)
            If (dic組織辞書.Exists(申請者所属)) Then
               A4A(i4, 6) = PickHeadWord(dic組織辞書, 申請者所属)
            Else
               A1A(i4, 6) = "*"
            End If
            i4 = i4 + 1
            q = True
         End If
         If (Right(c20, 2) = "特例") Then
            A5(i5) = c1
            i5 = i5 + 1
            q = True
         End If
         If ((Left(c17, 2) = "該当") Or (Left(c19, 2) = "該当")) And (Not q) Then
            A6(i6) = c1
            i6 = i6 + 1
         End If
      End If
   Next i

   ReDim Preserve A1(1 To i1 - 1)
   ReDim Preserve A2(1 To i2 - 1)
   ReDim Preserve A3(1 To i3 - 1)
   ReDim Preserve A4(1 To i4 - 1)
   ReDim Preserve A5(1 To i5 - 1)
   ReDim Preserve A6(1 To i6 - 1)
   ReDim A1V(1 To i1 - 1, 1 To 7)
   Dim r As Long
   Dim c As Long
   For r = 1 To i1 - 1
      For c = 1 To 7
         A1V(r, c) = A1A(r, c)
      Next c
   Next r
   ReDim A4V(1 To i4 - 1, 1 To 7)
   For r = 1 To i4 - 1
      For c = 1 To 7
         A4V(r, c) = A4A(r, c)
      Next c
   Next r
   
   ' Stop

   Dim AA() As Variant
   Dim iz As Long
   iz = 0
   If iz < (i1 - 1) Then iz = i1 - 1
   If iz < (i2 - 1) Then iz = i2 - 1
   If iz < (i3 - 1) Then iz = i3 - 1
   If iz < (i4 - 1) Then iz = i4 - 1
   If iz < (i5 - 1) Then iz = i5 - 1
   If iz < (i6 - 1) Then iz = i6 - 1
   ReDim AA(1 To iz, 1 To 6)
   For i = 1 To iz
      If UBound(A1, 1) < i Then
         AA(i, 1) = ""
      Else
         AA(i, 1) = A1(i)
      End If
      If UBound(A2, 1) < i Then
         AA(i, 2) = ""
      Else
         AA(i, 2) = A2(i)
      End If
      If UBound(A3, 1) < i Then
         AA(i, 3) = ""
      Else
         AA(i, 3) = A3(i)
      End If
      If UBound(A4, 1) < i Then
         AA(i, 4) = ""
      Else
         AA(i, 4) = A4(i)
      End If
      If UBound(A5, 1) < i Then
         AA(i, 5) = ""
      Else
         AA(i, 5) = A5(i)
      End If
      If UBound(A6, 1) < i Then
         AA(i, 6) = ""
      Else
         AA(i, 6) = A6(i)
      End If
   Next i

   ' Stop

   Call PrintArrayOnNamedRange("該当取引",AA,6,-4)
   Call PrintArrayOnNamedRange("特一包括┏",A1V,7,-4)
   Call PrintArrayOnNamedRange("少額特例┏",A4V,7,-4)
   ' ┗なんか、Root からFormat持ってこられてないような・・・
   '   PrintArrayOnNamedRangeのバグ？
   
End Sub

'
' ┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
' ┃┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
' ┃┃一般化プロシジャ（おもに Private 関数）
' ┃┗

Private Sub Col2CIonSTrng(rngC As Range, _
                       ByRef CI() As Long, _
                       ByRef ST() As String)
   ' 『組織』と名前付けした範囲を読み取り：
   ' ・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・
   ' １：第１引数＜rngC＞が示す範囲（１列×複数行）を読み取り
   ' ２：第２引数＜CI()＞として指定した配列に、範囲の ColorIndex をLong型で返す。
   ' ３：第３引数＜ST()＞として指定した配列に、範囲の セルの値 をString型で返す。
   ' ・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・
   '
   Dim aryC() As Variant
   ' 　┗第１引数の文字列で名付けられた範囲のセルの値を格納する配列。範囲由来の
   '     配列なので次元は２。各次元の要素数は範囲に依存するので不明。要素数を
   '     ReDim などで明示的に指定する扱いはしない。※
   '     要素の型はVariant としている。
   ' ※動的配列として
   ' ・Q:範囲の大きさがわかれば、ReDimで明示的に指定してもよい？
   ' ・A:『範囲のカラム数』や『範囲の行数』は、.Rows.Count などで手に入るが、セルの
   ' 値を手に入れるのはかなり面倒 （組織略称_S(1, 1) = 組織.Cells(1, 1).Value） で
   ' ある。そのため推奨される方法ではない。
   ' （まだこの時点では次元も大きさも未定）としてVariant にしておくのがよい（String
   ' 　にはできない）
   aryC = rngC.Value
   ' ┗組織略称(i,j) = 組織.Cells(i,j)
   ' 範囲-Range-の　組織　は１列なのだが、代入により生成される配列は１次元配列では
   ' なく、２次元配列になることに注意！！
   Dim m As Long
   m = UBound(aryC, 1)
   ' Debug.Print LBound(aryC, 1)
   '   ┗範囲から読み込んだ配列（aryC = rngC.Value）なので　aryC(1,1)　が左上の
   '   　（最初の）セルの値の入る要素になる。
   '     ※　(0,0)ではない
   ' Debug.Print m
   '   ┗行方向（１列のみ）の配列なので、第１の次元の上限値を求めておく。
   ' Debug.Print aryC.Cells(1, 1)
   ' Debug.Print aryC(1)
   ' Debug.Print aryC(1).Cells(1, 1)
   '   ┗『aryC』は２次元配列である。これらのアクセスのしかたはすべて誤り
   ' Debug.Print aryC(1, 1)
   '   ┗範囲の左上の（最初の）セルの値が入っていることを確認できる。
   ' Dim b As Long
   ' b = rngC.Cells(1, 1).Row
   ' Debug.Print b
   ' ┗『rngC』は 範囲-Range- なので .Cell メソッドで行と列によってアクセスする。
   ' 　また、もとの表で何行目であるか（ .Row メソッド ）、などの情報も持っている。
   ReDim CI(m)
   '     ┗『rngC』としてもっているセルの背景色情報を格納する配列を用意する。
   '       範囲を代入するのではないため、明示的に次元と大きさを指定しなければ
   '       ならない。そこで、動的配列として宣言したあと、組織略称（範囲としての
   '       組織から複製した２次元配列）の行数ぶんの要素を持つ１次元の配列を設定
   '       しておく。
   ReDim ST(m)
   '     ┗『aryC』としてもっているセルの値を格納する配列を用意する。
   '
   For r = 1 To m
      CI(r) = rngC.Cells(r, 1).Interior.ColorIndex
      ST(r) = aryC(r, 1)
   Next r
   '
End Sub

' 組織辞書初期化のなかで str組織辞書 を生成するために呼ぶ
'
Private Sub CIonST2Arr(ByRef CI() As Long, _
                       ByRef ST() As String, _
                       ByRef varstrArr() As Variant)
   Dim m As Long
   m = UBound(ST, 1)
   For r = 1 To m
      ' Debug.Print ST(r) & ":" & CI(r)
   Next r
   Dim varArr() As Variant
   ReDim varArr(m, m)
   '     ┗varArrの列がとりうる最大は m 行がとりうる最大は m である。
   '     （ここで定義するとき 0 行 や 0 列 があるが使わず参照もしない）
   Dim i As Long
   Dim j As Long
   Dim k As Long
   Dim RCI As Long
   i = 0: j = 0: k = 0
   RCI = CI(1)
   '┗・・・第１行の色を上位組織の判定基準として使う rootCI
   For r = 1 To m
      If CI(r) = RCI Then
         If k < i Then k = i
         i = 1
         j = j + 1
      Else
         i = i + 1
      End If
      varArr(j, i) = ST(r)
   Next r
   ' varArrは j 行 k 列の配列ということになる。
   ReDim varstrArr(1 To j, 1 To k)
   For q = 1 To k
      For p = 1 To j
         varstrArr(p, q) = varArr(p, q)
      Next p
   Next q
End Sub

' 組織別個別審査件数集計　の中で、表の範囲を更新するために呼ぶ
'
Private Sub 組織集計個別審査件数更新(ByRef 組織別個別審査件数() As Long)
   ' 組織集計個別審査件数クリア
   Dim R組織集計件数列 As Range
   If False Then
      ' 以下使わなくなった。
      Dim strName As String
      strName = "組織集計"
      Dim R組織集計 As Range
      Set R組織集計 = ThisWorkbook.Names(strName).RefersToRange
      r0 = R組織集計.Row
      r1 = 列の最終行(strName,, 1)
      ' c0 = R組織集計.Column + 2
      ' Dim R組織集計件数列 As Range
      ' stop
      ' Set R組織集計件数列 = Range(Cells(r0, c0), Cells(r1, c0))
      ' ┣・・・名付けた範囲をもとに新たな範囲を指定する。
      Set R組織集計件数列 = R組織集計.Offset(0, 2).Resize((r1 - r0 + 1), 1)
      ' R組織集計件数列.Clear
      ' 問答無用で、いまあるところを最後まで消す、というやつにしたほうがいいだろう
      Call ClearColumnRowEnd(R組織集計件数列)
      ' ---
   Else
      '＜↓変更↓＞
      ' Dim R組織集計件数列 As Range
      Set R組織集計件数列 = Range_列の最終行_namedrange("組織集計",, 1).Offset(0, 2)
      R組織集計件数列.Clear
   End If
   ' ---
   R組織集計件数列.Font.Name = "BIZ UDゴシック"
   R組織集計件数列 = 組織別個別審査件数
End Sub

Sub ClearColumnRowEnd(ByVal Ro1C As Range)
   ' 指定された範囲の一番上のセルから空白なしで連続する範
   ' 囲のセルをクリアする
   ' 第１引数：Ro1C - Range of 1 Column - レンジ型の範囲
   ' Ro1Cの一番上のセルから、列の最終行までの範囲の内容を
   ' クリアする。
   ' 
   Dim r0 As Long
   Dim r1 As Long
   r0 = Ro1C.Row
   r1 = 列の最終行_Range(Ro1C,, 1)
   Set Ro1C = Ro1C.Resize((r1 - r0 + 1), 1)
   Ro1C.Clear
End Sub

Sub FillColumnRowEnd(ByVal Ro1C As Range, ByRef Ary As Variant)
   ' 指定された範囲の一番上のセルから空白なしで連続する範
   ' 囲のセルに、複数行×１列の配列を書き込む
   ' 第１引数：Ro1C - Range of 1 Column - レンジ型の範囲
   ' Ro1Cの一番上のセルから、
   ' 第２引数：Ary 複数行×１列の配列を書き込む
   ' 
   Dim rs As Long
   rs = UBound(Ary, 1) - LBound(Ary, 1) + 1
   Set Ro1C = Ro1C.Resize(rs, 1)
   Ro1C.Font.Name = "BIZ UDゴシック"
   Ro1C = Ary
End Sub

' 組織別個別審査件数集計　の中で、承認記録を読み取るために呼び出す
' 仕向地集計　の中でも呼び出す。
'
Private Sub 承認記録読み取り(str_承認記録() As String)
   '
   ' ・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・
   ' 『承認記録』と名前付けした範囲を読み取り：
   ' 　引数として指定した配列に、範囲の セルの値 をString型で返す。
   ' ・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・
   '
   Call NamedRangeSQ2ArrStr("承認記録",str_承認記録)
End Sub

Private Sub NamedRangeSQ2ArrStr(strName As String, _
                                str_承認記録() As String)
   ' 第１引数：同じ長さの複数の行からなる範囲　に名付けた名前
   ' 第２引数：上記の範囲を格納する配列
   ' 領域の大きさを確認（mr, mc）
   Dim mr As Long
   Dim mc As Long
   mr = 列の最終行(strName,, 1)
   mc = 行の最終列(strName,, 1)
   Dim V_承認記録() As Variant
   Dim 承認記録 As Range
   Set 承認記録 = ThisWorkbook.Names("承認記録").RefersToRange.Resize(mr, mc)
   V_承認記録 = 承認記録
   ' ┗V_承認記録(i,j) = 承認記録.Cells(i,j)
   ' Dim str_承認記録 As string
   ReDim str_承認記録(1 To mr, 1 To mc)
   For i = 1 To mr
      For j = 1 To mc
         str_承認記録(i, j) = V_承認記録(i, j)
      Next j
   Next i
   '
End Sub

' 組織別個別審査件数集計　の中で、組織辞書を読み取るために呼び出す
'
Private Sub 組織辞書読み取り(ByRef str組織辞書() As String)
   ' Unused
   ' この関数は使われなくなったはず。確認のために以下にアサートする
   ' ようにしておく。
   Debug.Print("使われなくなった関数が呼ばれました：組織辞書読み取り")
   Exit Sub
   '
   ' ・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・
   ' 『Range_組織辞書』と名前付けした範囲を読み取り：
   ' 　引数として指定した配列に、範囲の セルの値 をString型で返す。
   ' ・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・
   '
   ' 組織辞書のシートは手作業での追記も想定されている。そのため
   ' 領域の大きさを確認（wr, wc）する必要があるので確認。
   Dim strName As String
   ' strName = "Range_組織辞書"
   ' ＜↓変更↓＞
   strName = "組織辞書┏"

   ' Dim r0 As Long
   ' Dim c0 As Long
   ' Dim rz As Long
   ' Dim cz As Long
   ' Dim wr As Long
   ' Dim wc As Long
   ' Dim R組織辞書 As Range
   ' Set R組織辞書 = ThisWorkbook.Names(strName).RefersToRange
   ' r0 = R組織辞書.Row
   ' c0 = R組織辞書.Column
   ' rz = 列の最終行(strName)
   ' cz = 行の最終列(strName)
   ' wr = rz - r0 + 1
   ' wc = cz - c0 + 1
   ' Set R組織辞書 = R組織辞書.Resize(wr, wc)
   ' ＜↓変更↓＞
   Dim R組織辞書 As Range
   Set R組織辞書 = range_連続列最大行_namedrange(strName)

   Dim V組織辞書() As Variant
   V組織辞書 = R組織辞書
   ' 引数として返す配列の大きさをここで設定
   ' ＜↓変更↓＞
   wr = R組織辞書.Rows.Count
   wc = R組織辞書.Columns.Count
   ReDim str組織辞書(1 To wr, 1 To wc)
   For i = 1 To wr
      For j = 1 To wc
         str組織辞書(i, j) = V組織辞書(i, j)
      Next j
   Next i
   Stop
   '
End Sub

' 組織別個別審査件数集計　のなかで　組織別辞書　を構成するために呼ぶ
'
Sub 組織辞書構成(ByRef str辞書() As String, ByRef dic辞書 As Dictionary)
   Call SingleHomeDict(str辞書(), dic辞書)
End Sub

Sub SingleHomeDict(ByRef str辞書() As String, _
                   ByRef dic辞書 As Dictionary)
    '
    ' 『str辞書()』
    '  上位組織を先頭として、その組織に帰属する下位組織を以降に並べた行を
    '  上位組織の数だけならべた配列
    ' 『dic辞書』
    '  str辞書に１から始まる行番号を与えてこれを Value とし、
    '  各行の上位組織および下位組織を key とする辞書
    '
   Dim U1 As Long
   U1 = UBound(str辞書, 1)
   U2 = UBound(str辞書, 2)
   For i = 1 To U1
      d = str辞書(i, 1)
      d0 = i
      dic辞書.Add d, d0
      For j = 2 To U2
         d = str辞書(i, j)
         If d = "" Then Exit For
         If dic辞書.Exists(d) Then
            MsgBox "key複数所属検出：再帰的に辞書を生成して格納"
            ' 複数の仕向地などの対応のため
         Else
            dic辞書.Add d, d0
         End If
      Next j
   Next i
End Sub

' 組織別個別審査件数集計　のなかで集計のための期間に入っているレコードか
' 判定するために使う
'
Function p有効期間(strDate As String, _
                   Optional R_n As String = "集計期間") As Boolean
   '
   ' 『集計期間』と名付けられた２セルの列から期間の開始日と終了日を得て
   '  strDate がその期間に含まれている場合には True を返す。そうでない
   '  なら False を返す。
   '
   Dim r As Boolean
   Dim r1 As Boolean
   Dim r2 As Boolean
   Dim strD As String
   Dim strD1 As String
   Dim strD2 As String
   strD = CDate(strDate)
   strD1 = CDate(ThisWorkbook.Names(R_n).RefersToRange.Cells(1, 1))
   strD2 = CDate(ThisWorkbook.Names(R_n).RefersToRange.Cells(2, 1))
   r1 = (0 <= DateDiff("d", strD1, strD))
   r2 = (0 <= DateDiff("d", strD, strD2))
   r = r1 And r2
   p有効期間 = r
End Function

' 組織集計_非ゼロ更新 で使用している。
'
Private Sub 組織集計_非ゼロ抽出(ByRef 組織集計_非ゼロ() As Variant)
   '
   ' 『組織集計』で名付けた範囲を件数を含むように拡張し、件数が０でない
   ' 　行で構成された配列を返す
   ' ▼引数に参照で返す。
   ' ▼文字列ではなく数値として返したい場合もあるので引数は Variant とした。
   '
   Call NZrowCompaction("組織集計", 3, 組織集計_非ゼロ)
End Sub

Private Sub 取引集計_非ゼロ抽出(ByRef 取引集計_非ゼロ() As Variant)
   '
   Call NZrowCompaction("取引集計", 3, 取引集計_非ゼロ)
   '
End Sub

' 組織集計_非ゼロ抽出 で使用している。
' 取引集計_非ゼロ抽出 で使用している。
'
Private Sub NZrowCompaction(strName As String, _
                            cols As Long, _
                            ByRef NZC() As Variant)
   '
   ' 第１引数『strName』：対象となる表範囲の左上のセルに与えられた名前
   ' 第２引数   『cols』：上記の表範囲の列数
   ' 第３引数    『NZC』：返す配列
   '
   ' Dim strName As String
   ' strName = "組織集計"
   Dim R_n As Range
   Set R_n = ThisWorkbook.Names(strName).RefersToRange
   r0 = R_n.Row
   c0 = R_n.Column
   r1 = 列の最終行(strName,, 1)
   c1 = c0 + cols - 1
   ' ┗・・・年間登録件数と個別審査件数の欄まで拡張
   ' Dim strCV() As String
   ' strCV = Range(Cells(r0, c0), Cells(r1, c1)).Value
   Dim strCV() As Variant
   Set R_n = Range(Cells(r0, c0), Cells(r1, c1))
   strCV = R_n.Value
   Dim strNZ() As String
   ReDim strNZ(1 To (UBound(strCV, 1) - LBound(strCV, 1) + 1), 1 To cols)
   ReDim NZC(1 To (UBound(strCV, 1) - LBound(strCV, 1) + 1), 1 To cols)
   Dim j As Long
   Dim c As Long
   Dim z As Boolean
   j = 0
   For i = LBound(strCV, 1) To UBound(strCV, 1)
      z = True
      For c = 2 To cols
         z = z And (Val(strCV(i, LBound(strCV, 2) + (c - 1))) = 0)
      Next c
      If Not (z) Then
         j = j + 1
         For c = 1 To cols
            strNZ(j, c) = strCV(i, LBound(strCV, 2) + (c - 1))
         Next c
         ' strNZ(j, 1) = strCV(i, LBound(strCV, 2))
         ' strNZ(j, 2) = strCV(i, LBound(strCV, 2) + 1)
         ' strNZ(j, 3) = strCV(i, LBound(strCV, 2) + 2)
      End If
   Next i
   ' Dim NZC() As Variant
   ReDim NZC(1 To j, 1 To cols)
   For i = 1 To j
      NZC(i, 1) = strNZ(i, 1)
      For c = 2 To cols
         NZC(i, c) = Val(strNZ(i, c))
      Next c
      ' NZC(i, 1) = strNZ(i, 1)
      ' NZC(i, 2) = Val(strNZ(i, 2))
      ' NZC(i, 3) = Val(strNZ(i, 3))
   Next i
   ' 組織集計_非ゼロ = NZC
   '
End Sub

' 組織集計_非ゼロ更新 で使用している。
'
Private Sub a_組織集計_非ゼロ書出(ByRef 組織集計_非ゼロ() As Variant)
   '
   ' 配列『組織集計＿非ゼロ』（文字列の配列）を受け取って書き出す。
   ' 引数である配列は、その要素が文字列ではなくて数値の場合にも同様
   ' に機能してほしいことから、Variant とした。
   '
   Dim strName As String
   strName = "組織集計"
   r1 = 列の最終行(strName,, 1)
   ' ┗・・・全集計名の最後の行数を得る
   strName = "組織集計＿非ゼロ"
   Dim R組織集計_非ゼロ As Range
   Set R組織集計_非ゼロ = ThisWorkbook.Names(strName).RefersToRange
   r0 = R組織集計_非ゼロ.Row
   c0 = R組織集計_非ゼロ.Column
   c1 = c0 + 2
   ' ┗・・・年間登録件数と個別審査件数の欄まで拡張
   Set R組織集計_非ゼロ = Range(Cells(r0, c0), Cells(r1, c1))
   R組織集計_非ゼロ.Clear
   R組織集計_非ゼロ.Font.Name = "BIZ UDゴシック"
   ' ┗・・・消すのはこの領域の最大の行数
   '
   r1 = UBound(組織集計_非ゼロ, 1) - LBound(組織集計_非ゼロ, 1) + r0
   Set R組織集計_非ゼロ = Range(Cells(r0, c0), Cells(r1, c1))
   R組織集計_非ゼロ = 組織集計_非ゼロ
   ' ┗・・・書き出すのは行列の行数だけ
End Sub
' ↓
' これ修正したほうがいい。以下修正版
'
Private Sub 組織集計_非ゼロ書出(ByRef 組織集計_非ゼロ() As Variant)
   '
   ' 配列『組織集計＿非ゼロ』（文字列の配列）を受け取って書き出す。
   ' 引数である配列は、その要素が文字列ではなくて数値の場合にも同様
   ' に機能してほしいことから、Variant とした。
   '
   Call PrintNZrowCompaction("組織集計＿非ゼロ", 3, 組織集計_非ゼロ)
   '
End Sub

Private Sub 取引集計_非ゼロ書出(ByRef 取引集計_非ゼロ() As Variant)
   '
   Call PrintNZrowCompaction("取引集計＿非ゼロ", 3, 取引集計_非ゼロ, -4)
   '
End Sub

' 組織集計_非ゼロ書き出し で使用している。
' 取引集計_非ゼロ書き出し で使用している。
'
Private Sub PrintNZrowCompaction(strName As String, _
                                 cols As Long, _
                                 ByRef NZC() As Variant, _
                                 Optional ROOT As Long = 0)
   '
   Call PrintArrayOnNamedRange(strName, NZC, cols, ROOT)
   '
End sub
                                 
Private Sub Copy1Dto2CAry(ByRef iAry() As Variant, _
                        ByRef oAry() As Variant)
   ' iAry() が１次元配列であるときだけ作用する。
   ' oAry() は列数が１の２次元配列である。
End Sub

' ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   '
   ' 記入対象の配列と同じ大きさの範囲をいったん読み込む。『*』 以外
   ' の文字がある要素を、記入対象の配列の要素で置き換える。
   ' 『*』も含めて読み込み、『*』はそのままにして、配列全体を
   ' そのまま書き出してもよいかも←変更予定

   ' 配列の内容を名付けた範囲に書き込む。列ごとの書式を設定する
   ' ことができる。列ごとの書式は、各列の書き込む範囲の上方向の
   ' セル（オフセットが ROOT で指定：負の値）に予め設定しておく。
   ' 範囲の名前付けは、範囲の左上のセルを名付ける。
   ' 範囲の行数は、名付けたセルの下方向（行番号の増える方向）に
   ' 内容を持つセルの連続する範囲で拡張される。
   ' 範囲の列数は、『cols』で与える。
   ' x 『cols』が与えない（デフォルト値０が与えられた）ときは行の
   ' x 範囲内の最も長い（列番号の増える方向で内容を持つセルの連続
   ' x する）範囲で拡張される。
   ' x ┗このような列の拡張は未実装。
   ' ＜↓変更↓＞
   ' セルに記載する列数の指定である『cols』が与えられない（デフォルト
   ' 値０が与えられた）ときは配列全体（配列がもつ列数のすべて）を記載
   ' する。与えられる Ary は２次元配列なので、２次元めのインデクスの
   ' 取りうる値の種類数が cols になる。
   ' 
   ' 第１引数『strName』：書き込む範囲につけた名前
   ' 第２引数    『Ary』：書き込む内容を保持した配列
   ' 第３引数（オプション）『cols』：書き込む範囲の列数
   ' 第４引数（オプション）『ROOT(RowOffsetOfTemplate)』
   '
   ' 第２引数『Ary』が１次元配列のときは１行ではなくて１列配列
   ' として取り扱うようにしてみたが、うまく行っていない。
   ' 現状では、第２引数は、必ず２次元配列であること。
   '
Private Sub PrintArrayOnNamedRange(strName As String, _
                                   ByRef aAry() As Variant, _
                                   Optional cols As Long = 0, _
                                   Optional ROOT As Long = 0)
   Stop
   '
   Dim R_n As Range
   Set R_n = range_連続列最大行_namedrange(strName)
   Call PrintArrayOnRange(R_n, aAry, cols ,ROOT)
End Sub

Private Sub PrintArrayOnRange(R_n As Range, _
                              ByRef aAry() As Variant, _
                              Optional cols As Long = 0, _
                              Optional ROOT As Long = 0)
   ' cols > 1 のとき：
   ' 　　　┗cols で指定された列数をセルに書き出す
   ' cols = 0 のとき：
   ' 　　　┗与えられた配列の内容をすべてセルに書き出す
   ' cols = -1 のとき：
   ' 　　　┗すでに書き込まれている範囲のみセルに書き出す
   ' ==フォーマット文字操作==
   ' 名前で指定された範囲（セル）に連続列最大行の検出を適用して、書き出す
   ' 範囲を列も行も確定する。
   ' ただし『*』のセルには何もしない（『*』のままにする）
   ' 
   ' ▼１行Ｎ列の配列はＮ列１行に変換
   Dim Ary As Variant
   Ary = Ary1C_If_Ary1R(aAry)
   ' ▼R_n を『書かれているまま』から変えるか
   ' 　cols を特別に指定するか
   ' 　┏左上の角のセルをもとにして連続列の最大行領域を対象として設定する
   Set R_n = range_連続列最大行_range(R_n.Cells(1,1), 1)
   Dim rowsA As Long
   Dim rowsR As Long
   rowsA = UBound(Ary, 1) - LBound(Ary, 1) + 1
   rowsR = R_n.Rows.Count
   If rowsA < rowsR Then rowsA = rowsR
   Dim colsA As Long
   colsA = UBound(Ary, 2) - LBound(Ary, 2) + 1
   colsR = R_n.Columns.Count
   If colsA < colsR Then colsA = colsR
   '
   Dim vAry As Variant
   Dim pvAry As Boolean
   pvAry = False
   ' ▼colsで特殊な値が指定された場合の書き込む列数 cols の調整
   ' 　と書き込む範囲の取得と消去
   Select Case cols
      Case 0
         '┗Ary の大きさに合わせて範囲に書き込み
         Set R_n = R_n.Resize(rowsA, colsA)
      Case -1
         '┗すでに書き込まれている範囲の大きさに合わせて書き込み
         vAry = R_n.Value
         pVary = True
         rowsA = R_n.Rows.Count
         colsA = R_n.Columns.Count
      Case Else
         '┗colsで指定された列数
         Set R_n = R_n.Resize(rowsA, cols)
         colsA = cols
   End Select
   R_n.Clear
   ' ▼列ごとに書式を指定しながら書き込み
   Dim AryC As Variant
   ' ReDim AryC(LBound(Ary, 1) To UBound(Ary, 1), 1 To 1)
   ReDim AryC(1 To rowsA, 1 To 1)
   Dim r As Long
   Dim c As Long
   Dim nfl As String '.NumberFormatLocal
   Dim bls As Long   '.Borders.LineStyle
   ' Dim rw As Long
   Dim R_0 As Range  '書式を有するセルを示す Range
   Dim R_c As Range  '書き込む範囲の列を示す Range
   ' rw = UBound(Ary, 1) - LBound(Ary, 1) + 1
   ' ┗・・・書き込む行数（配列の行数）で
   ' Set R_n = R_n.Resize(rw, 1)
   Set R_n = R_n.Resize(rowsA, 1)
   ' ┗・・・表の更新する範囲を決める。１列のみ。
   '
'   Dim Lc As Long
'   Dim Uc As Long
'   Dim A1 As Boolean
'   On Error GoTo ARY1D
'   Lc = LBound(Ary, 2)
'   Uc = UBound(Ary, 2)
'   A1 = False
'   GoTo ARY1DEND
'ARY1D:
'   Lc = 1
'   Uc = 1
'   A1 = True
'ARY1DEND:
'
'ARYEND:
   ' For c = LBound(Ary, 2) To UBound(Ary, 2)
'   Stop
   
   ' For c = Lc To Uc
   For c = 1 To colsA
      ' ┗・・・配列の最初の列から最後の列まで
      Set R_0 = R_n.Resize(1, 1).Offset(ROOT, (c - 1))
      If ROOT < 0 Then
         nfl = R_0.NumberFormatLocal
         ' ┗・・・書き込む列の先頭行セルからみて ROOT（-1なら１つ上）の
         ' ・・・・オフセットのセルに設定された書式を持ってくる
         bls = xlLineStyleNone ' エラーのときの既定値
         On Error Resume Next
         bls = R_0.Borders.LineStyle
         On Error GoTo 0
      Else
         nfl = ""
         bls = xlLineStyleNone ' 指定されている行がないときの既定値
      End If
      Set R_c = R_n.Offset(0, (c - 1))
      R_c.Font.Name = "BIZ UDゴシック"
      R_c.NumberFormatLocal = nfl
      R_c.Borders.LineStyle = bls
      ' ┗・・・列ごとに書式を設定する
      ' For r = LBound(Ary, 1) To UBound(Ary, 1)
      For r = 1 To rowsA
         If pvAry Then
            If vAry(r, c) = "*" Then
               AryC(r, 1) = "*"
            Else
               ' 範囲よりも配列が小さいときにはエラーになるが
               ' エラーの場合のデフォルトとして空文字列を設定
               AryC(r, 1) = ""
               On Error Resume Next
               AryC(r, 1) = Ary(r, c)
               On Error GoTo 0
            End If
         Else
'         If A1 Then
'            AryC(r, 1) = Ary(r)
'         Else
            ' 範囲よりも配列が小さいときにはエラーになるが
            ' エラーの場合のデフォルトとして空文字列を設定
            AryC(r, 1) = ""
            On Error Resume Next
            AryC(r, 1) = Ary(r, c)
            On Error GoTo 0
'         End If
         End If
      Next r
      R_n.Offset(0, (c - 1)) = AryC
   Next c
End Sub

Function Ary1C_If_Ary1R(ByRef Ary As Variant) As Variant
   ' 引数が１行Ｎ列（添字が１次元）の配列だったときだけ、
   ' Ｎ行１列（添字が２次元）の配列に変換
   ' そうでないときはそのまま
   Dim LB As Long
   Dim UB As Long
   LB = LBound(Ary, 1)
   UB = UBound(Ary, 1)
   '
   On Error GoTo MAIN
   LB = LBound(Ary, 2)
   Ary1C_If_Ary1R = Ary
   Exit Function
   '
MAIN:
   On Error GoTo 0
   '
   Dim vAry() As Variant
   ReDim vAry(LB To UB, 1)
   For i = LB To UB
      vAry(i, 1) = Ary(i)
   Next i
   Ary1C_If_Ary1R = vAry
End Function

' 取引区分別個別審査件数集計　の中で、表の範囲を更新するために呼ぶ
'
Private Sub 取引区分集計個別審査件数更新(ByRef 取引区分別個別審査件数() As Long)
   ' Call 取引区分集計個別審査件数クリア
   Dim r0 As Long
   Dim r1 As Long
   Dim strName As String
   strName = "取引集計"
   Dim R取引集計 As Range
   Set R取引集計 = ThisWorkbook.Names(strName).RefersToRange
   r0 = R取引集計.Row
   r1 = 列の最終行(strName,, 1)
   Dim R取引集計件数列 As Range
   ' stop
   Set R取引集計件数列 = R取引集計.Offset(0, 1).Resize((r1 - r0 + 1), 1)
   '   ┗←・・名付けた範囲・・┛をもとに新たな範囲を指定する。
   R取引集計件数列.Clear
   R取引集計件数列.Font.Name = "BIZ UDゴシック"
   R取引集計件数列.Borders.LineStyle = xlContinuous
   R取引集計件数列 = 取引区分別個別審査件数
End Sub

' 取引区分別個別審査件数集計　の中で、表の範囲を更新するために呼ぶ
'
Private Sub 取引区分集計個別審査金額更新(ByRef 取引区分別個別審査金額() As Long)
   ' Call 取引区分集計個別審査件数クリア
   Dim r0 As Long
   Dim r1 As Long
   Dim strName As String
   strName = "取引集計"
   Dim R取引集計 As Range
   Set R取引集計 = ThisWorkbook.Names(strName).RefersToRange
   r0 = R取引集計.Row
   r1 = 列の最終行(strName,, 1)
   Dim R取引集計金額列 As Range
   ' stop
   Set R取引集計金額列 = R取引集計.Offset(0, 2).Resize((r1 - r0 + 1), 1)
   '   ┗←・・名付けた範囲・・┛をもとに新たな範囲を指定する。
   R取引集計金額列.Clear
   R取引集計金額列.Font.Name = "BIZ UDゴシック"
   R取引集計金額列.NumberFormatLocal = "#,##0,"
   ' ┗・・・金額を千円単位で表示する
   R取引集計金額列.Borders.LineStyle = xlContinuous
   R取引集計金額列 = 取引区分別個別審査金額
End Sub

' 仕向地別集計　で利用
'
Private Sub 仕向地辞書生成(strName As String, _
                           ByRef nClass As Long, _
                           ByRef dicDIN As Dictionary)
   ' 仕向地辞書生成("仕向地別名", dic仕向番号)
   ' dicDN - Distination ID Number
   Call MultiHomeDict_namedRange(strName, nClass, dicDIN)
End Sub

' 仕向地別集計　で利用
'
Private Sub 仕向地集計名配列生成(strName As String, _
                                 ByRef nClass As Long, _
                                 ByRef aryNTM As Variant)
   'ary集計名件数金額 - Name, cont of Times, amount of Money
   ReDim aryNTM(1 To nClass, 1 To 3)
   Dim aryN() As Variant
   ReDim aryN(1 To nClass, 1 To 1)
   Call NamedRange2ary(strName, aryN)
   ' Stop
   For i = LBound(aryN, 1) To UBound(aryN, 1)
      aryNTM(i, 1) = aryN(i, 1)
   Next i
   ' Stop
   '
End Sub

' ┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
' ┃┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
' ┃┃汎用プロシジャ（後にこのプロジェクト以外でも転用する見込みのあるもの）
' ┃┗

Private Sub 組織略称クリア()
   Dim 組織辞書┏ As Range
   ' Set 組織辞書┏ = range_連続列最大行_namedrange（"組織辞書┏"）
   ' 組織辞書┏.Clear
   range_連続列最大行_namedrange（"組織辞書┏"）.Clear
   Exit Sub
   '
   ' Unused
   ' 以下は、すべて、使われなくなった
   Dim strName As String
   strName = "Range_組織辞書"
   Dim Range_組織辞書 As Range
   On Error Resume Next
   ' Clear を呼びつつ Range を返して再代入、ってのが間違い。
   ' 間違いだったので、エラーが起きているのを On Error で
   ' 囲っている、という間違い・・・。なんとも。
   Set Range_組織辞書 = _
       ThisWorkbook.Names(strName).RefersToRange.Clear
   Call updateRDofNamedRange(strName, 1, 1)
   On Error GoTo 0
End Sub

Private Sub 組織略称読み取り(ByRef 組織略称CI() As Long, _
                             ByRef 組織略称ST() As String)
   ' 『組織』と名前付けた範囲（１列×複数行）を読み取り：
   ' ・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・
   ' １：第１引数として指定した配列に、範囲の ColorIndex をLong型で返す。
   ' ２：第２引数として指定した配列に、範囲の セルの値 をString型で返す。
   ' ・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・
   Call Col2CIonST("組織", 組織略称CI(), 組織略称ST())
End Sub

Private Sub Col2CIonST(strName As String, _
                       ByRef CI() As Long, _
                       ByRef ST() As String)
   ' 『組織』と名前付けした範囲を読み取り：
   ' ・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・
   ' １：第１引数の文字列で名付けられた範囲（１列×複数行）を読み取り
   ' ２：第２引数として指定した配列に、範囲の ColorIndex をLong型で返す。
   ' ３：第３引数として指定した配列に、範囲の セルの値 をString型で返す。
   ' ・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・
   '
   Dim rngC As Range
   Set rngC = ThisWorkbook.Names(strName).RefersToRange
   ' 　┗名前付きの範囲（Named Range）を配列代入可能な範囲-Range- へ変換するメソッド
   '     .RefersToRange
   ' 第１引数の文字列（たとえば『組織』）は任意のセル範囲であり、もとのシートの１行目
   ' からの範囲ではないのだが、aryC(1) に最初の行が入る。
   ' たとえばもとのシートの６行目からの範囲であれば、そこ（６行目）へのアクセスは、
   ' 『組織』の１行目にアクセスすればよい。
   Call Col2CIonSTrng(rngC, CI(), ST())
   '                  ┗与えた範囲（列）の背景色と内容をそれぞれ１列の配列に得る。
End Sub

' 名前付きの範囲を抽象化して取り扱うためには、まずは、範囲に対する名前定義
' の変更（名前の付け替え）を備えておきたい。
' そもそも、名前付きの領域を名前から得るようにしておくと、名前の付け替え
' がしやすい。
' 変数とワークシートの入出力は名前をつけたセルの範囲を使って実現する。
'
Private Sub newName2Range(rng As Range, strName As String)
   '
   ' 範囲に名前を与える。名前を手がかりにつかわない。
   ' もし名前がついていたら古い名前は消去する。
   '
   Dim nm As Name
   For Each nm In Names
      If rng.Address = nm.RefersToRange.Address Then
         nm.Delete
      End If
   Next
   rng.Parent.Parent.Names.Add _
      Name:=strName, _
      RefersToLocal:="=" & rng.Address(External:=True)
End Sub

Private Sub newName2NamedRange(orgName As String, newName As String)
   '
   ' 名前つきの範囲にあらたな名前を与える。
   ' すでに名前がついている範囲を名前で指定して、
   ' 新しい名前を与え、古い名前は消去する。
   '
   Dim aRange As Range
   On Error Resume Next ' エラーが発生しても次の行から実行.
   Set aRange = ThisWorkbook.Names(orgName).RefersToRange
   On Error GoTo 0 ' On Error Resume Next を使用して有効にしたエラー処理を無効にする.
    
   If aRange Is Nothing Then
      Debug.Print "範囲のもとの名前：" & orgName & "　が無いようです。"
      Exit Sub
   End If
   Debug.Print "範囲のもとの名前：" & orgName
   ThisWorkbook.Names(orgName).Delete
   ThisWorkbook.Names.Add Name:=newName, RefersTo:=aRange
   ' ThisWorkbook ではなくて、 aRange.Parent.Parent を使うとよりよい？
   '
End Sub

' 名前付きの範囲を抽象化して取り扱えると見通しのいい記述ができると思うので
' これも定義。
Private Sub updateRDofNamedRange(strName As String, _
                                 Nrows As Long, _
                                 Ncolumns As Long)
   '
   ' 名前つきの範囲《strName》の右下(RDend)の位置の指定（下図の■）を引数
   ' 《Nrows》と《Ncolumnss》に更新する。
   ' ┏…←□→
   ' ：　　：
   ' ↑　　↑
   ' □…←■→
   ' ↓　　↓
   ' 左上は基準点で（１，１）となる。
   ' 上図の■の位置を指定する 《Nrows》と《Ncolumnss》は、
   ' 負でない整数値であり、 0 は現在の指定を変えないものとして取り扱われる。
   '
   Dim aRange As Range
   Set aRange = ThisWorkbook.Names(strName).RefersToRange
   Debug.Print "範囲のもとの行数：" & aRange.Rows.count
   Debug.Print "範囲のもとの列数：" & aRange.Columns.count
   If Nrows = 0 Then
      If Ncolumns = 0 Then
         Exit Sub
      Else
         Set aRange = aRange.Resize(, Ncolumns)
      End If
   ElseIf Ncolumns = 0 Then
      Set aRange = aRange.Resize(Nrows)
   Else
      Set aRange = aRange.Resize(Nrows, Ncolumns)
   End If
   '
   Debug.Print "範囲の新たな行数：" & aRange.Rows.count
   Debug.Print "範囲の新たな列数：" & aRange.Columns.count
   ThisWorkbook.Names.Add Name:=strName, RefersTo:=aRange
End Sub

' 組織辞書初期化のなかで str組織辞書 を書き出すために呼ぶ
'
   '    .Resize(＜行＞,＜列＞) で書き出し範囲を変えられる┛
   '    （『シートの集計名と別名』という範囲が変わる＜ここでは拡張される＞
   '    　ので、はみ出した部分を切り落とされることなく書き出せる）
   '    ただし、これだけではシートに定義された名前も更新されたわけではない。
   '
   ' 名前をつけた範囲についても更新しておく。
   ' （もとの名前『strName』に、更新された参照範囲『シートの集計名と別名』を
   ' 　割り当てることになるので、もとの名前の定義に上書きされる＜別の名前を
   ' 　つけると、もとの名前も残ってしまう点に注意＞）
   '
Private Sub Arr2ReNamedRange(ByRef varstrArr() As Variant, _
                             strName As String)
   '
   ' <1 to UBound(strArr,1)> x <1 to UBound(strArr,2)>  の大きさを持つ
   ' 配列 strArr を
   ' strName で名付けた範囲に書き出す。
   ' 範囲の大きさは、 strArr に合わせて拡縮（再設定）される。
   ' さらに、名付けも拡縮された範囲に再設定される。
   ' Spillに類似機能を一般的に使えるプロシジャ
   '
   Dim j As Long
   Dim k As Long
   j = UBound(varstrArr, 1)
   k = UBound(varstrArr, 2)
   Dim R_n As Range
   Set R_n = _
       ThisWorkbook.Names(strName).RefersToRange.Resize(j, k)
   '    ただし、これだけではシートに定義された名前も更新されたわけではない。
   '
   R_n = varstrArr
   '
   ' 名前をつけた範囲についても更新しておく。
   ActiveWorkbook.Names.Add Name:=strName, RefersTo:=R_n
   '
End Sub

' --- 範囲をDictionaryに変換する
' 範囲を各セルの内容を key として記載された行を value とする Dictionary に変換する
' 組織名や取引区分の集計辞書で、前者であれば複数列、後者は１列の辞書変換
'
Private Sub SingleHomeDict_namedRange(strName As String, _
                                      ByRef nClass As Long, _
                                      ByRef dic辞書 As Dictionary, _
                                      Optional cx As Long = 1)
   Call MultiHomeDict_namedRange(strName, _
                                 nClass, _
                                 dic辞書, _
                                 cx, _
                                 True)
End Sub

' --- 範囲をDictionaryに変換する
' 範囲を各セルの内容を key として記載された行を value とする Dictionary に変換する
' 同じ key が複数の行に現れる場合を許し、その場合のために、value は記載された行番号
' のリストとする辞書。
' 第５引数 pS を True にすると、value は記載された行番号そのものとなる。
'
Private Sub MultiHomeDict_namedRange(strName As String, _
                                     ByRef nClass As Long, _
                                     ByRef dic辞書 As Dictionary, _
                                     Optional cx As Long = 0, _
                                     Optional pS As Boolean = False)
   ' 値の階級数を返す必要がある。
   ' ここでの辞書の作成の目的は、複数の key に同じ Value を返すしくみを
   ' 簡単に実装することなので、何種類の Value を返すことになっているのか
   ' については、辞書を構成したときにわかるものとして返すことが求められる。
   ' 第２引数として参照渡ししてもらっておいて返す。
   '
   ' 第１引数：辞書にする内容を記載してある範囲に名付けた名前（文字列）
   ' 第２引数：辞書の持つ Value のクラス数を返すための変数（参照渡し）
   ' 　　　　　別名辞書が何行の範囲で構成されているかが返される。
   ' 第３引数：生成される辞書
   ' ・・・・・┣範囲の各セルの内容→ key
   ' ・・・・・┗それが記載された行番号（範囲内での行番号）→ value
   ' 第４引数：（オプショナル）辞書範囲の列数を制限するときその列数
   ' ・・・・・制限しないときには『 0 』（1より小さい値）とする。
   ' 第１引数がセル１個だけのときには、
   ' 範囲は『列の最終行』と『複数行の最終列_range』まで拡大される。
   ' 第５引数： cx - 範囲のカラム数を固定値で指定するときその値。
   ' 　　　　　指定されないときは値 0 とみなす。最も長い行のカラム数が
   ' 　　　　　指定されたものとみなす。
   ' 第６引数： pS - SingleHomeDict として呼ぶときに True にする。
   ' 　　　　　指定されないときは値 False とみなす。複数の行番号への帰属
   ' 　　　　　がありえるものとして（かりに行番号が１つだけでも、要素が１つの）
   ' 　　　　　配列を返す辞書となる。
   '
   ' Debug.Print strName
   ' Stop
   Dim R_n As Range
   Set R_n = ThisWorkbook.Names(strName).RefersToRange
   Set R_n = range_連続列最大行_range(R_n)
   '
   Dim var辞書() As Variant
   var辞書 = R_n.Value
   nClass = R_n.Rows.count
   '
   Dim U1 As Long
   U1 = UBound(var辞書, 1)
   U2 = UBound(var辞書, 2)
   If (cx > 0) Then
      U2 = cx
   End If
   
   Dim d0 As Long
   If pS Then
      For i = 1 To U1
         d0 = i
         ' ┗ この辞書での value となるd0 は値。
         For j = 1 To U2
            d = var辞書(i, j)
            If d = "" Then Exit For
            If dic辞書.Exists(d) Then
               ' この辞書での key である d が
               ' すでに登録されているときは、
               ' つまり２回目以降は、無視する。
            Else
               dic辞書.Add d, d0
            End If
         Next j
      Next i
   Else
      For i = 1 To U1
         d0 = i
         ' ┗ この辞書での value の要素となるd0 は値。
         For j = 1 To U2
            d = var辞書(i, j)
            If d = "" Then Exit For
            Dim dv() As Variant ' 再定義が前提
            If dic辞書.Exists(d) Then
               dv = dic辞書.Item(d)
               k = UBound(dv, 1)
               ReDim Preserve dv(1 To k + 1)
               dv(k + 1) = d0
               dic辞書.Item(d) = dv
            Else
               ReDim Preserve dv(1 To 1)
               dv(0 + 1) = d0
               dic辞書.Add d, dv
            End If
            Erase dv ' つぎの回のために消す
         Next j
      Next i
   End if
End Sub

' 仕向地集計名配列生成　で利用
' 
' 名付けられた範囲の列を行方向に拡張した範囲を配列に格納
'
Private Sub NamedRange2ary(strName As String, _
                           ByRef aryV As Variant)
   ' 第１引数：範囲につけた名前　※１　
   ' 第２引数：上記の範囲を拡張して内容を格納して返す配列
   ' ※１　その名前がなかったらそれなりのエラーを返したいところ
   ' 
   Dim rngC As Range
   Set rngC = ThisWorkbook.Names(strName).RefersToRange
   ' Set rngC = 列の最終行
   ' stop
   Set rngC = range_列の最終行_range(rngC,, 1)
   aryV = rngC
End Sub

