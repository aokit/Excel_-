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
   rz = 列の最終行_range(Cells(r0, c1), , 3) ' 最初の空白行の手前の行
   ' ┗・・・なぜかここ、パラメータ（上記の３）を増やさないといけなかった。
   ' 　　　　どうしてかは未解明
   ' stop
   For r = r0 To rz
      Cells(r, c2).Value = "=isref(" & Cells(r, c1).Value & ")"
   Next r
   ' 表示のための名前付け範囲生成：
   ' その下の空白につづいて状況表示用のセルとその名前を配置する。
   r0 = 列の最終行_range(Cells(rz, c1), , 4)
   rz = 列の最終行_range(Cells(rz, c1), , 5)
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
   ' 名前『Range_組織辞書』で定義した範囲を組織の辞書として使う
   ' ・組織集計の第１列、集計名を初期化するとき
   ' ・同第３列、件数を集計するとき
   '
   ' なお初期化で生成される組織辞書は初期値であって、書き換えることが想定されて
   ' いる。ただし、直接セルを書き換えることは想定しない。Range_組織辞書（範囲）
   ' の更新、途中に空白がない、多重帰属がない、などの整合性を維持したり、操作の
   ' 利便性（２セルの選択とボタンクリックで登録）のために別途プロシジャを用意す
   ' る予定。
   '
   Call 組織略称クリア
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
   ' ３┏・・・名付けした範囲に配列を書き出し範囲を広げて名付けを更新する
   Dim strName As String
   strName = "Range_組織辞書"
   Call Arr2ReNamedRange(str組織辞書(), strName)
   '
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
   '
   Dim strName As String
   strName = "Range_組織辞書"
   Dim Range_組織辞書 As Range
   Set Range_組織辞書 = _
       ThisWorkbook.Names(strName).RefersToRange
   Dim str組織辞書() As Variant
   str組織辞書 = Range_組織辞書.Value
   Dim j As Long
   j = UBound(str組織辞書, 1)
   strName = "組織集計"
   Dim Range_組織集計1列 As Range
   Set Range_組織集計1列 = _
       ThisWorkbook.Names(strName).RefersToRange.Resize(j, 1)
   Range_組織集計1列.Clear
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
      Dim str組織辞書() As String
      Call 組織辞書読み取り(str組織辞書)
      Call 組織辞書構成(str組織辞書, dic組織辞書)
      U1 = UBound(str組織辞書, 1)
   Else
      ' 以下のに置き換えてみる
      Call SingleHomeDict_namedRange("Range_組織辞書", U1, dic組織辞書, 0)
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
End Sub

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
   Call 仕向地辞書生成("仕向地別名", dic仕向番号)
   Stop
   ' ┗・・・ここで　Debug.Print(dic仕向番号("イタリヤ")(1))　とすれば
   ' 　　　　dic仕向番号に読み込まれていることがわかる
   ' Call 仕向地集計名配列生成(ary集計名件数金額)
   ' Call 仕向地集計("承認記録", 11, 10, dic仕向番号, ary集計名件数金額)
   ' Call 仕向地別名回数表示("仕向地別名回数")
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
   Dim strName As String
   strName = "組織集計"
   Dim R組織集計 As Range
   Set R組織集計 = ThisWorkbook.Names(strName).RefersToRange
   r0 = R組織集計.Row
   r1 = 列の最終行(strName)
   ' c0 = R組織集計.Column + 2
   Dim R組織集計件数列 As Range
   ' stop
   ' Set R組織集計件数列 = Range(Cells(r0, c0), Cells(r1, c0))
   ' ┣・・・名付けた範囲をもとに新たな範囲を指定する。
   Set R組織集計件数列 = R組織集計.Offset(0, 2).Resize((r1 - r0 + 1), 1)
   R組織集計件数列.Clear
   R組織集計件数列.Font.Name = "BIZ UDゴシック"
   R組織集計件数列 = 組織別個別審査件数
End Sub

' 組織別個別審査件数集計　の中で、承認記録を読み取るために呼び出す
'
Private Sub 承認記録読み取り(str_承認記録() As String)
   '
   ' ・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・
   ' 『承認記録』と名前付けした範囲を読み取り：
   ' 　引数として指定した配列に、範囲の セルの値 をString型で返す。
   ' ・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・
   '
   ' 領域の大きさを確認（mr, mc）
   Dim mr As Long
   Dim mc As Long
   mr = 列の最終行("承認記録")
   mc = 行の最終列("承認記録")
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
   '
   ' ・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・
   ' 『Range_組織辞書』と名前付けした範囲を読み取り：
   ' 　引数として指定した配列に、範囲の セルの値 をString型で返す。
   ' ・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・
   '
   ' 組織辞書のシートは手作業での追記も想定されている。そのため
   ' 領域の大きさを確認（wr, wc）する必要があるので確認。
   Dim strName As String
   strName = "Range_組織辞書"
   Dim r0 As Long
   Dim c0 As Long
   Dim rz As Long
   Dim cz As Long
   Dim wr As Long
   Dim wc As Long
   Dim R組織辞書 As Range
   Set R組織辞書 = ThisWorkbook.Names(strName).RefersToRange
   r0 = R組織辞書.Row
   c0 = R組織辞書.Column
   rz = 列の最終行(strName)
   cz = 行の最終列(strName)
   wr = rz - r0 + 1
   wc = cz - c0 + 1
   Set R組織辞書 = R組織辞書.Resize(wr, wc)
   Dim V組織辞書() As Variant
   V組織辞書 = R組織辞書
   ' 引数として返す配列の大きさをここで設定
   ReDim str組織辞書(1 To wr, 1 To wc)
   For i = 1 To wr
      For j = 1 To wc
         str組織辞書(i, j) = V組織辞書(i, j)
      Next j
   Next i
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
   r1 = 列の最終行(strName)
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
   r1 = 列の最終行(strName)
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
                                 
Private Sub PrintArrayOnNamedRange(strName As String, _
                                   ByRef Ary() As Variant, _
                                   cols As Long, _
                                   Optional ROOT As Long = 0)
   '
   ' 配列の内容を名付けた範囲に書き込む。列ごとの書式を設定する
   ' ことができる。列ごとの書式は、各列の書き込む範囲の上方向の
   ' セル（オフセットが ROOT で指定：負の値）に予め設定しておく。
   ' 範囲の名前付けは、範囲の左上のセルを名付ける。
   ' 範囲の行数は、名付けたセルの下方向（行番号の増える方向）に
   ' 内容を持つセルの連続する範囲で拡張される。
   ' 範囲の列数は、『cols』で与える。
   ' 『cols』が与えない（デフォルト値０が与えられた）ときは行の
   ' 範囲内の最も長い（列番号の増える方向で内容を持つセルの連続
   ' する）範囲で拡張される。
   ' 列の拡張は未実装。
   ' 
   ' 第１引数『strName』：書き込む範囲につけた名前
   ' 第２引数    『Ary』：書き込む内容を保持した配列
   ' 第３引数（オプション）『cols』：書き込む範囲の列数
   ' 第４引数（オプション）『ROOT(RowOffsetOfTemplate)』
   '
   ' ▼範囲を拡張して範囲内にあるセルの内容を消去
   Dim c0 As Long
   Dim c1 As Long
   Dim r0 As Long
   Dim r1 As Long
   Dim R_n As Range
   Set R_n = ThisWorkbook.Names(strName).RefersToRange
   r0 = R_n.Row
   c0 = R_n.Column
   c1 = c0 + cols - 1
   ' ┗・・・列数が cols になるように最終カラム番号を拡張
   r1 = 列の最終行(strName)
   Set R_n = R_n.Resize((r1 - r0 + 1), (c1 - c0 + 1))
   R_n.Clear
   ' ┗・・・消すのはこの領域の最大の行数
   ' ▼列ごとに書式を指定しながら書き込み
   Dim AryR As Variant
   ReDim AryR(LBound(Ary, 1) To UBound(Ary, 1), 1 To 1)
   Dim r As Long
   Dim c As Long
   Dim nfl As String '.NumberFormatLocal
   Dim bls As Long   '.Borders.LineStyle
   Dim rw As Long
   Dim R_0 As Range  '書式のみもつセルを示す Range
   Dim R_c As Range  '書き込む範囲の列を示す Range
   rw = UBound(Ary, 1) - LBound(Ary, 1) + 1
   ' ┗・・・書き込む行数（配列の行数）で
   Set R_n = R_n.Resize(rw, 1)
   ' ┗・・・表の更新する範囲を決める。１列のみ。
   For c = LBound(Ary, 2) To UBound(Ary, 2)
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
      For r = LBound(Ary, 1) To UBound(Ary, 1)
         AryR(r, 1) = Ary(r, c)
      Next r
      R_n.Offset(0, (c - 1)) = AryR
   Next c
End Sub

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
   r1 = 列の最終行(strName)
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
   r1 = 列の最終行(strName)
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

Private Sub 仕向地辞書生成(strName As String, _
                           ByRef dicDIN As Dictionary)
   ' 仕向地辞書生成("仕向地別名", dic仕向番号)
   ' dicDN - Distination ID Number
   Dim nClass As Long
   nClass = 0
   Call MultiHomeDict_namedRange(strName, nClass, dicDIN)
End Sub

' ┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
' ┃┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
' ┃┃汎用プロシジャ（後にこのプロジェクト以外でも転用する見込みのあるもの）
' ┃┗

Private Sub 組織略称クリア()
   Dim strName As String
   strName = "Range_組織辞書"
   Dim Range_組織辞書 As Range
   On Error Resume Next
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

' --- 範囲をDictionaryに変換する
' 範囲を各セルの内容を key として記載された行を value とする Dictionary に変換する
'
Private Sub x_SingleHomeDict_namedRange(strName As String, _
                                      ByRef nClass As Long, _
                                      ByRef dic辞書 As Dictionary, _
                                      Optional cx As Long = 1)
   ' 値の階級数を返す必要がある。
   ' ここでの辞書の作成の目的は、複数の key に同じ Value を返すしくみを
   ' 簡単に実装することなので、何種類の Value を返すことになっているのか
   ' については、辞書を構成したときにわかるものとして返すことが求められる。
   ' 第２引数として参照渡ししてもらっておいて返す。
   '
   ' 第１引数：辞書にする内容を記載してある範囲に名付けた名前（文字列）
   ' 第２引数：辞書の持つ Value のクラス数を返すための変数（参照渡し）
   ' 第３引数：生成される辞書
   ' ・・・・・┣範囲の各セルの内容→ key
   ' ・・・・・┗それが記載された行番号（範囲内での行番号）→ value
   ' 第４引数：（オプショナル）辞書範囲の列数を制限するときその列数
   ' ・・・・・制限しないときには『 0 』（1より小さい値）とする。
   ' 第１引数がセル１個だけのときには、
   ' 範囲は『列の最終行』と『複数行の最終列_range』まで拡大される。
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
   
   For i = 1 To U1
      d = var辞書(i, 1)
      d0 = i
      
      dic辞書.Add d, d0

      For j = 2 To U2
         d = var辞書(i, j)
         If d = "" Then Exit For
         If dic辞書.Exists(d) Then
         Else
            dic辞書.Add d, d0
         End If
      Next j
   Next i
End Sub


