' -*- coding:shift_jis -*-

'./x2p》ダッシュボード.bas

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
   Dim rZ As Long
   ' チェックリスト生成：
   ' ボタンの左下のセルから名前付けに用意した文字列がセルに格納してあるので
   ' それらの名前について、範囲が割り当てられているか表示するような式を隣の
   ' セルに与える。
   c2 = BA.TopLeftCell.Column
   c1 = c2 - 1
   r0 = BA.TopLeftCell.Row + 1
   rZ = 列の最終行_range(Cells(r0, c1), , 2) ' 最初の空白行の手前の行
   For r = r0 To rZ
      Cells(r, c2).Value = "=isref(" & Cells(r, c1).Value & ")"
   Next r
   ' 表示のための名前付け範囲生成：
   ' その下の空白につづいて状況表示用のセルとその名前を配置する。
   r0 = 列の最終行_range(Cells(rZ, c1), , 2)
   rZ = 列の最終行_range(Cells(rZ, c1), , 3)
   For r = r0 To rZ
      Call newName2Range(Cells(r, c2), Cells(r, c1).Value)
   Next r
   Call 終了時解放
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

' ・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・
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

' --- 配列をDictionaryに変換する
' 上位組織を先頭として、その組織に帰属する下位組織を以降に並べた行を
' 上位組織の数だけならべた配列を
' 下位組織を key とし、上位組織を Value とする辞書に変換する。
' 複数帰属（同じ下位組織が複数の上位組織に所属する。
' 結果、key が複数の Value を持つ）はないものとする。
' 
Sub SingleHomeDict(ByRef str辞書() As String, _
                   ByRef dic辞書 As Dictionary)
    '
    ' 『str辞書()』
    '  上位組織を先頭として、その組織に帰属する下位組織を以降に並べた行を
    '  上位組織の数だけならべた配列
    ' 『dic辞書』
    '  下位組織を key とし、上位組織を Value とする辞書
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

Sub 組織別個別審査件数集計()
   Dim str承認記録() As String
   Call 承認記録読み取り(str承認記録)
   Dim str組織辞書() As String
   Call 組織辞書読み取り(str組織辞書)
   Dim dic組織辞書 As New Dictionary
   Call 組織辞書構成(str組織辞書, dic組織辞書)
   '
   Dim U1 As Long
   Dim 組織別個別審査件数() As Long
   U1 = UBound(str組織辞書, 1)
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
      Debug.Print 組織別個別審査件数(i, 1)
   Next i
   Call 組織集計個別審査件数更新(組織別個別審査件数)
End Sub

Sub 組織辞書構成(ByRef str辞書() As String, ByRef dic辞書 As Dictionary)
   Call SingleHomeDict(str辞書(), dic辞書)
End Sub

Sub 集計名_組織_初期化()
   ' Call 組織略称初期化
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
   Call Col2CIonSTrng(rngC, CI() ,ST())
   '                  ┗与えた範囲（列）の背景色と内容をそれぞれ１列の配列に得る。
End Sub

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

Private Sub 組織略称読み取り(ByRef 組織略称CI() As Long, _
                             ByRef 組織略称ST() As String)
   ' 『組織』と名前付けた範囲（１列×複数行）を読み取り：
   ' ・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・
   ' １：第１引数として指定した配列に、範囲の ColorIndex をLong型で返す。
   ' ２：第２引数として指定した配列に、範囲の セルの値 をString型で返す。
   ' ・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・
   Call Col2CIonST("組織",組織略称CI(), 組織略称ST())
End Sub

' 組織辞書初期化のなかで str組織辞書 を生成するために呼ぶ
' 
Private Sub CIonST2Arr(ByRef CI() As Long, _
                       ByRef ST() As String, _
                       ByRef varstrArr() As Variant)
   Dim m As Long
   m = UBound(ST,1)
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
   Dim 組織略称CI() As Long
   Dim 組織略称ST() As String
   Call 組織略称読み取り(組織略称CI(), 組織略称ST())
   '
   Dim str組織辞書() As Variant
   Call CIonST2Arr(組織略称CI(), 組織略称ST(), str組織辞書())
   '
   Dim strName As String
   strName = "Range_組織辞書"
   Call Arr2ReNamedRange(str組織辞書(), strName)
   '
End Sub


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

Private Sub 組織集計個別審査件数更新(ByRef 組織別個別審査件数() As Long)
   ' 組織集計個別審査件数クリア
   Dim strName As String
   strName = "組織集計"
   Dim R組織集計 As Range
   Set R組織集計 = ThisWorkbook.Names(strName).RefersToRange
   r0 = R組織集計.Row
   r1 = 列の最終行(strName)
   c0 = R組織集計.Column + 2
   Dim R組織集計件数列 As Range
   Set R組織集計件数列 = Range(Cells(r0, c0), Cells(r1, c0))
   R組織集計件数列.Clear
   R組織集計件数列.Font.Name = "BIZ UDゴシック"
   R組織集計件数列 = 組織別個別審査件数
End Sub

Function 組織集計名() As String
   '
   ' 『組織集計』で名付けた範囲から　集計名　を取り出して配列として返す
   ' 　文字列と数値が混在しているが、ひとまず文字列として読む。
   '
   Dim strName As String
   strName = "組織集計"
   Dim R集計名1 As Range
   Set R集計名1 = ThisWorkbook.Names(strName).RefersToRange
   r0 = R集計名1.Row
   c0 = R集計名1.Column
   r1 = 列の最終行(strName)
   c1 = c0 + 2
   ' ┗・・・年間登録件数と個別審査件数の欄まで拡張
   Set R集計名1 = Range(Cells(r0, c0), Cells(r1, c1))
   Dim str集計名() As String
   str集計名 = R集計名1
   組織集計名 = str集計名()
   ' ┗・・・配列を関数の返す値にするときには『()』が必要
   '
End Function

Sub 組織集計_非ゼロ抽出(ByRef 組織集計_非ゼロ() As Variant)
   '
   ' 『組織集計』で名付けた範囲を件数を含むように拡張し、件数が０でない
   ' 　行で構成された配列を返す
   ' ▼引数に参照で返す。
   ' ▼文字列ではなく数値として返したい場合もあるので引数は Variant とした。
   '
   Dim strName As String
   strName = "組織集計"
   Dim R組織集計 As Range
   Set R組織集計 = ThisWorkbook.Names(strName).RefersToRange
   r0 = R組織集計.Row
   c0 = R組織集計.Column
   r1 = 列の最終行(strName)
   c1 = c0 + 2
   ' ┗・・・年間登録件数と個別審査件数の欄まで拡張
   ' Dim strCV() As String
   ' strCV = Range(Cells(r0, c0), Cells(r1, c1)).Value
   Dim strCV() As Variant
   Set R組織集計 = Range(Cells(r0, c0), Cells(r1, c1))
   strCV = R組織集計.Value
   Dim strNZ() As String
   ReDim strNZ(1 To (UBound(strCV, 1) - LBound(strCV, 1) + 1), 1 To 3)
   j = 0
   For i = LBound(strCV, 1) To UBound(strCV, 1)
      If (Val(strCV(i, LBound(strCV, 2) + 1)) > 0) _
            Or (Val(strCV(i, LBound(strCV, 2) + 2)) > 0) Then
         j = j + 1
         strNZ(j, 1) = strCV(i, LBound(strCV, 2))
         strNZ(j, 2) = strCV(i, LBound(strCV, 2) + 1)
         strNZ(j, 3) = strCV(i, LBound(strCV, 2) + 2)
      End If
   Next i
   Dim NZC() As Variant
   ReDim NZC(1 To j, 1 To 3)
   For i = 1 To j
      NZC(i, 1) = strNZ(i, 1)
      NZC(i, 2) = Val(strNZ(i, 2))
      NZC(i, 3) = Val(strNZ(i, 3))
   Next i
   組織集計_非ゼロ = NZC
   '
End Sub

Private Sub 組織集計_非ゼロ書出(ByRef 組織集計_非ゼロ() As Variant)
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

Sub 組織集計_非ゼロ更新()
   Dim 組織集計_非ゼロ() As Variant
   ' ┗・・・連絡用の配列は、文字列以外も渡せるように Variant としておく。
   '
   ReDim 組織集計_非ゼロ(1 To 3, 1 To 3)
   組織集計_非ゼロ(1, 1) = "AA"
   組織集計_非ゼロ(1, 2) = "AB"
   組織集計_非ゼロ(1, 3) = "AC"
   組織集計_非ゼロ(2, 1) = "BA"
   組織集計_非ゼロ(2, 2) = "BB"
   組織集計_非ゼロ(2, 3) = "BC"
   組織集計_非ゼロ(3, 1) = "CA"
   組織集計_非ゼロ(3, 2) = "CB"
   組織集計_非ゼロ(3, 3) = "CC"
   '
   ' Set 組織集計_非ゼロ = 組織集計_非ゼロ抽出()
   Call 組織集計_非ゼロ抽出(組織集計_非ゼロ)
   Call 組織集計_非ゼロ書出(組織集計_非ゼロ)
End Sub


Private Sub 配列からセルへ書き出す(strName As String, ByRef 配列() As String)
   '
   ' Dim strName As String
   ' strName = "集計名と別名"
   Dim 集計名と別名() As String
   '
   ReDim 集計名と別名(1 To 4, 1 To 4)
   '                 ┗明示的に『１』から始める。デフォルトで０から始まると
   '                 　ずれてしまう。
   '
   集計名と別名(1, 1) = "┏左上"
   集計名と別名(3, 3) = "右下┛"
   集計名と別名(4, 4) = "範囲外"
   '
   ' 『仕向地』のシートに『集計名と別名』という名前で、3x3の範囲を設定した。
   ' 　最初に設定してあっても、書き出す配列の大きさに変更しないと
   ' ・はみ出している範囲は書き出されない
   ' ・不足していると『#N/A』が書き出される
   '
   Dim シートの集計名と別名 As Range
   Set シートの集計名と別名 = _
       ThisWorkbook.Names(strName).RefersToRange.Resize(4, 4)
   '    .Resize(＜行＞,＜列＞) で書き出し範囲を変えられる┛
   '    （『シートの集計名と別名』という範囲が変わる＜ここでは拡張される＞
   '    　ので、はみ出した部分を切り落とされることなく書き出せる）
   '    ただし、これだけではシートに定義された名前も更新されたわけではない。
   '
   シートの集計名と別名 = 集計名と別名
   '
   ' 名前をつけた範囲についても更新しておく。
   ' （もとの名前『strName』に、更新された参照範囲『シートの集計名と別名』を
   ' 　割り当てることになるので、もとの名前の定義に上書きされる＜別の名前を
   ' 　つけると、もとの名前も残ってしまう点に注意＞）
   ActiveWorkbook.Names.Add Name:=strName, RefersTo:=シートの集計名と別名
   '
End Sub


Sub テストMultiHomeDict()
   Dim strD() As String
   ReDim strD(1 To 4, 1 To 4)
   strD(1, 1) = "A"
   strD(2, 1) = "B"
   strD(3, 1) = "C"
   strD(4, 1) = "D"
   '
   strD(2, 2) = "B2"
   strD(2, 3) = "23or32"
   strD(2, 4) = "B4"
   '
   strD(3, 2) = "23or32"
   strD(3, 3) = "C3"
   '
   strD(4, 2) = "D2"
   '
   Dim dictMH As New Dictionary
   ' dictMH = MultiHomeDict(strD)
   Set dictMH = MultiHomeDict(strD)
   ' 2
   Dim aryVal() As String
   Debug.Print dictMH.Item("23or32")(2)
   aryVal = dictMH.Item("23or32")
   Debug.Print UBound(aryVal, 1)
   Erase aryVal
   '
End Sub

Function MultiHomeDict(ByRef strHomeMember() As String) As Dictionary
   '
   ' 各行の第１列に Home が記載され、第２列以降にその Home に属する Member が
   ' あれば、所定の個数ぶん記載されている不定長の行をあつめた（配列としては最長
   ' の行を格納できる列数の）1..N 行 1..M 列の配列である strHomeMember を引数
   ' とし、
   ' Member を key として Home を Value とする辞書を構成して返す関数。
   ' 実際には、Home が単一のときには、ただ一つの要素を持つ配列が Value となる
   ' 配列である。つまり、ある Member が 複数の Home に記載されている場合には、
   ' その Member を Key としてアクセスすると 複数の Home を要素として持つ配列
   ' が、 Value となる。
   '
   Dim dictMH As New Dictionary
   ' Dim strValue() As String
   ' ReDim strValue(1 To UBound(strHomeMember,1)) ' Homeの種類数が最大要素
   Dim k As Long
   Dim i As Long
   Dim j As Long
   Dim c As String
   Dim d As String
   For i = 1 To UBound(strHomeMember, 1)
      d = strHomeMember(i, 1)
      For j = 1 To UBound(strHomeMember, 2)
         c = strHomeMember(i, j)
         If c = "" Then Exit For
         Dim strValue() As String ' 再定義ができるかどうか
         If dictMH.Exists(c) Then
            strValue = dictMH.Item(c) ' 大きさのわからない配列が値
            k = UBound(strValue, 1)
            ReDim Preserve strValue(1 To k + 1)
            strValue(k + 1) = d
            dictMH.Item(c) = strValue
         Else
            ReDim strValue(1 To 1)
            strValue(0 + 1) = d
            dictMH.Add c, strValue
         End If
         Erase strValue
      Next j
   Next i
   ' MultiHomeDict = dictMH
   Set MultiHomeDict = dictMH
   '
End Function

Function Extent1arr(ByRef strValue() As String, c As String) As Variant
   Dim xstrValue() As String
   Dim U As Long
   U = UBound(strValue, 1)
   ReDim xstrValue(1 To (U + 1))
   xstrValue(U + 1) = c
   Extent1arr = xstrValue
End Function


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
   Dim rZ As Long
   Dim cZ As Long
   Dim wr As Long
   Dim wc As Long
   Dim R組織辞書 As Range
   Set R組織辞書 = ThisWorkbook.Names(strName).RefersToRange
   r0 = R組織辞書.Row
   c0 = R組織辞書.Column
   rZ = 列の最終行(strName)
   cZ = 行の最終列(strName)
   wr = rZ - r0 + 1
   wc = cZ - c0 + 1
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

Private Sub 組織略称初期化_■test■()
   ' Dim 組織略称() As String
   ' Dim 組織略称() As Variant
   Dim 組織略称() As Variant
   Dim 組織略称_S() As String
   ' Dim 組織略称 As Range
   Dim 組織 As Range
   Set 組織 = ThisWorkbook.Names("組織").RefersToRange
   ' Debug.Print "UBound(組織,1)：" & UBound(組織, 1)
   ' Debug.Print "UBound(組織,2)：" & UBound(組織, 2)
   nr = 組織.Rows.count
   nc = 組織.Columns.count
   Debug.Print "組織.Rows.count：" & nr
   Debug.Print "組織.Columns.count：" & nc
   ReDim 組織略称_S(nr, nc)
   ' 名前付きの範囲（Named Range）から配列へ：.RefersToRange
   ' 組織略称 = 組織.Cells(1, 1)
   組織略称 = 組織
   組織略称_S(1, 1) = 組織.Cells(1, 1).Value
   ' 『組織』は１行目からの範囲ではないのだが、『組織略称』の（１）に最初の行が入る
   m = UBound(組織略称)
   Debug.Print m
   ' Debug.Print 組織略称.Cells(1, 1)
   ' Debug.Print 組織略称(1)
   ' Debug.Print 組織略称(1).Cells(1, 1)
   Debug.Print 組織略称(1, 1)
   Dim b As Long
   b = 組織.Cells(1, 1).Row
   Debug.Print b
   Dim g As Long
   g = UBound(組織略称, 1)
   Debug.Print g
   ' Dim 組織略称ColorIndex(116) As Long
   Dim 組織略称CI() As Long
   ReDim 組織略称CI(g)
   ' ReDim 組織略称ColorIndex(m)
   For r = 1 To g
      ' 組織略称CI(r) = 組織.Cells(b + r - 1, 1).Interior.ColorIndex
      組織略称CI(r) = 組織.Cells(r, 1).Interior.ColorIndex
   Next r
   
   ' For r = 0 To m - 1
      ' 組織略称ColorIndex(r + 1) = 組織.Cells(b + r, 1).Interior.ColorIndex
   ' Next r
   For r = 1 To m
      ' Debug.Print 組織略称(r) & ":" & 組織略称ColorIndex(r)
      ' ┗　これは２次元配列になっていないので間違い
      ' Debug.Print 組織略称(r, 1) & ":" & 組織略称CI(r)
   Next r
End Sub

Private Sub ■2次元配列再定義■実験■()
   ' 最後の次元の定義は増減いずれも変えることができるが、それ以外の次元は変えられない。
   Dim a() As Variant
   ReDim Preserve a(3, 3)
   a(2, 2) = 3
   a(3, 3) = 5
   Debug.Print a(3, 3)
   ReDim Preserve a(3, 4)
   a(3, 4) = 7
   Debug.Print a(3, 4)
   ' ReDim Preserve a(4, 4)
   ' a(4, 4) = 11
   Debug.Print a(3, 3)
   Debug.Print a(3, 4)
   ' Debug.Print a(4, 4)
   ReDim Preserve a(3, 2)
   Debug.Print a(2, 2)
End Sub

' ------END

