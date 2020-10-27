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

Sub unused()
   ' 範囲の名前付けが行われているかどうか
   Range("B6").Value = "=isref(" & Range("A6").Value & ")"
   Range("B7").Value = "=isref(" & Range("A7").Value & ")"
   Range("B8").Value = "=isref(" & Range("A8").Value & ")"
   Range("B9").Value = "=isref(" & Range("A9").Value & ")"
   Range("B10").Value = "=isref(" & Range("A10").Value & ")"
   Range("B11").Value = "=isref(" & Range("A11").Value & ")"
   Range("B12").Value = "=isref(" & Range("A12").Value & ")"
   Range("B13").Value = "=isref(" & Range("A13").Value & ")"
   Range("B14").Value = "=isref(" & Range("A14").Value & ")"
   Range("B15").Value = "=isref(" & Range("A15").Value & ")"
   Range("B16").Value = "=isref(" & Range("A16").Value & ")"
   ' 集計状況の各範囲の名前定義
   Call newName2Range(Range("B20"), Range("A20").Value)
   Call newName2Range(Range("B21"), Range("A21").Value)
   Call newName2Range(Range("B22"), Range("A22").Value)
   Call newName2Range(Range("B23"), Range("A23").Value)
   Call newName2Range(Range("B24"), Range("A24").Value)
   Call newName2Range(Range("B25"), Range("A25").Value)
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

Function p有効期間(strDate As String, _
                    Optional R_n As String = "集計期間") As Boolean
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

Sub テスト()
   ' Call updateRDofNamedRange("集計名と別名", 3, 3)
   ' Call updateRDofNamedRange("集計名と別名", 0, 2)
   ' Call updateRDofNamedRange("集計名と別名", 4, 0)
   ' Call updateRDofNamedRange("集計名と別名", 0, 0)
   ' Call newName2NamedRange("集計名と別名", "『集計名』と別名")
   ' Call newName2NamedRange("『集計名』と別名", "集計名と別名")
   ' Call 配列からセルへ書き出す■実験■
   Dim str_承認記録() As String
   Call 承認記録読み取り(str_承認記録)
   ' Debug.Print str_承認記録(1, 1)
   ' Debug.Print str_承認記録(240, 1)
End Sub

Sub 組織別個別審査集計()
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

' --- 配列をDictionaryに変換する
Sub 組織辞書構成(ByRef str辞書() As String, ByRef dic辞書 As Dictionary)
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

Sub 集計名_組織_初期化()
   Call 組織略称初期化
End Sub

Private Sub 組織略称初期化()
   '
   Call 組織略称クリア
   '
   Dim 組織略称CI() As Long
   Dim 組織略称ST() As String
   Call 組織略称読み取り(組織略称CI, 組織略称ST)
   Dim m As Long
   m = UBound(組織略称ST)
   For r = 1 To m
      Debug.Print 組織略称ST(r) & ":" & 組織略称CI(r)
   Next r
   Dim 組織構成() As Variant
   ReDim 組織構成(m, m)
   Dim i As Long
   Dim j As Long
   Dim k As Long
   Dim RCI As Long
   i = 0: j = 0: k = 0
   RCI = 組織略称CI(1)
   For r = 1 To m
      If 組織略称CI(r) = RCI Then
         If k < i Then k = i
         i = 1
         j = j + 1
      Else
         i = i + 1
      End If
      組織構成(j, i) = 組織略称ST(r)
   Next r
   ' 組織構成は j 行 k 列の配列ということになる。
   Dim str組織辞書() As String
   ReDim str組織辞書(1 To j, 1 To k)
   For q = 1 To k
      For p = 1 To j
         str組織辞書(p, q) = 組織構成(p, q)
      Next p
   Next q
   '
   Dim strName As String
   strName = "Range_組織辞書"
   Dim Range_組織辞書 As Range
   Set Range_組織辞書 = _
       ThisWorkbook.Names(strName).RefersToRange.Resize(j, k)
   '    .Resize(＜行＞,＜列＞) で書き出し範囲を変えられる┛
   '    （『シートの集計名と別名』という範囲が変わる＜ここでは拡張される＞
   '    　ので、はみ出した部分を切り落とされることなく書き出せる）
   '    ただし、これだけではシートに定義された名前も更新されたわけではない。
   '
   Range_組織辞書 = str組織辞書
   '
   ' 名前をつけた範囲についても更新しておく。
   ' （もとの名前『strName』に、更新された参照範囲『シートの集計名と別名』を
   ' 　割り当てることになるので、もとの名前の定義に上書きされる＜別の名前を
   ' 　つけると、もとの名前も残ってしまう点に注意＞）
   ActiveWorkbook.Names.Add Name:=strName, RefersTo:=Range_組織辞書
   '
   ' 『組織集計』で名前付けされた左上セルから、１列の範囲を生成する。
   '  Range_組織集計１列
   ' 『組織集計』
   strName = "組織集計"
   Dim Range_組織集計1列 As Range
   Set Range_組織集計1列 = _
       ThisWorkbook.Names(strName).RefersToRange.Resize(j, 1)
   Range_組織集計1列.Clear
   Range_組織集計1列.Font.Name = "BIZ UDゴシック"
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

Private Sub 組織略称読み取り(組織略称CI() As Long, 組織略称ST() As String)
   ' ・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・
   ' 『組織』と名前付けした範囲を読み取り：
   ' 　１：第１引数として指定した配列に、範囲の ColorIndex をLong型で返す。
   ' 　２：第２引数として指定した配列に、範囲の セルの値 をString型で返す。
   ' ・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・
   '
   Dim 組織略称() As Variant
   ' 　　　　　　　　┗『組織略称』は名前付き範囲から変換した 範囲-Range- を代入する。
   '                範囲なので次元は２で各次元の要素数は不明。また、要素の型も
   '                Variant としている。
   ' 動的配列として
   ' ・Q:範囲の大きさがわかれば、ReDimで明示的に指定してもよい？
   ' ・A:『範囲のカラム数』や『範囲の行数』は、.Rows.Count などで手に入るが、セルの
   ' 値を手に入れるのはかなり面倒 （組織略称_S(1, 1) = 組織.Cells(1, 1).Value） で
   ' ある。そのため推奨される方法ではない。
   ' （まだこの時点では次元も大きさも未定）としてVariant にしておくのがよい（String
   ' 　にはできない）
   Dim 組織 As Range
   Set 組織 = ThisWorkbook.Names("組織").RefersToRange
   ' 　┗名前付きの範囲（Named Range）を配列代入可能な範囲-Range- へ変換するメソッド
   '     .RefersToRange
   ' 『組織』は１行目からの範囲ではないのだが、『組織略称』の（１）に最初の行が入る。
   ' たとえばもとのシートの６行目からの範囲であれば、そこ（６行目）へのアクセスは、
   ' 『組織』の１行目にアクセスすればよい。
   組織略称 = 組織
   ' ┗組織略称(i,j) = 組織.Cells(i,j)
   ' 範囲-Range-の　組織　は１列なのだが、代入により生成される配列は１次元配列では
   ' なく、２次元配列になることに注意！！
   Dim m As Long
   m = UBound(組織略称, 1)
   Debug.Print m
   '   ┗行方向（１列のみ）の配列なので、第１の次元の上限値を求めておく。
   ' Debug.Print 組織略称.Cells(1, 1)
   ' Debug.Print 組織略称(1)
   ' Debug.Print 組織略称(1).Cells(1, 1)
   ' ┗『組織略称』は２次元配列である。これらのアクセスのしかたはすべて誤り
   Debug.Print 組織略称(1, 1)
   ' Dim b As Long
   ' b = 組織.Cells(1, 1).Row
   ' Debug.Print b
   ' ┗『組織』は 範囲-Range- なので .Cell メソッドで行と列によってアクセスする。
   ' 　また、もとの表で何行目であるか（ .Row メソッド ）、などの情報も持っている。
   ' Dim 組織略称ColorIndex(116) As Long
   ' Dim 組織略称CI() As Long
   ReDim 組織略称CI(m)
   '     ┗『組織』としてもっているセルの背景色情報を格納する配列を用意する。
   '       範囲を代入するのではないため、明示的に次元と大きさを指定しなければ
   '       ならない。そこで、動的配列として宣言したあと、組織略称（範囲としての
   '       組織から複製した２次元配列）の行数ぶんの要素を持つ１次元の配列を設定
   '       しておく。
   ' Dim 組織略称ST() As String
   ReDim 組織略称ST(m)
   '
   For r = 1 To m
      組織略称CI(r) = 組織.Cells(r, 1).Interior.ColorIndex
      組織略称ST(r) = 組織略称(r, 1)
   Next r
   '
   For r = 1 To m
      ' Debug.Print 組織略称(r, 1) & ":" & 組織略称CI(r)
      ' 組織構成(r) = 組織略称(r, 1)
      Debug.Print 組織略称ST(r) & ":" & 組織略称CI(r)
   Next r
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
         Dim strValue() As String
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
      Next j
   Next i
   MultiHomeDict = dictMH
   '
End Function

Private Sub 組織辞書読み取り(str組織辞書() As String)
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
      Debug.Print 組織略称(r, 1) & ":" & 組織略称CI(r)
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



