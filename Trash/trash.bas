' 使わなくなった関数を集めておく

' --- 名前付き範囲が単一セルのときには、範囲を拡大する。
' cx を１以上の値で指定すると、カラム数は cx で固定される。
' cx が１より小さいときには、カラム数は『複数行の最終列』から決まる。
Private Sub ExpandRangeCont(ByRef R_n As Range, strName As String, cx As Long)
   Dim ra As Long
   Dim ca As Long
   ra = R_n.Row
   ca = R_n.Column
   Dim rz As Long
   rz = 列の最終行(strName)
   ' Set R_n = Range(Cells(ra, ca), Cells(rz, ca))
   ' ┗これでは違うシートをみてしまう。
   ' Set R_n = R_n.Resize(rz - ra + 1, 1)
   Dim cz As Long
   If cx < 1 Then
      cz = 複数行の最終列_range(R_n)
   Else
      cz = ca + cx - 1
   End If
   ' Set R_n = Range(Cells(ra, ca), Cells(rz, cz))
   ' ┗これでは違うシートをみてしまう。
   Set R_n = R_n.Resize(rz - ra + 1, cz - ca + 1)
End Sub

' ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝

' ・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・
'

Sub テスト◆SingleHomeDict_range()
   Dim strName As String
   Dim dic辞書 As New Dictionary
   strName = "取引集計"
   ' Call SingleHomeDict_range(strName, dic辞書, 1)
   Call SingleHomeDict_range(strName, dic辞書)
   ' ┗これでも同じはず。
   Stop
   ' Erase dic辞書
   dic辞書.RemoveAll
   ' Dim dic辞書 As New Dictionary
   ' strName = "Range_組織辞書"
   strName = "左上cell_組織辞書"
   Call SingleHomeDict_range(strName, dic辞書, 0)
   Stop
End Sub

' --- 配列をDictionaryに変換する
' 上位組織を先頭として、その組織に帰属する下位組織を以降に並べた行を
' 上位組織の数だけならべた配列を
' 上位組織自身および下位組織を key とし、
' 上位組織の記載された行番号を Value とする辞書に変換する。
' 複数帰属（同じ下位組織が複数の上位組織に所属する。
' 結果、key が複数の Value を持つ）はないものとする。
'
Sub SingleHomeDict_var(ByRef var辞書() As Variant, _
                   ByRef dic辞書 As Dictionary)
   Dim U1 As Long
   U1 = UBound(str辞書, 1)
   U2 = UBound(str辞書, 2)
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

Sub テスト取引辞書構成()
   Dim var辞書() As Variant
   Dim dic辞書 As New Dictionary
   var辞書 = varNamedRange2Arr("取引集計")
   ' Call SingleHomeDict(CStr(var辞書()), dic辞書)
   Call 取引辞書構成(var辞書(), dic辞書)
End Sub

Sub 取引辞書構成(ByRef var辞書() As Variant, ByRef dic辞書 As Dictionary)
   ' Call SingleHomeDict(CStr(var辞書()), dic辞書)
   Call SingleHomeDict_var(var辞書(), dic辞書)
End Sub

Sub 集計名_組織_初期化()
   ' Call 組織略称初期化
End Sub

' この関数は実は使っていない（組織集計名は組織辞書から生成し、組織集計の
' 個別審査件数も、組織辞書に基づいて集計するので）
Function 組織集計() As Variant
   '
   ' 『組織集計』で名付けた範囲（左上セル）から　集計名　とその集計名
   ' 　の件数（のべ３列）の範囲を取り出して配列として返す
   ' 　文字列と数値が混在となるので Variant として返す
   '
   組織集計 = varNamedRange2Arr("組織集計", 3)
   ' Debug.Print ("Done")
End Function

Function 取引集計() As Variant
   '
   ' 『取引集計』で名付けた範囲（左上セル）から　集計名　とその集計名
   ' 　の件数（のべ３列）の範囲を取り出して配列として返す
   ' 　文字列と数値が混在となるので Variant として返す
   '
   取引集計 = varNamedRange2Arr("取引集計", 3)
   ' Debug.Print ("Done")
End Function

Function varNamedRange2Arr(strName As String, _
                           Optional nC As Long = 1) As Variant
   '
   ' 第１引数 strName で名付けた範囲を『列の最終行』で拡張して、オプショナルの
   ' 第２引数 nC の列数（指定しなければ１列のみ）の範囲を配列として返す
   '
   Dim R_n As Range
   Set R_n = ThisWorkbook.Names(strName).RefersToRange
   Dim ra As Long
   Dim ca As Long
   Dim rz As Long
   Dim cz As Long
   ra = R_n.Row
   ca = R_n.Column
   rz = 列の最終行(strName)
   cz = ca + nC - 1
   ' Set R_n = Range(Cells(ra, ca), Cells(rz, cz))
   ' ┣・・・基準の範囲からの相対で指定するとシート間違いがない。
   Set R_n = R_n.Resize((rz - ra + 1), (cz - ca + 1))
   Dim varArr() As Variant
   varArr = R_n.Value
   varNamedRange2Arr = varArr()
   ' ┗・・・配列を関数の返す値にするときには『()』が必要
   '
End Function

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
   nC = 組織.Columns.count
   Debug.Print "組織.Rows.count：" & nr
   Debug.Print "組織.Columns.count：" & nC
   ReDim 組織略称_S(nr, nC)
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


