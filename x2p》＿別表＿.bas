Option Explicit
' -*- coding:shift_jis -*-

'./x2p》＿別表＿.bas

'
' いちおう別表１については期待どおりで動く。ただし、中で使っている
' 連続列最大行　の関数が正しくない。　１　を指定しても最後まで飛んで
' しまう。つまり、E列が正解なのに、M列を指してしまう。
' range_連続列最大行_range を直接変更すると、副作用の解消によるバグ
' の発生になってしまうので、range_2_連続列最大行_range という関数を
' 別に用意して、動作を確認したいところ。
'

Private Function range_2_連続列最大行_range(R_n As Range) As Range
   '
   ' このモジュールの中だけ、まずは、Call range_連続列最大行_range を
   ' これに置き換える。他のところは随時置き換えて確認。最終的にもとの
   ' を使わなくなったところで、もとのを論理的には消す。
   '
End Function

' ┏━━
' ┃▼９
'
Sub 別表１２３への転記()
   '
   ' ダッシュボード上の範囲を指定して別表１から別表３までのそれぞれの範囲へ
   ' 転記する。
   Dim cas(1 To 2) As Long
   cas(1) = 1
   cas(2) = 3
   Call Sheet2colTableCopy("組織集計", cas(), "別表１", 1)
   cas(1) = 1
   cas(2) = 2
   Call Sheet2colTableCopy("取引集計", cas(), "別表２", 1)
   cas(1) = 5
   cas(2) = 6
   Call Sheet2colTableCopy("仕向地集計", cas(), "別表３", 1)
   '
   ' Range("表１").Cells(7,4) = Range("期間の審査件数")
   ' ┗名付けた範囲がこのスクリプトと同じシートになければこの記法は
   ' 　使えない。なので、以下、プロシジャに。
   Call SheetNamedCellCopy("包括許可適用件数", "表１", 3, 4)
   Call SheetNamedCellCopy("少額特例適用件数", "表１", 4, 4)
   Call SheetNamedCellCopy("該当国内取引件数", "表１", 5, 4)
   Call SheetNamedCellCopy("リスト規制非該当件数", "表１", 6, 4)
   Call SheetNamedCellCopy("期間の審査件数", "表１", 7, 4)
   '
End Sub

Sub SheetNamedCellCopy(strName1 As String, _
                       strName2 As String, _
                       r0 As Long, _
                       c0 As Long)
   ' strName1 と名付けたセルの内容を strName2 と名付けた範囲のセル
   ' (r0,c0)に書き込む。
   Dim r1 As Range
   Dim r2 As Range
   Set r1 = ThisWorkbook.Names(strName1).RefersToRange
   Set r2 = ThisWorkbook.Names(strName2).RefersToRange.Cells(r0, c0)
   ' r2.Value = CStr(r1.Value)
   r2.Value = r1.Value
End Sub

Sub Sheet2colTableCopy(strName1 As String, cas() As Long, _
                       strName2 As String, rOff As Long)
   '
   ' シート上での２列の表のコピー：
   ' 　strName1 で指定された範囲に２列で記載されている表を strName2 で
   ' 　指定された範囲にコピーする。配列 cas(1 To 2)で、cas(i)=j により
   ' 　指定することでコピー先の i 列目にコピー元の j 列をコピーする。
   ' ・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・
   ' ：cas - column assign
   ' Dim cas(1 To "c1") As Long
   ' 別表１の場合：cas(1)=1, cas(2)=3
   ' 別表２の場合：cas(1)=1, cas(2)=2
   ' 別表３の場合：cas(1)=5, cas(2)=6
   '
   Dim R_1 As Range
   ' Set R_1 = ThisWorkbook.Names(strName2).RefersToRange.Offset(1,0).Resize(1)
   ' ┗範囲の最初の行を渡してもOK
   Set R_1 = ThisWorkbook.Names(strName2).RefersToRange.Offset(1,0)
   ' ひょっとして、以下でも大丈夫？
   ' Set R_1 = Range(strName2).Offset(rOff,0)
   ' このシートで呼ぶときは大丈夫だろうけど他のシートで呼ぶと見えないはず。
   Debug.Print(strName2)
   Debug.Print(R_1.Address)
   Dim r1 As Long
   Dim c1 As Long
   r1 = R_1.Rows.Count
   ' c1 = R_1.Columns.Count
   c1 = 2
   Dim VT As Variant
   Dim NT As String
   NT = strName1
   Call NamedRange2Ary(NT, VT, r1, cas(c1))
   Dim L1 As Long
   Dim U1 As Long
   L1 = LBound(VT, 1)
   U1 = UBound(VT, 1)
   Dim i As Long
   Dim VP As Variant
   ReDim VP(L1 To U1, 1 To 2)
   For i = L1 To U1
      VP(i, 1) = VT(i, cas(1))
      VP(i, 2) = VT(i, cas(2))
   Next i
   ' Stop
   Call PrintArrayOnRange(VP, R_1, 0, 2)
   '
End Sub

Private Sub PrintArrayOnRange(ByRef Ary As Variant, _
                              R_n As Range, _
                              Optional nr As Long = 0, _
                              Optional nc As Long = 0, _
                              Optional BC As String ="*")
   '
   ' 配列 Ary を範囲 R_n に印刷する。
   ' nr, nc いずれについても、１以上の値が指定されていれば
   ' 範囲の行・列が明示的に指定されているものとする。
   ' nr = 0 の場合には列数を、nc = 0 の場合には行数を
   ' 連続列最大行によって拡張する。
   ' nr = -1 の場合には列数を、nc = -1 の場合には行数を
   ' 配列の列数、行数によって決める。
   '
   ' Set R_n = range_連続列最大行_range(R_n, 1)
   ' Set R_n = range_TabBottom_range(R_n, 1)
   ' Set R_n = range_n_TabBottom_range(R_n)
   Set R_n = range_n_TabWiden_range(R_n)
   ' これはまだ、行の拡張のみ。列の拡張は対応していない。
   ' Stop
   ' ここで上記で止めてみて、? R_n.address すると M 列まで
   ' 入ってしまっていることがわかる。
   Debug.Print(R_n.address)
   ' Debug.Print("上記が M 列まで含んでいるようならここでは影響しないが、バグ。")
   If nr > 0 Then Set R_n = R_n.Resize(nr)
   If nc > 0 Then Set R_n = R_n.Resize(,nc)
   Dim L1 As Long
   Dim U1 As Long
   Dim L2 As Long
   Dim U2 As Long
   L1 = LBound(Ary, 1)
   U1 = UBound(Ary, 1)
   L2 = LBound(Ary, 2)
   U2 = UBound(Ary, 2)
   On Error GoTo OUTOFRANGE
   If nr = -1 Then Set R_n = R_n.Resize(U1 - L1 + 1)
   If nc = -1 Then Set R_n = R_n.Resize(,U2 - L2 + 1)
   On Error GoTo 0
   Dim nrR As Long
   nrR = R_n.Rows.Count
   Dim ncR As Long
   ncR = R_n.Columns.Count
   Dim PAry As Variant
   Dim BAry As Variant
   ReDim PAry(1 To nrR, 1 To ncR)
   ReDim Bary(1 To nrR, 1 To ncR)
   Dim r As Long
   Dim c As Long
   For r = 1 To nrR
      For c = 1 To ncR
         PAry(r, c) = ""
         On Error Resume Next
         PAry(r, c) = Ary(L1 + r - 1, L2 + c - 1)
         On Error GoTo 0
         On Error Resume Next
         If BAry(r, c) = BC Then PAry(r, c) = BC
         On Error GoTo 0
      Next c
   Next r
   R_n = PAry
   Exit Sub
OUTOFRANGE:
   Debug.Print("PrintArrayOnRange - 引数として与えられた配列の大きさが不正です")
End Sub

Sub NamedRange2Ary(ByVal strName As String, _
                   ByRef Ary As Variant, _
                   Optional nr As Long = 0, _
                   Optional nc As Long = 0)
   '
   ' strNameで指定した範囲（の左上のセル）を左上とする nr 行 nc 列の
   ' 範囲を引数の配列に格納する。
   '
   Dim R_n As Range
   Set R_n = ThisWorkbook.Names(strName).RefersToRange
   On Error Resume Next
   ' ここは、以下だとエラー。なんでかな。
   Set R_n = Range(strName)
   ' 答え： strName がこのシートの名前ではないから。
   On Error GoTo 0
   Call Range2Ary(R_n, Ary, nr, nc)
   '
End Sub

Sub Range2Ary(R_n As Range, _
              ByRef Ary As Variant, _
              Optional nr As Long = 0, _
              Optional nc As Long = 0)
   '
   ' R_n の範囲（の左上のセル）を左上とする nr 行 nc 列の
   ' 範囲を引数の配列に格納する。
   '
   ' Ary じゃなくて Ary() と書かないとだめ？
   '
   Dim r0 As Long
   Dim c0 As Long
   r0 = R_n.Row
   c0 = R_n.Column
   Set R_n = range_連続列最大行_range(R_n, 1)
   If nr > 0 Then Set R_n = R_n.Resize(nr)
   If nc > 0 Then Set R_n = R_n.Resize(,nc)
   Ary = R_n
End Sub
