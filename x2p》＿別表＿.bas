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
   '
   Dim VT As Variant
   Dim NT As String
   NT = "組織集計"
   Call NamedRange2Ary(NT, VT,, 3)
   Stop
   ' このあと、VTから印刷用の配列に転記して、印刷用の配列を
   ' PrintArrayOnRangeで印刷する。
   Dim L1 As Long
   Dim U1 As Long
   L1 = LBound(VT, 1)
   U1 = UBound(VT, 1)
   Dim i As Long
   Dim VP As Variant
   ReDim VP(L1 To U1, 1 To 2)
   For i = L1 To U1
      VP(i, 1) = VT(i, 1)
      VP(i, 2) = VT(i, 3)
   Next i
   Stop
   Dim R_1 As Range
   Set R_1 = ThisWorkbook.Names("別表１").RefersToRange.Cells(1,1).Offset(1,0)
   Call PrintArrayOnRange(VP, R_1, 0, 2)
   Stop
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
   ' 連続列最大行によって決める。
   ' nr = -1 の場合には列数を、nc = -1 の場合には行数を
   ' 配列の列数、行数によって決める。
   '
   Set R_n = range_連続列最大行_range(R_n, 1)
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
