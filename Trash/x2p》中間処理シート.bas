' -*- coding:shift_jis -*-

'./x2p》中間処理シート.bas

Sub マクロ実行欄取得()
  Application.Calculation = xlCalculationManual
  Application.ScreenUpdating = False
  
  On Error GoTo Setting_Error
  D_start = Me.Range("期間＿開始").Value
  D_end = Me.Range("期間＿終了").Value
  i1_col = Me.Range("BU抽出").Value
  i2_col = Me.Range("区分抽出").Value
  i3_col = Me.Range("仕向地抽出").Value
  i4_col = Me.Range("包括抽出").Value
  i5_col = Me.Range("特例抽出").Value
  i6_col = Me.Range("非許可特例抽出").Value
  tmp = Me.Range("期間の審査数").Value
  ' 上記の領域が定義されていないと、エラーが生じる
  On Error GoTo 0
  
  ActiveWorkbook.Sheets("yushutsu_kobetsu").Activate
  n1_col = 15 ' BU
  n2_col = 8  ' 取引区分
  n3_col = 11 ' 仕向地
  
  A_yk = ActiveSheet.UsedRange.Value
  Dim A_B1() As String ' BU
  Dim A_B2() As String ' 取引区分
  Dim A_B3() As String ' 仕向地
  Dim A_B4() As String ' 包括
  Dim A_B5() As String ' 特例
  Dim A_B6() As String ' 該当だが許可特例適用なし
  ReDim A_B1(UBound(A_yk, 1))
  ReDim A_B2(UBound(A_yk, 1))
  ReDim A_B3(UBound(A_yk, 1))
  ReDim A_B4(UBound(A_yk, 1))
  ReDim A_B5(UBound(A_yk, 1))
  ReDim A_B6(UBound(A_yk, 1))
  
  Debug.Print A_yk(2, n1_col)
  Debug.Print D_start
  Debug.Print D_end
  
  j = 0
  k = 0
  m = 0
  n = 0
  For i = 2 To UBound(A_yk, 1)
    ' 許可特例が当てられていれば True にする。
    q = False
    If (D_start <= A_yk(i, 2)) And (A_yk(i, 2) <= D_end) Then
      A_B1(j) = A_yk(i, n1_col)
      A_B2(j) = A_yk(i, n2_col)
      A_B3(j) = A_yk(i, n3_col)
      j = j + 1
      ' ReDim A_B(j)
      ' 包括の取り扱いを抽出
      If (A_yk(i, 18) = "包括許可適用") Or (A_yk(i, 20) = "包括許可適用") Then
        A_B4(k) = A_yk(i, 1)
        Debug.Print "包括許可適用" & A_yk(i, 1)
        k = k + 1
        q = True
      End If
      ' 特例の取り扱いを抽出
      p_rr2 = (Right(A_yk(i, 18), 2) = "特例")
      p_tr2 = (Right(A_yk(i, 20), 2) = "特例")
      If p_rr2 Or p_tr2 Then
        A_B5(m) = A_yk(i, 1)
        Debug.Print "・・特例" & A_yk(i, 1)
        m = m + 1
        q = True
      End If
      ' 該当だが許可特例適用なしを抽出
      p_ql2 = (Left(A_yk(i, 17), 2) = "該当")
      p_sl2 = (Left(A_yk(i, 19), 2) = "該当")
      If (p_ql2 Or p_sl2) And (Not q) Then
        A_B6(n) = A_yk(i, 1)
        Debug.Print "・・・・国内" & A_yk(i, 1)
        n = n + 1
      End If
    End If
  Next i
  ReDim Preserve A_B1(j - 1)
  ReDim Preserve A_B2(j - 1)
  ReDim Preserve A_B3(j - 1)
  ReDim Preserve A_B4(k - 1)
  ReDim Preserve A_B5(m - 1)
  ReDim Preserve A_B6(n - 1)
  
  Debug.Print (j - 1)
  Debug.Print A_B1(1)
  Debug.Print A_B3(j - 1)
  
  Me.Activate
  i1_col = Me.Range("BU抽出").Value
  i2_col = Me.Range("区分抽出").Value
  i3_col = Me.Range("仕向地抽出").Value
  i4_col = Me.Range("包括抽出").Value
  i5_col = Me.Range("特例抽出").Value
  i6_col = Me.Range("非許可特例抽出").Value
  Me.Range("期間の審査数").Value = j - 1 + 1 ' 0 (j-1)なので
  
  i_row = Me.UsedRange.Columns(i1_col).Rows.count
  Me.Range(Cells(2, i1_col), Cells(2 + i_row - 1, i1_col)).Clear
  Me.Range(Cells(2, i1_col), Cells(2 + UBound(A_B1, 1), i1_col)) = WorksheetFunction.Transpose(A_B1)
  i_row = Me.UsedRange.Columns(i2_col).Rows.count
  Me.Range(Cells(2, i2_col), Cells(2 + i_row - 1, i2_col)).Clear
  Me.Range(Cells(2, i2_col), Cells(2 + UBound(A_B2, 1), i2_col)) = WorksheetFunction.Transpose(A_B2)
  i_row = Me.UsedRange.Columns(i3_col).Rows.count
  Me.Range(Cells(2, i3_col), Cells(2 + i_row - 1, i3_col)).Clear
  Me.Range(Cells(2, i3_col), Cells(2 + UBound(A_B3, 1), i3_col)) = WorksheetFunction.Transpose(A_B3)
  i_row = Me.UsedRange.Columns(i4_col).Rows.count
  Me.Range(Cells(2, i4_col), Cells(2 + i_row - 1, i4_col)).Clear
  Me.Range(Cells(2, i4_col), Cells(2 + UBound(A_B4, 1), i4_col)) = WorksheetFunction.Transpose(A_B4)
  i_row = Me.UsedRange.Columns(i5_col).Rows.count
  Me.Range(Cells(2, i5_col), Cells(2 + i_row - 1, i5_col)).Clear
  Me.Range(Cells(2, i5_col), Cells(2 + UBound(A_B5, 1), i5_col)) = WorksheetFunction.Transpose(A_B5)
  i_row = Me.UsedRange.Columns(i6_col).Rows.count
  Me.Range(Cells(2, i6_col), Cells(2 + i_row - 1, i6_col)).Clear
  Me.Range(Cells(2, i6_col), Cells(2 + UBound(A_B6, 1), i6_col)) = WorksheetFunction.Transpose(A_B6)

  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic

  Exit Sub

Setting_Error:
  Debug.Print "シートの名前定義に不足があります。"
  MsgBox "シートの名前定義に不足があります。"
  
End Sub


' Private
Sub belong2whom(組織定義の領域 As Range, 参照領域 As Range, 所属領域 As Range）
   ' 組織名変換のためのサブルーチンを作ることにした。
   ' ＝＝＝＝＝＝＝＝
   ' belong2whom(組織定義の領域,参照領域,所属領域）
   ' ＝＝＝＝
   ' 参照領域にある文字列を、組織定義の領域で検索し、帰属領域に結果を書き込む。
   ' 
   ' 《組織定義の領域》
   ' 下位組織名Bが帰属する上位組織名A（さらには、上位組織名Aが保有する下位組織名B）
   ' を以下の正規表現で列に配置した領域。上位組織名Aと同じ組織名を下位組織名Bとして
   ' 有する場合があるが明記しない。また、下位組織名Bの無い上位組織名Aもありうる。
   ' (A(B*))+
   ' 上位組織Aと下位組織Bはセルの背景色によって識別される。
   ' 領域の第１行の背景色によって、上位組織が定義される。
   ' 
   ' 《参照領域》
   ' 帰属する上位組織名Aを得たい下位組織名Bを列に配置したもの
   ' 
   ' 《帰属領域》
   ' 帰属することがわかった上位組織名Aを書き込む領域
   ' 
   ' なお、参照領域と所属領域は同じ行数の１列領域であること。
   ' ＝＝＝＝＝＝＝＝
   ' 
End sub