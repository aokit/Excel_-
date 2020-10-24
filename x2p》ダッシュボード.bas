' -*- coding:shift_jis -*-

'./x2p》ダッシュボード.bas

Sub 名前の定義確認の生成()
   Call 開始時抑制
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
   Call 集計状況の各範囲の名前定義
   Call 終了時解放
End Sub

Private Sub 集計状況の各範囲の名前定義()
   Call newName2Range(Range("B18"), Range("A18").Value)
   Call newName2Range(Range("B19"), Range("A19").Value)
   Call newName2Range(Range("B20"), Range("A20").Value)
   Call newName2Range(Range("B21"), Range("A21").Value)
   Call newName2Range(Range("B22"), Range("A22").Value)
   Call newName2Range(Range("B23"), Range("A23").Value)
End Sub

' 名前をつけた範囲にあらたな名前をつける
' （まだ名前をつけていなければはじめて名前をつける）
Private Sub newName2Range(rng As Range, strName As String)
   Dim nm As Name
   For Each nm In Names
      If rng.Address = nm.RefersToRange.Address Then
         nm.Delete
      End If
   Next
   rng.Parent.Parent.Names.Add _
      Name := strName, _
      RefersToLocal := "=" & rng.Address(External := True)
End Sub

Sub 集計名_組織_初期化()
   Call 組織略称初期化
End Sub

Private Sub 組織略称初期化_改良前()
   ' 改良前のコードはダサいので消すことにする。
   Dim cc() As Long
   Dim cs() As String
   Dim 組織略称 As Variant
   Dim 組織 As Range
   Set 組織 = ThisWorkbook.Names("組織").RefersToRange
   ' 名前付きの範囲（Named Range）から配列へ：.RefersToRange
   組織略称 = 組織
   Dim n As Long
   Dim m As Long
   m = UBound(組織略称)
   Debug.Print m
   Debug.Print LBound(組織略称)
   n = 0
   ' ReDim cc(n)
   ' ReDim cs(n)
   Dim i As Long, j As Long, k As Long
   i = 0
   j = m ' 最小値が m ということは無いはずなのでシードにする。
   k = 0
   For Each c In 組織
      ' Debug.Print c.Row + 0 & ":" & セルの固定色(c) & ":" & c.Value
      ' Debug.Print n > c.Row
      i = c.Row
      If i < j Then j = i
      If i > k Then k = i
      If n < i Then
         n = n + m
         ReDim cc(n)
         ReDim cs(n)
      End If
      cc(i) = セルの固定色(c)
      cs(i) = c.Value
   Next c
   Debug.Print j
   Debug.Print k
End Sub

Private Sub 組織略称初期化_test()
   ' Dim 組織略称() As String
   ' Dim 組織略称() As Variant
   Dim 組織略称() As Variant
   ' Dim 組織略称 As Range
   Dim 組織 As Range
   Set 組織 = ThisWorkbook.Names("組織").RefersToRange
   ' 名前付きの範囲（Named Range）から配列へ：.RefersToRange
   ' 組織略称 = 組織.Cells(1, 1)
   組織略称 = 組織
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

Private Sub 組織略称初期化()
   Dim 組織略称() As Variant
   ' 　　　　　　　　┗組織略称は名前付き範囲から変換された 範囲-Range- が代入され
   ' 　　　　　　　　　るので、動的配列（まだこの時点では次元も大きさも未定）として
   ' 　　　　　　　　　は、Variant にしておかないといけない（Stringにはできない）
   Dim 組織 As Range
   Set 組織 = ThisWorkbook.Names("組織").RefersToRange
   ' 　┗名前付きの範囲（Named Range）から配列へ：.RefersToRange
   ' 『組織』は１行目からの範囲ではないのだが、『組織略称』の（１）に最初の行が入る。
   ' たとえばもとのシートの６行目からの範囲であれば、そこ（６行目）へのアクセスは、
   ' 『組織』の１行目にアクセスすればよい。
   組織略称 = 組織
   ' ┗組織略称(i,j) = 組織.Cells(i,j)
   ' 範囲-Range-の　組織　は１列なのだが、代入により生成される配列は１次元配列では
   ' なく、２次元配列になることに注意！！
   Dim m As Long
   m = UBound(組織略称,1)
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
   Dim 組織略称CI() As Long
   ReDim 組織略称CI(m)
   '     ┗『組織』としてもっているセルの背景色情報を格納する配列を用意する。
   '       範囲を代入するのではないため、明示的に次元と大きさを指定しなければ
   '       ならない。そこで、動的配列として宣言したあと、組織略称（範囲としての
   '       組織から複製した２次元配列）の行数ぶんの要素を持つ１次元の配列を設定
   '       しておく。
   For r = 1 To m
      組織略称CI(r) = 組織.Cells(r, 1).Interior.ColorIndex
   Next r
   '
   For r = 1 To m
      Debug.Print 組織略称(r, 1) & ":" & 組織略称CI(r)
   Next r
   '
End Sub

Sub ■2次元配列再定義()
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

Function カラムの最終行(n As String, k) As Long
   ' n - 範囲に与えた名前（文字列）
   ' k - その範囲の中のカラム番号
   Dim r1 As Long
   Dim r2 As Long
   Dim mr As Long
   Dim s As Variant
   mr = Rows.count
   Set s = Range(n).Columns(k).End(xlDown)
   r1 = s.Row
   If r1 = mr Then
      カラムの最終行 = 0
      Exit Function
   End If
   Do While Not (r1 = mr)
      ' Debug.Print s.Value
      r2 = r1
      Set s = s.End(xlDown)
      r1 = s.Row
   Loop
   カラムの最終行 = r2
End Function

