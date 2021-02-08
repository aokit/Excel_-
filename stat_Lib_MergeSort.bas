Attribute VB_Name = "stat_Lib_MergeSort"
' -*- coding:shift_jis-dos -*-

' https://qiita.com/HiroshiAkutsu/items/0c6b3a1c1ab54e3f9fe1

' --- 昇順の場合

Private Sub merge_sort2(ByRef arr As Variant, ByVal col As Long)
    Dim irekae As Variant
    Dim indexer As Variant
    Dim tmp1() As Variant
    Dim tmp2() As Variant
    Dim i As Long
    ReDim irekae(LBound(arr, 1) To UBound(arr, 1))
    ReDim indexer(LBound(arr, 1) To UBound(arr, 1))
    ReDim tmp1(LBound(arr, 1) To UBound(arr, 1))
    ReDim tmp2(LBound(arr, 1) To UBound(arr, 1))
    For i = LBound(arr, 1) To UBound(arr, 1) Step 2
        If i + 1 > UBound(arr, 1) Then
            irekae(i) = arr(i, col)
            indexer(i) = i
            Exit For
        End If
        If arr(i + 1, col) < arr(i, col) Then
            irekae(i) = arr(i + 1, col)
            irekae(i + 1) = arr(i, col)
            indexer(i) = i + 1
            indexer(i + 1) = i
        Else
            irekae(i) = arr(i, col)
            irekae(i + 1) = arr(i + 1, col)
            indexer(i) = i
            indexer(i + 1) = i + 1
        End If
    Next
    Dim st1 As Long
    Dim en1 As Long
    Dim st2 As Long
    Dim en2 As Long
    Dim n As Long
    i = 1
    Do While i * 2 <= UBound(arr, 1)
        i = i * 2
        n = 0
        Do While en2 + i - 1 < UBound(arr, 1)
            n = n + 1
            st1 = i * 2 * (n - 1) + LBound(arr, 1)
            en1 = i * 2 * (n - 1) + i - 1 + LBound(arr, 1)
            st2 = en1 + 1
            en2 = IIf(st2 + i - 1 >= UBound(arr, 1), UBound(arr, 1), st2 + i - 1)
            Call merge2(irekae, indexer, tmp1, tmp2, st1, en1, st2, en2)
        Loop
        en2 = 0
    Loop
    Dim ret As Variant
    ReDim ret(LBound(arr, 1) To UBound(arr, 1), LBound(arr, 2) To UBound(arr, 2))
    For i = LBound(arr, 1) To UBound(arr, 1)
        For n = LBound(arr, 2) To UBound(arr, 2)
            If IsObject(arr(indexer(i), n)) Then
                Set ret(i, n) = arr(indexer(i), n)
            Else
                ret(i, n) = arr(indexer(i), n)
            End If
        Next
    Next
    arr = ret
End Sub

Private Sub merge2(ByRef irekae As Variant, _
ByRef indexer As Variant, _
ByRef tmpArr() As Variant, _
ByRef tmpIndexer() As Variant, _
ByVal st1 As Long, _
ByVal en1 As Long, _
ByVal st2 As Long, _
ByVal en2 As Long)
    Dim j As Long
    Dim n As Long
    Dim i As Long
    For i = st1 To en2
        tmpArr(i) = irekae(i)
        tmpIndexer(i) = indexer(i)
    Next
    j = st1
    n = st2
    Do While (j < en1 + 1 Or n < en2 + 1)
        If n >= en2 + 1 Then
            irekae(j + n - st2) = tmpArr(j)
            indexer(j + n - st2) = tmpIndexer(j)
            j = j + 1
        ElseIf j < en1 + 1 And tmpArr(j) <= tmpArr(n) Then
            irekae(j + n - st2) = tmpArr(j)
            indexer(j + n - st2) = tmpIndexer(j)
            j = j + 1
        Else
            irekae(j + n - st2) = tmpArr(n)
            indexer(j + n - st2) = tmpIndexer(n)
            n = n + 1
        End If
    Loop
End Sub

' --- 降順の場合
Sub 降順ソート(ByRef Arr As Variant, ByVal Col As Long)
   Call merge_sort2_desc(Arr, Col)
End Sub

Private Sub merge_sort2_desc(ByRef Arr As Variant, ByVal Col As Long)
    Dim irekae As Variant
    Dim indexer As Variant
    Dim tmp1() As Variant
    Dim tmp2() As Variant
    Dim i As Long
    ReDim irekae(LBound(Arr, 1) To UBound(Arr, 1))
    ReDim indexer(LBound(Arr, 1) To UBound(Arr, 1))
    ReDim tmp1(LBound(Arr, 1) To UBound(Arr, 1))
    ReDim tmp2(LBound(Arr, 1) To UBound(Arr, 1))
    For i = LBound(Arr, 1) To UBound(Arr, 1) Step 2
        If i + 1 > UBound(Arr, 1) Then
            irekae(i) = Arr(i, Col)
            indexer(i) = i
            Exit For
        End If
        If Arr(i + 1, Col) > Arr(i, Col) Then
            irekae(i) = Arr(i + 1, Col)
            irekae(i + 1) = Arr(i, Col)
            indexer(i) = i + 1
            indexer(i + 1) = i
        Else
            irekae(i) = Arr(i, Col)
            irekae(i + 1) = Arr(i + 1, Col)
            indexer(i) = i
            indexer(i + 1) = i + 1
        End If
    Next
    Dim st1 As Long
    Dim en1 As Long
    Dim st2 As Long
    Dim en2 As Long
    Dim n As Long
    i = 1
    Do While i * 2 <= UBound(Arr, 1)
        i = i * 2
        n = 0
        Do While en2 + i - 1 < UBound(Arr, 1)
            n = n + 1
            st1 = i * 2 * (n - 1) + LBound(Arr, 1)
            en1 = i * 2 * (n - 1) + i - 1 + LBound(Arr, 1)
            st2 = en1 + 1
            en2 = IIf(st2 + i - 1 >= UBound(Arr, 1), UBound(Arr, 1), st2 + i - 1)
            Call merge2desc(irekae, indexer, tmp1, tmp2, st1, en1, st2, en2)
        Loop
        en2 = 0
    Loop
    Dim ret As Variant
    ReDim ret(LBound(Arr, 1) To UBound(Arr, 1), LBound(Arr, 2) To UBound(Arr, 2))
    For i = LBound(arr, 1) To UBound(arr, 1)
        For n = LBound(arr, 2) To UBound(arr, 2)
            If IsObject(arr(indexer(i), n)) Then
                Set ret(i, n) = arr(indexer(i), n)
            Else
                ret(i, n) = arr(indexer(i), n)
            End If
        Next
    Next
    Arr = ret
End Sub

Private Sub merge2desc(ByRef irekae As Variant, _
ByRef indexer As Variant, _
ByRef tmpArr() As Variant, _
ByRef tmpIndexer() As Variant, _
ByVal st1 As Long, _
ByVal en1 As Long, _
ByVal st2 As Long, _
ByVal en2 As Long)
    Dim j As Long
    Dim n As Long
    Dim i As Long
    For i = st1 To en2
        tmpArr(i) = irekae(i)
        tmpIndexer(i) = indexer(i)
    Next
    j = st1
    n = st2
    Do While (j < en1 + 1 Or n < en2 + 1)
        If n >= en2 + 1 Then
            irekae(j + n - st2) = tmpArr(j)
            indexer(j + n - st2) = tmpIndexer(j)
            j = j + 1
        ElseIf j < en1 + 1 And tmpArr(j) >= tmpArr(n) Then
            irekae(j + n - st2) = tmpArr(j)
            indexer(j + n - st2) = tmpIndexer(j)
            j = j + 1
        Else
            irekae(j + n - st2) = tmpArr(n)
            indexer(j + n - st2) = tmpIndexer(n)
            n = n + 1
        End If
    Loop
End Sub

' ---

' Dim a As Variant
' Range("A1:B1048576").Select
' a = Selection.Value
' Call merge_sort2(a, 2) 'aの二次元配列に対して、2列目をキーにして並べ替えを行う。