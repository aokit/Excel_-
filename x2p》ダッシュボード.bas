' -*- coding:shift_jis -*-

'./x2p�t�_�b�V���{�[�h.bas

Sub ���O�̒�`�m�F�̐���()
   Call �J�n���}��
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
   Call �W�v�󋵂̊e�͈̖͂��O��`
   Call �I�������
End Sub

Private Sub �W�v�󋵂̊e�͈̖͂��O��`()
   Call newName2Range(Range("B18"), Range("A18").Value)
   Call newName2Range(Range("B19"), Range("A19").Value)
   Call newName2Range(Range("B20"), Range("A20").Value)
   Call newName2Range(Range("B21"), Range("A21").Value)
   Call newName2Range(Range("B22"), Range("A22").Value)
   Call newName2Range(Range("B23"), Range("A23").Value)
End Sub

' ���O�������͈͂ɂ��炽�Ȗ��O������
' �i�܂����O�����Ă��Ȃ���΂͂��߂Ė��O������j
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

Sub �W�v��_�g�D_������()
   Call �g�D���̏�����
End Sub

Private Sub �g�D���̏�����()
   Dim cc() As Long
   Dim cs() As String
   Dim �g�D���� As Variant
   Dim �g�D As Range
   Set �g�D = ThisWorkbook.Names("�g�D").RefersToRange
   ' ���O�t���͈̔́iNamed Range�j����z��ցF.RefersToRange
   �g�D���� = �g�D
   Dim n As Long
   Dim m As Long
   m = UBound(�g�D����)
   Debug.Print m
   Debug.Print LBound(�g�D����)
   n = 0
   ' ReDim cc(n)
   ' ReDim cs(n)
   Dim i As Long, j As Long, k As Long
   i = 0
   j = m ' �ŏ��l�� m �Ƃ������Ƃ͖����͂��Ȃ̂ŃV�[�h�ɂ���B
   k = 0
   For Each c In �g�D
      ' Debug.Print c.Row + 0 & ":" & �Z���̌Œ�F(c) & ":" & c.Value
      ' Debug.Print n > c.Row
      i = c.Row
      If i < j Then j = i
      If i > k Then k = i
      If n < i Then
         n = n + m
         ReDim cc(n)
         ReDim cs(n)
      End If
      cc(i) = �Z���̌Œ�F(c)
      cs(i) = c.Value
   Next c
   Debug.Print j
   Debug.Print k
End Sub

Function �J�����̍ŏI�s(n As String, k) As Long
   ' n - �͈͂ɗ^�������O�i������j
   ' k - ���͈̔͂̒��̃J�����ԍ�
   Dim r1 As Long
   Dim r2 As Long
   Dim mr As Long
   Dim s As Variant
   mr = Rows.count
   Set s = Range(n).Columns(k).End(xlDown)
   r1 = s.Row
   If r1 = mr Then
      �J�����̍ŏI�s = 0
      Exit Function
   End If
   Do While Not (r1 = mr)
      ' Debug.Print s.Value
      r2 = r1
      Set s = s.End(xlDown)
      r1 = s.Row
   Loop
   �J�����̍ŏI�s = r2
End Function

