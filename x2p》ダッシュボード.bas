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

Private Sub newName2Range(rng As Range, strName As String)
   '
   ' ���O���͈̔͂ɂ��炽�Ȗ��O��^����
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

Private Sub NamedRangeUpdate(strName As String, rng As Range)
   '
   ' ���O���͈̔͂�V���Ȕ͈͂ɍX�V����
   '
   With ActiveWorkbook
      .Names.Add Name:=strName, RefersTo:=rng
   End With
   '
   Dim nm As Name
   For Each nm In Names
      If nm = strName Then
         nm.Parent.Parent.
      If rng.Address = nm.RefersToRange.Address Then
         nm.Delete
      End If
   Next
   rng.Parent.Parent.Names.Add _
      Name:=strName, _
      RefersToLocal:="=" & rng.Address(External:=True)
End Sub

Sub �W�v��_�g�D_������()
   Call �g�D���̏�����
End Sub

Private Sub �g�D���̏�����()
   Dim �g�D����CI() As Long
   Dim �g�D����ST() As String
   Call �g�D���̓ǂݎ��(�g�D����CI, �g�D����ST)
   Dim m As Long
   m = UBound(�g�D����ST)
   For r = 1 To m
      Debug.Print �g�D����ST(r) & ":" & �g�D����CI(r)
   Next r
End Sub

Private Sub �g�D���̍\��(�g�D����CI() As Long, �g�D����ST() As String)
   Dim �W�v���ƕʖ�() As String
   '
   Dim �V�[�g�̏W�v���ƕʖ� As Range
   Set �V�[�g�̏W�v���ƕʖ� = ThisWorkbook.Names("�W�v���ƕʖ�").RefersToRange
   '
   �V�[�g�̏W�v���ƕʖ� = �W�v���ƕʖ�
   '
End Sub

Private Sub �z�񂩂�Z���֏����o����������()
   '
   Dim strName As String
   strName = "�W�v���ƕʖ�"
   Dim �W�v���ƕʖ�() As String
   '
   ReDim �W�v���ƕʖ�(1 To 4, 1 To 4)
   '                 �������I�Ɂw�P�x����n�߂�B�f�t�H���g�łO����n�܂��
   '                 �@����Ă��܂��B
   '
   �W�v���ƕʖ�(1, 1) = "������"
   �W�v���ƕʖ�(3, 3) = "�E����"
   �W�v���ƕʖ�(4, 4) = "�͈͊O"
   '
   ' �w�d���n�x�̃V�[�g�Ɂw�W�v���ƕʖ��x�Ƃ������O�ŁA3x3�͈̔͂�ݒ肵���B
   ' �@�ŏ��ɐݒ肵�Ă����Ă��A�����o���z��̑傫���ɕύX���Ȃ���
   ' �E�͂ݏo���Ă���͈͂͏����o����Ȃ�
   ' �E�s�����Ă���Ɓw#N/A�x�������o�����
   ' 
   Dim �V�[�g�̏W�v���ƕʖ� As Range
   Set �V�[�g�̏W�v���ƕʖ� = _
       ThisWorkbook.Names(strName).RefersToRange.Resize(4,4)
   '    .Resize(���s��,����) �ŏ����o���͈͂�ς����鄮
   '
   �V�[�g�̏W�v���ƕʖ� = �W�v���ƕʖ�
   '
   With ActiveWorkbook
      .Names.Add Name:=strName, RefersTo:=�V�[�g�̏W�v���ƕʖ�
   End With

   '
End Sub

Private Sub �g�D���̓ǂݎ��(�g�D����CI() As Long, �g�D����ST() As String)
   ' �E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E
   ' �w�͈́x�Ɩ��O�t�������͈͂�ǂݎ��F
   ' �@�P�F��P�����Ƃ��Ďw�肵���z��ɁA�͈͂� ColorIndex ��Long�^�ŕԂ��B
   ' �@�Q�F��Q�����Ƃ��Ďw�肵���z��ɁA�͈͂� �Z���̒l ��String�^�ŕԂ��B
   ' �E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E
   '
   Dim �g�D����() As Variant
   ' �@�@�@�@�@�@�@�@���w�g�D���́x�͖��O�t���͈͂���ϊ����� �͈�-Range- ��������B
   '                �͈͂Ȃ̂Ŏ����͂Q�Ŋe�����̗v�f���͕s���B�܂��A�v�f�̌^��
   '                Variant �Ƃ��Ă���B
   ' ���I�z��Ƃ���
   ' �EQ:�͈͂̑傫�����킩��΁AReDim�Ŗ����I�Ɏw�肵�Ă��悢�H
   ' �EA:�w�͈͂̃J�������x��w�͈͂̍s���x�́A.Rows.Count �ȂǂŎ�ɓ��邪�A�Z����
   ' �l����ɓ����̂͂��Ȃ�ʓ| �i�g�D����_S(1, 1) = �g�D.Cells(1, 1).Value�j ��
   ' ����B���̂��ߐ����������@�ł͂Ȃ��B
   ' �i�܂����̎��_�ł͎������傫��������j�Ƃ���Variant �ɂ��Ă����̂��悢�iString
   ' �@�ɂ͂ł��Ȃ��j
   Dim �g�D As Range
   Set �g�D = ThisWorkbook.Names("�g�D").RefersToRange
   ' �@�����O�t���͈̔́iNamed Range�j��z�����\�Ȕ͈�-Range- �֕ϊ����郁�\�b�h
   '     .RefersToRange
   ' �w�g�D�x�͂P�s�ڂ���͈̔͂ł͂Ȃ��̂����A�w�g�D���́x�́i�P�j�ɍŏ��̍s������B
   ' ���Ƃ��΂��Ƃ̃V�[�g�̂U�s�ڂ���͈̔͂ł���΁A�����i�U�s�ځj�ւ̃A�N�Z�X�́A
   ' �w�g�D�x�̂P�s�ڂɃA�N�Z�X����΂悢�B
   �g�D���� = �g�D
   ' ���g�D����(i,j) = �g�D.Cells(i,j)
   ' �͈�-Range-�́@�g�D�@�͂P��Ȃ̂����A����ɂ�萶�������z��͂P�����z��ł�
   ' �Ȃ��A�Q�����z��ɂȂ邱�Ƃɒ��ӁI�I
   Dim m As Long
   m = UBound(�g�D����, 1)
   Debug.Print m
   '   ���s�����i�P��̂݁j�̔z��Ȃ̂ŁA��P�̎����̏���l�����߂Ă����B
   ' Debug.Print �g�D����.Cells(1, 1)
   ' Debug.Print �g�D����(1)
   ' Debug.Print �g�D����(1).Cells(1, 1)
   ' ���w�g�D���́x�͂Q�����z��ł���B�����̃A�N�Z�X�̂������͂��ׂČ��
   Debug.Print �g�D����(1, 1)
   ' Dim b As Long
   ' b = �g�D.Cells(1, 1).Row
   ' Debug.Print b
   ' ���w�g�D�x�� �͈�-Range- �Ȃ̂� .Cell ���\�b�h�ōs�Ɨ�ɂ���ăA�N�Z�X����B
   ' �@�܂��A���Ƃ̕\�ŉ��s�ڂł��邩�i .Row ���\�b�h �j�A�Ȃǂ̏��������Ă���B
   ' Dim �g�D����ColorIndex(116) As Long
   ' Dim �g�D����CI() As Long
   ReDim �g�D����CI(m)
   '     ���w�g�D�x�Ƃ��Ă����Ă���Z���̔w�i�F�����i�[����z���p�ӂ���B
   '       �͈͂�������̂ł͂Ȃ����߁A�����I�Ɏ����Ƒ傫�����w�肵�Ȃ����
   '       �Ȃ�Ȃ��B�����ŁA���I�z��Ƃ��Đ錾�������ƁA�g�D���́i�͈͂Ƃ��Ă�
   '       �g�D���畡�������Q�����z��j�̍s���Ԃ�̗v�f�����P�����̔z���ݒ�
   '       ���Ă����B
   ' Dim �g�D����ST() As String
   ReDim �g�D����ST(m)
   '
   For r = 1 To m
      �g�D����CI(r) = �g�D.Cells(r, 1).Interior.ColorIndex
      �g�D����ST(r) = �g�D����(r, 1)
   Next r
   '
   For r = 1 To m
      ' Debug.Print �g�D����(r, 1) & ":" & �g�D����CI(r)
      ' �g�D�\��(r) = �g�D����(r, 1)
      Debug.Print �g�D����ST(r) & ":" & �g�D����CI(r)
   Next r
   '
End Sub

Private Sub �g�D���̏�����_test()
   ' Dim �g�D����() As String
   ' Dim �g�D����() As Variant
   Dim �g�D����() As Variant
   Dim �g�D����_S() As String
   ' Dim �g�D���� As Range
   Dim �g�D As Range
   Set �g�D = ThisWorkbook.Names("�g�D").RefersToRange
   ' Debug.Print "UBound(�g�D,1)�F" & UBound(�g�D, 1)
   ' Debug.Print "UBound(�g�D,2)�F" & UBound(�g�D, 2)
   nr = �g�D.Rows.count
   nc = �g�D.Columns.count
   Debug.Print "�g�D.Rows.count�F" & nr
   Debug.Print "�g�D.Columns.count�F" & nc
   ReDim �g�D����_S(nr, nc)
   ' ���O�t���͈̔́iNamed Range�j����z��ցF.RefersToRange
   ' �g�D���� = �g�D.Cells(1, 1)
   �g�D���� = �g�D
   �g�D����_S(1, 1) = �g�D.Cells(1, 1).Value
   ' �w�g�D�x�͂P�s�ڂ���͈̔͂ł͂Ȃ��̂����A�w�g�D���́x�́i�P�j�ɍŏ��̍s������
   m = UBound(�g�D����)
   Debug.Print m
   ' Debug.Print �g�D����.Cells(1, 1)
   ' Debug.Print �g�D����(1)
   ' Debug.Print �g�D����(1).Cells(1, 1)
   Debug.Print �g�D����(1, 1)
   Dim b As Long
   b = �g�D.Cells(1, 1).Row
   Debug.Print b
   Dim g As Long
   g = UBound(�g�D����, 1)
   Debug.Print g
   ' Dim �g�D����ColorIndex(116) As Long
   Dim �g�D����CI() As Long
   ReDim �g�D����CI(g)
   ' ReDim �g�D����ColorIndex(m)
   For r = 1 To g
      ' �g�D����CI(r) = �g�D.Cells(b + r - 1, 1).Interior.ColorIndex
      �g�D����CI(r) = �g�D.Cells(r, 1).Interior.ColorIndex
   Next r
   
   ' For r = 0 To m - 1
      ' �g�D����ColorIndex(r + 1) = �g�D.Cells(b + r, 1).Interior.ColorIndex
   ' Next r
   For r = 1 To m
      ' Debug.Print �g�D����(r) & ":" & �g�D����ColorIndex(r)
      Debug.Print �g�D����(r, 1) & ":" & �g�D����CI(r)
   Next r
End Sub

Private Sub ��2�����z��Ē�`��������()
   ' �Ō�̎����̒�`�͑�����������ς��邱�Ƃ��ł��邪�A����ȊO�̎����͕ς����Ȃ��B
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

