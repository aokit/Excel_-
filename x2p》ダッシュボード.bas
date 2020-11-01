' -*- coding:shift_jis -*-

'./x2p�t�_�b�V���{�[�h.bas

Sub ���O�̒�`�m�F�̐���()
   '
   ' �W�v�����ŎQ�Ɓ^�\������͈͂��w�肷�邽�߂ɖ��O�t�����ς�ł��邩��
   ' �`�F�b�N���X�g�𐶐�����B
   ' �W�v�����̌��ʂ�󋵂��܂Ƃ߂ĕ\�����閼�O�t���͈͂𐶐�����B
   '
   Call �J�n���}��
   Dim BA As Variant
   Set BA = ActiveSheet.Shapes(Application.Caller)
   '   ���֐����N�������{�^���̂���Z���͈͂��m�ۂ��Ă���
   Dim c1 As Long
   Dim c2 As Long
   Dim r As Long
   Dim r0 As Long
   Dim rZ As Long
   ' �`�F�b�N���X�g�����F
   ' �{�^���̍����̃Z�����疼�O�t���ɗp�ӂ��������񂪃Z���Ɋi�[���Ă���̂�
   ' �����̖��O�ɂ��āA�͈͂����蓖�Ă��Ă��邩�\������悤�Ȏ���ׂ�
   ' �Z���ɗ^����B
   c2 = BA.TopLeftCell.Column
   c1 = c2 - 1
   r0 = BA.TopLeftCell.Row + 1
   rZ = ��̍ŏI�s_range(Cells(r0, c1), , 2) ' �ŏ��̋󔒍s�̎�O�̍s
   For r = r0 To rZ
      Cells(r, c2).Value = "=isref(" & Cells(r, c1).Value & ")"
   Next r
   ' �\���̂��߂̖��O�t���͈͐����F
   ' ���̉��̋󔒂ɂÂ��ď󋵕\���p�̃Z���Ƃ��̖��O��z�u����B
   r0 = ��̍ŏI�s_range(Cells(rZ, c1), , 2)
   rZ = ��̍ŏI�s_range(Cells(rZ, c1), , 3)
   For r = r0 To rZ
      Call newName2Range(Cells(r, c2), Cells(r, c1).Value)
   Next r
   Call �I�������
End Sub

' ���O�t���͈̔͂𒊏ۉ����Ď�舵�����߂ɂ́A�܂��́A�͈͂ɑ΂��閼�O��`
' �̕ύX�i���O�̕t���ւ��j������Ă��������B
' ���������A���O�t���̗̈�𖼑O���瓾��悤�ɂ��Ă����ƁA���O�̕t���ւ�
' �����₷���B
' �ϐ��ƃ��[�N�V�[�g�̓��o�͖͂��O�������Z���͈̔͂��g���Ď�������B
'
Private Sub newName2Range(rng As Range, strName As String)
   '
   ' �͈͂ɖ��O��^����B���O���肪����ɂ���Ȃ��B
   ' �������O�����Ă�����Â����O�͏�������B
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
   ' ���O���͈̔͂ɂ��炽�Ȗ��O��^����B
   ' ���łɖ��O�����Ă���͈͂𖼑O�Ŏw�肵�āA
   ' �V�������O��^���A�Â����O�͏�������B
   '
   Dim aRange As Range
   On Error Resume Next ' �G���[���������Ă����̍s������s.
   Set aRange = ThisWorkbook.Names(orgName).RefersToRange
   On Error GoTo 0 ' On Error Resume Next ���g�p���ėL���ɂ����G���[�����𖳌��ɂ���.
    
   If aRange Is Nothing Then
      Debug.Print "�͈͂̂��Ƃ̖��O�F" & orgName & "�@�������悤�ł��B"
      Exit Sub
   End If
   Debug.Print "�͈͂̂��Ƃ̖��O�F" & orgName
   ThisWorkbook.Names(orgName).Delete
   ThisWorkbook.Names.Add Name:=newName, RefersTo:=aRange
   ' ThisWorkbook �ł͂Ȃ��āA aRange.Parent.Parent ���g���Ƃ��悢�H
   '
End Sub

' ���O�t���͈̔͂𒊏ۉ����Ď�舵����ƌ��ʂ��̂����L�q���ł���Ǝv���̂�
' �������`�B
Private Sub updateRDofNamedRange(strName As String, _
                                 Nrows As Long, _
                                 Ncolumns As Long)
   '
   ' ���O���͈̔́sstrName�t�̉E��(RDend)�̈ʒu�̎w��i���}�́��j������
   ' �sNrows�t�ƁsNcolumnss�t�ɍX�V����B
   ' ���c������
   ' �F�@�@�F
   ' ���@�@��
   ' ���c������
   ' ���@�@��
   ' ����͊�_�Łi�P�C�P�j�ƂȂ�B
   ' ��}�́��̈ʒu���w�肷�� �sNrows�t�ƁsNcolumnss�t�́A
   ' ���łȂ������l�ł���A 0 �͌��݂̎w���ς��Ȃ����̂Ƃ��Ď�舵����B
   '
   Dim aRange As Range
   Set aRange = ThisWorkbook.Names(strName).RefersToRange
   Debug.Print "�͈͂̂��Ƃ̍s���F" & aRange.Rows.count
   Debug.Print "�͈͂̂��Ƃ̗񐔁F" & aRange.Columns.count
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
   Debug.Print "�͈͂̐V���ȍs���F" & aRange.Rows.count
   Debug.Print "�͈͂̐V���ȗ񐔁F" & aRange.Columns.count
   ThisWorkbook.Names.Add Name:=strName, RefersTo:=aRange
End Sub

' �E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E
'
Function p�L������(strDate As String, _
                   Optional R_n As String = "�W�v����") As Boolean
   '
   ' �w�W�v���ԁx�Ɩ��t����ꂽ�Q�Z���̗񂩂���Ԃ̊J�n���ƏI�����𓾂�
   '  strDate �����̊��ԂɊ܂܂�Ă���ꍇ�ɂ� True ��Ԃ��B�����łȂ�
   '  �Ȃ� False ��Ԃ��B
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
   p�L������ = r
End Function

' --- �z���Dictionary�ɕϊ�����
' ��ʑg�D��擪�Ƃ��āA���̑g�D�ɋA�����鉺�ʑg�D���ȍ~�ɕ��ׂ��s��
' ��ʑg�D�̐������Ȃ�ׂ��z���
' ���ʑg�D�� key �Ƃ��A��ʑg�D�� Value �Ƃ��鎫���ɕϊ�����B
' �����A���i�������ʑg�D�������̏�ʑg�D�ɏ�������B
' ���ʁAkey �������� Value �����j�͂Ȃ����̂Ƃ���B
' 
Sub SingleHomeDict(ByRef str����() As String, _
                   ByRef dic���� As Dictionary)
    '
    ' �wstr����()�x
    '  ��ʑg�D��擪�Ƃ��āA���̑g�D�ɋA�����鉺�ʑg�D���ȍ~�ɕ��ׂ��s��
    '  ��ʑg�D�̐������Ȃ�ׂ��z��
    ' �wdic�����x
    '  ���ʑg�D�� key �Ƃ��A��ʑg�D�� Value �Ƃ��鎫��
    '
   Dim U1 As Long
   U1 = UBound(str����, 1)
   U2 = UBound(str����, 2)
   For i = 1 To U1
      d = str����(i, 1)
      d0 = i
      dic����.Add d, d0
      For j = 2 To U2
         d = str����(i, j)
         If d = "" Then Exit For
         If dic����.Exists(d) Then
            MsgBox "key�����������o�F�ċA�I�Ɏ����𐶐����Ċi�["
            ' �����̎d���n�Ȃǂ̑Ή��̂���
         Else
            dic����.Add d, d0
         End If
      Next j
   Next i
End Sub

Sub �g�D�ʌʐR�������W�v()
   Dim str���F�L�^() As String
   Call ���F�L�^�ǂݎ��(str���F�L�^)
   Dim str�g�D����() As String
   Call �g�D�����ǂݎ��(str�g�D����)
   Dim dic�g�D���� As New Dictionary
   Call �g�D�����\��(str�g�D����, dic�g�D����)
   '
   Dim U1 As Long
   Dim �g�D�ʌʐR������() As Long
   U1 = UBound(str�g�D����, 1)
   ReDim �g�D�ʌʐR������(1 To U1, 1 To 1)
   '
   Dim �\���ҏ��� As String
   Dim IX As Long
   Dim U2 As Long
   U2 = UBound(str���F�L�^, 1)
   For i = 2 To U2
      ' Debug.Print str���F�L�^(i, 2)
      If p�L������(str���F�L�^(i, 2)) Then
         ' 15 - �\���ҏ���
         ' Debug.Print str���F�L�^(i, 15)
         �\���ҏ��� = str���F�L�^(i, 15)
         If (dic�g�D����.Exists(�\���ҏ���)) Then
            IX = dic�g�D����.Item(�\���ҏ���)
            ' Debug.Print IX
            �g�D�ʌʐR������(IX, 1) = �g�D�ʌʐR������(IX, 1) + 1
         End If
      End If
   Next i
   For i = 1 To U1
      Debug.Print �g�D�ʌʐR������(i, 1)
   Next i
   Call �g�D�W�v�ʐR�������X�V(�g�D�ʌʐR������)
End Sub

Sub �g�D�����\��(ByRef str����() As String, ByRef dic���� As Dictionary)
   Call SingleHomeDict(str����(), dic����)
End Sub

Sub �W�v��_�g�D_������()
   ' Call �g�D���̏�����
End Sub

Private Sub Col2CIonST(strName As String, _
                       ByRef CI() As Long, _
                       ByRef ST() As String)
   ' �w�g�D�x�Ɩ��O�t�������͈͂�ǂݎ��F
   ' �E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E
   ' �P�F��P�����̕�����Ŗ��t����ꂽ�͈́i�P��~�����s�j��ǂݎ��
   ' �Q�F��Q�����Ƃ��Ďw�肵���z��ɁA�͈͂� ColorIndex ��Long�^�ŕԂ��B
   ' �R�F��R�����Ƃ��Ďw�肵���z��ɁA�͈͂� �Z���̒l ��String�^�ŕԂ��B
   ' �E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E
   '
   Dim rngC As Range
   Set rngC = ThisWorkbook.Names(strName).RefersToRange
   ' �@�����O�t���͈̔́iNamed Range�j��z�����\�Ȕ͈�-Range- �֕ϊ����郁�\�b�h
   '     .RefersToRange
   ' ��P�����̕�����i���Ƃ��΁w�g�D�x�j�͔C�ӂ̃Z���͈͂ł���A���Ƃ̃V�[�g�̂P�s��
   ' ����͈̔͂ł͂Ȃ��̂����AaryC(1) �ɍŏ��̍s������B
   ' ���Ƃ��΂��Ƃ̃V�[�g�̂U�s�ڂ���͈̔͂ł���΁A�����i�U�s�ځj�ւ̃A�N�Z�X�́A
   ' �w�g�D�x�̂P�s�ڂɃA�N�Z�X����΂悢�B
   Call Col2CIonSTrng(rngC, CI() ,ST())
   '                  ���^�����͈́i��j�̔w�i�F�Ɠ��e�����ꂼ��P��̔z��ɓ���B
End Sub

Private Sub Col2CIonSTrng(rngC As Range, _
                       ByRef CI() As Long, _
                       ByRef ST() As String)
   ' �w�g�D�x�Ɩ��O�t�������͈͂�ǂݎ��F
   ' �E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E
   ' �P�F��P������rngC���������͈́i�P��~�����s�j��ǂݎ��
   ' �Q�F��Q������CI()���Ƃ��Ďw�肵���z��ɁA�͈͂� ColorIndex ��Long�^�ŕԂ��B
   ' �R�F��R������ST()���Ƃ��Ďw�肵���z��ɁA�͈͂� �Z���̒l ��String�^�ŕԂ��B
   ' �E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E
   '
   Dim aryC() As Variant
   ' �@����P�����̕�����Ŗ��t����ꂽ�͈͂̃Z���̒l���i�[����z��B�͈͗R����
   '     �z��Ȃ̂Ŏ����͂Q�B�e�����̗v�f���͔͈͂Ɉˑ�����̂ŕs���B�v�f����
   '     ReDim �ȂǂŖ����I�Ɏw�肷�鈵���͂��Ȃ��B��
   '     �v�f�̌^��Variant �Ƃ��Ă���B
   ' �����I�z��Ƃ���
   ' �EQ:�͈͂̑傫�����킩��΁AReDim�Ŗ����I�Ɏw�肵�Ă��悢�H
   ' �EA:�w�͈͂̃J�������x��w�͈͂̍s���x�́A.Rows.Count �ȂǂŎ�ɓ��邪�A�Z����
   ' �l����ɓ����̂͂��Ȃ�ʓ| �i�g�D����_S(1, 1) = �g�D.Cells(1, 1).Value�j ��
   ' ����B���̂��ߐ����������@�ł͂Ȃ��B
   ' �i�܂����̎��_�ł͎������傫��������j�Ƃ���Variant �ɂ��Ă����̂��悢�iString
   ' �@�ɂ͂ł��Ȃ��j
   aryC = rngC.Value
   ' ���g�D����(i,j) = �g�D.Cells(i,j)
   ' �͈�-Range-�́@�g�D�@�͂P��Ȃ̂����A����ɂ�萶�������z��͂P�����z��ł�
   ' �Ȃ��A�Q�����z��ɂȂ邱�Ƃɒ��ӁI�I
   Dim m As Long
   m = UBound(aryC, 1)
   ' Debug.Print LBound(aryC, 1)
   '   ���͈͂���ǂݍ��񂾔z��iaryC = rngC.Value�j�Ȃ̂Ł@aryC(1,1)�@�������
   '   �@�i�ŏ��́j�Z���̒l�̓���v�f�ɂȂ�B
   '     ���@(0,0)�ł͂Ȃ�
   ' Debug.Print m
   '   ���s�����i�P��̂݁j�̔z��Ȃ̂ŁA��P�̎����̏���l�����߂Ă����B
   ' Debug.Print aryC.Cells(1, 1)
   ' Debug.Print aryC(1)
   ' Debug.Print aryC(1).Cells(1, 1)
   '   ���waryC�x�͂Q�����z��ł���B�����̃A�N�Z�X�̂������͂��ׂČ��
   ' Debug.Print aryC(1, 1)
   '   ���͈͂̍���́i�ŏ��́j�Z���̒l�������Ă��邱�Ƃ��m�F�ł���B
   ' Dim b As Long
   ' b = rngC.Cells(1, 1).Row
   ' Debug.Print b
   ' ���wrngC�x�� �͈�-Range- �Ȃ̂� .Cell ���\�b�h�ōs�Ɨ�ɂ���ăA�N�Z�X����B
   ' �@�܂��A���Ƃ̕\�ŉ��s�ڂł��邩�i .Row ���\�b�h �j�A�Ȃǂ̏��������Ă���B
   ReDim CI(m)
   '     ���wrngC�x�Ƃ��Ă����Ă���Z���̔w�i�F�����i�[����z���p�ӂ���B
   '       �͈͂�������̂ł͂Ȃ����߁A�����I�Ɏ����Ƒ傫�����w�肵�Ȃ����
   '       �Ȃ�Ȃ��B�����ŁA���I�z��Ƃ��Đ錾�������ƁA�g�D���́i�͈͂Ƃ��Ă�
   '       �g�D���畡�������Q�����z��j�̍s���Ԃ�̗v�f�����P�����̔z���ݒ�
   '       ���Ă����B
   ReDim ST(m)
   '     ���waryC�x�Ƃ��Ă����Ă���Z���̒l���i�[����z���p�ӂ���B
   '
   For r = 1 To m
      CI(r) = rngC.Cells(r, 1).Interior.ColorIndex
      ST(r) = aryC(r, 1)
   Next r
   '
End Sub

Private Sub �g�D���̓ǂݎ��(ByRef �g�D����CI() As Long, _
                             ByRef �g�D����ST() As String)
   ' �w�g�D�x�Ɩ��O�t�����͈́i�P��~�����s�j��ǂݎ��F
   ' �E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E
   ' �P�F��P�����Ƃ��Ďw�肵���z��ɁA�͈͂� ColorIndex ��Long�^�ŕԂ��B
   ' �Q�F��Q�����Ƃ��Ďw�肵���z��ɁA�͈͂� �Z���̒l ��String�^�ŕԂ��B
   ' �E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E
   Call Col2CIonST("�g�D",�g�D����CI(), �g�D����ST())
End Sub

' �g�D�����������̂Ȃ��� str�g�D���� �𐶐����邽�߂ɌĂ�
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
   '     ��varArr�̗񂪂Ƃ肤��ő�� m �s���Ƃ肤��ő�� m �ł���B
   '     �i�����Œ�`����Ƃ� 0 �s �� 0 �� �����邪�g�킸�Q�Ƃ����Ȃ��j
   Dim i As Long
   Dim j As Long
   Dim k As Long
   Dim RCI As Long
   i = 0: j = 0: k = 0
   RCI = CI(1)
   '���E�E�E��P�s�̐F����ʑg�D�̔����Ƃ��Ďg�� rootCI
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
   ' varArr�� j �s k ��̔z��Ƃ������ƂɂȂ�B
   ReDim varstrArr(1 To j, 1 To k)
   For q = 1 To k
      For p = 1 To j
         varstrArr(p, q) = varArr(p, q)
      Next p
   Next q
End Sub
   
' �g�D�����������̂Ȃ��� str�g�D���� �������o�����߂ɌĂ�
'
   '    .Resize(���s��,����) �ŏ����o���͈͂�ς����鄮
   '    �i�w�V�[�g�̏W�v���ƕʖ��x�Ƃ����͈͂��ς�遃�����ł͊g������遄
   '    �@�̂ŁA�͂ݏo����������؂藎�Ƃ���邱�ƂȂ������o����j
   '    �������A���ꂾ���ł̓V�[�g�ɒ�`���ꂽ���O���X�V���ꂽ�킯�ł͂Ȃ��B
   '
   ' ���O�������͈͂ɂ��Ă��X�V���Ă����B
   ' �i���Ƃ̖��O�wstrName�x�ɁA�X�V���ꂽ�Q�Ɣ͈́w�V�[�g�̏W�v���ƕʖ��x��
   ' �@���蓖�Ă邱�ƂɂȂ�̂ŁA���Ƃ̖��O�̒�`�ɏ㏑������遃�ʂ̖��O��
   ' �@����ƁA���Ƃ̖��O���c���Ă��܂��_�ɒ��Ӂ��j
   '
Private Sub Arr2ReNamedRange(ByRef varstrArr() As Variant, _
                             strName As String)
   '
   ' <1 to UBound(strArr,1)> x <1 to UBound(strArr,2)>  �̑傫��������
   ' �z�� strArr ��
   ' strName �Ŗ��t�����͈͂ɏ����o���B
   ' �͈͂̑傫���́A strArr �ɍ��킹�Ċg�k�i�Đݒ�j�����B
   ' ����ɁA���t�����g�k���ꂽ�͈͂ɍĐݒ肳���B
   ' Spill�ɗގ��@�\����ʓI�Ɏg����v���V�W��
   '
   Dim j As Long
   Dim k As Long
   j = UBound(varstrArr, 1)
   k = UBound(varstrArr, 2)
   Dim R_n As Range
   Set R_n = _
       ThisWorkbook.Names(strName).RefersToRange.Resize(j, k)
   '    �������A���ꂾ���ł̓V�[�g�ɒ�`���ꂽ���O���X�V���ꂽ�킯�ł͂Ȃ��B
   '
   R_n = varstrArr
   '
   ' ���O�������͈͂ɂ��Ă��X�V���Ă����B
   ActiveWorkbook.Names.Add Name:=strName, RefersTo:=R_n
   '
End Sub

Sub �g�D����������()
   '
   ' �g�D�\�i���O�w�g�D�x�Œ�`�����͈�-Range-�@���܂̂Ƃ���A
   ' �g�D���́E���́E�p���ď̂̃V�[�g�ɂ���j�̕\�L���e�ɏ]����
   ' ���O�wRange_�g�D�����x�Œ�`�����͈͂ɑg�D������W�J����B
   ' ���O�wRange_�g�D�����x�Œ�`�����͈͍͂ŏ��͍���ƂȂ�P�Z�������ł��邪
   ' �������ɂ���āA�K�v�ȑ傫���͈̔͂ɏ�����������B
   ' ���O�wRange_�g�D�����x�Œ�`�����͈͂�g�D�̎����Ƃ��Ďg��
   ' �E�g�D�W�v�̑�P��A�W�v��������������Ƃ�
   ' �E����R��A�������W�v����Ƃ�
   '
   ' �Ȃ��������Ő��������g�D�����͏����l�ł����āA���������邱�Ƃ��z�肳���
   ' ����B�������A���ڃZ�������������邱�Ƃ͑z�肵�Ȃ��BRange_�g�D�����i�͈́j
   ' �̍X�V�A�r���ɋ󔒂��Ȃ��A���d�A�����Ȃ��A�Ȃǂ̐��������ێ�������A�����
   ' ���֐��i�Q�Z���̑I���ƃ{�^���N���b�N�œo�^�j�̂��߂ɕʓr�v���V�W����p�ӂ�
   ' ��\��B
   '
   Call �g�D���̃N���A
   '
   Dim �g�D����CI() As Long
   Dim �g�D����ST() As String
   Call �g�D���̓ǂݎ��(�g�D����CI(), �g�D����ST())
   '
   Dim str�g�D����() As Variant
   Call CIonST2Arr(�g�D����CI(), �g�D����ST(), str�g�D����())
   '
   Dim strName As String
   strName = "Range_�g�D����"
   Call Arr2ReNamedRange(str�g�D����(), strName)
   '
End Sub


Sub �g�D�W�v1�񏉊���()
   '
   ' ���O�wRange_�g�D�����x�Ŗ��t�����͈͂ɑg�D�����ɂ��ƂÂ��āA
   ' ���O�w�g�D�W�v�x�Ŗ��t�����͈́i�W�v���̍���Z������K�v�ȍs���̂P���
   ' �͈͂𐶐�����B
   '
   ' ���O�w�g�D�W�v�x�Ŗ��t�����͈͂����Ƃ�
   ' ���O�wRange_�g�D�W�v�P��x�Ŗ��t�����͈͂��\������B
   '
   Dim strName As String
   strName = "Range_�g�D����"
   Dim Range_�g�D���� As Range
   Set Range_�g�D���� = _
       ThisWorkbook.Names(strName).RefersToRange
   Dim str�g�D����() As Variant
   str�g�D���� = Range_�g�D����.Value
   Dim j As Long
   j = UBound(str�g�D����, 1)
   strName = "�g�D�W�v"
   Dim Range_�g�D�W�v1�� As Range
   Set Range_�g�D�W�v1�� = _
       ThisWorkbook.Names(strName).RefersToRange.Resize(j, 1)
   Range_�g�D�W�v1��.Clear
   Range_�g�D�W�v1��.Font.Name = "BIZ UD�S�V�b�N"
   ' Range_�g�D�W�v1��.Value = str�g�D����
   Range_�g�D�W�v1�� = str�g�D����
   '
End Sub

Private Sub �g�D���̃N���A()
   Dim strName As String
   strName = "Range_�g�D����"
   Dim Range_�g�D���� As Range
   On Error Resume Next
   Set Range_�g�D���� = _
       ThisWorkbook.Names(strName).RefersToRange.Clear
   Call updateRDofNamedRange(strName, 1, 1)
   On Error GoTo 0
End Sub

Private Sub �g�D�W�v�ʐR�������X�V(ByRef �g�D�ʌʐR������() As Long)
   ' �g�D�W�v�ʐR�������N���A
   Dim strName As String
   strName = "�g�D�W�v"
   Dim R�g�D�W�v As Range
   Set R�g�D�W�v = ThisWorkbook.Names(strName).RefersToRange
   r0 = R�g�D�W�v.Row
   r1 = ��̍ŏI�s(strName)
   c0 = R�g�D�W�v.Column + 2
   Dim R�g�D�W�v������ As Range
   Set R�g�D�W�v������ = Range(Cells(r0, c0), Cells(r1, c0))
   R�g�D�W�v������.Clear
   R�g�D�W�v������.Font.Name = "BIZ UD�S�V�b�N"
   R�g�D�W�v������ = �g�D�ʌʐR������
End Sub

Function �g�D�W�v��() As String
   '
   ' �w�g�D�W�v�x�Ŗ��t�����͈͂���@�W�v���@�����o���Ĕz��Ƃ��ĕԂ�
   ' �@������Ɛ��l�����݂��Ă��邪�A�ЂƂ܂�������Ƃ��ēǂށB
   '
   Dim strName As String
   strName = "�g�D�W�v"
   Dim R�W�v��1 As Range
   Set R�W�v��1 = ThisWorkbook.Names(strName).RefersToRange
   r0 = R�W�v��1.Row
   c0 = R�W�v��1.Column
   r1 = ��̍ŏI�s(strName)
   c1 = c0 + 2
   ' ���E�E�E�N�ԓo�^�����ƌʐR�������̗��܂Ŋg��
   Set R�W�v��1 = Range(Cells(r0, c0), Cells(r1, c1))
   Dim str�W�v��() As String
   str�W�v�� = R�W�v��1
   �g�D�W�v�� = str�W�v��()
   ' ���E�E�E�z����֐��̕Ԃ��l�ɂ���Ƃ��ɂ́w()�x���K�v
   '
End Function

Sub �g�D�W�v_��[�����o(ByRef �g�D�W�v_��[��() As Variant)
   '
   ' �w�g�D�W�v�x�Ŗ��t�����͈͂��������܂ނ悤�Ɋg�����A�������O�łȂ�
   ' �@�s�ō\�����ꂽ�z���Ԃ�
   ' �������ɎQ�ƂŕԂ��B
   ' ��������ł͂Ȃ����l�Ƃ��ĕԂ������ꍇ������̂ň����� Variant �Ƃ����B
   '
   Dim strName As String
   strName = "�g�D�W�v"
   Dim R�g�D�W�v As Range
   Set R�g�D�W�v = ThisWorkbook.Names(strName).RefersToRange
   r0 = R�g�D�W�v.Row
   c0 = R�g�D�W�v.Column
   r1 = ��̍ŏI�s(strName)
   c1 = c0 + 2
   ' ���E�E�E�N�ԓo�^�����ƌʐR�������̗��܂Ŋg��
   ' Dim strCV() As String
   ' strCV = Range(Cells(r0, c0), Cells(r1, c1)).Value
   Dim strCV() As Variant
   Set R�g�D�W�v = Range(Cells(r0, c0), Cells(r1, c1))
   strCV = R�g�D�W�v.Value
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
   �g�D�W�v_��[�� = NZC
   '
End Sub

Private Sub �g�D�W�v_��[�����o(ByRef �g�D�W�v_��[��() As Variant)
   '
   ' �z��w�g�D�W�v�Q��[���x�i������̔z��j���󂯎���ď����o���B
   ' �����ł���z��́A���̗v�f��������ł͂Ȃ��Đ��l�̏ꍇ�ɂ����l
   ' �ɋ@�\���Ăق������Ƃ���AVariant �Ƃ����B
   '
   Dim strName As String
   strName = "�g�D�W�v"
   r1 = ��̍ŏI�s(strName)
   ' ���E�E�E�S�W�v���̍Ō�̍s���𓾂�
   strName = "�g�D�W�v�Q��[��"
   Dim R�g�D�W�v_��[�� As Range
   Set R�g�D�W�v_��[�� = ThisWorkbook.Names(strName).RefersToRange
   r0 = R�g�D�W�v_��[��.Row
   c0 = R�g�D�W�v_��[��.Column
   c1 = c0 + 2
   ' ���E�E�E�N�ԓo�^�����ƌʐR�������̗��܂Ŋg��
   Set R�g�D�W�v_��[�� = Range(Cells(r0, c0), Cells(r1, c1))
   R�g�D�W�v_��[��.Clear
   R�g�D�W�v_��[��.Font.Name = "BIZ UD�S�V�b�N"
   ' ���E�E�E�����̂͂��̗̈�̍ő�̍s��
   '
   r1 = UBound(�g�D�W�v_��[��, 1) - LBound(�g�D�W�v_��[��, 1) + r0
   Set R�g�D�W�v_��[�� = Range(Cells(r0, c0), Cells(r1, c1))
   R�g�D�W�v_��[�� = �g�D�W�v_��[��
   ' ���E�E�E�����o���͍̂s��̍s������
End Sub

Sub �g�D�W�v_��[���X�V()
   Dim �g�D�W�v_��[��() As Variant
   ' ���E�E�E�A���p�̔z��́A������ȊO���n����悤�� Variant �Ƃ��Ă����B
   '
   ReDim �g�D�W�v_��[��(1 To 3, 1 To 3)
   �g�D�W�v_��[��(1, 1) = "AA"
   �g�D�W�v_��[��(1, 2) = "AB"
   �g�D�W�v_��[��(1, 3) = "AC"
   �g�D�W�v_��[��(2, 1) = "BA"
   �g�D�W�v_��[��(2, 2) = "BB"
   �g�D�W�v_��[��(2, 3) = "BC"
   �g�D�W�v_��[��(3, 1) = "CA"
   �g�D�W�v_��[��(3, 2) = "CB"
   �g�D�W�v_��[��(3, 3) = "CC"
   '
   ' Set �g�D�W�v_��[�� = �g�D�W�v_��[�����o()
   Call �g�D�W�v_��[�����o(�g�D�W�v_��[��)
   Call �g�D�W�v_��[�����o(�g�D�W�v_��[��)
End Sub


Private Sub �z�񂩂�Z���֏����o��(strName As String, ByRef �z��() As String)
   '
   ' Dim strName As String
   ' strName = "�W�v���ƕʖ�"
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
       ThisWorkbook.Names(strName).RefersToRange.Resize(4, 4)
   '    .Resize(���s��,����) �ŏ����o���͈͂�ς����鄮
   '    �i�w�V�[�g�̏W�v���ƕʖ��x�Ƃ����͈͂��ς�遃�����ł͊g������遄
   '    �@�̂ŁA�͂ݏo����������؂藎�Ƃ���邱�ƂȂ������o����j
   '    �������A���ꂾ���ł̓V�[�g�ɒ�`���ꂽ���O���X�V���ꂽ�킯�ł͂Ȃ��B
   '
   �V�[�g�̏W�v���ƕʖ� = �W�v���ƕʖ�
   '
   ' ���O�������͈͂ɂ��Ă��X�V���Ă����B
   ' �i���Ƃ̖��O�wstrName�x�ɁA�X�V���ꂽ�Q�Ɣ͈́w�V�[�g�̏W�v���ƕʖ��x��
   ' �@���蓖�Ă邱�ƂɂȂ�̂ŁA���Ƃ̖��O�̒�`�ɏ㏑������遃�ʂ̖��O��
   ' �@����ƁA���Ƃ̖��O���c���Ă��܂��_�ɒ��Ӂ��j
   ActiveWorkbook.Names.Add Name:=strName, RefersTo:=�V�[�g�̏W�v���ƕʖ�
   '
End Sub


Sub �e�X�gMultiHomeDict()
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
   ' �e�s�̑�P��� Home ���L�ڂ���A��Q��ȍ~�ɂ��� Home �ɑ����� Member ��
   ' ����΁A����̌��Ԃ�L�ڂ���Ă���s�蒷�̍s�����߂��i�z��Ƃ��Ă͍Œ�
   ' �̍s���i�[�ł���񐔂́j1..N �s 1..M ��̔z��ł��� strHomeMember ������
   ' �Ƃ��A
   ' Member �� key �Ƃ��� Home �� Value �Ƃ��鎫�����\�����ĕԂ��֐��B
   ' ���ۂɂ́AHome ���P��̂Ƃ��ɂ́A������̗v�f�����z�� Value �ƂȂ�
   ' �z��ł���B�܂�A���� Member �� ������ Home �ɋL�ڂ���Ă���ꍇ�ɂ́A
   ' ���� Member �� Key �Ƃ��ăA�N�Z�X����� ������ Home ��v�f�Ƃ��Ď��z��
   ' ���A Value �ƂȂ�B
   '
   Dim dictMH As New Dictionary
   ' Dim strValue() As String
   ' ReDim strValue(1 To UBound(strHomeMember,1)) ' Home�̎�ސ����ő�v�f
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
         Dim strValue() As String ' �Ē�`���ł��邩�ǂ���
         If dictMH.Exists(c) Then
            strValue = dictMH.Item(c) ' �傫���̂킩��Ȃ��z�񂪒l
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


Private Sub �g�D�����ǂݎ��(ByRef str�g�D����() As String)
   '
   ' �E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E
   ' �wRange_�g�D�����x�Ɩ��O�t�������͈͂�ǂݎ��F
   ' �@�����Ƃ��Ďw�肵���z��ɁA�͈͂� �Z���̒l ��String�^�ŕԂ��B
   ' �E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E
   '
   ' �g�D�����̃V�[�g�͎��Ƃł̒ǋL���z�肳��Ă���B���̂���
   ' �̈�̑傫�����m�F�iwr, wc�j����K�v������̂Ŋm�F�B
   Dim strName As String
   strName = "Range_�g�D����"
   Dim r0 As Long
   Dim c0 As Long
   Dim rZ As Long
   Dim cZ As Long
   Dim wr As Long
   Dim wc As Long
   Dim R�g�D���� As Range
   Set R�g�D���� = ThisWorkbook.Names(strName).RefersToRange
   r0 = R�g�D����.Row
   c0 = R�g�D����.Column
   rZ = ��̍ŏI�s(strName)
   cZ = �s�̍ŏI��(strName)
   wr = rZ - r0 + 1
   wc = cZ - c0 + 1
   Set R�g�D���� = R�g�D����.Resize(wr, wc)
   Dim V�g�D����() As Variant
   V�g�D���� = R�g�D����
   ' �����Ƃ��ĕԂ��z��̑傫���������Őݒ�
   ReDim str�g�D����(1 To wr, 1 To wc)
   For i = 1 To wr
      For j = 1 To wc
         str�g�D����(i, j) = V�g�D����(i, j)
      Next j
   Next i
End Sub

Private Sub ���F�L�^�ǂݎ��(str_���F�L�^() As String)
   '
   ' �E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E
   ' �w���F�L�^�x�Ɩ��O�t�������͈͂�ǂݎ��F
   ' �@�����Ƃ��Ďw�肵���z��ɁA�͈͂� �Z���̒l ��String�^�ŕԂ��B
   ' �E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E
   '
   ' �̈�̑傫�����m�F�imr, mc�j
   Dim mr As Long
   Dim mc As Long
   mr = ��̍ŏI�s("���F�L�^")
   mc = �s�̍ŏI��("���F�L�^")
   Dim V_���F�L�^() As Variant
   Dim ���F�L�^ As Range
   Set ���F�L�^ = ThisWorkbook.Names("���F�L�^").RefersToRange.Resize(mr, mc)
   V_���F�L�^ = ���F�L�^
   ' ��V_���F�L�^(i,j) = ���F�L�^.Cells(i,j)
   ' Dim str_���F�L�^ As string
   ReDim str_���F�L�^(1 To mr, 1 To mc)
   For i = 1 To mr
      For j = 1 To mc
         str_���F�L�^(i, j) = V_���F�L�^(i, j)
      Next j
   Next i
   '
End Sub

Private Sub �g�D���̏�����_��test��()
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
      ' ���@����͂Q�����z��ɂȂ��Ă��Ȃ��̂ŊԈႢ
      ' Debug.Print �g�D����(r, 1) & ":" & �g�D����CI(r)
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

' ------END

