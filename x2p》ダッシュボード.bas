' -*- coding:shift_jis -*-

'./x2p�t�_�b�V���{�[�h.bas

' ������
' �����O
' 
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
   Dim rz As Long
   ' �`�F�b�N���X�g�����F
   ' �{�^���̍����̃Z�����疼�O�t���ɗp�ӂ��������񂪃Z���Ɋi�[���Ă���̂�
   ' �����̖��O�ɂ��āA�͈͂����蓖�Ă��Ă��邩�\������悤�Ȏ���ׂ�
   ' �Z���ɗ^����B
   c2 = BA.TopLeftCell.Column
   c1 = c2 - 1
   r0 = BA.TopLeftCell.Row + 1
   rz = ��̍ŏI�s_range(Cells(r0, c1), , 2) ' �ŏ��̋󔒍s�̎�O�̍s
   For r = r0 To rz
      Cells(r, c2).Value = "=isref(" & Cells(r, c1).Value & ")"
   Next r
   ' �\���̂��߂̖��O�t���͈͐����F
   ' ���̉��̋󔒂ɂÂ��ď󋵕\���p�̃Z���Ƃ��̖��O��z�u����B
   r0 = ��̍ŏI�s_range(Cells(rz, c1), , 4)
   rz = ��̍ŏI�s_range(Cells(rz, c1), , 5)
   For r = r0 To rz
      Call newName2Range(Cells(r, c2), Cells(r, c1).Value)
   Next r
   Call �I�������
End Sub

' ������
' �����P
' 
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
   ' �P���E�E�E�w�i�F���ʂō��ꂽ�g�D�\�̔w�i�F�Ɠ��e��z��ɂ��ꂼ��i�[����
   Dim �g�D����CI() As Long
   Dim �g�D����ST() As String
   Call �g�D���̓ǂݎ��(�g�D����CI(), �g�D����ST())
   '
   ' �Q���E�E�E�w�i�F���ʂō��ꂽ�g�D�\�̔z����s�P�ʂ̑g�D�\�ɕϊ�����
   Dim str�g�D����() As Variant
   Call CIonST2Arr(�g�D����CI(), �g�D����ST(), str�g�D����())
   '
   ' �R���E�E�E���t�������͈͂ɔz��������o���͈͂��L���Ė��t�����X�V����
   Dim strName As String
   strName = "Range_�g�D����"
   Call Arr2ReNamedRange(str�g�D����(), strName)
   '
End Sub

' ������
' �����Q
' 
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

' ������
' �����R
' 
Sub �g�D�ʌʐR�������W�v()
   Dim str���F�L�^() As String
   Call ���F�L�^�ǂݎ��(str���F�L�^)
   Dim dic�g�D���� As New Dictionary
   Dim U1 As Long
   If False Then 'True then
      Dim str�g�D����() As String
      Call �g�D�����ǂݎ��(str�g�D����)
      Call �g�D�����\��(str�g�D����, dic�g�D����)
      U1 = UBound(str�g�D����, 1)
   Else
      ' �ȉ��̂ɒu�������Ă݂�
      Call SingleHomeDict_namedRange("Range_�g�D����",U1,dic�g�D����,0)
   End If
   '
   Dim �g�D�ʌʐR������() As Long
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
      ' Debug.Print �g�D�ʌʐR������(i, 1)
   Next i
   Call �g�D�W�v�ʐR�������X�V(�g�D�ʌʐR������)
End Sub



' ������������������������������������������������������������������������
' ��������������������������������������������������������������������������
' ������ʉ��v���V�W��
' ����

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

' �g�D�����������̂Ȃ��� str�g�D���� �𐶐����邽�߂ɌĂ�
'
Private Sub CIonST2Arr(ByRef CI() As Long, _
                       ByRef ST() As String, _
                       ByRef varstrArr() As Variant)
   Dim m As Long
   m = UBound(ST, 1)
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

' �g�D�ʌʐR�������W�v�@�̒��ŁA�\�͈̔͂��X�V���邽�߂ɌĂ�
' 
Private Sub �g�D�W�v�ʐR�������X�V(ByRef �g�D�ʌʐR������() As Long)
   ' �g�D�W�v�ʐR�������N���A
   Dim strName As String
   strName = "�g�D�W�v"
   Dim R�g�D�W�v As Range
   Set R�g�D�W�v = ThisWorkbook.Names(strName).RefersToRange
   r0 = R�g�D�W�v.Row
   R1 = ��̍ŏI�s(strName)
   ' c0 = R�g�D�W�v.Column + 2
   Dim R�g�D�W�v������ As Range
   ' stop
   ' Set R�g�D�W�v������ = Range(Cells(r0, c0), Cells(r1, c0))
   ' ���E�E�E���t�����͈͂����ƂɐV���Ȕ͈͂��w�肷��B
   Set R�g�D�W�v������ = R�g�D�W�v.Offset(0, 2).Resize((R1 - r0 + 1), 1)
   R�g�D�W�v������.Clear
   R�g�D�W�v������.Font.Name = "BIZ UD�S�V�b�N"
   R�g�D�W�v������ = �g�D�ʌʐR������
End Sub

' �g�D�ʌʐR�������W�v�@�̒��ŁA���F�L�^��ǂݎ�邽�߂ɌĂяo��
' 
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

' �g�D�ʌʐR�������W�v�@�̒��ŁA�g�D������ǂݎ�邽�߂ɌĂяo��
' 
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
   Dim rz As Long
   Dim cz As Long
   Dim wr As Long
   Dim wc As Long
   Dim R�g�D���� As Range
   Set R�g�D���� = ThisWorkbook.Names(strName).RefersToRange
   r0 = R�g�D����.Row
   c0 = R�g�D����.Column
   rz = ��̍ŏI�s(strName)
   cz = �s�̍ŏI��(strName)
   wr = rz - r0 + 1
   wc = cz - c0 + 1
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

' �g�D�ʌʐR�������W�v�@�̂Ȃ��Ł@�g�D�ʎ����@���\�����邽�߂ɌĂ�
'
Sub �g�D�����\��(ByRef str����() As String, ByRef dic���� As Dictionary)
   Call SingleHomeDict(str����(), dic����)
End Sub

Sub SingleHomeDict(ByRef str����() As String, _
                   ByRef dic���� As Dictionary)
    '
    ' �wstr����()�x
    '  ��ʑg�D��擪�Ƃ��āA���̑g�D�ɋA�����鉺�ʑg�D���ȍ~�ɕ��ׂ��s��
    '  ��ʑg�D�̐������Ȃ�ׂ��z��
    ' �wdic�����x
    '  str�����ɂP����n�܂�s�ԍ���^���Ă���� Value �Ƃ��A
    '  �e�s�̏�ʑg�D����щ��ʑg�D�� key �Ƃ��鎫��
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

' �g�D�ʌʐR�������W�v�@�̂Ȃ��ŏW�v�̂��߂̊��Ԃɓ����Ă��郌�R�[�h��
' ���肷�邽�߂Ɏg��
'
Function p�L������(strDate As String, _
                   Optional R_n As String = "�W�v����") As Boolean
   '
   ' �w�W�v���ԁx�Ɩ��t����ꂽ�Q�Z���̗񂩂���Ԃ̊J�n���ƏI�����𓾂�
   '  strDate �����̊��ԂɊ܂܂�Ă���ꍇ�ɂ� True ��Ԃ��B�����łȂ�
   '  �Ȃ� False ��Ԃ��B
   '
   Dim r As Boolean
   Dim R1 As Boolean
   Dim r2 As Boolean
   Dim strD As String
   Dim strD1 As String
   Dim strD2 As String
   strD = CDate(strDate)
   strD1 = CDate(ThisWorkbook.Names(R_n).RefersToRange.Cells(1, 1))
   strD2 = CDate(ThisWorkbook.Names(R_n).RefersToRange.Cells(2, 1))
   R1 = (0 <= DateDiff("d", strD1, strD))
   r2 = (0 <= DateDiff("d", strD, strD2))
   r = R1 And r2
   p�L������ = r
End Function

' ������������������������������������������������������������������������
' ��������������������������������������������������������������������������
' �����ėp�v���V�W��
' ����

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

Private Sub �g�D���̓ǂݎ��(ByRef �g�D����CI() As Long, _
                             ByRef �g�D����ST() As String)
   ' �w�g�D�x�Ɩ��O�t�����͈́i�P��~�����s�j��ǂݎ��F
   ' �E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E
   ' �P�F��P�����Ƃ��Ďw�肵���z��ɁA�͈͂� ColorIndex ��Long�^�ŕԂ��B
   ' �Q�F��Q�����Ƃ��Ďw�肵���z��ɁA�͈͂� �Z���̒l ��String�^�ŕԂ��B
   ' �E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E
   Call Col2CIonST("�g�D", �g�D����CI(), �g�D����ST())
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
   Call Col2CIonSTrng(rngC, CI(), ST())
   '                  ���^�����͈́i��j�̔w�i�F�Ɠ��e�����ꂼ��P��̔z��ɓ���B
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

' --- �͈͂�Dictionary�ɕϊ�����
' �͈͂��e�Z���̓��e�� key �Ƃ��ċL�ڂ��ꂽ�s�� value �Ƃ��� Dictionary �ɕϊ�����
'
Private Sub SingleHomeDict_namedRange(strName As String, _
                                      ByRef nClass As Long, _
                                      ByRef dic���� As Dictionary, _
                                      Optional cx As Long = 1)
   ' �l�̊K������Ԃ��K�v������B
   ' �����ł̎����̍쐬�̖ړI�́A������ key �ɓ��� Value ��Ԃ������݂�
   ' �ȒP�Ɏ������邱�ƂȂ̂ŁA����ނ� Value ��Ԃ����ƂɂȂ��Ă���̂�
   ' �ɂ��ẮA�������\�������Ƃ��ɂ킩����̂Ƃ��ĕԂ����Ƃ����߂���B
   ' ��Q�����Ƃ��ĎQ�Ɠn�����Ă�����Ă����ĕԂ��B
   '
   ' ��P�����F�����ɂ�����e���L�ڂ��Ă���͈͂ɖ��t�������O�i������j
   ' ��Q�����F�����̎��� Value �̃N���X����Ԃ����߂̕ϐ��i�Q�Ɠn���j
   ' ��R�����F��������鎫��
   ' �E�E�E�E�E���͈͂̊e�Z���̓��e�� key
   ' �E�E�E�E�E�����ꂪ�L�ڂ��ꂽ�s�ԍ��i�͈͓��ł̍s�ԍ��j�� value
   ' ��S�����F�i�I�v�V���i���j�����͈̗͂񐔂𐧌�����Ƃ����̗�
   ' �E�E�E�E�E�������Ȃ��Ƃ��ɂ́w 0 �x�i1��菬�����l�j�Ƃ���B
   ' ��P�������Z���P�����̂Ƃ��ɂ́A
   ' �͈͂́w��̍ŏI�s�x�Ɓw�����s�̍ŏI��_range�x�܂Ŋg�傳���B
   '
   ' Debug.Print strName
   ' Stop
   Dim R_n As Range
   Set R_n = ThisWorkbook.Names(strName).RefersToRange
   Set R_n = range_�A����ő�s_range(R_n)
   '
   Dim var����() As Variant
   var���� = R_n.Value
   nClass = R_n.Rows.Count
   '
   Dim U1 As Long
   U1 = UBound(var����, 1)
   U2 = UBound(var����, 2)
   For i = 1 To U1
      d = var����(i, 1)
      d0 = i
      dic����.Add d, d0
      For j = 2 To U2
         d = var����(i, j)
         If d = "" Then Exit For
         If dic����.Exists(d) Then
         Else
            dic����.Add d, d0
         End If
      Next j
   Next i
End Sub

