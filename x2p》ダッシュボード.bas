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
   ' rz = ��̍ŏI�s_range(Cells(r0, c1), , 2) ' �ŏ��̋󔒍s�̎�O�̍s
   ' rz = ��̍ŏI�s_range(Cells(r0, c1), , 3) ' �ŏ��̋󔒍s�̎�O�̍s
   ' ���E�E�E�Ȃ��������A�p�����[�^�i��L�̂R�j�𑝂₳�Ȃ��Ƃ����Ȃ������B
   ' �@�@�@�@�ǂ����Ă��͖���
   '�����ύX����
   rz = ��̍ŏI�s_range(Cells(r0, c1),, 1) ' �ŏ��̋󔒍s�̎�O�̍s
   ' stop
   For r = r0 To rz
      Cells(r, c2).Value = "=isref(" & Cells(r, c1).Value & ")"
   Next r
   ' �\���̂��߂̖��O�t���͈͐����F
   ' ���̉��̋󔒂ɂÂ��ď󋵕\���p�̃Z���Ƃ��̖��O��z�u����B
   ' r0 = ��̍ŏI�s_range(Cells(rz, c1), , 4)
   ' rz = ��̍ŏI�s_range(Cells(rz, c1), , 5)
   '�����ύX����
   r0 = ��̍ŏI�s_range(Cells(rz, c1),, 1)
   rz = ��̍ŏI�s_range(Cells(r0, c1),, 1) ' �ŏ��̋󔒍s�̎�O�̍s
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
   ' �����ύX����
   ' ���Q�ƁF�u�b�N���̖��O�w�g�D�x
   ' �@�@�@�@���R�����̑g�D�����P��ɕ\�L�B
   ' �@�@�@�@����ʑg�D�͐F�t���̃Z���E���ʑg�D�͐F�Ȃ��̃Z���ŕ\���B
   ' �@�@�@�@���V�[�g�w�g�D���́E���́E�p���ď́x�ɒu���B
   ' �@�@�@�@�����̖��O�A���̃V�[�g�ł����Ă����Ȃ��B
   ' �����ʁF�u�b�N���̖��O�w�g�D�������x
   ' �@�@�@�@������̂P�Z�������̖��O�Ƃ��Ē�`�B�@�g�D�����@�Ƃ��ēW�J�B
   ' �@�@�@�@���g�D�����@�́A�P��ڂ���ʑg�D�A�Q��߈ȍ~�������鉺�ʑg�D�B
   ' �@�@�@�@���V�[�g��ł́A�P��ȏ�́A�s�蒷�i�s��񐔁j�A���s�@
   ' �@�@�@�@�����̂��ׂẴv���V�W���ł́@�g�D�����@�̎Q�Ƃ́A���̖��O��
   ' �@�@�@�@�@����čs���B���O����@�s�蒷�A���s�@�͈̔͂��m�肷��@�\��
   ' �@�@�@�@�@���ꂼ��̃v���V�W�����p�ӂ���B
   ' ���O�wRange_�g�D�����x�Œ�`�����͈͂�g�D�̎����Ƃ��Ďg��
   ' �E�g�D�W�v�̑�P��A�W�v��������������Ƃ�
   ' �E����R��A�������W�v����Ƃ�
   ' �����ύX����
   ' ���O�w�g�D�������x�̓��e�́A���O�w�g�D�W�v�x������i�_�b�V���{�[�h
   ' �ɂ���j���i�w�g�D�x�ɂ��ƂÂ��āj�X�V����r���ōX�V�����B
   ' �܂��A���O�w�g�D�W�v�x�̗���P��Ƃ�����R��i�W�v�j���X�V����Ƃ�
   ' �ɁA�W�v�����ŎQ�Ƃ����i���ʑg�D����ʑg�D�ŏW�v���邽�߂ɕK�v�j
   ' 
   ' �Ȃ��������Ő��������g�D�����͏����l�ł����āA���������邱�Ƃ��z�肳���
   ' ����B�������A���ڃZ�������������邱�Ƃ͑z�肵�Ȃ��BRange_�g�D�����i�͈́j
   ' �̍X�V�A�r���ɋ󔒂��Ȃ��A���d�A�����Ȃ��A�Ȃǂ̐��������ێ�������A�����
   ' ���֐��i�Q�Z���̑I���ƃ{�^���N���b�N�œo�^�j�̂��߂ɕʓr�v���V�W����p�ӂ�
   ' ��\��B
   ' �����ύX����
   ' �w�g�D�������x�ō���Z�����`�����A�s�蒷�A���s�͈̔͂ł���g�D�����́A
   ' �w���P�x�Łw�g�D�x���琶�������B�g�D�����́A�w�g�D�W�v�x�̑�P��̃Z����
   ' �̋L�ڂƁA��R��̃Z���ւ̏W�v�����ɗ��p�����B
   ' �g�D�����ł́A���d�A���i�����̏�ʑg�D�ɓ���̉��ʑg�D���������邱�Ɓj��
   ' �N����Ȃ����̂Ƃ���B����́A�w�g�D�x�������̋@�ւ�����肵�����A�g�D�ύX
   ' �̍ۂɁA���炽�Ɂw�g�D�x�̓��e����肷�邱�Ƃ��ł��Ȃ��ꍇ�ɂ��A�g�D������
   ' �͂Ȃ��w�g�D�x��ҏW�i���݂̐ݒ���R�s�[�A�ҏW���āA�w�g�D�x�Ɩ��O��`�j
   ' ������̂Ƃ����B
   '
   ' Call �g�D���̃N���A
   ' �����ύX����
   ' ������������
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
   Stop
   
   ' �R���E�E�E���t�������͈͂ɔz��������o���͈͂��L���Ė��t�����X�V����
   Dim strName As String
   ' strName = "Range_�g�D����"
   ' Call Arr2ReNamedRange(str�g�D����(), strName)
   ' �����ύX����
   strName = "�g�D������"
   Call PrintArrayOnNamedRange(strName, str�g�D����)

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
   ' �܂��A�����̗̈���N���A����悤�ɂ����B�����̗̈悪�傫���Ă��@�\����B
   '
   Dim strName As String
   ' strName = "Range_�g�D����"
   ' Dim Range_�g�D���� As Range
   ' Set Range_�g�D���� = _
   '     ThisWorkbook.Names(strName).RefersToRange
   ' Dim str�g�D����() As Variant
   ' str�g�D���� = Range_�g�D����.Value
   '������s���~�P��̔z��
   ' �����ύX����
   strName = "�g�D������"
   Dim rh As long
   Dim �g�D���� As Range
   Set �g�D���� = ThisWorkbook.Names(strName).RefersToRange
   rh = Range_��̍ŏI�s_Range(�g�D����,, 1).Rows.Count
   Set �g�D���� = �g�D����.Resize(rh,1)
   Dim str�g�D����() As Variant
   str�g�D���� = �g�D����.Value
   '������s���~�P��̔z��
   
   ' Dim j As Long
   ' j = UBound(str�g�D����, 1)
   ' strName = "�g�D�W�v"
   ' Dim Range_�g�D�W�v1�� As Range
   ' Dim r0 As Long
   ' Dim r1 As Long
   ' Set Range_�g�D�W�v1�� = _
   '     ThisWorkbook.Names(strName).RefersToRange
   ' �����ύX����
   Dim j As Long
   j = UBound(str�g�D����, 1)
   strName = "�g�D�W�v"
   Dim Range_�g�D�W�v1�� As Range
   Set Range_�g�D�W�v1�� = _
       ThisWorkbook.Names(strName).RefersToRange

   ' r0 = Range_�g�D�W�v1��.Row
   ' r1 = ��̍ŏI�s(strName)
   ' Set Range_�g�D�W�v1�� = Range_�g�D�W�v1��.Resize((r1 - r0 + 1), 1)
   ' Range_�g�D�W�v1��.Clear
   Set Range_�g�D�W�v1�� = Range_�g�D�W�v1��.Resize(j, 1)
   Call ClearColumnRowEnd(Range_�g�D�W�v1��)
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
      ' Unused
      ' �ȉ��� True�� �͎g���Ȃ��Ȃ����͂��B���t�@�N�^�����O�Ńo�T�b��
      ' ������悤�Ɉȉ��̃A�T�[�g�̎c���B
      Debug.Print("�g���Ȃ��Ȃ���True�߂����s����܂����B")
      Exit Sub
      '
      Dim str�g�D����() As String
      Call �g�D�����ǂݎ��(str�g�D����)
      Call �g�D�����\��(str�g�D����, dic�g�D����)
      U1 = UBound(str�g�D����, 1)
   Else
      ' �ȉ��̂ɒu�������Ă݂�
      ' Call SingleHomeDict_namedRange("Range_�g�D����", U1, dic�g�D����, 0)
      '�����ύX����
      Call SingleHomeDict_namedRange("�g�D������", U1, dic�g�D����, 0)
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
   '
   ' Stop
   '
   Call �g�D�W�v�ʐR�������]�L()
End Sub

Sub �g�D�W�v�ʐR�������]�L()
   Dim R_1 As Range
   Dim R_2 As Range
   Dim R_3 As range
   Set R_1 = range_��̍ŏI�s_namedrange("�g�D�W�v",, 1)
   Set R_2 = R_1.Offset(0, 2) ' �g�D�W�v�̏W�v�l
   Dim v1() As Variant
   Dim v2() As Variant
   Dim v3() As Variant ' - v3 �͂Q�����z��ł��邱�Ƃ𖾎����邱��
   ' �@�@�������́w()�x���Ȃ��ƁA�R���p�C�����z��łȂ��Ɣ��f���ăG���[�B
   v1 = R_1.value
   v2 = R_2.value
   Dim LB1 As Long
   Dim UB1 As Long
   LB1 = LBound(v1, 1)
   UB1 = UBound(v1, 1)
   If LB1 <> 1 Then Debug.Print("�g�D�W�v�Ɉُ킪����܂��B���]�L")
   If LB1 <> LBound(v2, 1) Then Debug.Print("�g�D�W�v�Ɉُ킪����܂��B���]�L")
   If UB1 <> UBound(v2, 1) Then Debug.Print("�g�D�W�v�Ɉُ킪����܂��B���]�L")
   ReDim v3(1 To UB1, 1 To 2)
   For i = 1 To UB1
      v3(i, 1) = v1(i, 1)
      v3(i, 2) = v2(i, 1)
   Next i
   ' Set R_3 = range_��̍ŏI�s_namederange("�ʕ\�P").Offset(1,0)
   
   Set R_3 = ThisWorkBook.Names("�ʕ\�P").RefersToRange.Offset(1, 0).Cells(1, 1)

   ' Call PrintArrayOnRange(R_3, v3, 2)

   ' Call PrintArrayOnRange(R_3, v3, 0)
   Call PrintArrayOnRange(R_3, v3, -1)
   
End Sub

'��������������
Function PickHeadWord(ByRef dicSyn As Dictionary, _
                 ByVal word As String) As String
   ' �ԍ��ŗޕʂ��ꂽ�ދ`�ꎫ���̌��o�����Ԃ�
   '
   PickHeadWord = ""
   If Not dicSyn.Exists(word) Then Exit Function
   Dim ix As Long
   ix = dicSyn(word)
   Dim i  As Long
   For i = 0 To dicSyn.Count - 1
      If dicSyn(dicSyn.Keys(i)) = ix Then
         PickHeadWord = dicSyn.Keys(i)
         Exit For
      End If
   Next i   
End Function
'��������������

' ������
' �����S
'
Sub �g�D�W�v_��[���X�V()
   Dim �g�D�W�v_��[��() As Variant
   ' ���E�E�E���̔z��́A������ȊO�������n���ړI�� Variant �Ƃ��Ă����B
   Call �g�D�W�v_��[�����o(�g�D�W�v_��[��)
   Call �g�D�W�v_��[�����o(�g�D�W�v_��[��)
End Sub

' ������
' �����T
'
Sub ����敪�ʌʐR�������W�v()
   '
   Dim str���F�L�^() As String
   Call ���F�L�^�ǂݎ��(str���F�L�^)
   '
   Dim dic����敪���� As New Dictionary
   Dim U1 As Long
   Call SingleHomeDict_namedRange("����W�v", U1, dic����敪����, 1)
   ' ���w����W�v�x�Ɩ��t�����Z���̒����P��� dic����敪���� �Ƃ��Ċm��
   '
   Dim ����敪�ʌʐR������() As Long
   ReDim ����敪�ʌʐR������(1 To U1, 1 To 1)
   Dim ����敪�ʌʐR�����z() As Long
   ReDim ����敪�ʌʐR�����z(1 To U1, 1 To 1)
   '
   Dim ����敪 As String ' �\�ł́w������e�敪�x
   Dim ���z As Long ' �\�ł́w���z�i�~�j�x
   Dim IX As Long
   Dim U2 As Long
   U2 = UBound(str���F�L�^, 1)
   For i = 2 To U2
      If p�L������(str���F�L�^(i, 2)) Then
         ' 8 - �\�ł́w������e�敪�x
         ' 10 - �\�ł́w���z�i�~�j�x
         ' Debug.Print str���F�L�^(i, 8)
         ' Debug.Print str���F�L�^(i, 10)
         ����敪 = str���F�L�^(i, 8)
         ���z = str���F�L�^(i, 10)
         If (dic����敪����.Exists(����敪)) Then
            IX = dic����敪����.Item(����敪)
            ' Debug.Print IX
            ����敪�ʌʐR������(IX, 1) = ����敪�ʌʐR������(IX, 1) + 1
            ����敪�ʌʐR�����z(IX, 1) = ����敪�ʌʐR�����z(IX, 1) + ���z
         End If
      End If
   Next i
   ' stop
   Call ����敪�W�v�ʐR�������X�V(����敪�ʌʐR������)
   Call ����敪�W�v�ʐR�����z�X�V(����敪�ʌʐR�����z)
End Sub

' ������
' �����U
'
Sub ����W�v_��[���X�V()
   Dim ����W�v_��[��() As Variant
   ' ���E�E�E���̔z��́A������ȊO�������n���ړI�� Variant �Ƃ��Ă����B
   Call ����W�v_��[�����o(����W�v_��[��)
   Call ����W�v_��[�����o(����W�v_��[��)
End Sub

' ������
' �����V
'
' ���d���n�����i�W�v���^�ʖ��ˍs�ԍ��j�̐���
' ���@�E���t���͈́w�d���n�ʖ��x����
' ���d���n�W�v���z��i�s�ԍ����W�v���j�̐���
' ���d���n�W�v
' ���d���n�ʖ��񐔕\���i�O��ƂQ��ȏ�𖼕t���͈́w�d���n�ʖ��񐔁x�ɕ\���j
' �@�@�E���t���͈́w�d���n�ʖ��x�Ɠ����V�[�g�Ɂw�d���n�ʖ��񐔁x�i�Q��j��ݒ�
' �@�@�@�d���n�̕ʖ��Ɋւ��āA�ʖ��ɋL�ڂ���Ă��Ȃ����́i�O��j�@��
' �@�@�@�ʖ��ɕ�����i�Q��ȏ�j�L�ڂ���Ă�����́@���񐔂ƂƂ��ɕ\�����Ă���B
'
Sub �d���n�ʏW�v()
   Dim dic�d���ԍ� As New Dictionary
   Dim ary�W�v��() As String
   Dim ary�W�v���������z() As Variant
   ' Dim ary�������d��() As String
   Dim ary�������d��() As Variant
   Dim nClass As Long
   Call �d���n��������("�d���n�ʖ�", nClass, dic�d���ԍ�)
   ' Stop
   ' ���E�E�E�����Ł@Debug.Print(dic�d���ԍ�("�C�^����")(1))�@�Ƃ����
   ' �@�@�@�@dic�d���ԍ��ɓǂݍ��܂�Ă��邱�Ƃ��킩��
   Call �d���n�W�v���z�񐶐�("�d���n�ʖ�", nClass, ary�W�v���������z)
   ' ���E�E�E
   Call �d���n�W�v("���F�L�^", 11, 10, dic�d���ԍ�, ary�W�v���������z, ary�������d��)
   ' Stop
   ' ���E�E�E���̂��Ɓ@�W�v���������z�@�Ɓ@�������d���@��\������
   ' ���d���n�W�v���������z�\��
   Call PrintArrayOnNamedRange("�d���n�W�v", ary�W�v���������z, 3, -4)
   ' ���������ʖ��\��
   ' Stop
   'Call PrintArrayOnNamedRange("�������ʖ�", ary�W�v���������z, 1)
   '
   ' �\�[�g���A����͈̔͂ɋL��
   ' 
   Call �~���\�[�g(ary�W�v���������z, 2) '�����~��
   Dim ary�W�v������() As Variant
   ReDim ary�W�v������(1 To nClass, 1 To 2)
   For i = 1 To nClass
      ary�W�v������(i, 1) = ary�W�v���������z(i, 1)
      ary�W�v������(i, 2) = ary�W�v���������z(i, 2)
   Next i
   Call PrintArrayOnNamedRange("�d���n�����Q�g�b�v", ary�W�v������, 2, -4)
   
   Call �~���\�[�g(ary�W�v���������z, 3) '���z�~��
   Dim ary�W�v�����z() As Variant
   ReDim ary�W�v�����z(1 To nClass, 1 To 2)
   For i = 1 To nClass
      ary�W�v�����z(i, 1) = ary�W�v���������z(i, 1)
      ary�W�v�����z(i, 2) = ary�W�v���������z(i, 3) / 1000
   Next i
   Call PrintArrayOnNamedRange("�d���n���z�Q�g�b�v", ary�W�v�����z, 2, -4)

   '
   Call PrintArrayOnNamedRange("�������ʖ�", ary�������d��, 1)
   'Call PrintArrayOnNamedRange("�������ʖ�", Transpose(ary�������d��), 1)
   ' Dim aryT As Variant
   ' Dim Ur As Long
   ' Dim Lr As Long
   ' Ur = UBound(ary�������d��, 1)
   ' Lr = LBound(ary�������d��, 1)
   ' ReDim aryT(Lr To Ur, 1 To 1)
   ' For r = Lr To Ur
   '    aryT(r, 1) = ary�������d��(r)
   ' Next r
   ' Stop
   ' Call PrintArrayOnNamedRange("�������ʖ�", aryT, 1)
   ' 
   ' Stop
   ' Call �d���n�ʖ��񐔕\��("�d���n�ʖ���")
End Sub

Private Sub �d���n�W�v(strNameD As String, _
                       iC_�d���n As Long, _
                       iC_���z As Long, _
                       ByRef dicDIN As Dictionary, _
                       ByRef aryNTM() As Variant, _
                       ByRef aryYAD() As Variant)
   '...................ByRef aryYAD() As String)
   ' aryYAD�@�������d���@�̈���
   ' �\��������Ƃ��� aryYAD �� Variant �łȂ��ƌ^����v�����R���p�C���G���[
   ' ���N�����Ă��܂��̂ŁA�Ȃ񂾂��D�ɗ����Ȃ��� Variant �ɂ����B
   ' �܂��A�\�ɖ߂��ۂɁA�P�������ƍs�ɂȂ��Ă��܂��̂ŁA�Q�����̂P��z���
   ' �ڂ��ւ��Ă���B
   '
   ' ��P�����FstrNameD�F�̖��O���^����ꂽ�͈͂��珳�F�f�[�^��ǂݎ��B
   ' ��Q�����FiC_�d���n�F���F�f�[�^�ɂ�����d���n�̃J�����ԍ�
   ' ��R�����FiC_���z�F���F�f�[�^�ɂ�������z�̃J�����ԍ�
   ' ��S�����FdicDIN�F�d���n�̕ʖ�����
   ' ��T�����FaryNTM�F�i�Ԃ��l�j�d���n���E���v���z�E���v�񐔂̕\
   ' ��U�����FaryYAD�F�i�Ԃ��l�j�����������ł��Ȃ������d���n�̕\
   ' 
   ' ("���F�L�^", 11, 10, dic�d���ԍ�, ary�W�v���������z)
   ' ��U�����F�����蓖�Ďd���n��Ԃ��i�j
   '
   ' ReDim aryYAD(1 To 200)
   ' �Q�����ɂ��Ă���
   ' ReDim aryYAD(1 To 200, 1 To 1)
   Dim TaryYAD(1 To 200)
   ' ���E�E�E�������̎d�����i�[����i�Ԃ��Ƃ��ɂ͖��g�p�����j
   Dim nYAD As Long
   nYAD = 0
   Dim str���F�L�^() As String
   ' Call ���F�L�^�ǂݎ��(str���F�L�^)
   ' Call NamedRangeSQ2ArrStr("���F�L�^", str���F�L�^)
   Call NamedRangeSQ2ArrStr(strNameD, str���F�L�^)
   ' Stop
   ' ��������@�z��@str���F�L�^�@�ɑ΂��ā@���R�^���T���Q�l�ɂ��ďW�v�������s���B
   Dim U2 As Long
   ' �����Ƌ��z���������iEmpty�ł͂Ȃ�0�Ɂj
   U2 = UBound(aryNTM, 1)
   For i = 1 To U2
      aryNTM(i, 2) = 0
      aryNTM(i, 3) = 0
   Next i
   ' str���F�L�^�@�̑S�s���X�L����
   U2 = UBound(str���F�L�^, 1)
   Dim iC_���F�� As Long
   iC_���F�� = 2
   Dim IX As Long
   Dim str�d���n As String
   Dim str���z As String
   For i = 2 To U2
      If p�L������(str���F�L�^(i, iC_���F��)) Then
         str�d���n = str���F�L�^(i, iC_�d���n)
         str���z = str���F�L�^(i, iC_���z)
         If dicDIN.Exists(str�d���n) Then
            Dim aryIX() As Variant
            aryIX = dicDIN.Item(str�d���n)
            ' ���E�E�E�d���n�� Identification Number�i�������������蓾��̂ł̔z��j
            Dim U22 As Long
            u22 = UBound(aryIX, 1)
            For j = 1 To U22
               If j > 1 Then Debug.Print j
               IX = aryIX(j)
               aryNTM(IX, 2) = aryNTM(IX, 2) + 1
               ' aryNTM(IX, 3) = aryNTM(IX, 3) + Val(str���z)
               aryNTM(IX, 3) = aryNTM(IX, 3) + Val(str���z) / U22
               ' ���E�E�E�ʖ������񐔂ŋϓ��������ĐώZ
            Next j
            Erase aryIX
         Else
            Debug.Print(str�d���n)
            ' ���W�v���ɂ��ʖ��ɂ��Ȃ��d���n�͂ЂƂ܂�������Ă����B
            ' �@���t���͈́w�ʖ��񐔁x�ɂO��Ƃ��ĕ\�����邽�߂�
            ' �@�z����\�����Ă����\��B
            nYAD = nYAD + 1
            TaryYAD(nYAD) = str�d���n
         End If
      End If
   Next i
   If nYAD = 0 Then
      ' �������d�����܂��������������Ƃ��A�z����Ȃ����Ă��܂��Ɨ�O����
      ' ���ʓ|�Ȃ̂ŁA���ʂɂP�v�f�ŋ󕶎���̔z��ɂ��Ă����B
      ReDim aryYAD(1 To 1, 1 To 1)
      aryYAD(1, 1) = ""
   Else
      ReDim aryYAD(1 To nYAD, 1 To 1)
      ' ���E�E�E�������d���̔z��ŏ�������ł��Ȃ��Ƃ����؂藎�Ƃ�
      ' �@�@�@�@���t���͈́w�ʖ��񐔁x�ɂO��Ƃ��ĕ\������Ώ�
      For j = 1 To nYAD
         aryYAD(j, 1) = TaryYAD(j)
      Next j
   End If
   ' Stop
   '
End Sub

' ������
' �����W
'
Sub �����ᓙ���o()
   '
   '
   ' ���R�̃R�[�h�𗬗p
   Dim str���F�L�^() As String
   Call ���F�L�^�ǂݎ��(str���F�L�^)
   Dim U2 As Long
   U2 = UBound(str���F�L�^, 1)

   Dim dic�g�D���� As New Dictionary
   Dim U1 As Long

   If False Then
      ' �g���Ȃ��Ȃ����i�͈͂̕ς��@Range_�g�D�����@�ł͂Ȃ�
      ' �P��Z���́@�g�D�������@���g���悤�ɂ����j
      ' Call SingleHomeDict_namedRange("Range_�g�D����", U1, dic�g�D����, 0)
   Else
      Call SingleHomeDict_namedRange("�g�D������", U1, dic�g�D����, 0)
   End If
   
   Dim �\���ҏ��� As String

   Dim A1() As String ' �����
   Dim A2() As String ' ����i�ݕ��j
   Dim A3() As String ' ����i�𖱁j
   Dim A4() As String ' ���z����
   Dim A5() As String ' ���m����
   Dim A6() As String ' �Y������
   Dim A1A() As String
   Dim A4A() As String
   Dim A1V() As Variant
   Dim A4V() As variant
      
   ReDim A1(1 To U2)
   ReDim A2(1 To U2)
   ReDim A3(1 To U2)
   ReDim A4(1 To U2)
   ReDim A5(1 To U2)
   ReDim A6(1 To U2)
   ReDim A1A(1 To U2, 1 To 7)
   ReDim A4A(1 To U2, 1 To 7)
   
   Dim i1 As Long
   Dim i2 As Long
   Dim i3 As Long
   Dim i4 As Long
   Dim i5 As Long
   Dim i6 As Long

   i1 = 1
   i2 = 1
   i3 = 1
   i4 = 1
   i5 = 1
   i6 = 1
   
   Dim c1 As String
   Dim c17 As String
   Dim c18 As String
   Dim c19 As String
   Dim c20 As String
   Dim q As Boolean

   Dim yen As Long
   
   For i = 2 To U2
      q = False
      If p�L������(str���F�L�^(i, 2)) Then
         c1 = str���F�L�^(i, 1)
         c17 = str���F�L�^(i, 17)
         c18 = str���F�L�^(i, 18)
         c19 = str���F�L�^(i, 19)
         c20 = str���F�L�^(i, 20)
         If (c18 = "����K�p") Or (c20 = "����K�p") Then
            A1(i1) = c1
            ' �����ˑ������F�L�^�̃t�B�[���h�\��
            A1A(i1, 1) = str���F�L�^(i, 1) ' �Ǘ��ԍ�
            A1A(i1, 2) = str���F�L�^(i, 7) ' ����
            A1A(i1, 3) = str���F�L�^(i, 11) ' �d���n
            A1A(i1, 4) = str���F�L�^(i, 13) ' �ڋq�E�_���
            A1A(i1, 5) = str���F�L�^(i, 14) ' �ŏI���v��
            ' A1A(i1, 6) = str���F�L�^(i, 15) ' �\���ҏ����Q�{���ϊ��O
            �\���ҏ��� = str���F�L�^(i, 15) ' �\���ҏ����Q�{���ϊ��O
            yen = str���F�L�^(i, 10) ' ���z�Q�~
            A1A(i1, 7) = CStr(CLng(yen) / 1000)
            If (dic�g�D����.Exists(�\���ҏ���)) Then
               A1A(i1, 6) = PickHeadWord(dic�g�D����, �\���ҏ���)
            Else
               A1A(i1, 6) = "*"
            End If
            i1 = i1 + 1
            q = True
         End If
         If (c18 = "����K�p") Then
            A2(i2) = c1
            i2 = i2 + 1
            q = True
         End If
         If (c20 = "����K�p") Then
            A3(i3) = c1
            i3 = i3 + 1
            q = True
         End If
         If (Right(c18, 2) = "����") Then ' ���z����
            A4(i4) = c1
            ' �����ˑ������F�L�^�̃t�B�[���h�\��
            A4A(i4, 1) = str���F�L�^(i, 1) ' �Ǘ��ԍ�
            A4A(i4, 2) = str���F�L�^(i, 7) ' ����
            A4A(i4, 3) = str���F�L�^(i, 11) ' �d���n
            A4A(i4, 4) = str���F�L�^(i, 13) ' �ڋq�E�_���
            A4A(i4, 5) = str���F�L�^(i, 14) ' �ŏI���v��
            ' A1A(i1, 6) = str���F�L�^(i, 15) ' �\���ҏ����Q�{���ϊ��O
            �\���ҏ��� = str���F�L�^(i, 15) ' �\���ҏ����Q�{���ϊ��O
            yen = str���F�L�^(i, 10) ' ���z�Q�~
            A4A(i4, 7) = CStr(CLng(yen) / 1000)
            If (dic�g�D����.Exists(�\���ҏ���)) Then
               A4A(i4, 6) = PickHeadWord(dic�g�D����, �\���ҏ���)
            Else
               A1A(i4, 6) = "*"
            End If
            i4 = i4 + 1
            q = True
         End If
         If (Right(c20, 2) = "����") Then
            A5(i5) = c1
            i5 = i5 + 1
            q = True
         End If
         If ((Left(c17, 2) = "�Y��") Or (Left(c19, 2) = "�Y��")) And (Not q) Then
            A6(i6) = c1
            i6 = i6 + 1
         End If
      End If
   Next i

   ReDim Preserve A1(1 To i1 - 1)
   ReDim Preserve A2(1 To i2 - 1)
   ReDim Preserve A3(1 To i3 - 1)
   ReDim Preserve A4(1 To i4 - 1)
   ReDim Preserve A5(1 To i5 - 1)
   ReDim Preserve A6(1 To i6 - 1)
   ReDim A1V(1 To i1 - 1, 1 To 7)
   Dim r As Long
   Dim c As Long
   For r = 1 To i1 - 1
      For c = 1 To 7
         A1V(r, c) = A1A(r, c)
      Next c
   Next r
   ReDim A4V(1 To i4 - 1, 1 To 7)
   For r = 1 To i4 - 1
      For c = 1 To 7
         A4V(r, c) = A4A(r, c)
      Next c
   Next r
   
   ' Stop

   Dim AA() As Variant
   Dim iz As Long
   iz = 0
   If iz < (i1 - 1) Then iz = i1 - 1
   If iz < (i2 - 1) Then iz = i2 - 1
   If iz < (i3 - 1) Then iz = i3 - 1
   If iz < (i4 - 1) Then iz = i4 - 1
   If iz < (i5 - 1) Then iz = i5 - 1
   If iz < (i6 - 1) Then iz = i6 - 1
   ReDim AA(1 To iz, 1 To 6)
   For i = 1 To iz
      If UBound(A1, 1) < i Then
         AA(i, 1) = ""
      Else
         AA(i, 1) = A1(i)
      End If
      If UBound(A2, 1) < i Then
         AA(i, 2) = ""
      Else
         AA(i, 2) = A2(i)
      End If
      If UBound(A3, 1) < i Then
         AA(i, 3) = ""
      Else
         AA(i, 3) = A3(i)
      End If
      If UBound(A4, 1) < i Then
         AA(i, 4) = ""
      Else
         AA(i, 4) = A4(i)
      End If
      If UBound(A5, 1) < i Then
         AA(i, 5) = ""
      Else
         AA(i, 5) = A5(i)
      End If
      If UBound(A6, 1) < i Then
         AA(i, 6) = ""
      Else
         AA(i, 6) = A6(i)
      End If
   Next i

   ' Stop

   Call PrintArrayOnNamedRange("�Y�����",AA,6,-4)
   Call PrintArrayOnNamedRange("������",A1V,7,-4)
   Call PrintArrayOnNamedRange("���z���ᄬ",A4V,7,-4)
   ' ���Ȃ񂩁ARoot ����Format�����Ă����ĂȂ��悤�ȁE�E�E
   '   PrintArrayOnNamedRange�̃o�O�H
   
End Sub

'
' ������������������������������������������������������������������������
' ��������������������������������������������������������������������������
' ������ʉ��v���V�W���i������ Private �֐��j
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
   Dim R�g�D�W�v������ As Range
   If False Then
      ' �ȉ��g��Ȃ��Ȃ����B
      Dim strName As String
      strName = "�g�D�W�v"
      Dim R�g�D�W�v As Range
      Set R�g�D�W�v = ThisWorkbook.Names(strName).RefersToRange
      r0 = R�g�D�W�v.Row
      r1 = ��̍ŏI�s(strName,, 1)
      ' c0 = R�g�D�W�v.Column + 2
      ' Dim R�g�D�W�v������ As Range
      ' stop
      ' Set R�g�D�W�v������ = Range(Cells(r0, c0), Cells(r1, c0))
      ' ���E�E�E���t�����͈͂����ƂɐV���Ȕ͈͂��w�肷��B
      Set R�g�D�W�v������ = R�g�D�W�v.Offset(0, 2).Resize((r1 - r0 + 1), 1)
      ' R�g�D�W�v������.Clear
      ' �ⓚ���p�ŁA���܂���Ƃ�����Ō�܂ŏ����A�Ƃ�����ɂ����ق����������낤
      Call ClearColumnRowEnd(R�g�D�W�v������)
      ' ---
   Else
      '�����ύX����
      ' Dim R�g�D�W�v������ As Range
      Set R�g�D�W�v������ = Range_��̍ŏI�s_namedrange("�g�D�W�v",, 1).Offset(0, 2)
      R�g�D�W�v������.Clear
   End If
   ' ---
   R�g�D�W�v������.Font.Name = "BIZ UD�S�V�b�N"
   R�g�D�W�v������ = �g�D�ʌʐR������
End Sub

Sub ClearColumnRowEnd(ByVal Ro1C As Range)
   ' �w�肳�ꂽ�͈͂̈�ԏ�̃Z������󔒂Ȃ��ŘA�������
   ' �͂̃Z�����N���A����
   ' ��P�����FRo1C - Range of 1 Column - �����W�^�͈̔�
   ' Ro1C�̈�ԏ�̃Z������A��̍ŏI�s�܂ł͈̔͂̓��e��
   ' �N���A����B
   ' 
   Dim r0 As Long
   Dim r1 As Long
   r0 = Ro1C.Row
   r1 = ��̍ŏI�s_Range(Ro1C,, 1)
   Set Ro1C = Ro1C.Resize((r1 - r0 + 1), 1)
   Ro1C.Clear
End Sub

Sub FillColumnRowEnd(ByVal Ro1C As Range, ByRef Ary As Variant)
   ' �w�肳�ꂽ�͈͂̈�ԏ�̃Z������󔒂Ȃ��ŘA�������
   ' �͂̃Z���ɁA�����s�~�P��̔z�����������
   ' ��P�����FRo1C - Range of 1 Column - �����W�^�͈̔�
   ' Ro1C�̈�ԏ�̃Z������A
   ' ��Q�����FAry �����s�~�P��̔z�����������
   ' 
   Dim rs As Long
   rs = UBound(Ary, 1) - LBound(Ary, 1) + 1
   Set Ro1C = Ro1C.Resize(rs, 1)
   Ro1C.Font.Name = "BIZ UD�S�V�b�N"
   Ro1C = Ary
End Sub

' �g�D�ʌʐR�������W�v�@�̒��ŁA���F�L�^��ǂݎ�邽�߂ɌĂяo��
' �d���n�W�v�@�̒��ł��Ăяo���B
'
Private Sub ���F�L�^�ǂݎ��(str_���F�L�^() As String)
   '
   ' �E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E
   ' �w���F�L�^�x�Ɩ��O�t�������͈͂�ǂݎ��F
   ' �@�����Ƃ��Ďw�肵���z��ɁA�͈͂� �Z���̒l ��String�^�ŕԂ��B
   ' �E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E
   '
   Call NamedRangeSQ2ArrStr("���F�L�^",str_���F�L�^)
End Sub

Private Sub NamedRangeSQ2ArrStr(strName As String, _
                                str_���F�L�^() As String)
   ' ��P�����F���������̕����̍s����Ȃ�͈́@�ɖ��t�������O
   ' ��Q�����F��L�͈̔͂��i�[����z��
   ' �̈�̑傫�����m�F�imr, mc�j
   Dim mr As Long
   Dim mc As Long
   mr = ��̍ŏI�s(strName,, 1)
   mc = �s�̍ŏI��(strName,, 1)
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
   ' Unused
   ' ���̊֐��͎g���Ȃ��Ȃ����͂��B�m�F�̂��߂Ɉȉ��ɃA�T�[�g����
   ' �悤�ɂ��Ă����B
   Debug.Print("�g���Ȃ��Ȃ����֐����Ă΂�܂����F�g�D�����ǂݎ��")
   Exit Sub
   '
   ' �E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E
   ' �wRange_�g�D�����x�Ɩ��O�t�������͈͂�ǂݎ��F
   ' �@�����Ƃ��Ďw�肵���z��ɁA�͈͂� �Z���̒l ��String�^�ŕԂ��B
   ' �E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E
   '
   ' �g�D�����̃V�[�g�͎��Ƃł̒ǋL���z�肳��Ă���B���̂���
   ' �̈�̑傫�����m�F�iwr, wc�j����K�v������̂Ŋm�F�B
   Dim strName As String
   ' strName = "Range_�g�D����"
   ' �����ύX����
   strName = "�g�D������"

   ' Dim r0 As Long
   ' Dim c0 As Long
   ' Dim rz As Long
   ' Dim cz As Long
   ' Dim wr As Long
   ' Dim wc As Long
   ' Dim R�g�D���� As Range
   ' Set R�g�D���� = ThisWorkbook.Names(strName).RefersToRange
   ' r0 = R�g�D����.Row
   ' c0 = R�g�D����.Column
   ' rz = ��̍ŏI�s(strName)
   ' cz = �s�̍ŏI��(strName)
   ' wr = rz - r0 + 1
   ' wc = cz - c0 + 1
   ' Set R�g�D���� = R�g�D����.Resize(wr, wc)
   ' �����ύX����
   Dim R�g�D���� As Range
   Set R�g�D���� = range_�A����ő�s_namedrange(strName)

   Dim V�g�D����() As Variant
   V�g�D���� = R�g�D����
   ' �����Ƃ��ĕԂ��z��̑傫���������Őݒ�
   ' �����ύX����
   wr = R�g�D����.Rows.Count
   wc = R�g�D����.Columns.Count
   ReDim str�g�D����(1 To wr, 1 To wc)
   For i = 1 To wr
      For j = 1 To wc
         str�g�D����(i, j) = V�g�D����(i, j)
      Next j
   Next i
   Stop
   '
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

' �g�D�W�v_��[���X�V �Ŏg�p���Ă���B
'
Private Sub �g�D�W�v_��[�����o(ByRef �g�D�W�v_��[��() As Variant)
   '
   ' �w�g�D�W�v�x�Ŗ��t�����͈͂��������܂ނ悤�Ɋg�����A�������O�łȂ�
   ' �@�s�ō\�����ꂽ�z���Ԃ�
   ' �������ɎQ�ƂŕԂ��B
   ' ��������ł͂Ȃ����l�Ƃ��ĕԂ������ꍇ������̂ň����� Variant �Ƃ����B
   '
   Call NZrowCompaction("�g�D�W�v", 3, �g�D�W�v_��[��)
End Sub

Private Sub ����W�v_��[�����o(ByRef ����W�v_��[��() As Variant)
   '
   Call NZrowCompaction("����W�v", 3, ����W�v_��[��)
   '
End Sub

' �g�D�W�v_��[�����o �Ŏg�p���Ă���B
' ����W�v_��[�����o �Ŏg�p���Ă���B
'
Private Sub NZrowCompaction(strName As String, _
                            cols As Long, _
                            ByRef NZC() As Variant)
   '
   ' ��P�����wstrName�x�F�ΏۂƂȂ�\�͈͂̍���̃Z���ɗ^����ꂽ���O
   ' ��Q����   �wcols�x�F��L�̕\�͈̗͂�
   ' ��R����    �wNZC�x�F�Ԃ��z��
   '
   ' Dim strName As String
   ' strName = "�g�D�W�v"
   Dim R_n As Range
   Set R_n = ThisWorkbook.Names(strName).RefersToRange
   r0 = R_n.Row
   c0 = R_n.Column
   r1 = ��̍ŏI�s(strName,, 1)
   c1 = c0 + cols - 1
   ' ���E�E�E�N�ԓo�^�����ƌʐR�������̗��܂Ŋg��
   ' Dim strCV() As String
   ' strCV = Range(Cells(r0, c0), Cells(r1, c1)).Value
   Dim strCV() As Variant
   Set R_n = Range(Cells(r0, c0), Cells(r1, c1))
   strCV = R_n.Value
   Dim strNZ() As String
   ReDim strNZ(1 To (UBound(strCV, 1) - LBound(strCV, 1) + 1), 1 To cols)
   ReDim NZC(1 To (UBound(strCV, 1) - LBound(strCV, 1) + 1), 1 To cols)
   Dim j As Long
   Dim c As Long
   Dim z As Boolean
   j = 0
   For i = LBound(strCV, 1) To UBound(strCV, 1)
      z = True
      For c = 2 To cols
         z = z And (Val(strCV(i, LBound(strCV, 2) + (c - 1))) = 0)
      Next c
      If Not (z) Then
         j = j + 1
         For c = 1 To cols
            strNZ(j, c) = strCV(i, LBound(strCV, 2) + (c - 1))
         Next c
         ' strNZ(j, 1) = strCV(i, LBound(strCV, 2))
         ' strNZ(j, 2) = strCV(i, LBound(strCV, 2) + 1)
         ' strNZ(j, 3) = strCV(i, LBound(strCV, 2) + 2)
      End If
   Next i
   ' Dim NZC() As Variant
   ReDim NZC(1 To j, 1 To cols)
   For i = 1 To j
      NZC(i, 1) = strNZ(i, 1)
      For c = 2 To cols
         NZC(i, c) = Val(strNZ(i, c))
      Next c
      ' NZC(i, 1) = strNZ(i, 1)
      ' NZC(i, 2) = Val(strNZ(i, 2))
      ' NZC(i, 3) = Val(strNZ(i, 3))
   Next i
   ' �g�D�W�v_��[�� = NZC
   '
End Sub

' �g�D�W�v_��[���X�V �Ŏg�p���Ă���B
'
Private Sub a_�g�D�W�v_��[�����o(ByRef �g�D�W�v_��[��() As Variant)
   '
   ' �z��w�g�D�W�v�Q��[���x�i������̔z��j���󂯎���ď����o���B
   ' �����ł���z��́A���̗v�f��������ł͂Ȃ��Đ��l�̏ꍇ�ɂ����l
   ' �ɋ@�\���Ăق������Ƃ���AVariant �Ƃ����B
   '
   Dim strName As String
   strName = "�g�D�W�v"
   r1 = ��̍ŏI�s(strName,, 1)
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
' ��
' ����C�������ق��������B�ȉ��C����
'
Private Sub �g�D�W�v_��[�����o(ByRef �g�D�W�v_��[��() As Variant)
   '
   ' �z��w�g�D�W�v�Q��[���x�i������̔z��j���󂯎���ď����o���B
   ' �����ł���z��́A���̗v�f��������ł͂Ȃ��Đ��l�̏ꍇ�ɂ����l
   ' �ɋ@�\���Ăق������Ƃ���AVariant �Ƃ����B
   '
   Call PrintNZrowCompaction("�g�D�W�v�Q��[��", 3, �g�D�W�v_��[��)
   '
End Sub

Private Sub ����W�v_��[�����o(ByRef ����W�v_��[��() As Variant)
   '
   Call PrintNZrowCompaction("����W�v�Q��[��", 3, ����W�v_��[��, -4)
   '
End Sub

' �g�D�W�v_��[�������o�� �Ŏg�p���Ă���B
' ����W�v_��[�������o�� �Ŏg�p���Ă���B
'
Private Sub PrintNZrowCompaction(strName As String, _
                                 cols As Long, _
                                 ByRef NZC() As Variant, _
                                 Optional ROOT As Long = 0)
   '
   Call PrintArrayOnNamedRange(strName, NZC, cols, ROOT)
   '
End sub
                                 
Private Sub Copy1Dto2CAry(ByRef iAry() As Variant, _
                        ByRef oAry() As Variant)
   ' iAry() ���P�����z��ł���Ƃ�������p����B
   ' oAry() �͗񐔂��P�̂Q�����z��ł���B
End Sub

' ��������������������������������������������������������������������
   '
   ' �L���Ώۂ̔z��Ɠ����傫���͈̔͂���������ǂݍ��ށB�w*�x �ȊO
   ' �̕���������v�f���A�L���Ώۂ̔z��̗v�f�Œu��������B
   ' �w*�x���܂߂ēǂݍ��݁A�w*�x�͂��̂܂܂ɂ��āA�z��S�̂�
   ' ���̂܂܏����o���Ă��悢�������ύX�\��

   ' �z��̓��e�𖼕t�����͈͂ɏ������ށB�񂲂Ƃ̏�����ݒ肷��
   ' ���Ƃ��ł���B�񂲂Ƃ̏����́A�e��̏������ޔ͈͂̏������
   ' �Z���i�I�t�Z�b�g�� ROOT �Ŏw��F���̒l�j�ɗ\�ߐݒ肵�Ă����B
   ' �͈̖͂��O�t���́A�͈͂̍���̃Z���𖼕t����B
   ' �͈͂̍s���́A���t�����Z���̉������i�s�ԍ��̑���������j��
   ' ���e�����Z���̘A������͈͂Ŋg�������B
   ' �͈̗͂񐔂́A�wcols�x�ŗ^����B
   ' x �wcols�x���^���Ȃ��i�f�t�H���g�l�O���^����ꂽ�j�Ƃ��͍s��
   ' x �͈͓��̍ł������i��ԍ��̑���������œ��e�����Z���̘A��
   ' x ����j�͈͂Ŋg�������B
   ' x �����̂悤�ȗ�̊g���͖������B
   ' �����ύX����
   ' �Z���ɋL�ڂ���񐔂̎w��ł���wcols�x���^�����Ȃ��i�f�t�H���g
   ' �l�O���^����ꂽ�j�Ƃ��͔z��S�́i�z�񂪂��񐔂̂��ׂāj���L��
   ' ����B�^������ Ary �͂Q�����z��Ȃ̂ŁA�Q�����߂̃C���f�N�X��
   ' ��肤��l�̎�ސ��� cols �ɂȂ�B
   ' 
   ' ��P�����wstrName�x�F�������ޔ͈͂ɂ������O
   ' ��Q����    �wAry�x�F�������ޓ��e��ێ������z��
   ' ��R�����i�I�v�V�����j�wcols�x�F�������ޔ͈̗͂�
   ' ��S�����i�I�v�V�����j�wROOT(RowOffsetOfTemplate)�x
   '
   ' ��Q�����wAry�x���P�����z��̂Ƃ��͂P�s�ł͂Ȃ��ĂP��z��
   ' �Ƃ��Ď�舵���悤�ɂ��Ă݂����A���܂��s���Ă��Ȃ��B
   ' ����ł́A��Q�����́A�K���Q�����z��ł��邱�ƁB
   '
Private Sub PrintArrayOnNamedRange(strName As String, _
                                   ByRef aAry() As Variant, _
                                   Optional cols As Long = 0, _
                                   Optional ROOT As Long = 0)
   Stop
   '
   Dim R_n As Range
   Set R_n = range_�A����ő�s_namedrange(strName)
   Call PrintArrayOnRange(R_n, aAry, cols ,ROOT)
End Sub

Private Sub PrintArrayOnRange(R_n As Range, _
                              ByRef aAry() As Variant, _
                              Optional cols As Long = 0, _
                              Optional ROOT As Long = 0)
   ' cols > 1 �̂Ƃ��F
   ' �@�@�@��cols �Ŏw�肳�ꂽ�񐔂��Z���ɏ����o��
   ' cols = 0 �̂Ƃ��F
   ' �@�@�@���^����ꂽ�z��̓��e�����ׂăZ���ɏ����o��
   ' cols = -1 �̂Ƃ��F
   ' �@�@�@�����łɏ������܂�Ă���͈͂̂݃Z���ɏ����o��
   ' ==�t�H�[�}�b�g��������==
   ' ���O�Ŏw�肳�ꂽ�͈́i�Z���j�ɘA����ő�s�̌��o��K�p���āA�����o��
   ' �͈͂����s���m�肷��B
   ' �������w*�x�̃Z���ɂ͉������Ȃ��i�w*�x�̂܂܂ɂ���j
   ' 
   ' ���P�s�m��̔z��͂m��P�s�ɕϊ�
   Dim Ary As Variant
   Ary = Ary1C_If_Ary1R(aAry)
   ' ��R_n ���w������Ă���܂܁x����ς��邩
   ' �@cols ����ʂɎw�肷�邩
   ' �@������̊p�̃Z�������Ƃɂ��ĘA����̍ő�s�̈��ΏۂƂ��Đݒ肷��
   Set R_n = range_�A����ő�s_range(R_n.Cells(1,1), 1)
   Dim rowsA As Long
   Dim rowsR As Long
   rowsA = UBound(Ary, 1) - LBound(Ary, 1) + 1
   rowsR = R_n.Rows.Count
   If rowsA < rowsR Then rowsA = rowsR
   Dim colsA As Long
   colsA = UBound(Ary, 2) - LBound(Ary, 2) + 1
   colsR = R_n.Columns.Count
   If colsA < colsR Then colsA = colsR
   '
   Dim vAry As Variant
   Dim pvAry As Boolean
   pvAry = False
   ' ��cols�œ���Ȓl���w�肳�ꂽ�ꍇ�̏������ޗ� cols �̒���
   ' �@�Ə������ޔ͈͂̎擾�Ə���
   Select Case cols
      Case 0
         '��Ary �̑傫���ɍ��킹�Ĕ͈͂ɏ�������
         Set R_n = R_n.Resize(rowsA, colsA)
      Case -1
         '�����łɏ������܂�Ă���͈͂̑傫���ɍ��킹�ď�������
         vAry = R_n.Value
         pVary = True
         rowsA = R_n.Rows.Count
         colsA = R_n.Columns.Count
      Case Else
         '��cols�Ŏw�肳�ꂽ��
         Set R_n = R_n.Resize(rowsA, cols)
         colsA = cols
   End Select
   R_n.Clear
   ' ���񂲂Ƃɏ������w�肵�Ȃ��珑������
   Dim AryC As Variant
   ' ReDim AryC(LBound(Ary, 1) To UBound(Ary, 1), 1 To 1)
   ReDim AryC(1 To rowsA, 1 To 1)
   Dim r As Long
   Dim c As Long
   Dim nfl As String '.NumberFormatLocal
   Dim bls As Long   '.Borders.LineStyle
   ' Dim rw As Long
   Dim R_0 As Range  '������L����Z�������� Range
   Dim R_c As Range  '�������ޔ͈̗͂������ Range
   ' rw = UBound(Ary, 1) - LBound(Ary, 1) + 1
   ' ���E�E�E�������ލs���i�z��̍s���j��
   ' Set R_n = R_n.Resize(rw, 1)
   Set R_n = R_n.Resize(rowsA, 1)
   ' ���E�E�E�\�̍X�V����͈͂����߂�B�P��̂݁B
   '
'   Dim Lc As Long
'   Dim Uc As Long
'   Dim A1 As Boolean
'   On Error GoTo ARY1D
'   Lc = LBound(Ary, 2)
'   Uc = UBound(Ary, 2)
'   A1 = False
'   GoTo ARY1DEND
'ARY1D:
'   Lc = 1
'   Uc = 1
'   A1 = True
'ARY1DEND:
'
'ARYEND:
   ' For c = LBound(Ary, 2) To UBound(Ary, 2)
'   Stop
   
   ' For c = Lc To Uc
   For c = 1 To colsA
      ' ���E�E�E�z��̍ŏ��̗񂩂�Ō�̗�܂�
      Set R_0 = R_n.Resize(1, 1).Offset(ROOT, (c - 1))
      If ROOT < 0 Then
         nfl = R_0.NumberFormatLocal
         ' ���E�E�E�������ޗ�̐擪�s�Z������݂� ROOT�i-1�Ȃ�P��j��
         ' �E�E�E�E�I�t�Z�b�g�̃Z���ɐݒ肳�ꂽ�����������Ă���
         bls = xlLineStyleNone ' �G���[�̂Ƃ��̊���l
         On Error Resume Next
         bls = R_0.Borders.LineStyle
         On Error GoTo 0
      Else
         nfl = ""
         bls = xlLineStyleNone ' �w�肳��Ă���s���Ȃ��Ƃ��̊���l
      End If
      Set R_c = R_n.Offset(0, (c - 1))
      R_c.Font.Name = "BIZ UD�S�V�b�N"
      R_c.NumberFormatLocal = nfl
      R_c.Borders.LineStyle = bls
      ' ���E�E�E�񂲂Ƃɏ�����ݒ肷��
      ' For r = LBound(Ary, 1) To UBound(Ary, 1)
      For r = 1 To rowsA
         If pvAry Then
            If vAry(r, c) = "*" Then
               AryC(r, 1) = "*"
            Else
               ' �͈͂����z�񂪏������Ƃ��ɂ̓G���[�ɂȂ邪
               ' �G���[�̏ꍇ�̃f�t�H���g�Ƃ��ċ󕶎����ݒ�
               AryC(r, 1) = ""
               On Error Resume Next
               AryC(r, 1) = Ary(r, c)
               On Error GoTo 0
            End If
         Else
'         If A1 Then
'            AryC(r, 1) = Ary(r)
'         Else
            ' �͈͂����z�񂪏������Ƃ��ɂ̓G���[�ɂȂ邪
            ' �G���[�̏ꍇ�̃f�t�H���g�Ƃ��ċ󕶎����ݒ�
            AryC(r, 1) = ""
            On Error Resume Next
            AryC(r, 1) = Ary(r, c)
            On Error GoTo 0
'         End If
         End If
      Next r
      R_n.Offset(0, (c - 1)) = AryC
   Next c
End Sub

Function Ary1C_If_Ary1R(ByRef Ary As Variant) As Variant
   ' �������P�s�m��i�Y�����P�����j�̔z�񂾂����Ƃ������A
   ' �m�s�P��i�Y�����Q�����j�̔z��ɕϊ�
   ' �����łȂ��Ƃ��͂��̂܂�
   Dim LB As Long
   Dim UB As Long
   LB = LBound(Ary, 1)
   UB = UBound(Ary, 1)
   '
   On Error GoTo MAIN
   LB = LBound(Ary, 2)
   Ary1C_If_Ary1R = Ary
   Exit Function
   '
MAIN:
   On Error GoTo 0
   '
   Dim vAry() As Variant
   ReDim vAry(LB To UB, 1)
   For i = LB To UB
      vAry(i, 1) = Ary(i)
   Next i
   Ary1C_If_Ary1R = vAry
End Function

' ����敪�ʌʐR�������W�v�@�̒��ŁA�\�͈̔͂��X�V���邽�߂ɌĂ�
'
Private Sub ����敪�W�v�ʐR�������X�V(ByRef ����敪�ʌʐR������() As Long)
   ' Call ����敪�W�v�ʐR�������N���A
   Dim r0 As Long
   Dim r1 As Long
   Dim strName As String
   strName = "����W�v"
   Dim R����W�v As Range
   Set R����W�v = ThisWorkbook.Names(strName).RefersToRange
   r0 = R����W�v.Row
   r1 = ��̍ŏI�s(strName,, 1)
   Dim R����W�v������ As Range
   ' stop
   Set R����W�v������ = R����W�v.Offset(0, 1).Resize((r1 - r0 + 1), 1)
   '   �����E�E���t�����͈́E�E�������ƂɐV���Ȕ͈͂��w�肷��B
   R����W�v������.Clear
   R����W�v������.Font.Name = "BIZ UD�S�V�b�N"
   R����W�v������.Borders.LineStyle = xlContinuous
   R����W�v������ = ����敪�ʌʐR������
End Sub

' ����敪�ʌʐR�������W�v�@�̒��ŁA�\�͈̔͂��X�V���邽�߂ɌĂ�
'
Private Sub ����敪�W�v�ʐR�����z�X�V(ByRef ����敪�ʌʐR�����z() As Long)
   ' Call ����敪�W�v�ʐR�������N���A
   Dim r0 As Long
   Dim r1 As Long
   Dim strName As String
   strName = "����W�v"
   Dim R����W�v As Range
   Set R����W�v = ThisWorkbook.Names(strName).RefersToRange
   r0 = R����W�v.Row
   r1 = ��̍ŏI�s(strName,, 1)
   Dim R����W�v���z�� As Range
   ' stop
   Set R����W�v���z�� = R����W�v.Offset(0, 2).Resize((r1 - r0 + 1), 1)
   '   �����E�E���t�����͈́E�E�������ƂɐV���Ȕ͈͂��w�肷��B
   R����W�v���z��.Clear
   R����W�v���z��.Font.Name = "BIZ UD�S�V�b�N"
   R����W�v���z��.NumberFormatLocal = "#,##0,"
   ' ���E�E�E���z���~�P�ʂŕ\������
   R����W�v���z��.Borders.LineStyle = xlContinuous
   R����W�v���z�� = ����敪�ʌʐR�����z
End Sub

' �d���n�ʏW�v�@�ŗ��p
'
Private Sub �d���n��������(strName As String, _
                           ByRef nClass As Long, _
                           ByRef dicDIN As Dictionary)
   ' �d���n��������("�d���n�ʖ�", dic�d���ԍ�)
   ' dicDN - Distination ID Number
   Call MultiHomeDict_namedRange(strName, nClass, dicDIN)
End Sub

' �d���n�ʏW�v�@�ŗ��p
'
Private Sub �d���n�W�v���z�񐶐�(strName As String, _
                                 ByRef nClass As Long, _
                                 ByRef aryNTM As Variant)
   'ary�W�v���������z - Name, cont of Times, amount of Money
   ReDim aryNTM(1 To nClass, 1 To 3)
   Dim aryN() As Variant
   ReDim aryN(1 To nClass, 1 To 1)
   Call NamedRange2ary(strName, aryN)
   ' Stop
   For i = LBound(aryN, 1) To UBound(aryN, 1)
      aryNTM(i, 1) = aryN(i, 1)
   Next i
   ' Stop
   '
End Sub

' ������������������������������������������������������������������������
' ��������������������������������������������������������������������������
' �����ėp�v���V�W���i��ɂ��̃v���W�F�N�g�ȊO�ł��]�p���錩���݂̂�����́j
' ����

Private Sub �g�D���̃N���A()
   Dim �g�D������ As Range
   ' Set �g�D������ = range_�A����ő�s_namedrange�i"�g�D������"�j
   ' �g�D������.Clear
   range_�A����ő�s_namedrange�i"�g�D������"�j.Clear
   Exit Sub
   '
   ' Unused
   ' �ȉ��́A���ׂāA�g���Ȃ��Ȃ���
   Dim strName As String
   strName = "Range_�g�D����"
   Dim Range_�g�D���� As Range
   On Error Resume Next
   ' Clear ���Ăт� Range ��Ԃ��čđ���A���Ă̂��ԈႢ�B
   ' �ԈႢ�������̂ŁA�G���[���N���Ă���̂� On Error ��
   ' �͂��Ă���A�Ƃ����ԈႢ�E�E�E�B�Ȃ�Ƃ��B
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
' �g�D�������敪�̏W�v�����ŁA�O�҂ł���Ε�����A��҂͂P��̎����ϊ�
'
Private Sub SingleHomeDict_namedRange(strName As String, _
                                      ByRef nClass As Long, _
                                      ByRef dic���� As Dictionary, _
                                      Optional cx As Long = 1)
   Call MultiHomeDict_namedRange(strName, _
                                 nClass, _
                                 dic����, _
                                 cx, _
                                 True)
End Sub

' --- �͈͂�Dictionary�ɕϊ�����
' �͈͂��e�Z���̓��e�� key �Ƃ��ċL�ڂ��ꂽ�s�� value �Ƃ��� Dictionary �ɕϊ�����
' ���� key �������̍s�Ɍ����ꍇ�������A���̏ꍇ�̂��߂ɁAvalue �͋L�ڂ��ꂽ�s�ԍ�
' �̃��X�g�Ƃ��鎫���B
' ��T���� pS �� True �ɂ���ƁAvalue �͋L�ڂ��ꂽ�s�ԍ����̂��̂ƂȂ�B
'
Private Sub MultiHomeDict_namedRange(strName As String, _
                                     ByRef nClass As Long, _
                                     ByRef dic���� As Dictionary, _
                                     Optional cx As Long = 0, _
                                     Optional pS As Boolean = False)
   ' �l�̊K������Ԃ��K�v������B
   ' �����ł̎����̍쐬�̖ړI�́A������ key �ɓ��� Value ��Ԃ������݂�
   ' �ȒP�Ɏ������邱�ƂȂ̂ŁA����ނ� Value ��Ԃ����ƂɂȂ��Ă���̂�
   ' �ɂ��ẮA�������\�������Ƃ��ɂ킩����̂Ƃ��ĕԂ����Ƃ����߂���B
   ' ��Q�����Ƃ��ĎQ�Ɠn�����Ă�����Ă����ĕԂ��B
   '
   ' ��P�����F�����ɂ�����e���L�ڂ��Ă���͈͂ɖ��t�������O�i������j
   ' ��Q�����F�����̎��� Value �̃N���X����Ԃ����߂̕ϐ��i�Q�Ɠn���j
   ' �@�@�@�@�@�ʖ����������s�͈̔͂ō\������Ă��邩���Ԃ����B
   ' ��R�����F��������鎫��
   ' �E�E�E�E�E���͈͂̊e�Z���̓��e�� key
   ' �E�E�E�E�E�����ꂪ�L�ڂ��ꂽ�s�ԍ��i�͈͓��ł̍s�ԍ��j�� value
   ' ��S�����F�i�I�v�V���i���j�����͈̗͂񐔂𐧌�����Ƃ����̗�
   ' �E�E�E�E�E�������Ȃ��Ƃ��ɂ́w 0 �x�i1��菬�����l�j�Ƃ���B
   ' ��P�������Z���P�����̂Ƃ��ɂ́A
   ' �͈͂́w��̍ŏI�s�x�Ɓw�����s�̍ŏI��_range�x�܂Ŋg�傳���B
   ' ��T�����F cx - �͈͂̃J���������Œ�l�Ŏw�肷��Ƃ����̒l�B
   ' �@�@�@�@�@�w�肳��Ȃ��Ƃ��͒l 0 �Ƃ݂Ȃ��B�ł������s�̃J��������
   ' �@�@�@�@�@�w�肳�ꂽ���̂Ƃ݂Ȃ��B
   ' ��U�����F pS - SingleHomeDict �Ƃ��ČĂԂƂ��� True �ɂ���B
   ' �@�@�@�@�@�w�肳��Ȃ��Ƃ��͒l False �Ƃ݂Ȃ��B�����̍s�ԍ��ւ̋A��
   ' �@�@�@�@�@�����肦����̂Ƃ��āi����ɍs�ԍ����P�����ł��A�v�f���P�́j
   ' �@�@�@�@�@�z���Ԃ������ƂȂ�B
   '
   ' Debug.Print strName
   ' Stop
   Dim R_n As Range
   Set R_n = ThisWorkbook.Names(strName).RefersToRange
   Set R_n = range_�A����ő�s_range(R_n)
   '
   Dim var����() As Variant
   var���� = R_n.Value
   nClass = R_n.Rows.count
   '
   Dim U1 As Long
   U1 = UBound(var����, 1)
   U2 = UBound(var����, 2)
   If (cx > 0) Then
      U2 = cx
   End If
   
   Dim d0 As Long
   If pS Then
      For i = 1 To U1
         d0 = i
         ' �� ���̎����ł� value �ƂȂ�d0 �͒l�B
         For j = 1 To U2
            d = var����(i, j)
            If d = "" Then Exit For
            If dic����.Exists(d) Then
               ' ���̎����ł� key �ł��� d ��
               ' ���łɓo�^����Ă���Ƃ��́A
               ' �܂�Q��ڈȍ~�́A��������B
            Else
               dic����.Add d, d0
            End If
         Next j
      Next i
   Else
      For i = 1 To U1
         d0 = i
         ' �� ���̎����ł� value �̗v�f�ƂȂ�d0 �͒l�B
         For j = 1 To U2
            d = var����(i, j)
            If d = "" Then Exit For
            Dim dv() As Variant ' �Ē�`���O��
            If dic����.Exists(d) Then
               dv = dic����.Item(d)
               k = UBound(dv, 1)
               ReDim Preserve dv(1 To k + 1)
               dv(k + 1) = d0
               dic����.Item(d) = dv
            Else
               ReDim Preserve dv(1 To 1)
               dv(0 + 1) = d0
               dic����.Add d, dv
            End If
            Erase dv ' ���̉�̂��߂ɏ���
         Next j
      Next i
   End if
End Sub

' �d���n�W�v���z�񐶐��@�ŗ��p
' 
' ���t����ꂽ�͈̗͂���s�����Ɋg�������͈͂�z��Ɋi�[
'
Private Sub NamedRange2ary(strName As String, _
                           ByRef aryV As Variant)
   ' ��P�����F�͈͂ɂ������O�@���P�@
   ' ��Q�����F��L�͈̔͂��g�����ē��e���i�[���ĕԂ��z��
   ' ���P�@���̖��O���Ȃ������炻��Ȃ�̃G���[��Ԃ������Ƃ���
   ' 
   Dim rngC As Range
   Set rngC = ThisWorkbook.Names(strName).RefersToRange
   ' Set rngC = ��̍ŏI�s
   ' stop
   Set rngC = range_��̍ŏI�s_range(rngC,, 1)
   aryV = rngC
End Sub

