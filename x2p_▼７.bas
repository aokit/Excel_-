Attribute VB_Name = "x2p_���V"
Option Explicit

' -*- coding:shift_jis-dos -*-

'./x2p_���V.bas

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
   '
   ' Stop
   '
   Dim str���F�L�^() As String
   Dim dic�d���ԍ� As New Dictionary
   Dim ary�W�v��() As String
   Dim ary�W�v���������z() As Variant 'Long
   ' Dim ary�������d��() As String
   Dim ary�������d��() As Variant
   Dim nClass As Long
   '
   Call ���F�L�^�z��i�[("���F�L�^", str���F�L�^)
   Call �d���n�W�v���ԃt�B���^(str���F�L�^, 3)
   '
   Stop
   '
   Call �d���n��������("�d���n�ʖ�", nClass, dic�d���ԍ�)
   ' ���E�E�E�����Ł@Debug.Print(dic�d���ԍ�("�C�^����")(1))�@�Ƃ����
   ' �@�@�@�@dic�d���ԍ��ɓǂݍ��܂�Ă��邱�Ƃ��킩��
   Stop
   '
   Call �d���n�W�v���z�񐶐�("�d���n�ʖ�", nClass, ary�W�v���������z)
   ' ���E�E�E���̂��ƁAstr���F�L�^(2,3) �ȂǃA�N�Z�X����Ɨ�����B
   ' ����ł͂Ȃ��āA�P�����x���āAExcel���̂�������B
   '
   Stop
   '
   Call �d���n�W�v(str���F�L�^, 11, 10, dic�d���ԍ�, _
                   ary�W�v���������z, ary�������d��, 3)
   ' Stop
   ' ���E�E�E���̂��Ɓ@�W�v���������z�@�Ɓ@�������d���@��\������
   '
   Stop
   '
   ' ���d���n�W�v���������z�\��
   Call PrintArrayOnNamedRange("�d���n�W�v", ary�W�v���������z, 3, -4)
   ' ���������ʖ��\��
   ' Stop
   'Call PrintArrayOnNamedRange("�������ʖ�", ary�W�v���������z, 1)
   '
   ' �\�[�g���A����͈̔͂ɋL��
   ' 
   Dim i As Long
   
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

Private Sub ���F�L�^�z��i�[(strNameD As String, _
                             str���F�L�^() As String)
   Call NamedRangeSQ2ArrStr(strNameD, str���F�L�^)
End Sub

Private Sub �d���n�W�v���ԃt�B���^(ByRef str���F�L�^() As String, _
                                   ByVal iC_pTerm As Long)
   '
   Dim U2 As Long
   U2 = UBound(str���F�L�^, 1)
   Dim iC_���F�� As Long
   iC_���F�� = 2
   Dim pTerm As Boolean
   ' Dim iC_pTerm As Long
   ' iC_pTerm = 3
   '
   Dim i As Long
   For i = 2 To U2
      pTerm = p�L������(str���F�L�^(i, iC_���F��))
      ' ���ꂾ���ŗ����邩�H�� Yes
      ' ��p�����ɂ��ƁA���ꂾ���ŗ�����B������� F8 �łQ���قǂ����
      ' ���̂��� F5 �ł����v�E�E�E�Ȃ񂾂낱��B
      ' strx = str���F�L�^(i, iC_���F��)
      ' ���ꂾ���ŗ����邩�H�� No
      str���F�L�^(i, iC_pTerm) = Str(pTerm)
   Next i
End Sub


Sub �d���n�W�v(ByRef str���F�L�^() As String, _
                       ByVal iC_�d���n As Long, _
                       ByVal iC_���z As Long, _
                       ByRef dicDIN As Dictionary, _
                       ByRef aryNTM() As Variant, _
                       ByRef aryYAD() As Variant, _
                       ByVal iC_pTerm)
   '
   Stop
   '
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
   Dim TaryYAD(1 To 200) As String
   ' ���E�E�E�������̎d�����i�[����i�Ԃ��Ƃ��ɂ͖��g�p�����j
   Dim nYAD As Long
   nYAD = 0
   '
   ' Dim str���F�L�^() As String
   ' Call ���F�L�^�ǂݎ��(str���F�L�^)
   ' Call NamedRangeSQ2ArrStr("���F�L�^", str���F�L�^)
   ' Call NamedRangeSQ2ArrStr(strNameD, str���F�L�^)
   '
   ' Stop
   '
   ' ��������@�z��@str���F�L�^�@�ɑ΂��ā@���R�^���T���Q�l�ɂ��ďW�v�������s���B
   ' Dim U2 As Long
   ' �����Ƌ��z���������iEmpty�ł͂Ȃ�0�Ɂj
   ' U2 = UBound(aryNTM, 1)
   ' Dim i As Long
   ' For i = 1 To U2
   '    aryNTM(i, 2) = 0
   '    aryNTM(i, 3) = 0
   ' Next i
   ' str���F�L�^�@�̑S�s���X�L����
   U2 = UBound(str���F�L�^, 1)
   ' Dim iC_���F�� As Long
   ' iC_���F�� = 2
   Dim IX As Long
   Dim str�d���n As String
   Dim str���z As String
   Dim strx As String
   Dim pterm As Boolean
   Dim aryIX() As Variant
   ' Dim iC_pTerm As Long
   ' iC_pTerm = 3
   '
   Stop
   ' --- ��������X�e�b�v���s�Ŋm�F
   ' For i = 2 To U2
   '    pTerm = p�L������(str���F�L�^(i, iC_���F��))
   '    ' ���ꂾ���ŗ����邩�H�� Yes
   '    ' ��p�����ɂ��ƁA���ꂾ���ŗ�����B������� F8 �łQ���قǂ����
   '    ' ���̂��� F5 �ł����v�E�E�E�Ȃ񂾂낱��B
   '    ' strx = str���F�L�^(i, iC_���F��)
   '    ' ���ꂾ���ŗ����邩�H�� No
   '    str���F�L�^(i, iC_pTerm) = Str(pTerm)
   ' Next i
   '
   ' �����܂łł��łɖ�肪�N���Ă���E�E�E
   '
   ' Stop
   '
   For i = 2 To U2
      ' pTerm = p�L������(str���F�L�^(i, iC_���F��))
      pTerm = CStr(str���F�L�^(i, iC_pTerm))
      If pTerm Then
      End If
   Next i
   '
   Stop
   '
   For i = 2 To U2
      pTerm = CStr(str���F�L�^(i, iC_pTerm))
      If pTerm Then
      ' p�L�������@���Ă΂Ȃ���Α��v���H
      ' ��΂Ȃ���΂����ł͗����Ȃ����A���Ƃ��痎���邱�Ƃ������B
      ' If True Then
         str�d���n = str���F�L�^(i, iC_�d���n)
         str���z = str���F�L�^(i, iC_���z)
         '
         Stop
         '
         If dicDIN.Exists(str�d���n) Then
            ' Dim aryIX() As Variant
            aryIX = dicDIN.Item(str�d���n)
            ' ���E�E�E�d���n�� Identification Number�i�������������蓾��̂ł̔z��j
            Dim U22 As Long
            u22 = UBound(aryIX, 1)
            Dim j As Long
            For j = 1 To U22
               If j > 1 Then Debug.Print j
               IX = aryIX(j)
               aryNTM(IX, 2) = aryNTM(IX, 2) + 1
               ' aryNTM(IX, 3) = aryNTM(IX, 3) + Val(str���z)
               aryNTM(IX, 3) = aryNTM(IX, 3) + Val(str���z) / U22
               ' ���E�E�E�ʖ������񐔂ŋϓ��������ĐώZ
            Next j
            ' Erase aryIX
            ' �����ύX����
            ''' �s�v�ł́H
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
   ' --- �����܂ł̎��s�ɂȂɂ���肪����B
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
