Attribute VB_Name = "x2p_Lib1"
' -*- coding:shift_jis-dos -*-

'./x2p_Lib1.bas
' �����Ӂ�
' �@���̃X�N���v�g�ł͊֐���`���Ă���̂ŕW�����W���[���ɒu���Ȃ��Ƌ@�\���Ȃ��B
' �@���̂��߃t�@�C���̐擪�Ɂ@Attribute VB_Name = "x2p_Lib1"�@�ƋL�q���Ă����B
' �@�i�O���t�@�C���Ƃ��ĕҏW���Ă���Ƃ��̂݌�����j

Function �Z���̌Œ�F(�Z��)
     Dim a
     '// a = �Z��.Interior.ColorIndex
     '// a = �Z��.FormatConditions.Interior.Color ' - ����͂��܂��s���Ȃ�
     '// �@�@�����t���w�i�F�𓾂����Ƃ��Ă�
     '// �@�@���[�U�֐��Ƃ��Ď��s����ƃG���[���N����̂Ŏg���Ȃ�
     '// a = �Z��.DisplayFormat.Interior.Color

     a = �Z��.Interior.ColorIndex
     �Z���̌Œ�F = a
End Function

Public Sub �J�n���}��()
    Application.ScreenUpdating = False '��ʕ`����~
    Application.Cursor = xlWait '�E�G�C�g�J�[�\��
    Application.EnableEvents = False '�C�x���g��}�~
    Application.DisplayAlerts = False '�m�F���b�Z�[�W��}�~
    Application.Calculation = xlCalculationManual '�v�Z���蓮��
End Sub

Public Sub �I�������()
    Application.StatusBar = False '�X�e�[�^�X�o�[������
    Application.Calculation = xlCalculationAutomatic '�v�Z��������
    Application.DisplayAlerts = True '�m�F���b�Z�[�W���J�n
    Application.EnableEvents = True '�C�x���g���J�n
    Application.Cursor = xlDefault '�W���J�[�\��
    Application.ScreenUpdating = True '��ʕ`����J�n
End Sub

Function ��̍ŏI�s(n As String, _
                          Optional k As Long = 1, _
                          Optional q As Long = 0) As Long
   ' n - �J�n����Z���i�͈͂ł��悢�j�ɖ��t�������O�i������j
   ' k - ���I�v�V���������͈̔͂̒��̗�ԍ�
   ' q - ���I�v�V����������ڂ̋󔒂��I���Ƃ݂Ȃ����F0���Ɩ�����
   Dim R_n As Range
   Set R_n = ThisWorkbook.Names(n).RefersToRange
   If IsMissing(k) Then
      If IsMissing(q) Then
         ��̍ŏI�s = ��̍ŏI�s_range(R_n)
      Else
         ��̍ŏI�s = ��̍ŏI�s_range(R_n, , q)
      End If
   Else
      If IsMissing(q) Then
         ��̍ŏI�s = ��̍ŏI�s_range(R_n, k)
      Else
         ��̍ŏI�s = ��̍ŏI�s_range(R_n, k, q)
      End If
   End If
End Function

Function �s�̍ŏI��(n As String, _
                          Optional k As Long = 1, _
                          Optional q As Long = 0) As Long
   ' n - �J�n����Z���i�͈͂ł��悢�j�ɖ��t�������O�i������j
   ' k - ���I�v�V���������͈̔͂̒��̍s�ԍ�
   ' q - ���I�v�V����������ڂ̋󔒂��I���Ƃ݂Ȃ����F0���Ɩ�����
   Dim R_n As Range
   Set R_n = ThisWorkbook.Names(n).RefersToRange
   If IsMissing(k) Then
      If IsMissing(q) Then
         �s�̍ŏI�� = �s�̍ŏI��_range(R_n)
      Else
         �s�̍ŏI�� = �s�̍ŏI��_range(R_n, , q)
      End If
   Else
      If IsMissing(q) Then
         �s�̍ŏI�� = �s�̍ŏI��_range(R_n, k)
      Else
         �s�̍ŏI�� = �s�̍ŏI��_range(R_n, k, q)
      End If
   End If
End Function

Function ��̍ŏI�s_range(ByRef R_n As Range, _
                          Optional k As Long = 1, _
                          Optional q As Long = 0) As Long
   ' R_n - �J�n����Z�����܂ޔ͈�-Range-
   ' k - ���I�v�V���������͈̔͂̒��̗�ԍ�
   ' q - ���I�v�V����������ڂ̋󔒂��I���Ƃ݂Ȃ����F0���Ɩ�����
   Dim R1 As Long
   Dim r2 As Long
   Dim mr As Long
   Dim s As Variant
   mr = Rows.count ' �s�̍ő�l�E�E�E�����ŖO�a����B
   Set s = R_n.Columns(k)
   r2 = s.Row
   q = q - 1
   Do
      R1 = r2
      Set s = s.End(xlDown)
      r2 = s.Row
      q = q - 1
   Loop While Not ((r2 >= mr) Or (q = 0))
   ��̍ŏI�s_range = R1
End Function

Function range_��̍ŏI�s_range(ByRef R_n As Range, _
                                Optional k As Long = 1, _
                                Optional q As Long = 0) As Range
   ' R_n - �J�n����Z�����܂ޔ͈�-Range-
   ' k - ���I�v�V���������͈̔͂̒��̗�ԍ�
   ' q - ���I�v�V����������ڂ̋󔒂��I���Ƃ݂Ȃ����F0���Ɩ�����
   Dim r0 As Long
   Dim r1 As Long
   Dim r2 As Long
   Dim mr As Long
   Dim s As Variant
   r0 = R_n.Row
   mr = Rows.count ' �s�̍ő�l�E�E�E�����ŖO�a����B
   Set s = R_n.Columns(k)
   r2 = s.Row
   q = q - 1
   Do
      r1 = r2
      Set s = s.End(xlDown)
      r2 = s.Row
      q = q - 1
   Loop While Not ((r2 >= mr) Or (q = 0))
   ' ��̍ŏI�s_range = R1
   Set range_��̍ŏI�s_range = R_n.resize((r1 - r0 + 1))
   '   �@�@�@�@�@�@�@�@�@�@�@�@�@�@��͏ȗ����čs�̂݊g����
   '���߂�l���͈͂܂�w�I�u�W�F�N�g�x�Ȃ̂� Set ���g���I�I
End Function

Function �s�̍ŏI��_range(R_n As Range, _
                          Optional k As Long = 1, _
                          Optional q As Long = 0) As Long
   ' R_n - �J�n����Z�����܂ޔ͈�-Range-
   ' k - ���I�v�V���������͈̔͂̒��̍s�ԍ�
   ' q - ���I�v�V����������ڂ̋󔒂��I���Ƃ݂Ȃ����F0���Ɩ�����
   Dim c1 As Long
   Dim c2 As Long
   Dim mc As Long
   Dim s As Variant
   mc = Columns.count ' ��̍ő�l�E�E�E�����ŖO�a����B
   Set s = R_n.Rows(k)
   c1 = 0
   c2 = s.Column
   Do
      c1 = c2
      Set s = s.End(xlToRight)
      c2 = s.Column
      q = q - 1
   Loop While Not ((c2 >= mc) Or (q = 0))
   �s�̍ŏI��_range = c1
End Function

Function �����s�̍ŏI��_range(R_n As Range, _
                              Optional q As Long = 0) As Long
   ' R_n - �J�n����Z�����܂ޔ͈�-Range-
   ' q - ���I�v�V����������ڂ̋󔒂��I���Ƃ݂Ȃ����F0���Ɩ�����
   Dim c1 As Long
   Dim c2 As Long
   Dim mc As Long
   Dim s As Range ' yet Variant
   mc = Columns.count ' ��̍ő�l�E�E�E�����ŖO�a����B
   Dim k As Long
   Dim cx As Long
   cx = 0
   For k = 1 To R_n.Rows.count
      Set s = R_n.Rows(k)
      c1 = 0
      ' c2 = 0
      c2 = s.Column
      ' �����l�͂����Őݒ肵�Ă����Ȃ��Ƃ����Ȃ����B
      Do
         c1 = c2
         Set s = s.End(xlToRight)
         c2 = s.Column
         q = q - 1
      Loop While Not ((c2 >= mc) Or (q = 0))
      If cx < c1 Then cx = c1
   Next k
   �����s�̍ŏI��_range = cx
End Function

Function range_�A����ő�s_range(R_n As Range, _
                                    Optional q As Long = 0) As Range
   '
   ' ������̍ŏI�s���g���āA�͈͂��Ђ낰��
   ' �E�i�擪��̍s���j�~�i�Œ��s�̗񐔁j��͈͂Ƃ���
   ' R_n - �J�n����Z�����܂ޔ͈�-Range-
   ' q - ���I�v�V����������ڂ̋󔒂��I���Ƃ݂Ȃ����F0���Ɩ�����
   Dim r0 As Long
   Dim c0 As Long
   r0 = R_n.Row
   c0 = R_n.Column
   Dim nr As Long
   Dim nC As Long
   nr = R_n.Rows.count
   nC = R_n.Columns.count
   If (nr = 1) And (nC = 1) Then
      ' Call ExpandRangeCont(R_n, strName, cx)
      Dim rz As Long
      rz = ��̍ŏI�s_range(R_n)
      Set R_n = R_n.Resize((rz - r0 + 1), 1)
      Dim cz As Long
      cz = �����s�̍ŏI��_range(R_n)
      ' cz �� 0 �ɂȂ��Ă��܂��̂͂Ȃ��B
      Set R_n = R_n.Resize((rz - r0 + 1), (cz - c0 + 1))
   End If
   Set range_�A����ő�s_range = R_n
End Function

