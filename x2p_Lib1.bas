Attribute VB_Name = "x2p_Lib1"
Option Explicit

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
   ��̍ŏI�s = ��̍ŏI�s_range(R_n, k, q)
End Function

Function a_��̍ŏI�s(n As String, _
                          Optional k As Long = 1, _
                          Optional ByVal q As Long = 0) As Long
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
   '
   Dim R_n As Range
   Set R_n = ThisWorkbook.Names(n).RefersToRange
   �s�̍ŏI�� = �s�̍ŏI��_range(R_n, k, q)
End Function

Function ��̍ŏI�s_range(ByRef R_n As Range, _
                          Optional k As Long = 1, _
                          Optional g As Long = 0) As Long
   ' R_n - �J�n����Z�����܂ޔ͈�-Range-
   ' k - ���I�v�V���������͈̔͂̒��̗�ԍ�
   ' q - ���I�v�V����������ڂ̋󔒂��I���Ƃ݂Ȃ����F0���Ɩ�����
   ' q=1 �̂Ƃ��F
   ' �E���̃Z�����󔒂łȂ��Ƃ��A�l�̂���Z���̘A���̍Ō�̃Z���Ɉړ�
   ' �@�i���̃Z�����󔒂̃Z���ɂȂ�Z���Ɉړ��j
   ' �E���̃Z�����󔒂̂Ƃ��A�l�̂���Z���̘A���̍ŏ��̃Z���Ɉړ�
   ' �@�i��̃Z�����󔒂̃Z���ɂȂ�Z���Ɉړ��j
   ' �@�^�܂��́A�X�v���b�h�V�[�g�̘_����̍ő�s�̃Z���Ɉړ�
   Dim q As Long
   q = g
   ' ---
   Dim r1 As Long
   Dim r2 As Long
   Dim mr As Long
   Dim s As Range
   Dim w As Long
   w = 0
   mr = Rows.count ' �s�̍ő�l�E�E�E�����ŖO�a����B
   Set s = R_n.Columns(k)
   r2 = s.Row
   If (mr = r2) Or (mr = s.Rows.Count) Then
      Debug.Print("��̍ŏI�s_range�ɗ^����ꂽ�͈͂��V�[�g�̍ŉ��s�ɒB���Ă��܂�")
      ��̍ŏI�s_range = mr
      Exit Function
   End If
   '
   Do
      r1 = r2
      ' ---
      ' Set s = s.End(xlDown)
      ' �����ύX����
      ' s �������Z���^�s�͈̔͂̂��Ƃ�����̂ŁF
      If "" = s.Cells((s.Rows.Count + 1), 1).Value Then
         If w > 0 Then
            w = 0
            Set s = s.End(xlDown)
         Else
            w = 1
         End If
      Else
         w = 0
         Set s = s.End(xlDown)
      End If
      ' ---
      r2 = s.Row
      q = q - 1
      ' ......... ((r2 >= mr) Or (q = 0) Or (r1 = r2))
      ' �����ύX����
   Loop While Not ((r2 >= mr) Or (q = 0))
   ��̍ŏI�s_range = r2
End Function

Function range_��̍ŏI�s_range(ByRef R_n As Range, _
                          Optional k As Long = 1, _
                          Optional q As Long = 0) As Range
   ' R_n - �J�n����Z�����܂ޔ͈�-Range-
   ' k - ���I�v�V���������͈̔͂̒��̗�ԍ�
   ' q - ���I�v�V����������ڂ̋󔒂��I���Ƃ݂Ȃ����F0���Ɩ�����
   Dim r2 As Long
   Dim s As Range
   Set s = R_n.Columns(k)
   r2 = s.Row
   ' Set range_��̍ŏI�s_range = s.Offset((��̍ŏI�s_range(R_n, k, q) - r2), 0)
   Set range_��̍ŏI�s_range = s.Resize((��̍ŏI�s_range(R_n, k, q) - r2 + 1), 1)
   '
End Function

Function range_��̍ŏI�s_namedrange(strName As String, _
                                     Optional k As Long = 1, _
                                     Optional q As Long = 0) As Range
   Dim R_n As range
   Set R_n = ThisWorkbook.Names(strName).RefersToRange
   Set range_��̍ŏI�s_namedrange = _
        range_��̍ŏI�s_range(R_n, k, q)
   '
End Function

Function a_range_��̍ŏI�s_range(ByRef R_n As Range, _
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

Function �ύX�O_�s�̍ŏI��_range(R_n As Range, _
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

Function �s�̍ŏI��_range(R_n As Range, _
                          Optional k As Long = 1, _
                          Optional g As Long = 0) As Long
   ' R_n - �J�n����Z�����܂ޔ͈�-Range-
   ' k - ���I�v�V���������͈̔͂̒��̍s�ԍ�
   ' q - ���I�v�V����������ڂ̋󔒂��I���Ƃ݂Ȃ���
   ' �F�w�肳��Ȃ���� 1 �w�肪 0 ���Ɩ������B������͂܂��������B
   ' ���ǂ����B
   '
   ' �����ł��邪�AByVal �Ƃ��Ă����Ȃ��ƁA
   ' �Ăяo�������ƁAq �̒l���ς���Ă��܂��̂ŁA�J��Ԃ��̂Ȃ���
   ' q ���w�肵�ČĂяo���ƁA�\�����ʌ��ʂɂȂ�B
   ' �֐��Ȃ̂ɁA�����ɂ��Ă����A�f�t�H���g�� ByRef �Ƃ����E�E�E
   '
   Dim q As Long
   q = g
   ' ---
   Dim c1 As Long
   Dim c2 As Long
   Dim mc As Long
   Dim s As Range
   Dim w As long
   w = 0
   mc = Columns.count ' ��̍ő�l�E�E�E�����ŖO�a����B
   Set s = R_n.Rows(k)
   c2 = s.Column
   If (mc = c2) Or (mc = s.Columns.Count) Then
      Debug.Print("�s�̍ŏI��_range�ɗ^����ꂽ�͈͂��V�[�g�̍ŉE��ɒB���Ă��܂�")
      �s�̍ŏI��_range = mc
      Exit Function
   End If
   '
   Do
      c1 = c2
      ' ---
      ' Set s = s.End(xlToRight)
      ' �����ύX����
      ' s �������Z���^�񂩂�Ȃ�͈͂ł��邱�Ƃ�����̂ŁF
      If "" = s.Cells(1, (s.Columns.Count + 1)).Value Then
         If w > 0 Then
            w = 0
            Set s = s.End(xlToRight)
         Else
            w =1
         End If
      Else
         w = 0
         Set s = s.End(xlToRight)
      End If
      ' ---
      c2 = s.Column
      q = q - 1
      ' ......... ((c2 >= mc) Or (q = 0) Or (c1 = c2))
   ' �����ύX����
   Loop While Not ((c2 >= mc) Or (q = 0))
   �s�̍ŏI��_range = c2
End Function

Function range_�s�̍ŏI��_range(ByRef R_n As Range, _
                                Optional k As Long = 1, _
                                Optional q As Long = 0) As Range
   ' R_n - �J�n����Z�����܂ޔ͈�-Range-
   ' k - ���I�v�V���������͈̔͂̒��̍s�ԍ�
   ' q - ���I�v�V����������ڂ̋󔒂��I���Ƃ݂Ȃ����F0���Ɩ�����
   Dim c2 As Long
   Dim s As Range
   Set s = R_n.Rows(k)
   c2 = s.Column
   ' Set range_�s�̍ŏI��_range = s.Offset(0, (�s�̍ŏI��_range(R_n, k, q) - c2))
   Set range_�s�̍ŏI��_range = s.Resize(1, (�s�̍ŏI��_range(R_n, k, q) - c2 + 1))
   '
End Function

Function range_�s�̍ŏI��_namedrange(strName As String, _
                                     Optional k As Long = 1, _
                                     Optional q As Long = 0) As Range
   Dim R_n As range
   Set R_n = ThisWorkbook.Names(strName).RefersToRange
   Set range_�s�̍ŏI��_namedrange = _
        range_�s�̍ŏI��_range(R_n, k, q)
   '
End Function

Function a_�����s�̍ŏI��_range(R_n As Range, _
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
      qi = q
      Set s = R_n.Rows(k)
      ' c1 = 0
      ' c2 = 0
      c2 = s.Column
      ' �����l�͂����Őݒ肵�Ă����Ȃ��Ƃ����Ȃ����B
      Do
         c1 = c2
         Set s = s.End(xlToRight)
         c2 = s.Column
         qi = qi - 1
      Loop While Not ((c2 >= mc) Or (qi = 0))
      ' If cx < c1 Then cx = c1
      If cx < c2 Then cx = c2
   Next k
   �����s�̍ŏI��_range = cx
End Function

Function �����s�̍ŏI��_range(R_n As Range, _
                              Optional q As Long = 0) As Long
   ' R_n - �J�n����Z�����܂ޔ͈�-Range-
   ' q - ���I�v�V����������ڂ̋󔒂��I���Ƃ݂Ȃ����F0���Ɩ�����
   '
   Dim cx As Long
   cx = 0
   Dim c2 As Long
   c2 = 0
   Dim r2 As Long
   r2 = R_n.Row
   ' R_n.Rows.Count > 1 �ł��邱�Ƃ�����̂ŁF
   If R_n.Rows.Count > 1 then
      Set R_n = R_n.Resize((��̍ŏI�s_range(R_n.Cells(1, 1), 1, 1) - r2), 1)
   End If
   ' ��R_n.Rows.Count = 1 �Ȃ� R_n�͂��̂܂܁B
   '
   Dim k As Long
   For k = 1 To R_n.Rows.Count
      c2 = �s�̍ŏI��_range(R_n, k, q)
      If cx < c2 Then cx = c2
   Next k
   '
   �����s�̍ŏI��_range = cx
End Function

Function range_�A����ő�s_range(R_n As Range, _
                                  Optional q As Long = 1) As Range
   '
   ' ������̍ŏI�s���g���āA�͈͂��Ђ낰��
   ' �E�i�擪��̍s���j�~�i�Œ��s�̗񐔁j��͈͂Ƃ���
   ' R_n - �J�n����Z�����܂ޔ͈�-Range-
   ' q - ���I�v�V����������ڂ̋󔒂��I���Ƃ݂Ȃ����F0���Ɩ�����
   '
   ' �A����ő�s�����߂�ꍇ�ɂ́A�f�t�H���g�� q = 1 �ɌŒ肵�Ă݂�B
   ' 
   Dim r0 As Long
   Dim c0 As Long
   r0 = R_n.Row
   c0 = R_n.Column
   Dim nr As Long
   Dim nC As Long
   nr = R_n.Rows.count
   nC = R_n.Columns.count
   Dim qi As long
   If (nr = 1) And (nC = 1) Then
      ' Call ExpandRangeCont(R_n, strName, cx)
      qi = q
      Dim rz As Long
      rz = ��̍ŏI�s_range(R_n, , qi)
      Set R_n = R_n.Resize((rz - r0 + 1), 1)
      qi = q
      Dim cz As Long
      cz = �����s�̍ŏI��_range(R_n, qi)
      ' cz �� 0 �ɂȂ��Ă��܂��̂͂Ȃ��B
      Set R_n = R_n.Resize((rz - r0 + 1), (cz - c0 + 1))
   End If
   Set range_�A����ő�s_range = R_n
End Function

Function range_�A����ő�s_namedrange(strRangeName As String, _
                                    Optional q As Long = 1) As Range
   '
   ' ������̍ŏI�s���g���āA�͈͂��Ђ낰��
   ' �E�i�擪��̍s���j�~�i�Œ��s�̗񐔁j��͈͂Ƃ���
   ' R_n - �J�n����Z�����܂ޔ͈�-Range-
   ' q - ���I�v�V����������ڂ̋󔒂��I���Ƃ݂Ȃ����F0���Ɩ�����
   '
   ' �A����ő�s�����߂�ꍇ�ɂ́A�f�t�H���g�� q = 1 �ɌŒ肵�Ă݂�B
   ' 
   ' ���O����������Z������A�s��񐔂̘A���s�͈̔͂Ɋg�����Ĕ͈͂�Ԃ�
   Dim R_n As Range
   Set R_n = ThisWorkbook.Names(strRangeName).RefersToRange
   ' Set R_n =...
   Set range_�A����ő�s_namedrange = range_�A����ő�s_range(R_n, q)
End Function
