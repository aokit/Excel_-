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
�@�@Application.ScreenUpdating = False '��ʕ`����~
�@�@Application.Cursor = xlWait '�E�G�C�g�J�[�\��
�@�@Application.EnableEvents = False '�C�x���g��}�~
�@�@Application.DisplayAlerts = False '�m�F���b�Z�[�W��}�~
�@�@Application.Calculation = xlCalculationManual '�v�Z���蓮��
End Sub

Public Sub �I�������()
�@�@Application.StatusBar = False '�X�e�[�^�X�o�[������
�@�@Application.Calculation = xlCalculationAutomatic '�v�Z��������
�@�@Application.DisplayAlerts = True '�m�F���b�Z�[�W���J�n
�@�@Application.EnableEvents = True '�C�x���g���J�n
�@�@Application.Cursor = xlDefault '�W���J�[�\��
�@�@Application.ScreenUpdating = True '��ʕ`����J�n
End Sub

Function ��̍ŏI�s(n As String, Optional k As Long = 1) As Long
   ' n - �J�n����Z���i�͈͂ł��悢�j�ɖ��t�������O�i������j
   ' k - ���͈̔͂̒��̗�ԍ��i�I�v�V�����j
   Dim r1 As Long
   Dim r2 As Long
   Dim mr As Long
   Dim s As Variant
   Dim R_n As Range
   Set R_n = ThisWorkbook.Names(n).RefersToRange
   r1 = R_n.Row
   mr = Rows.count ' �s�̍ő�l�E�E�E�����ŖO�a����B
   Set s = R_n.Columns(k).End(xlDown)
   r2 = s.Row
   If r2 = mr Then
      ��̍ŏI�s = r1
      Exit Function
   End If
   Do While Not (r2 = mr)
      ' Debug.Print s.Value
      r1 = r2
      Set s = s.End(xlDown)
      r2 = s.Row
   Loop
   ��̍ŏI�s = r1
End Function

Function �s�̍ŏI��(n As String, Optional k As Long = 1) As Long
   ' n - �J�n����Z���i�͈͂ł��悢�j�ɖ��t�������O�i������j
   ' k - ���͈̔͂̒��̍s�ԍ��i�I�v�V�����j
   Dim c1 As Long
   Dim c2 As Long
   Dim mc As Long
   Dim s As Variant
   Dim R_n As Range
   Set R_n = ThisWorkbook.Names(n).RefersToRange
   c1 = R_n.Column
   mc = Columns.count ' �s�̍ő�l�E�E�E�����ŖO�a����B
   Set s = R_n.Rows(k).End(xlToRight)
   c2 = s.Column
   If c2 = mc Then
      �s�̍ŏI�� = c1
      Exit Function
   End If
   Do While Not (c2 = mc)
      ' Debug.Print s.Value
      c1 = c2
      Set s = s.End(xlToRight)
      c2 = s.Column
   Loop
   �s�̍ŏI�� = c1
End Function
