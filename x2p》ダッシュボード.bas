' -*- coding:shift_jis -*-

'./x2p�t�_�b�V���{�[�h.bas

' �͈͂ɑ΂��ā@.End(xlDown)�@�ł��炽�Ȕ͈͂𓾂�B�����
' ���e������Z���̘A���̍Ō�̍s�̃Z���ł���B
' �J��Ԃ��K�p�����Ƃ��ɁA�s�� Rows.Count�i�s�̎d�l�ő�l�j
' �ɂȂ����Ƃ��A���̓K�p�̑O�̍s��Ԃ��B

Function �J�����̍ŏI�s(n As String, k) As Long
   ' n - �͈͗^�������O�i������j
   ' k - ���͈̔͂̒��̃J�����ԍ�
   Dim r1 As Long
   Dim r2 As Long
   Dim mr As Long
   Dim s As Variant
   mr = Rows.Count
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

