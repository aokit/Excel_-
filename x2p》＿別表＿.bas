Option Explicit
' -*- coding:shift_jis -*-

'./x2p�t�Q�ʕ\�Q.bas

'
' ���������ʕ\�P�ɂ��Ă͊��҂ǂ���œ����B�������A���Ŏg���Ă���
' �A����ő�s�@�̊֐����������Ȃ��B�@�P�@���w�肵�Ă��Ō�܂Ŕ���
' ���܂��B�܂�AE�񂪐����Ȃ̂ɁAM����w���Ă��܂��B
' range_�A����ő�s_range �𒼐ڕύX����ƁA����p�̉����ɂ��o�O
' �̔����ɂȂ��Ă��܂��̂ŁArange_2_�A����ő�s_range �Ƃ����֐���
' �ʂɗp�ӂ��āA������m�F�������Ƃ���B
'

Private Function range_2_�A����ő�s_range(R_n As Range) As Range
   '
   ' ���̃��W���[���̒������A�܂��́ACall range_�A����ő�s_range ��
   ' ����ɒu��������B���̂Ƃ���͐����u�������Ċm�F�B�ŏI�I�ɂ��Ƃ�
   ' ���g��Ȃ��Ȃ����Ƃ���ŁA���Ƃ̂�_���I�ɂ͏����B
   '
End Function

' ������
' �����X
'
Sub �ʕ\�P�Q�R�ւ̓]�L()
   '
   ' �_�b�V���{�[�h��͈̔͂��w�肵�ĕʕ\�P����ʕ\�R�܂ł̂��ꂼ��͈̔͂�
   ' �]�L����B
   '
   Dim VT As Variant
   Dim NT As String
   NT = "�g�D�W�v"
   Call NamedRange2Ary(NT, VT,, 3)
   Stop
   ' ���̂��ƁAVT�������p�̔z��ɓ]�L���āA����p�̔z���
   ' PrintArrayOnRange�ň������B
   Dim L1 As Long
   Dim U1 As Long
   L1 = LBound(VT, 1)
   U1 = UBound(VT, 1)
   Dim i As Long
   Dim VP As Variant
   ReDim VP(L1 To U1, 1 To 2)
   For i = L1 To U1
      VP(i, 1) = VT(i, 1)
      VP(i, 2) = VT(i, 3)
   Next i
   Stop
   Dim R_1 As Range
   Set R_1 = ThisWorkbook.Names("�ʕ\�P").RefersToRange.Cells(1,1).Offset(1,0)
   Call PrintArrayOnRange(VP, R_1, 0, 2)
   Stop
   '
End Sub

Private Sub PrintArrayOnRange(ByRef Ary As Variant, _
                              R_n As Range, _
                              Optional nr As Long = 0, _
                              Optional nc As Long = 0, _
                              Optional BC As String ="*")
   '
   ' �z�� Ary ��͈� R_n �Ɉ������B
   ' nr, nc ������ɂ��Ă��A�P�ȏ�̒l���w�肳��Ă����
   ' �͈͂̍s�E�񂪖����I�Ɏw�肳��Ă�����̂Ƃ���B
   ' nr = 0 �̏ꍇ�ɂ͗񐔂��Anc = 0 �̏ꍇ�ɂ͍s����
   ' �A����ő�s�ɂ���Č��߂�B
   ' nr = -1 �̏ꍇ�ɂ͗񐔂��Anc = -1 �̏ꍇ�ɂ͍s����
   ' �z��̗񐔁A�s���ɂ���Č��߂�B
   '
   Set R_n = range_�A����ő�s_range(R_n, 1)
   If nr > 0 Then Set R_n = R_n.Resize(nr)
   If nc > 0 Then Set R_n = R_n.Resize(,nc)
   Dim L1 As Long
   Dim U1 As Long
   Dim L2 As Long
   Dim U2 As Long
   L1 = LBound(Ary, 1)
   U1 = UBound(Ary, 1)
   L2 = LBound(Ary, 2)
   U2 = UBound(Ary, 2)
   On Error GoTo OUTOFRANGE
   If nr = -1 Then Set R_n = R_n.Resize(U1 - L1 + 1)
   If nc = -1 Then Set R_n = R_n.Resize(,U2 - L2 + 1)
   On Error GoTo 0
   Dim nrR As Long
   nrR = R_n.Rows.Count
   Dim ncR As Long
   ncR = R_n.Columns.Count
   Dim PAry As Variant
   Dim BAry As Variant
   ReDim PAry(1 To nrR, 1 To ncR)
   ReDim Bary(1 To nrR, 1 To ncR)
   Dim r As Long
   Dim c As Long
   For r = 1 To nrR
      For c = 1 To ncR
         PAry(r, c) = ""
         On Error Resume Next
         PAry(r, c) = Ary(L1 + r - 1, L2 + c - 1)
         On Error GoTo 0
         On Error Resume Next
         If BAry(r, c) = BC Then PAry(r, c) = BC
         On Error GoTo 0
      Next c
   Next r
   R_n = PAry
   Exit Sub
OUTOFRANGE:
   Debug.Print("PrintArrayOnRange - �����Ƃ��ė^����ꂽ�z��̑傫�����s���ł�")
End Sub

Sub NamedRange2Ary(ByVal strName As String, _
                   ByRef Ary As Variant, _
                   Optional nr As Long = 0, _
                   Optional nc As Long = 0)
   '
   ' strName�Ŏw�肵���͈́i�̍���̃Z���j������Ƃ��� nr �s nc ���
   ' �͈͂������̔z��Ɋi�[����B
   '
   Dim R_n As Range
   Set R_n = ThisWorkbook.Names(strName).RefersToRange
   Call Range2Ary(R_n, Ary, nr, nc)
   '
End Sub

Sub Range2Ary(R_n As Range, _
              ByRef Ary As Variant, _
              Optional nr As Long = 0, _
              Optional nc As Long = 0)
   '
   ' R_n �͈̔́i�̍���̃Z���j������Ƃ��� nr �s nc ���
   ' �͈͂������̔z��Ɋi�[����B
   '
   ' Ary ����Ȃ��� Ary() �Ə����Ȃ��Ƃ��߁H
   '
   Dim r0 As Long
   Dim c0 As Long
   r0 = R_n.Row
   c0 = R_n.Column
   Set R_n = range_�A����ő�s_range(R_n, 1)
   If nr > 0 Then Set R_n = R_n.Resize(nr)
   If nc > 0 Then Set R_n = R_n.Resize(,nc)
   Ary = R_n
End Sub
