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
   Dim cas(1 To 2) As Long
   cas(1) = 1
   cas(2) = 3
   Call Sheet2colTableCopy("�g�D�W�v", cas(), "�ʕ\�P", 1)
   cas(1) = 1
   cas(2) = 2
   Call Sheet2colTableCopy("����W�v", cas(), "�ʕ\�Q", 1)
   cas(1) = 5
   cas(2) = 6
   Call Sheet2colTableCopy("�d���n�W�v", cas(), "�ʕ\�R", 1)
   '
   ' Range("�\�P").Cells(7,4) = Range("���Ԃ̐R������")
   ' �����t�����͈͂����̃X�N���v�g�Ɠ����V�[�g�ɂȂ���΂��̋L�@��
   ' �@�g���Ȃ��B�Ȃ̂ŁA�ȉ��A�v���V�W���ɁB
   Call SheetNamedCellCopy("����K�p����", "�\�P", 3, 4)
   Call SheetNamedCellCopy("���z����K�p����", "�\�P", 4, 4)
   Call SheetNamedCellCopy("�Y�������������", "�\�P", 5, 4)
   Call SheetNamedCellCopy("���X�g�K����Y������", "�\�P", 6, 4)
   Call SheetNamedCellCopy("���Ԃ̐R������", "�\�P", 7, 4)
   '
End Sub

Sub SheetNamedCellCopy(strName1 As String, _
                       strName2 As String, _
                       r0 As Long, _
                       c0 As Long)
   ' strName1 �Ɩ��t�����Z���̓��e�� strName2 �Ɩ��t�����͈͂̃Z��
   ' (r0,c0)�ɏ������ށB
   Dim r1 As Range
   Dim r2 As Range
   Set r1 = ThisWorkbook.Names(strName1).RefersToRange
   Set r2 = ThisWorkbook.Names(strName2).RefersToRange.Cells(r0, c0)
   ' r2.Value = CStr(r1.Value)
   r2.Value = r1.Value
End Sub

Sub Sheet2colTableCopy(strName1 As String, cas() As Long, _
                       strName2 As String, rOff As Long)
   '
   ' �V�[�g��ł̂Q��̕\�̃R�s�[�F
   ' �@strName1 �Ŏw�肳�ꂽ�͈͂ɂQ��ŋL�ڂ���Ă���\�� strName2 ��
   ' �@�w�肳�ꂽ�͈͂ɃR�s�[����B�z�� cas(1 To 2)�ŁAcas(i)=j �ɂ��
   ' �@�w�肷�邱�ƂŃR�s�[��� i ��ڂɃR�s�[���� j ����R�s�[����B
   ' �E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E
   ' �Fcas - column assign
   ' Dim cas(1 To "c1") As Long
   ' �ʕ\�P�̏ꍇ�Fcas(1)=1, cas(2)=3
   ' �ʕ\�Q�̏ꍇ�Fcas(1)=1, cas(2)=2
   ' �ʕ\�R�̏ꍇ�Fcas(1)=5, cas(2)=6
   '
   Dim R_1 As Range
   ' Set R_1 = ThisWorkbook.Names(strName2).RefersToRange.Offset(1,0).Resize(1)
   ' ���͈͂̍ŏ��̍s��n���Ă�OK
   Set R_1 = ThisWorkbook.Names(strName2).RefersToRange.Offset(1,0)
   ' �Ђ���Ƃ��āA�ȉ��ł����v�H
   ' Set R_1 = Range(strName2).Offset(rOff,0)
   ' ���̃V�[�g�ŌĂԂƂ��͑��v���낤���Ǒ��̃V�[�g�ŌĂԂƌ����Ȃ��͂��B
   Debug.Print(strName2)
   Debug.Print(R_1.Address)
   Dim r1 As Long
   Dim c1 As Long
   r1 = R_1.Rows.Count
   ' c1 = R_1.Columns.Count
   c1 = 2
   Dim VT As Variant
   Dim NT As String
   NT = strName1
   Call NamedRange2Ary(NT, VT, r1, cas(c1))
   Dim L1 As Long
   Dim U1 As Long
   L1 = LBound(VT, 1)
   U1 = UBound(VT, 1)
   Dim i As Long
   Dim VP As Variant
   ReDim VP(L1 To U1, 1 To 2)
   For i = L1 To U1
      VP(i, 1) = VT(i, cas(1))
      VP(i, 2) = VT(i, cas(2))
   Next i
   ' Stop
   Call PrintArrayOnRange(VP, R_1, 0, 2)
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
   ' �A����ő�s�ɂ���Ċg������B
   ' nr = -1 �̏ꍇ�ɂ͗񐔂��Anc = -1 �̏ꍇ�ɂ͍s����
   ' �z��̗񐔁A�s���ɂ���Č��߂�B
   '
   ' Set R_n = range_�A����ő�s_range(R_n, 1)
   ' Set R_n = range_TabBottom_range(R_n, 1)
   ' Set R_n = range_n_TabBottom_range(R_n)
   Set R_n = range_n_TabWiden_range(R_n)
   ' ����͂܂��A�s�̊g���̂݁B��̊g���͑Ή����Ă��Ȃ��B
   ' Stop
   ' �����ŏ�L�Ŏ~�߂Ă݂āA? R_n.address ����� M ��܂�
   ' �����Ă��܂��Ă��邱�Ƃ��킩��B
   Debug.Print(R_n.address)
   ' Debug.Print("��L�� M ��܂Ŋ܂�ł���悤�Ȃ炱���ł͉e�����Ȃ����A�o�O�B")
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
   On Error Resume Next
   ' �����́A�ȉ����ƃG���[�B�Ȃ�ł��ȁB
   Set R_n = Range(strName)
   ' �����F strName �����̃V�[�g�̖��O�ł͂Ȃ�����B
   On Error GoTo 0
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
