Attribute VB_Name = "x2p_tabjump"
Option Explicit
' -*- coding:shift_jis-dos -*-

'./x2p_tabjump.bas

Function range_TabWiden_range(R_n As Range, _
                              Optional k As Long = 1) As Range
   '
   ' range_TabBottom_range �������B�w�肵���͈� R_n �� �� k ��
   ' ��Z���łȂ��s�@�܂Ł@�͈� R_n ���g�����ĕԂ��B
   ' �֐����Ԃ����͈͂ɍēx�֐���K�p���Ă��͈͕͂s�ςƂȂ邱��
   ' �͒�`���疾�炩�Ȃ̂ŁA�������̎��͈̔͂𓾂邽�߂ɂ́A
   ' �܂��A�Ԃ��ꂽ�͈͂̂ЂƂ��̍s�́@�� k �Ŏ���
   ' ��Z���łȂ��Z�� ��T�����Ƃ��K�v�ƂȂ�B�����
   ' range_n_TabWiden_range �irange_n_TabBottom_range �̕ʖ��j
   ' �ŋL�q���Ă���B
   ' ���̂��߁A range_TabBottom_range �ŎQ�Ƃ������ q �͕s�v
   ' �ƂȂ�B
   ' �� k �̒l�� R_n �̗񐔁iR_n.Rows.Count�j�𒴂��Ă��Ă��A
   ' �L���ł���B���̗�ɑΏۂƂ��ď������āA��p���āA�s�𓾂�
   ' ��ɂ��Ă� R_n �͈̔͂�Ԃ��B
   ' ���� R_n �� �� k �̉������ɋ�Z�������Ȃ��ꍇ�́A�͈͂�
   ' ���݂��Ȃ����̂Ƃ��āA�����Ŏw�肵���͈͂����̂܂ܕԂ��B
   '
   Dim r0 As Long
   Dim r1 As Long
   On Error GoTo RowError
   If (R_n.Cells(1,k).Value <> "") And _
      (R_n.Cells(2,k).Value = "") Then
      Set range_TabWiden_range = R_n
   Else
      r0 = R_n.Cells(1,k).Row
      ' r1 = R_n.Cells(1,k).End(xlDown).Row
      ' ���P�s�����ł͂Ȃ���
      ' �������s�͈̔͂��w�肳�ꂽ�Ƃ��́A�w�肳�ꂽ�͈͂�
      ' ���Ō�̍s���� .End(xlDown) ������B����ŁA
      ' ���w�肳�ꂽ�͈͓��� �� �������Ă��w�肳�ꂽ�͈�
      ' �������Ƃɔ͈͂�Ԃ��悤�ɂȂ�B
      r1 = R_n.Cells(R_n.Rows.Count,k).End(xlDown).Row
      If r1 = Rows.Count Then
         ' ���E�E�E��Z�������Ȃ�����
         Set range_TabWiden_range = R_n
      Else
         Set range_TabWiden_range = R_n.Resize((r1 - r0 + 1))
      End If
   End If
   Exit Function
RowError:
   Debug.Print("range_TabWiden_range�ŃG���[�F�w�肵���͈͂��ŉ��s�ɒB���Ă���Ȃ�")
   Set range_TabWiden_range = R_n
End Function

Function range_n_TabWiden_range(R_n As Range, _
                                Optional k As Long = 1, _
                                Optional n As Long = 1) As Range
   '
   ' �w�肵���͈� R_n �� k ��ڂ��牺�̕����ɒl�̂���Z����
   ' �A�����Ă���͈͂� R_n ���琔���āAn �߂͈̔͂�Ԃ��B
   ' k �� n ���L�ڂ��Ȃ��ꍇ�͂������ 1 �Ƃ���B
   ' �w��̍ő�s�x�ɑ�������B
   ' range_TabWiden_range �� n��w�J��Ԃ��āx�Ă�
   ' �E�E�E�E�E�E�E�E�E�E�J��Ԃ����߂ɁE�E�E�E�E�E�E�E�E�E
   ' �P��߂́A range_TabWiden_range �����s���Ĕ͈͂�Ԃ��B
   ' �@�͈͂̎��i���j�̃Z���i����Z���j�͈̔͂� Rt �łQ���
   ' �@�ֈ����p���B
   ' �Q��߈ȍ~�́ARt ���󂫃Z���̐擪�ֈړ����A
   ' �@���̌�A range_TabWiden_range �����s���Ĕ͈͂�Ԃ��B
   ' �@�͈͂̎��i���j�̃Z���i����Z���j�͈̔͂� Rt �Ŏ����
   ' �@�����p���B
   '
   Dim q As Long
   Dim Rt As Range
   Dim Rs As Range
   Set Rt = R_n
   Set range_n_TabWiden_range = R_n
   On Error GoTo EndOfRange
   For q = 1 To n
      If q > 1 Then
         Set Rt = Rt.Cells(1,k).End(xlDown).Offset(0,-(k-1)).Resize(1,Rt.Columns.Count)
         ' Set Rt = Rt.Cells(1,k).End(xlDown).Offset(0,-(k-1)).Resize(1,R_n.Columns.Count)
      End If
      Set Rs = range_n_TabWiden_range
      Set Rt = range_TabWiden_range(Rt, k)
      If (Rt.Row + Rt.Rows.Count - 1) = Rows.Count Then
         Set range_n_TabWiden_range = Rs
         Debug.Print("�͈͂��w�肳�ꂽ���ɑ���܂���ł����B")
         Exit Function
      Else
         Set range_n_TabWiden_range = Rt
         Set Rt = Rt.Offset(1,0).Cells(Rt.Rows.Count,1).Resize(1,Rt.Columns.Count)
         ' Set Rt = Rt.Offset(1,0).Cells(Rt.Rows.Count,1).Resize(1,R_n.Columns.Count)
      End If
   Next q
   On Error GoTo 0
   Exit Function
   '
EndOFRange:
   range_n_TabWiden_range = R_n
   Debug.Print("�͈͂��w�肳�ꂽ���ɑ���܂���ł����B")
End Function

' ========================================================================================
' ========================================================================================
' ========================================================================================

Function range_TabBottom_range(R_n As Range, _
                               Optional k As Long = 1, _
                               Optional q As Long = 1) As Range
   '
   ' R_n �ŗ^����ꂽ�͈͂� ��P�s�� k ��ځi�f�t�H���g�͂P��ځj
   ' ���牺�����ɒl������Z�������ǂ��Ĉ�ԉ��̃Z���܂Ŋ܂ނ悤��
   ' R_n �̗񐔂͂��̂܂܁A�s�������g�債�ĕԂ��B
   ' �� R_n �ŗ^����ꂽ�͈͂� ��P�s�� k ��ڂ̉������ł� ��Z��
   '    ��������A�g�債�Ȃ��B
   '
   ' R_n ���P�s�݂̂ŁAk ��̏�̃Z�������̃Z������Z���̏ꍇ�A
   ' �i�Q�s�ȏ゠��΁A�Q��̎��s�Ŏ��͈̔͂Ɉړ�����̂����j
   ' �P��̎��s�Ŏ��͈̔͂Ɉړ����Ă��܂��B�����}�����邽�߂�
   ' ������̌Ăяo���́A�񐔂��w�肷��ʂ̊֐��ɂ���ċL�q���āA
   ' ���̊֐����̂ɂ́A����ڂ̌Ăяo���ł��邩��`����悤�ɂ���B
   ' ���̂��߂� ���� q ���g���B�f�t�H���g�l�� 1 �ł���B
   ' �����Ď��s����Ƃ��ɂ́A�A������ڂł��邩�� q �ŗ^����B
   ' �i�A���R����s����Ƃ��́Aq �� 3,2,1 �ƕω�����j
   '
   Dim r0 As Long
   Dim r1 As Long
   On Error GoTo RowError
   If (R_n.Cells(1,k).Value <> "") And _
      (R_n.Cells(2,k).Value = "") And _
      ((q mod 2) = 1) Then
      Set range_TabBottom_range = R_n
   Else
      r0 = R_n.Cells(1,k).Row
      r1 = R_n.Cells(1,k).End(xlDown).Row
      If r1 = Rows.Count Then
         Set range_TabBottom_range = R_n
         ' ? range_TabBottom_range(Range("A4:B4"),2).Address
         ' $A$4:$B$1048576
         ' �� $A4:$B4
         ' �͈͂����݂��Ȃ��Ƃ��́A�����Ŏw�肵���͈͂�Ԃ��B
      Else
         Set range_TabBottom_range = R_n.Resize((r1 - r0 + 1))
      End If
   End If
   Exit Function
RowError:
   Set range_TabBottom_range = R_n
End Function

' ? range_n_TabBottom_range(Range("D4")).Address
' ? range_n_TabBottom_range(Range("D4"),,5).Address
' ? range_n_TabBottom_range(Range("C4"),2,5).Address
' ? range_n_TabBottom_range(Range("B4"),3,3).Address
' ? range_n_TabBottom_range(Range("B4:D4"),3,3).Address
' $B$30
' �� $B$30:$D$30
'
' ? range_TabBottom_range(Range("C4"), 2, 1).Address
' ? range_TabBottom_range(Range("C4"), 2, 1).Offset(1,1).Address
' ? range_TabBottom_range(Range("C4"), 2, 1).Offset(1,0+1).Cells(21,1).Address
' 0+1 = k
' ���̒l���g�������̂� Offset ���[�Ă�B
' 21 = range_TabBottom_range(Range("C4"), 2, 1).Rows.Count
' ? range_TabBottom_range(Range("C4"), 2, 1).Rows.Count
'
Function range_n_TabBottom_range(R_n As Range, _
                                 Optional k As Long = 1, _
                                 Optional n As Long = 1) As Range
   '
   ' �w�肵���͈� R_n �� k ��ڂ��牺�̕����ɒl�̂���Z����
   ' �A�����Ă���͈͂� R_n ���琔���āAn �߂͈̔͂�Ԃ��B
   ' k �� n ���L�ڂ��Ȃ��ꍇ�͂������ 1 �Ƃ���B
   ' �w��̍ő�s�x�ɑ�������B
   ' range_TabBottom_range �� n��w�J��Ԃ��āx�Ă�
   ' �E�E�E�E�E�E�E�E�E�E�J��Ԃ����߂ɁE�E�E�E�E�E�E�E�E�E
   ' �P��߂́A range_TabBottom_range �����s���Ĕ͈͂�Ԃ��B
   ' �@�͈͂̎��i���j�̃Z���i����Z���j�͈̔͂� Rt �łQ���
   ' �@�ֈ����p���B
   ' �Q��߈ȍ~�́ARt ���󂫃Z���̐擪�ֈړ����A
   ' �@���̌�A range_TabBottom_range �����s���Ĕ͈͂�Ԃ��B
   ' �@�͈͂̎��i���j�̃Z���i����Z���j�͈̔͂� Rt �Ŏ����
   ' �@�����p���B
   '
   Dim q As Long
   Dim Rt As Range
   Dim Rs As Range
   Set Rt = R_n
   Set range_n_TabBottom_range = R_n
   On Error GoTo EndOfRange
   For q = 1 To n
      ' If q > 1 Then Set Rt = Rt.Cells(1,k).End(xlDown)
      If q > 1 Then
         ' Set Rt = Rt.Cells(1,k).End(xlDown).Offset(0,-(k-1))
         Set Rt = Rt.Cells(1,k).End(xlDown).Offset(0,-(k-1)).Resize(1,Rt.Columns.Count)
         ' Set Rt = Rt.Cells(1,k).End(xlDown).Offset(0,-(k-1)).Resize(1,R_n.Columns.Count)
         ' Stop
      End If
      Set Rs = range_n_TabBottom_range
      ' Set Rt = range_TabBottom_range(Rt, k, q)
      Set Rt = range_TabBottom_range(Rt, k)
      ' Stop
      If (Rt.Row + Rt.Rows.Count - 1) = Rows.Count Then
         Set range_n_TabBottom_range = Rs
         Debug.Print("�͈͂��w�肳�ꂽ���ɑ���܂���ł����B")
         Exit Function
      Else
         Set range_n_TabBottom_range = Rt
         ' Set Rt = Rt.Cells(Rt.Rows.Count + 1, k)
         ' Set Rt = Rt.Cells(Rt.Rows.Count + 1, 1).Resize(1, Rt.Columns.Count)
         ' Set Rt = Rt.Cells(Rt.Rows.Count + 1, k).Offset(0, -(k-1)).Resize(1, Rt.Columns.Count)
         ' Set Rt = Rt.Offset(1,k).Cells(Rt.Rows.Count,1)
         Set Rt = Rt.Offset(1,0).Cells(Rt.Rows.Count,1).Resize(1,Rt.Columns.Count)
         ' Set Rt = Rt.Offset(1,0).Cells(Rt.Rows.Count,1).Resize(1,R_n.Columns.Count)
         ' Stop
         '
' ? range_TabBottom_range(Range("C4"), 2, 1).Offset(1,0+1).Cells(21,1).Address
' range_TabBottom_range(Range("C4"), 2, 1) = Rt
' 0+1 = k
' ���̒l���g�������̂� Offset ���[�Ă�B
' 21 = range_TabBottom_range(Range("C4"), 2, 1).Rows.Count
' ? range_TabBottom_range(Range("C4"), 2, 1).Rows.Count

         '
      End If
   Next q
   On Error GoTo 0
   Exit Function
   '
EndOFRange:
   ' range_n_TabBottom_range = Rs
   range_n_TabBottom_range = R_n
   Debug.Print("�͈͂��w�肳�ꂽ���ɑ���܂���ł����B")
End Function
   
' k ���@�\���Ă��Ȃ��悤�ȋC������B�͈͂𒴂��� k ���@�\�����邱�Ƃ͂ł��邩�H
' q �Ŋ���𕪂���K�v���Ȃ���������Ȃ����ǂ����B
' ? Range("C27").Resize(1,2).Cells(1,2).End(xlDown).Offset(0,-1).Resize(1,2).Address
' $C$28:$D$28

' ? Range("C27").Cells(1,2).End(xlDown).Offset(0,-1).Resize(1,2).Address