Attribute VB_Name = "x2p_tabjump"
Option Explicit
' -*- coding:shift_jis-dos -*-

'./x2p_tabjump.bas

Function range_TabBottom_range(R_n As Range, _
                               Optional k As Long = 1) As Range
   '
   ' R_n �ŗ^����ꂽ�͈͂� ��P�s�� k ��ځi�f�t�H���g�͂P��ځj
   ' ���牺�����ɒl������Z�������ǂ��Ĉ�ԉ��̃Z���܂Ŋ܂ނ悤��
   ' R_n �̗񐔂͂��̂܂܁A�s�������g�債�ĕԂ��B
   ' �� R_n �ŗ^����ꂽ�͈͂� ��P�s�� k ��ڂ̉������ł� ��Z��
   '    ��������A�g�債�Ȃ��B
   '
   Dim r0 As Long
   Dim r1 As Long
   On Error GoTo RowError
   If R_n.Cell(2,k) = "" Then
      Set range_TabBottom_range = R_n
   Else
      r0 = R_n.Cell(1,k).Row
      r1 = R_n.Cell(1,k).End(xlDown).Row
      Set range_TabBottom_range = R_n.Resize((r1 - r0 + 1))
   End If
RowError:
   Set range_TabBottom_range = R_n
End Function

Function range_TabsBottom_range(R_n As Range, _
                                Optional k As Long = 1, _
                                Optional q As Long = 1) As Range
End Function
   
