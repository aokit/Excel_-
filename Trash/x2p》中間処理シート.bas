' -*- coding:shift_jis -*-

'./x2p�t���ԏ����V�[�g.bas

Sub �}�N�����s���擾()
  Application.Calculation = xlCalculationManual
  Application.ScreenUpdating = False
  
  On Error GoTo Setting_Error
  D_start = Me.Range("���ԁQ�J�n").Value
  D_end = Me.Range("���ԁQ�I��").Value
  i1_col = Me.Range("BU���o").Value
  i2_col = Me.Range("�敪���o").Value
  i3_col = Me.Range("�d���n���o").Value
  i4_col = Me.Range("����o").Value
  i5_col = Me.Range("���ᒊ�o").Value
  i6_col = Me.Range("�񋖉��ᒊ�o").Value
  tmp = Me.Range("���Ԃ̐R����").Value
  ' ��L�̗̈悪��`����Ă��Ȃ��ƁA�G���[��������
  On Error GoTo 0
  
  ActiveWorkbook.Sheets("yushutsu_kobetsu").Activate
  n1_col = 15 ' BU
  n2_col = 8  ' ����敪
  n3_col = 11 ' �d���n
  
  A_yk = ActiveSheet.UsedRange.Value
  Dim A_B1() As String ' BU
  Dim A_B2() As String ' ����敪
  Dim A_B3() As String ' �d���n
  Dim A_B4() As String ' �
  Dim A_B5() As String ' ����
  Dim A_B6() As String ' �Y������������K�p�Ȃ�
  ReDim A_B1(UBound(A_yk, 1))
  ReDim A_B2(UBound(A_yk, 1))
  ReDim A_B3(UBound(A_yk, 1))
  ReDim A_B4(UBound(A_yk, 1))
  ReDim A_B5(UBound(A_yk, 1))
  ReDim A_B6(UBound(A_yk, 1))
  
  Debug.Print A_yk(2, n1_col)
  Debug.Print D_start
  Debug.Print D_end
  
  j = 0
  k = 0
  m = 0
  n = 0
  For i = 2 To UBound(A_yk, 1)
    ' �����Ⴊ���Ă��Ă���� True �ɂ���B
    q = False
    If (D_start <= A_yk(i, 2)) And (A_yk(i, 2) <= D_end) Then
      A_B1(j) = A_yk(i, n1_col)
      A_B2(j) = A_yk(i, n2_col)
      A_B3(j) = A_yk(i, n3_col)
      j = j + 1
      ' ReDim A_B(j)
      ' ��̎�舵���𒊏o
      If (A_yk(i, 18) = "����K�p") Or (A_yk(i, 20) = "����K�p") Then
        A_B4(k) = A_yk(i, 1)
        Debug.Print "����K�p" & A_yk(i, 1)
        k = k + 1
        q = True
      End If
      ' ����̎�舵���𒊏o
      p_rr2 = (Right(A_yk(i, 18), 2) = "����")
      p_tr2 = (Right(A_yk(i, 20), 2) = "����")
      If p_rr2 Or p_tr2 Then
        A_B5(m) = A_yk(i, 1)
        Debug.Print "�E�E����" & A_yk(i, 1)
        m = m + 1
        q = True
      End If
      ' �Y������������K�p�Ȃ��𒊏o
      p_ql2 = (Left(A_yk(i, 17), 2) = "�Y��")
      p_sl2 = (Left(A_yk(i, 19), 2) = "�Y��")
      If (p_ql2 Or p_sl2) And (Not q) Then
        A_B6(n) = A_yk(i, 1)
        Debug.Print "�E�E�E�E����" & A_yk(i, 1)
        n = n + 1
      End If
    End If
  Next i
  ReDim Preserve A_B1(j - 1)
  ReDim Preserve A_B2(j - 1)
  ReDim Preserve A_B3(j - 1)
  ReDim Preserve A_B4(k - 1)
  ReDim Preserve A_B5(m - 1)
  ReDim Preserve A_B6(n - 1)
  
  Debug.Print (j - 1)
  Debug.Print A_B1(1)
  Debug.Print A_B3(j - 1)
  
  Me.Activate
  i1_col = Me.Range("BU���o").Value
  i2_col = Me.Range("�敪���o").Value
  i3_col = Me.Range("�d���n���o").Value
  i4_col = Me.Range("����o").Value
  i5_col = Me.Range("���ᒊ�o").Value
  i6_col = Me.Range("�񋖉��ᒊ�o").Value
  Me.Range("���Ԃ̐R����").Value = j - 1 + 1 ' 0 (j-1)�Ȃ̂�
  
  i_row = Me.UsedRange.Columns(i1_col).Rows.count
  Me.Range(Cells(2, i1_col), Cells(2 + i_row - 1, i1_col)).Clear
  Me.Range(Cells(2, i1_col), Cells(2 + UBound(A_B1, 1), i1_col)) = WorksheetFunction.Transpose(A_B1)
  i_row = Me.UsedRange.Columns(i2_col).Rows.count
  Me.Range(Cells(2, i2_col), Cells(2 + i_row - 1, i2_col)).Clear
  Me.Range(Cells(2, i2_col), Cells(2 + UBound(A_B2, 1), i2_col)) = WorksheetFunction.Transpose(A_B2)
  i_row = Me.UsedRange.Columns(i3_col).Rows.count
  Me.Range(Cells(2, i3_col), Cells(2 + i_row - 1, i3_col)).Clear
  Me.Range(Cells(2, i3_col), Cells(2 + UBound(A_B3, 1), i3_col)) = WorksheetFunction.Transpose(A_B3)
  i_row = Me.UsedRange.Columns(i4_col).Rows.count
  Me.Range(Cells(2, i4_col), Cells(2 + i_row - 1, i4_col)).Clear
  Me.Range(Cells(2, i4_col), Cells(2 + UBound(A_B4, 1), i4_col)) = WorksheetFunction.Transpose(A_B4)
  i_row = Me.UsedRange.Columns(i5_col).Rows.count
  Me.Range(Cells(2, i5_col), Cells(2 + i_row - 1, i5_col)).Clear
  Me.Range(Cells(2, i5_col), Cells(2 + UBound(A_B5, 1), i5_col)) = WorksheetFunction.Transpose(A_B5)
  i_row = Me.UsedRange.Columns(i6_col).Rows.count
  Me.Range(Cells(2, i6_col), Cells(2 + i_row - 1, i6_col)).Clear
  Me.Range(Cells(2, i6_col), Cells(2 + UBound(A_B6, 1), i6_col)) = WorksheetFunction.Transpose(A_B6)

  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic

  Exit Sub

Setting_Error:
  Debug.Print "�V�[�g�̖��O��`�ɕs��������܂��B"
  MsgBox "�V�[�g�̖��O��`�ɕs��������܂��B"
  
End Sub


' Private
Sub belong2whom(�g�D��`�̗̈� As Range, �Q�Ɨ̈� As Range, �����̈� As Range�j
   ' �g�D���ϊ��̂��߂̃T�u���[�`������邱�Ƃɂ����B
   ' ����������������
   ' belong2whom(�g�D��`�̗̈�,�Q�Ɨ̈�,�����̈�j
   ' ��������
   ' �Q�Ɨ̈�ɂ��镶������A�g�D��`�̗̈�Ō������A�A���̈�Ɍ��ʂ��������ށB
   ' 
   ' �s�g�D��`�̗̈�t
   ' ���ʑg�D��B���A�������ʑg�D��A�i����ɂ́A��ʑg�D��A���ۗL���鉺�ʑg�D��B�j
   ' ���ȉ��̐��K�\���ŗ�ɔz�u�����̈�B��ʑg�D��A�Ɠ����g�D�������ʑg�D��B�Ƃ���
   ' �L����ꍇ�����邪���L���Ȃ��B�܂��A���ʑg�D��B�̖�����ʑg�D��A�����肤��B
   ' (A(B*))+
   ' ��ʑg�DA�Ɖ��ʑg�DB�̓Z���̔w�i�F�ɂ���Ď��ʂ����B
   ' �̈�̑�P�s�̔w�i�F�ɂ���āA��ʑg�D����`�����B
   ' 
   ' �s�Q�Ɨ̈�t
   ' �A�������ʑg�D��A�𓾂������ʑg�D��B���ɔz�u��������
   ' 
   ' �s�A���̈�t
   ' �A�����邱�Ƃ��킩������ʑg�D��A���������ޗ̈�
   ' 
   ' �Ȃ��A�Q�Ɨ̈�Ə����̈�͓����s���̂P��̈�ł��邱�ƁB
   ' ����������������
   ' 
End sub