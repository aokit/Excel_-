' -*- coding:shift_jis-dos -*-

'./x2p�t�ʕ\�P�`�R.bas

Sub �ʕ\1�ɒl��ݒ肷��_�e�L�X�g�{�b�N�X��()
    Const my_folder = "Z:\\tmp\"
    arrs = Worksheets("�ʕ\�P �R").Range("B2:D12").Value
    Set my_application = CreateObject("PowerPoint.Application")
    Set my_presentation = my_application.ActivePresentation
    Set lineIDs = CreateObject("System.Collections.ArrayList")
    lineIDs.Add ("�P")
    lineIDs.Add ("�Q")
    lineIDs.Add ("�R")
    lineIDs.Add ("�S")
    lineIDs.Add ("�T")
    lineIDs.Add ("�U")
    lineIDs.Add ("�V")
    lineIDs.Add ("�W")
    lineIDs.Add ("�X")
    lineIDs.Add ("�P�O")
    lineIDs.Add ("�P�P")
    '// lineIDs.Add ("�P")
    '// my_presentation.Slides(7).Shapes("�ʕ\�P�g�D�P").TextFrame.TextRange.Text = arrs(1, 1)
    '// my_presentation.Slides(7).Shapes("�ʕ\�P�ʂP").TextFrame.TextRange.Text = arrs(1, 2)
    '// my_presentation.Slides(7).Shapes("�ʕ\�P�N�ԂP").TextFrame.TextRange.Text = arrs(1, 3)
    For i = 1 To UBound(arrs, 1)
        �s = lineIDs(i - 1)
        my_presentation.Slides(3).Shapes("�ʕ\�P�g�D" & �s).TextFrame.TextRange.Text = arrs(i, 1)
        my_presentation.Slides(3).Shapes("�ʕ\�P��" & �s).TextFrame.TextRange.Text = arrs(i, 2)
        my_presentation.Slides(3).Shapes("�ʕ\�P�N��" & �s).TextFrame.TextRange.Text = arrs(i, 3)
    '//    my_presentation.Slides(1).Shapes("�e�L�X�g �{�b�N�X 3").TextFrame.TextRange.Text = arrs(i, 2)
    '//    my_presentation.Slides(2).Shapes("�e�L�X�g �{�b�N�X 5").TextFrame.TextRange.Text = arrs(i, 3)
    '//    my_presentation.SaveAs my_folder & "file" & arrs(i, 1)
    Next
End Sub


Sub �ʕ\2�ɒl��ݒ肷��()
  Set my_application = CreateObject("PowerPoint.Application")
  Set my_presentation = my_application.ActivePresentation
  arrs = Worksheets("�ʕ\�P �R").Range("F2:G20").Value
  
  '// With my_presentation.Slides(11).Shapes("Table 7")
  With my_presentation.Slides(3).Shapes("�ʕ\�Q")
    '// MsgBox ("aa")
    '// MsgBox (.Item(1).HasTble)
    Dim r As Long
    Dim c As Long
    With .Table
      For r = 1 To .Rows.count
        For c = 1 To .Columns.count
          '// .Cell(r, c).Shape.TextFrame.TextRange.Text = r & "�s" & c & "��"
          .Cell(r, c).Shape.TextFrame.TextRange.Text = arrs(r, c)
        Next c
      Next r
    End With
  End With
End Sub

Sub �ʕ\3�ɒl��ݒ肷��()
  Set my_application = CreateObject("PowerPoint.Application")
  Set my_presentation = my_application.ActivePresentation
  arrs = Worksheets("�ʕ\�P �R").Range("I2:J20").Value
  
  With my_presentation.Slides(3).Shapes("�ʕ\�R")
    '// MsgBox ("aa")
    '// MsgBox (.Item(1).HasTble)
    Dim r As Long
    Dim c As Long
    With .Table
      For r = 1 To .Rows.count
        For c = 1 To .Columns.count
          '// .Cell(r, c).Shape.TextFrame.TextRange.Text = r & "�s" & c & "��"
          .Cell(r, c).Shape.TextFrame.TextRange.Text = arrs(r, c)
        Next c
      Next r
    End With
  End With
End Sub

Sub �����K�p�ɒl��ݒ肷��()
  Set my_application = CreateObject("PowerPoint.Application")
  Set my_presentation = my_application.ActivePresentation
  arrs = Worksheets("��E����").Range("B3:H13").Value
  
  With my_presentation.Slides(4).Shapes("�����K�p")
    Dim r As Long
    Dim c As Long
    With .Table
      For r = 1 To .Rows.count
        For c = 1 To .Columns.count
          .Cell(r, c).Shape.TextFrame.TextRange.Text = arrs(r, c)
        Next c
      Next r
    End With
  End With
End Sub

Sub ���z����K�p�ɒl��ݒ肷��()
  Set my_application = CreateObject("PowerPoint.Application")
  Set my_presentation = my_application.ActivePresentation
  arrs = Worksheets("��E����").Range("B18:H24").Value
  
  With my_presentation.Slides(5).Shapes("���z����K�p")
    Dim r As Long
    Dim c As Long
    With .Table
      For r = 1 To .Rows.count
        For c = 1 To .Columns.count
          .Cell(r, c).Shape.TextFrame.TextRange.Text = arrs(r, c)
        Next c
      Next r
    End With
  End With
End Sub
