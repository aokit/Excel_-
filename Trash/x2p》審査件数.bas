' -*- coding:shift_jis-dos -*-

'./x2p�t�R������.bas

Sub �A�N�e�B�u�ȃX���C�h�ɕ\��}������()
  '// With ActiveWindow.Selection.SlideRange
  Set my_application = CreateObject("PowerPoint.Application")
  Set my_presentation = my_application.ActivePresentation
  With my_presentation.Slides(11)
    .Shapes.AddTable _
      NumRows:=5, _
      NumColumns:=3
  End With
  With my_presentation.Slides(11).Shapes(1).Table
  End With
End Sub

Sub �\��ǉ�����()
  ActivePresentation.Slides(11).Shapes.AddTable(3, 3)
End Sub

Sub �ʕ\2�ɒl��ݒ肷��()
  Set my_application = CreateObject("PowerPoint.Application")
  Set my_presentation = my_application.ActivePresentation
  
  '// With my_presentation.Slides(11).Shapes("Table 7")
  With my_presentation.Slides(3).Shapes("�ʕ\�Q")
    '// MsgBox ("aa")
    '// MsgBox (.Item(1).HasTble)
    Dim r As Long
    Dim c As Long
    With .Table
      For r = 1 To .Rows.count
        For c = 1 To .Columns.count
          .Cell(r, c).Shape.TextFrame.TextRange.Text = r & "�s" & c & "��"
        Next c
      Next r
    End With
  End With
End Sub

Sub ���Ԗڂ�Shape��Table������̂����ׂ�()
  Set my_application = CreateObject("PowerPoint.Application")
  Set my_presentation = my_application.ActivePresentation
  With my_presentation.Slides(7)
  For Each IShape In .Shapes
    If IShape.HasTable Then
      MsgBox ("Found " & IShape.Name)
    End If
  Next
  End With
  
End Sub
