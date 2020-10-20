' -*- coding:shift_jis-dos -*-

'./x2p》別表１～３.bas

Sub 別表1に値を設定する_テキストボックス版()
    Const my_folder = "Z:\\tmp\"
    arrs = Worksheets("別表１ ３").Range("B2:D12").Value
    Set my_application = CreateObject("PowerPoint.Application")
    Set my_presentation = my_application.ActivePresentation
    Set lineIDs = CreateObject("System.Collections.ArrayList")
    lineIDs.Add ("１")
    lineIDs.Add ("２")
    lineIDs.Add ("３")
    lineIDs.Add ("４")
    lineIDs.Add ("５")
    lineIDs.Add ("６")
    lineIDs.Add ("７")
    lineIDs.Add ("８")
    lineIDs.Add ("９")
    lineIDs.Add ("１０")
    lineIDs.Add ("１１")
    '// lineIDs.Add ("１")
    '// my_presentation.Slides(7).Shapes("別表１組織１").TextFrame.TextRange.Text = arrs(1, 1)
    '// my_presentation.Slides(7).Shapes("別表１個別１").TextFrame.TextRange.Text = arrs(1, 2)
    '// my_presentation.Slides(7).Shapes("別表１年間１").TextFrame.TextRange.Text = arrs(1, 3)
    For i = 1 To UBound(arrs, 1)
        行 = lineIDs(i - 1)
        my_presentation.Slides(3).Shapes("別表１組織" & 行).TextFrame.TextRange.Text = arrs(i, 1)
        my_presentation.Slides(3).Shapes("別表１個別" & 行).TextFrame.TextRange.Text = arrs(i, 2)
        my_presentation.Slides(3).Shapes("別表１年間" & 行).TextFrame.TextRange.Text = arrs(i, 3)
    '//    my_presentation.Slides(1).Shapes("テキスト ボックス 3").TextFrame.TextRange.Text = arrs(i, 2)
    '//    my_presentation.Slides(2).Shapes("テキスト ボックス 5").TextFrame.TextRange.Text = arrs(i, 3)
    '//    my_presentation.SaveAs my_folder & "file" & arrs(i, 1)
    Next
End Sub


Sub 別表2に値を設定する()
  Set my_application = CreateObject("PowerPoint.Application")
  Set my_presentation = my_application.ActivePresentation
  arrs = Worksheets("別表１ ３").Range("F2:G20").Value
  
  '// With my_presentation.Slides(11).Shapes("Table 7")
  With my_presentation.Slides(3).Shapes("別表２")
    '// MsgBox ("aa")
    '// MsgBox (.Item(1).HasTble)
    Dim r As Long
    Dim c As Long
    With .Table
      For r = 1 To .Rows.count
        For c = 1 To .Columns.count
          '// .Cell(r, c).Shape.TextFrame.TextRange.Text = r & "行" & c & "列"
          .Cell(r, c).Shape.TextFrame.TextRange.Text = arrs(r, c)
        Next c
      Next r
    End With
  End With
End Sub

Sub 別表3に値を設定する()
  Set my_application = CreateObject("PowerPoint.Application")
  Set my_presentation = my_application.ActivePresentation
  arrs = Worksheets("別表１ ３").Range("I2:J20").Value
  
  With my_presentation.Slides(3).Shapes("別表３")
    '// MsgBox ("aa")
    '// MsgBox (.Item(1).HasTble)
    Dim r As Long
    Dim c As Long
    With .Table
      For r = 1 To .Rows.count
        For c = 1 To .Columns.count
          '// .Cell(r, c).Shape.TextFrame.TextRange.Text = r & "行" & c & "列"
          .Cell(r, c).Shape.TextFrame.TextRange.Text = arrs(r, c)
        Next c
      Next r
    End With
  End With
End Sub

Sub 特一包括適用に値を設定する()
  Set my_application = CreateObject("PowerPoint.Application")
  Set my_presentation = my_application.ActivePresentation
  arrs = Worksheets("包括・特例").Range("B3:H13").Value
  
  With my_presentation.Slides(4).Shapes("特一包括適用")
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

Sub 少額特例適用に値を設定する()
  Set my_application = CreateObject("PowerPoint.Application")
  Set my_presentation = my_application.ActivePresentation
  arrs = Worksheets("包括・特例").Range("B18:H24").Value
  
  With my_presentation.Slides(5).Shapes("少額特例適用")
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
