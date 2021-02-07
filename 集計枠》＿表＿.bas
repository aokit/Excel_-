' -*- coding:shift_jis -*-

'   Dim Lsli As Long
'   Lsli = 1
'   Dim Snmt As String
'   Snmts = "表１"
'   Dim Snmt As String
'   Snmtp = "表１"
'
Sub SetTableTest()
   Dim Pppt As PowerPoint.Presentation
   Set Pppt = SetTableData2pptOpen("集計枠.pptx")
   Call SetTableData2ppt("表１",        3,Pppt)
   Call SetTableData2ppt("別表１",      4,Pppt)
   Call SetTableData2ppt("別表２",      4,Pppt)
   Call SetTableData2ppt("別表３",      4,Pppt)
   Call SetTableData2ppt("特一包括適用",5,Pppt)
   Call SetTableData2ppt("少額特例適用",6,Pppt)
End Sub

Function SetTableData2pptOpen(Spptfile As String) As Variant
   '
   Dim Appt As New PowerPoint.Application
   Dim Pppt As PowerPoint.Presentation
   ' pptファイルのファイル名確認
   If Spptfile Like "*.ppt*" Then
   Else
      Debug.Print ("pptファイルは拡張子も指定してください。")
      MsgBox ("pptファイルは拡張子も指定してください。")
      Exit Function
   End If
   ' 同じフォルダにあるpptファイルを名前指定で読み取り専用にて開く
   Set Pppt = Appt.Presentations.Open(ThisWorkbook.Path & "\" & Spptfile, True)
   Set SetTableData2pptOpen = Pppt
End Function

Sub SetTableData2ppt(Snmts As String, _
                     Lsli As Long, _
                     Optional Pppt As Variant, _
                     Optional Snmtp As String = "*")
   ' pptファイルの中での表の名前などが指定されていなかったらxlsのものに合わせる
   If Snmtp = "*" Then
      Snmtp = Snmts
   End If

   ' エクセルファイルにあるテーブルを名前(=Snmt)指定して値を収集
   Dim Rxls As Range
   On Error GoTo ErrUndefined1
   Set Rxls = ThisWorkbook.Names(Snmts).RefersToRange
   On Error GoTo 0
   Dim Vxls As Variant
   Vxls = Rxls

   Dim Lxrv As Long
   Dim Lxcv As Long
   Lxrv = UBound(Vxls, 1)
   Lxcv = UBound(Vxls, 2)

   Dim i As Long
   Dim j As Long
   Dim Lxrp As Long
   Dim Lxcp As Long
   
   ' 指定(=Lsli)枚めのスライドにあるテーブルを名前(=Snmt)指定して値を設定
   On Error GoTo ErrUndefined2
   With Pppt.Slides(Lsli).Shapes(Snmtp)
      On Error GoTo 0
      Lxrp = .Table.Rows.Count
      Lxcp = .Table.Columns.Count
      For i = 1 To Lxrv
         If i > Lxrp Then ' - Exit For
            '┏pptの表の行が不足していたら足すことにした
            .Table.Rows.Add
         End If
         For j = 1 To Lxcv
            If j > Lxcp Then Exit For
            If Vxls(i, j) = "*" Then
            Else
               .Table.Cell(i, j).Shape.TextFrame.TextRange.Text = Vxls(i, j)
            End If
         Next j
      Next i
   End With
   Exit Sub

ErrUndefined1:
   Debug.Print("表の名前 " & Snmts & " がエクセルで定義されていません")
   Exit Sub

ErrUndefined2:
   Debug.Print("表の名前 " & Snmtp & " がパワーポイントで定義されていません")
   Exit Sub

End Sub

