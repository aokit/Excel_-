' -*- coding:shift_jis -*-

'   Dim Lsli As Long
'   Lsli = 1
'   Dim Snmt As String
'   Snmts = "�\�P"
'   Dim Snmt As String
'   Snmtp = "�\�P"
'
Sub SetTableTest()
   Dim Pppt As PowerPoint.Presentation
   Set Pppt = SetTableData2pptOpen("�W�v�g.pptx")
   Call SetTableData2ppt("�\�P",        3,Pppt)
   Call SetTableData2ppt("�ʕ\�P",      4,Pppt)
   Call SetTableData2ppt("�ʕ\�Q",      4,Pppt)
   Call SetTableData2ppt("�ʕ\�R",      4,Pppt)
   Call SetTableData2ppt("�����K�p",5,Pppt)
   Call SetTableData2ppt("���z����K�p",6,Pppt)
End Sub

Function SetTableData2pptOpen(Spptfile As String) As Variant
   '
   Dim Appt As New PowerPoint.Application
   Dim Pppt As PowerPoint.Presentation
   ' ppt�t�@�C���̃t�@�C�����m�F
   If Spptfile Like "*.ppt*" Then
   Else
      Debug.Print ("ppt�t�@�C���͊g���q���w�肵�Ă��������B")
      MsgBox ("ppt�t�@�C���͊g���q���w�肵�Ă��������B")
      Exit Function
   End If
   ' �����t�H���_�ɂ���ppt�t�@�C���𖼑O�w��œǂݎ���p�ɂĊJ��
   Set Pppt = Appt.Presentations.Open(ThisWorkbook.Path & "\" & Spptfile, True)
   Set SetTableData2pptOpen = Pppt
End Function

Sub SetTableData2ppt(Snmts As String, _
                     Lsli As Long, _
                     Optional Pppt As Variant, _
                     Optional Snmtp As String = "*")
   ' ppt�t�@�C���̒��ł̕\�̖��O�Ȃǂ��w�肳��Ă��Ȃ�������xls�̂��̂ɍ��킹��
   If Snmtp = "*" Then
      Snmtp = Snmts
   End If

   ' �G�N�Z���t�@�C���ɂ���e�[�u���𖼑O(=Snmt)�w�肵�Ēl�����W
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
   
   ' �w��(=Lsli)���߂̃X���C�h�ɂ���e�[�u���𖼑O(=Snmt)�w�肵�Ēl��ݒ�
   On Error GoTo ErrUndefined2
   With Pppt.Slides(Lsli).Shapes(Snmtp)
      On Error GoTo 0
      Lxrp = .Table.Rows.Count
      Lxcp = .Table.Columns.Count
      For i = 1 To Lxrv
         If i > Lxrp Then ' - Exit For
            '��ppt�̕\�̍s���s�����Ă����瑫�����Ƃɂ���
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
   Debug.Print("�\�̖��O " & Snmts & " ���G�N�Z���Œ�`����Ă��܂���")
   Exit Sub

ErrUndefined2:
   Debug.Print("�\�̖��O " & Snmtp & " ���p���[�|�C���g�Œ�`����Ă��܂���")
   Exit Sub

End Sub

