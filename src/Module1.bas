Attribute VB_Name = "Module1"
Public Sub SelectionToMarkdown()

  Dim row As Range
  Dim cell As Range
  Dim markdown As String
  Dim header As Boolean
  Dim copyToClipboard As Boolean

  header = True
  copyToClipboard = True

  If TypeName(Selection) = "Range" Then

    ' �s���ƂɃ��[�v
    For Each row In Selection.Rows

      ' �s���̃Z�������[�v
      For Each cell In row.Cells
        markdown = markdown & Replace(Replace(cell.Value, "|", "\|"), vbLf, "<br>") & "|"
      Next cell

      ' �s���� "|" ���폜���A���s��ǉ�
      If Len(markdown) > 0 Then
        markdown = Left(markdown, Len(markdown) - 1)
      End If
    
      markdown = markdown & vbCrLf
      
      If header Then
        markdown = markdown & Repeat("-|", row.Cells.Count)
        markdown = Left(markdown, Len(markdown) - 1)
        markdown = markdown & vbCrLf
        header = False
      End If
      
    Next row

    ' Markdown���N���b�v�{�[�h�ɃR�s�[
    If copyToClipboard Then
      Dim dataObj As Object
      Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
      dataObj.SetText markdown
      dataObj.PutInClipboard
    Else
      MsgBox markdown
    End If
  End If
End Sub

Function Repeat(s, c)
    Dim i As Integer
    
    While i < c
        Repeat = Repeat & s
        i = i + 1
    Wend
End Function
