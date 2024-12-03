VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
  AddCopyToMarkdownMenu
End Sub

Sub AddCopyToMarkdownMenu()

  Dim ctxMenu As CommandBarControl

  ' �R���e�L�X�g���j���[���擾
  On Error Resume Next ' �G���[�𖳎�
  Set ctxMenu = Application.CommandBars("Cell").Controls("Copy As Markdown")
  On Error GoTo 0 ' �G���[�������ĊJ

  ' ���j���[�����ɑ��݂���ꍇ�͍폜
  If Not ctxMenu Is Nothing Then
    ctxMenu.Delete
  End If

  ' �V�������j���[���ڂ�ǉ�
  Set ctxMenu = Application.CommandBars("Cell").Controls.Add( _
    Type:=msoControlButton, _
    Before:=1, _
    Temporary:=True)

  ' ���j���[���ڂ̃v���p�e�B��ݒ�
  With ctxMenu
    .Caption = "Copy As Markdown"
    .OnAction = "SelectionToMarkdown"
  End With

End Sub