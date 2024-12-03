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

  ' コンテキストメニューを取得
  On Error Resume Next ' エラーを無視
  Set ctxMenu = Application.CommandBars("Cell").Controls("Copy As Markdown")
  On Error GoTo 0 ' エラー処理を再開

  ' メニューが既に存在する場合は削除
  If Not ctxMenu Is Nothing Then
    ctxMenu.Delete
  End If

  ' 新しいメニュー項目を追加
  Set ctxMenu = Application.CommandBars("Cell").Controls.Add( _
    Type:=msoControlButton, _
    Before:=1, _
    Temporary:=True)

  ' メニュー項目のプロパティを設定
  With ctxMenu
    .Caption = "Copy As Markdown"
    .OnAction = "SelectionToMarkdown"
  End With

End Sub
