VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private fp As Integer
Private temp_str As String
Private module_path As String

' ワークブックを開く時のイベント
Private Sub Workbook_Open()
    
    ' txtに書いてある外部ライブラリを読み込み
    load_from_conf ".\libdef.txt", False
       
End Sub

' 設定ファイルに書いてある外部ライブラリを読み込みます。
Sub load_from_conf(conf_path As String, flgMsg As Boolean)
    
    ' 絶対パスに変換
    conf_path = abs_path(conf_path)
    If Dir(conf_path) = "" Then
        If flgMsg = True Then
            MsgBox "外部ライブラリ定義" & conf_path & "が存在しません。"
        End If
        Exit Sub
    End If
    
    ' 全モジュールを削除
    clear_modules
    
    ' 読み取り
    fp = FreeFile
    Open conf_path For Input As #fp
    Do Until EOF(fp)
        ' １行ずつ
        Line Input #fp, temp_str
        If Len(temp_str) > 0 Then
            module_path = abs_path(temp_str)
            If Dir(module_path) = "" Then
                ' エラー
                MsgBox "モジュール" & module_path & "は存在しません。"
                Exit Do
            Else
                ' モジュールとして取り込み
                include module_path
            End If
        End If
    Loop
    Close #fp

    ThisWorkbook.Save
    
End Sub


' あるモジュールを外部から読み込みます。
' パスが.で始まる場合は，相対パスと解釈されます。
Sub include(file_path As String)
    ' 絶対パスに変換
    file_path = abs_path(file_path)
    
    ' 標準モジュールとして登録
    ThisWorkbook.VBProject.VBComponents.Import file_path
End Sub


' 全モジュールを初期化します。
Sub clear_modules()
    Dim component As Object
    For Each component In ThisWorkbook.VBProject.VBComponents
        If component.Type = 1 Then
            ' この標準モジュールを削除
            ThisWorkbook.VBProject.VBComponents.Remove component
        End If
    Next component
End Sub


' ファイルパスを絶対パスに変換します。
Function abs_path(file_path As String)

    ' 絶対パスに変換
    If Left(file_path, 1) = "." Then
        file_path = ThisWorkbook.Path & Mid(file_path, 2, Len(file_path) - 1)
    End If
    
    abs_path = file_path

End Function

'全モジュールをリロードします。
Sub reload_modules()
    
    load_from_conf ".\libdef.txt"

End Sub

