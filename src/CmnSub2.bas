Attribute VB_Name = "CmnSub2"
Sub prcAddinsAdd(addinName As String, extName As String)
  ' Add Addins
  Addins.Add Filename:="C:\Users\fumio\AppData\Roaming\Microsoft\Addins\" & _
                       addinName & "." & extName
End Sub

Sub prcAddinsInstall(addin As String)
  'install Addins
  Addins(addin).installed = True
  MsgBox "アドイン名＝" & addin & "をインストールしました。"
End Sub

Sub prcAddinsUninstall(addin As String)
  'install Addins
  Addins(addin).installed = False
  MsgBox "アドイン名＝" & addin & "をアンインストールしました。"
End Sub

Sub prcAddinsStatus(addin As String)
  Dim ans As Integer
  'check Addins
  If Addins(addin).installed = True Then
    ans = MsgBox("アドイン名＝" & addin & "はインストールされています。" & vbNewLine & _
                 "アンインストールしますか？", vbInformation + vbYesNo, "アンインストール確認")
    If ans = vbYes Then
      prcAddinsUninstall (addin)
    Else
      MsgBox "アドイン名＝" & addin & "はインストールされたままです。"
    End If
  Else
    ans = MsgBox("アドイン名＝" & addin & "はインストールされていません。" & vbNewLine & _
                 "インストールしますか？", vbInformation + vbYesNo, "インストール確認")
    If ans = vbYes Then
      prcAddinsInstall (addin)
    Else
      MsgBox "アドイン名＝" & addin & "はインストールされませんでした。"
    End If
  End If
End Sub

Sub Test()
  Call prcAddinsStatus("Vimxls_0.8.0")
End Sub

Sub AddinsStatusShortCutKey(addin As String)
  Dim strMacro  As String
  strMacro = "prcAddinsStatus(" & addin & ")"
  Application.MacroOptions Macro:=strMacro, ShortCutKey:="J"
  ' Press Ctrl + Shift + "J"
End Sub

Sub AddMacroShortCutKey()
  'Dim strMacro  As String
  'strMacro = "Test()"
  'Application.MacroOptions Macro:=strMacro, ShortCutKey:="J"
  Application.MacroOptions Macro:="Test", ShortCutKey:="J"
  ' Press Ctrl + Shift + "J"
End Sub
Sub RemoveMacroShortCutKey()
  Application.MacroOptions Macro:="Test()", ShortCutKey:=""
End Sub
