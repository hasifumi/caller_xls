Attribute VB_Name = "CmnSub2"
Sub prcAddinsAdd(addinName As String, extName As String)
  ' Add Addins
  Addins.Add Filename:="C:\Users\fumio\AppData\Roaming\Microsoft\Addins\" & _
                       addinName & "." & extName
End Sub

Sub prcAddinsInstall(addin As String)
  'install Addins
  Addins(addin).installed = True
  MsgBox "�A�h�C������" & addin & "���C���X�g�[�����܂����B"
End Sub

Sub prcAddinsUninstall(addin As String)
  'install Addins
  Addins(addin).installed = False
  MsgBox "�A�h�C������" & addin & "���A���C���X�g�[�����܂����B"
End Sub

Sub prcAddinsStatus(addin As String)
  Dim ans As Integer
  'check Addins
  If Addins(addin).installed = True Then
    ans = MsgBox("�A�h�C������" & addin & "�̓C���X�g�[������Ă��܂��B" & vbNewLine & _
                 "�A���C���X�g�[�����܂����H", vbInformation + vbYesNo, "�A���C���X�g�[���m�F")
    If ans = vbYes Then
      prcAddinsUninstall (addin)
    Else
      MsgBox "�A�h�C������" & addin & "�̓C���X�g�[�����ꂽ�܂܂ł��B"
    End If
  Else
    ans = MsgBox("�A�h�C������" & addin & "�̓C���X�g�[������Ă��܂���B" & vbNewLine & _
                 "�C���X�g�[�����܂����H", vbInformation + vbYesNo, "�C���X�g�[���m�F")
    If ans = vbYes Then
      prcAddinsInstall (addin)
    Else
      MsgBox "�A�h�C������" & addin & "�̓C���X�g�[������܂���ł����B"
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
