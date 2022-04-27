Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub 自動印刷()

 rc = MsgBox("実行しますか？", vbYesNo + vbQuestion, "確認")
 If rc = vbYes Then
  Dim startData As Integer
  Dim endData As Integer
  startData = Cells(2, 13).Value
  endData = Cells(2, 15).Value
  Dim HPurl As String
  For i = startData To endData
    HPurl = Cells(i, 1)
    CreateObject("WScript.Shell").Run ("chrome.exe -url " & HPurl)
    Call Sleep(2000)
    SendKeys "^P"
    Call Sleep(2000)
    SendKeys "{ENTER}"
    Call Sleep(2000)
    SendKeys "^W"
    Call Sleep(2000)
   Next
 Else
    MsgBox "中止しました"
 End If
 
End Sub