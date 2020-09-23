Attribute VB_Name = "modOnline"
Option Explicit
Public Const SW_SHOWNORMAL = 1

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub ExecuteLink(ByVal sLinkTo As String)
On Error Resume Next
Dim lRet As Long
Dim lOldCursor As Long
lOldCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass
lRet = ShellExecute(0, "open", sLinkTo, "", vbNull, SW_SHOWNORMAL)
If lRet >= 0 And lRet <= 0 Then
    MsgBox "There was an Error Opening the Web Link to" & vbCrLf & _
    sLinkTo, vbCritical
End If
Screen.MousePointer = vbDefault
    

End Sub

