Attribute VB_Name = "Hyperlink"
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
 (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
 ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) _
 As Long

Public Sub hypFile(frmName As Form, strFile As String)
  '
  Call ShellExecute(frmName.hwnd, "Open", strFile, "", "", 1)
  '
End Sub

Public Sub hypURL(frmName As Form, strWebAddress As String)
  '
  Call ShellExecute(frmName.hwnd, "Open", strWebAddress, "", "", 1)
  '
End Sub

Public Sub hypEmail(frmName As Form, strEmail As String)
  '
  Call ShellExecute(frmName.hwnd, "Open", ("mailto:" & strEmail), "", "", 1)
  '
End Sub

