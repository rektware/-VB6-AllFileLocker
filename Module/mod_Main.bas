Attribute VB_Name = "Module1"

Sub Main()

Dim ocxDir$
ocxDir = Environ("WinDir") & "\System32\FVUnicodeControl.ocx"
If (FileExists(ocxDir) = False) Then
Dim bytResourceData() As Byte
bytResourceData = LoadResData(101, "FVUnicodeControl.ocx")
Open ocxDir For Binary Shared As #1
Put #1, 1, bytResourceData
Close #1
Shell "regsvr32 /s " & ocxDir, vbHide
End If
frmMain.Show
End Sub

Public Function FileExists(sFile As String) As Boolean
On Error Resume Next
FileExists = ((GetAttr(sFile) And vbDirectory) = 0)
End Function
Public Function AppPath()
AppPath = App.Path
If Right(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
End Function






