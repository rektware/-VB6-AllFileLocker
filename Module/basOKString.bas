Attribute VB_Name = "OKString"
Public Function MaHoa(Data As String, Optional Depth As Integer) As String
Dim TempChar As String
Dim TempAsc As Long
Dim NewData As String
Dim vChar As Long
For vChar = 1 To Len(Data)
    TempChar = Mid$(Data, vChar, 1)
        TempAsc = Asc(TempChar)
        If Depth = 0 Then Depth = 40
        If Depth > 254 Then Depth = 254

        TempAsc = TempAsc + Depth
        If TempAsc > 255 Then TempAsc = TempAsc - 255
        TempChar = Chr(TempAsc)
        NewData = NewData & TempChar
Next vChar
MaHoa = StrReverse(NewData)

End Function
Public Function GiaiMa(Data As String, Optional Depth As Integer) As String
Dim TempChar As String
Dim TempAsc As Long
Dim NewData As String
Dim vChar As Long

For vChar = 1 To Len(Data)
    TempChar = Mid$(Data, vChar, 1)
        TempAsc = Asc(TempChar)
        If Depth = 0 Then Depth = 40
        If Depth > 254 Then Depth = 254
    TempAsc = TempAsc - Depth
        If TempAsc < 0 Then TempAsc = TempAsc + 255
        TempChar = Chr(TempAsc)
        NewData = NewData & TempChar
Next vChar
GiaiMa = StrReverse(NewData)

End Function
Public Function GetTempFolder()
Dim xTempFolder
Set KhoiTao = CreateObject("Shell.Application")
Set ThuMuc = KhoiTao.Namespace(&H15&)
Set xTempFolder = ThuMuc.Self
GetTempFolder = xTempFolder.Path
End Function

