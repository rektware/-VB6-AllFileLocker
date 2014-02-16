Attribute VB_Name = "modMaHoa"
Public Function MaHoa(Message As String) As String

    Randomize
    On Error GoTo errorcheck
    Dim tempmessage As String
    Dim basea As Integer
    Dim tempbasea As String
    Message = Reverse_String(Message)
    tempmessage = CStr(Message)
    basea = Int(Rnd * 75) + 25


    If basea < 0 Then
        tempbasea = CStr(basea)
        tempbasea = Right(tempbasea, Len(tempbasea) - 1)
        basea = CInt(tempbasea)
    End If

    basea = basea / 2
    MaHoa = CStr(basea) + ";"


    For x = 1 To Len(tempmessage)
        MaHoa = MaHoa + CStr(Asc(Left(tempmessage, x)) - basea) + ";"
        basea = basea + 1
        tempmessage = Right(tempmessage, Len(tempmessage) - 1)
    Next x

errorcheck:
End Function



Public Function GiaiMa(code As String) As String

    On Error GoTo errorcheck
    Dim basea As Integer
    Dim tempcode As String


    Do Until Left(code, 1) = ";"
        tempcode = tempcode + Left(code, 1)
        code = Right(code, Len(code) - 1)
    Loop

    basea = CInt(tempcode)
    tempcode = ""
    code = Right(code, Len(code) - 1)


    Do Until code = ""


        Do Until Left(code, 1) = ";"
            tempcode = tempcode + Left(code, 1)
            code = Right(code, Len(code) - 1)
        Loop

        GiaiMa = GiaiMa + Chr(CLng(tempcode) + basea)
        code = Right(code, Len(code) - 1)
        tempcode = ""
        basea = basea + 1
    Loop

    GiaiMa = Reverse_String(GiaiMa)
errorcheck:
End Function



Public Function Reverse_String(Message As String) As String



    For x = 1 To Len(Message)
        Reverse_String = Reverse_String + Left(Right(Message, x), 1)
    Next x

End Function
