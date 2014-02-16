Attribute VB_Name = "EDFILE"

Option Explicit

'--------------------- users public values -------------------------

Public UltraReturnValue     As Integer
Public UltraReturnString    As String
Public AbortUltraRun        As Boolean

'-----------------------------------------------------------------

Private Const PROGRESS_CALCFREQ = 3
Private Const PROGRESS_CALCCRC = 3
Private Const PROGRESS_ENCHUFF = 44
Private Const PROGRESS_DECHUFF = 45
Private Const PROGRESS_CHECKCRC = 5
Private Const PROGRESS_ENCRYPT = 50
Private Const PROGRESS_DECRYPT = 50

Private CurrProgresValue As Integer


Public Const strPCC = "AllFileLockerPCC"


Private Const FILE_VERSION = "ALL FILE LOCKER FINAL"
Private Const TEXT_BEGIN = "--- BEGIN AFL MESSAGE ---"
Private Const TEXT_VERSION = ""
Private Const TEXT_END = "END"
Private Const TEXT_MAXPERLINE = 60

Private K1(0 To 462)  As Integer
Private S1            As Integer
Private P1            As Integer

Private K2(0 To 250)  As Integer
Private P2            As Integer
Private S2            As Integer

Private K3(0 To 180)  As Integer
Private S3            As Integer
Private P3            As Integer

Private FEEDBACK      As Byte
Private SeedString As String


Private Const PR1 = 463
Private Const PR2 = 251
Private Const PR3 = 181

Private aDecTab(255)        As Integer
Private aEncTab(63)         As Byte
Private FileErrDescription  As String

Private Type HUFFMANTREE
  ParentNode As Integer
  RightNode As Integer
  LeftNode As Integer
  Value As Integer
  Weight As Long
End Type

Private Type byteArray
  Count As Byte
  Data() As Byte
End Type

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Function UltraText(ByVal aText As String, Key As String, PCC As String) As String
' Encode/decode text
Dim i As Integer
Dim Text As String
UltraReturnValue = 0
UltraReturnString = ""
FileErrDescription = ""
Text = TrimText(aText)
If Text = "" Then UltraReturnValue = 1: GoTo skip
If IsValidKey(Key) = False Then UltraReturnValue = 3: GoTo skip
i = CheckUltraText(Text)
Select Case i
Case 0
    UltraText = EncodeString(Text, Key, PCC)
Case 1
    UltraText = DecodeString(Text, Key, PCC)
Case 2
    UltraReturnValue = 10 'error unkwown version
Case 3
    UltraReturnValue = 30 'error crypto header
End Select
skip:
Call SetReturnString
If UltraReturnValue <> 0 Then UltraText = aText
End Function

' ------------------------------------------------------------
'    Progress Bar Picture sub (please adjust to your program code)
' ------------------------------------------------------------

Private Sub UpdateStatus(ByVal sngPercent As Single)
' IMPORTANT to use the progressbar:
' The following lines draw a progressbar in a picturebox
' called picProgress on a form called Form1
' change the names of form and picturebox to your own needs
' Set the picturebox Autoredraw property on TRUE !!!
' Set the picturebox Scalewidth property on 100 after sizing pic !!!
' Set the picturebox Forecolor at dark blue, the Backcolor at gray
' When the progressbar is not used you can speedup encryption process
' by deleting the marked code lines in the routines EncodeByteArray,
' DecodeByteArray, EncodeFile and DecodeFile, and by deleting all
' Updatestatus(x) lines.
frmMain.Bar1.Value = sngPercent
End Sub

' ------------------------------------------------------------
'                   Encryption algorithm functions
' ------------------------------------------------------------

Public Sub SetKey(ByVal aKey As String, ByVal aPCC As String)
Dim i           As Long
Dim j           As Long
Dim KEYLen      As Long
Dim KEY1()      As Byte
Dim KEY2(16)    As Byte
Dim KEY3(22)    As Byte
Dim KEYPCC()    As Byte
Dim tmp         As Integer
Dim PCCLen      As Integer
' setup key1 - variable
KEYLen = Len(aKey)
KEY1() = StrConv(aKey, vbFromUnicode)
For i = 0 To PR1 - 1
    K1(i) = i
Next
P1 = 0
S1 = 0
For i = 0 To PR1 - 1
    j = (j + K1(i) + KEY1(i Mod KEYLen)) Mod PR1
    tmp = K1(i)
    K1(i) = K1(j)
    K1(j) = tmp
Next
' setup key2 - 136 bits
For i = 0 To PR1 - 1
    KEY2(i Mod 17) = KEY2(i Mod 17) Xor (K1(i) And 255)
Next
For i = 0 To PR2 - 1
    K2(i) = i
Next
P2 = 0
S2 = 0
For i = 0 To PR2 - 1
    j = (j + K2(i) + KEY2(i Mod 17)) Mod PR2
    tmp = K2(i)
    K2(i) = K2(j)
    K2(j) = tmp
Next
' setup key3 - 184 bits
For i = 0 To PR2 - 1
    KEY3(i Mod 23) = KEY3(i Mod 23) Xor (K2(i) And 255)
Next
PCCLen = Len(aPCC)
KEYPCC() = StrConv(aPCC, vbFromUnicode)
If PCCLen > 0 Then
    For i = 0 To 22
        KEY3(i) = KEY3(i) Xor KEYPCC(i Mod PCCLen)
    Next
    End If
For i = 0 To PR3 - 1
    K3(i) = i
Next i
S2 = 0
P2 = 0
For i = 0 To PR3 - 1
    j = (j + K3(i) + KEY3(i Mod 23)) Mod PR3
    tmp = K3(i)
    K3(i) = K3(j)
    K3(j) = tmp
Next
S3 = 0
P3 = 0
FEEDBACK = 0
aKey = ""
aPCC = ""
End Sub

Private Function EncodeByte(aByte As Byte) As Byte
EncodeByte = aByte Xor FnULTRA(FEEDBACK)
FEEDBACK = EncodeByte
End Function

Private Function DecodeByte(aByte As Byte) As Byte
Dim tmpbyte As Byte
tmpbyte = aByte
DecodeByte = aByte Xor FnULTRA(FEEDBACK)
FEEDBACK = tmpbyte
End Function

Public Sub EncodeByteArray(byteArray() As Byte)
Dim ModVal As Integer
Dim i As Long
Dim ByteLen As Long
Dim NewProgress As Integer
ModVal = 5000
'use larger ModVal value to speedup when processing large amount of data
ByteLen = UBound(byteArray)
For i = 0 To ByteLen
    byteArray(i) = EncodeByte(byteArray(i))
    If i Mod ModVal = 0 Then
        DoEvents
        If AbortUltraRun = True Then Exit For
        '------------------------------------------------------
        'remove the following 5 lines if no progressbar is used
        NewProgress = i / ByteLen * PROGRESS_ENCHUFF + PROGRESS_CALCCRC + PROGRESS_CALCFREQ + PROGRESS_ENCRYPT
        If (NewProgress <> CurrProgresValue) Then
                CurrProgresValue = NewProgress
                Call UpdateStatus(CurrProgresValue)
            End If
        '------------------------------------------------------
        End If
Next i
End Sub

Public Sub DecodeByteArray(byteArray() As Byte)
Dim ModVal As Integer
Dim i As Long
Dim ByteLen As Long
Dim NewProgress As Integer
ModVal = 5000
'use larger ModVal value to speedup when processing large amount of data
ByteLen = UBound(byteArray)
For i = 0 To ByteLen
    byteArray(i) = DecodeByte(byteArray(i))
    If i Mod ModVal = 0 Then
        DoEvents
        If AbortUltraRun = True Then Exit For
        '------------------------------------------------------
        'remove the following 5 lines if no progressbar is used
        NewProgress = i / ByteLen * PROGRESS_DECRYPT
        If (NewProgress <> CurrProgresValue) Then
            CurrProgresValue = NewProgress
            Call UpdateStatus(CurrProgresValue)
            End If
        '------------------------------------------------------
        End If
Next i
End Sub

Private Function FnULTRA(FB As Byte) As Byte
Dim TS As Integer
Dim OUT1 As Byte
Dim OUT2 As Integer
Dim OUT3 As Integer
P1 = (P1 + 1) Mod PR1
S1 = (S1 + K1(P1) + FB) Mod PR1
TS = K1(P1)
K1(P1) = K1(S1)
K1(S1) = TS
OUT1 = K1((K1(P1) + K1(S1)) Mod PR1) Mod 256
P2 = (P2 + 1) Mod PR2
S2 = (S2 + K2(P2) + OUT1) Mod PR2
TS = K2(P2)
K2(P2) = K2(S2)
K2(S2) = TS
OUT2 = K2((K2(P2) + K2(S2)) Mod PR2) Mod 256
P3 = (P3 + 1) Mod PR3
S3 = (S3 + K3(P3) + OUT2) Mod PR3
TS = K3(P3)
K3(P3) = K3(S3)
K3(S3) = TS
OUT3 = K3((K3(P3) + K3(S3)) Mod PR3) Mod 256
FnULTRA = (OUT1 + OUT2 + OUT3) Mod 256
End Function

' ------------------------------------------------------------
'                  File encryption functions
' ------------------------------------------------------------

Public Function EncodeFile(ByVal SourceFile As String, ByVal TargetFile As String, ByVal xKeygen As String) As String
', ByVal Key As String, ByVal PCC As String
Dim Key As String
Dim PCC As String
Dim FileO       As Integer
Dim k           As Integer
Dim VersionBuffer() As Byte
Dim DummyBuffer() As Byte
Dim FileBuffer() As Byte
Dim OutBuffer() As Byte
Dim i As Long
Dim DummyString As String
Dim checkByte1 As Byte
Dim checkByte2 As Byte
Dim Extension  As String
Dim ModVal As Integer
Dim NewProgress As Integer
Dim ByteLen As Long
Dim tmpFile As String
Key = xKeygen
PCC = strPCC
On Error GoTo errHandler
ModVal = 5000
'use larger ModVal value to speedup when processing large amount of data
AbortUltraRun = False
'open file and read bytes into buffer array
FileO = FreeFile
Screen.MousePointer = 11
Open SourceFile For Binary As #FileO
    ReDim FileBuffer(0 To LOF(FileO) - 1)
    Get #FileO, , FileBuffer()
Close #FileO
Screen.MousePointer = 0
'start progress
CurrProgresValue = 0
'compress file
Call HuffEncodeByte(FileBuffer, UBound(FileBuffer) + 1)
If AbortUltraRun = True Then GoTo skip
'set version buffer
VersionBuffer = StrConv(FILE_VERSION, vbFromUnicode)
'set dummy
DummyString = RandomDummy
checkByte1 = Asc(Mid(DummyString, Len(DummyString) - 1, 1))
checkByte2 = Asc(Mid(DummyString, Len(DummyString), 1))
Extension = GetFileExt(SourceFile)
DummyString = DummyString + Extension + Chr(0)
Call SetKey(Key, PCC)
'encypt dummy+ext
DummyBuffer() = StrConv(DummyString, vbFromUnicode)
For i = 0 To UBound(DummyBuffer)
    DummyBuffer(i) = EncodeByte(DummyBuffer(i))
Next
'encrypt file
ByteLen = UBound(FileBuffer)
For i = 0 To ByteLen
    FileBuffer(i) = EncodeByte(FileBuffer(i))
    If i Mod ModVal = 0 Then
        DoEvents
        If AbortUltraRun = True Then Exit For
        '------------------------------------------------------
        'remove the following 5 lines if no progressbar is used
        NewProgress = i / ByteLen * PROGRESS_ENCHUFF + PROGRESS_CALCCRC + PROGRESS_CALCFREQ + PROGRESS_ENCRYPT
        If (NewProgress <> CurrProgresValue) Then '***
            CurrProgresValue = NewProgress '***
            Call UpdateStatus(CurrProgresValue) '***
            End If
        '------------------------------------------------------
        End If
Next
If AbortUltraRun = True Then GoTo skip
'encrypt sheckbytes
checkByte1 = EncodeByte(checkByte1)
checkByte2 = EncodeByte(checkByte2)
'save file
EncodeFile = TargetFile
If FileExists(EncodeFile) Then Kill EncodeFile
Screen.MousePointer = 11
Open EncodeFile For Binary As #FileO
    Put #FileO, , VersionBuffer()
    Put #FileO, , DummyBuffer()
    Put #FileO, , FileBuffer()
    Put #FileO, , checkByte1
    Put #FileO, , checkByte2
Close #FileO
Screen.MousePointer = 0
Call UpdateStatus(0)
If SourceFile = TargetFile Then
    'Kill SourceFile
    If FileExists(SourceFile) Then Kill SourceFile
    End If
skip:
If AbortUltraRun = True Then
    UltraReturnValue = 11 'encode aborted
    EncodeFile = SourceFile
    End If
Call UpdateStatus(0)
Screen.MousePointer = 0
Exit Function
errHandler:
Close #FileO
UltraReturnValue = 12 ' encode file error
FileErrDescription = Err.Description
EncodeFile = SourceFile
Screen.MousePointer = 0
Call UpdateStatus(0)
End Function

Public Function DecodeFile(ByVal SourceFile As String, ByVal TargetFile As String, ByVal xKeygen As String) As String
Dim Key As String
Dim PCC As String
Dim i As Long
Dim DataStart As Long
Dim DummySize As Integer
Dim DummyStart As Integer
Dim Umax As Long
Dim FileBuffer() As Byte
Dim RndByte As Byte
Dim ExtByte As Byte
Dim ExtCount As Integer
Dim checkByte1 As Byte
Dim checkByte2 As Byte
Dim checkbyteA As Byte
Dim checkbyteB As Byte
Dim tmpASC As Integer
Dim SizeDummy As Byte
Dim FileO As Integer
Dim offSet As Integer
Dim TargetExt As String
Dim ModVal As Integer
Dim NewProgress As Integer
Dim ByteLen As Long
Dim tmpFile As String
Key = xKeygen
PCC = strPCC
On Error GoTo errHandler
ModVal = 5000
'increase ModVal value to speedup when processing large amount of data
AbortUltraRun = False
FileO = FreeFile
Screen.MousePointer = 11
Open SourceFile For Binary As #FileO
    'check if there is data
    ReDim FileBuffer(0 To LOF(FileO) - 1)
    Get #FileO, , FileBuffer()
Close #FileO
Screen.MousePointer = 0
Call SetKey(Key, PCC)
DummyStart = Len(FILE_VERSION)
'decrypt dummy bytes
DummySize = DecodeByte(FileBuffer(DummyStart))
If (DummySize + DummyStart) > UBound(FileBuffer) Then GoTo errHandlerCrypto
'decrypt dummy's
For i = 2 To DummySize
    RndByte = DecodeByte(FileBuffer(DummyStart + i - 1))
    'get checkbytes
    If i = DummySize - 1 Then checkByte1 = RndByte
    If i = DummySize Then checkByte2 = RndByte
Next
offSet = Len(FILE_VERSION) + DummySize
'decrypt ext
TargetExt = ""
Do
    ExtByte = DecodeByte(FileBuffer(offSet + ExtCount))
    If ExtByte <> 0 Then TargetExt = TargetExt & Chr(ExtByte)
    ExtCount = ExtCount + 1
Loop Until ExtByte = 0
If TargetExt <> "" Then TargetExt = "." & TargetExt
offSet = DummyStart + DummySize + ExtCount
CurrProgresValue = 0
ByteLen = UBound(FileBuffer) - offSet - 2
For i = 0 To ByteLen
    FileBuffer(i) = DecodeByte(FileBuffer(i + offSet))
    If i Mod ModVal = 0 Then
        DoEvents
        If AbortUltraRun = True Then Exit For
        '------------------------------------------------------
        'remove the following 5 lines if no progressbar is used
        NewProgress = i / ByteLen * PROGRESS_DECRYPT
        If (NewProgress <> CurrProgresValue) Then
            CurrProgresValue = NewProgress
            Call UpdateStatus(CurrProgresValue)
            End If
        '------------------------------------------------------
        End If
Next
If AbortUltraRun = True Then GoTo skip
checkbyteA = FileBuffer(UBound(FileBuffer) - 1)
checkbyteB = FileBuffer(UBound(FileBuffer))
checkbyteA = DecodeByte(checkbyteA)
checkbyteB = DecodeByte(checkbyteB)
If checkByte1 <> checkbyteA Or checkByte2 <> checkbyteB Then
    GoTo errHandlerCrypto
    End If
ReDim Preserve FileBuffer(UBound(FileBuffer) - offSet - 2)
'decompress file
Call HuffDecodeByte(FileBuffer, UBound(FileBuffer) + 1)
If AbortUltraRun = True Then GoTo skip
If UltraReturnValue <> 0 Then GoTo skip
'save file
DecodeFile = TargetFile
If FileExists(DecodeFile) Then Kill DecodeFile
'save the file
FileO = FreeFile
Screen.MousePointer = 11
Open DecodeFile For Binary As #FileO
    Put #FileO, , FileBuffer()
Close #FileO
Screen.MousePointer = 0
If SourceFile = TargetFile Then
    'overwrit source
    If FileExists(SourceFile) Then Kill SourceFile
    End If
skip:
'decode ok
Call UpdateStatus(0)
If AbortUltraRun = True Then
    UltraReturnValue = 21 'decode aborted
    DecodeFile = SourceFile
    End If
Screen.MousePointer = 0
Exit Function
errHandler:
Call UpdateStatus(0)
UltraReturnValue = 22 ' decode file error
FileErrDescription = Err.Description
Screen.MousePointer = 0
Exit Function
errHandlerCrypto:
Call UpdateStatus(0)
UltraReturnValue = 23 ' decode crypto error
Screen.MousePointer = 0
End Function

Public Function CheckUltraFile(ByVal Source As String) As Integer
' 0 = not encrypted
' 1 = ultra
' 2 = unknown version
Dim VersionBuffer() As Byte
Dim strVersion As String
Dim FileO As Integer
On Error Resume Next
'read crypto info from file
FileO = FreeFile
Open Source For Binary As #FileO
ReDim VersionBuffer(0 To Len(FILE_VERSION) - 1)
Get #FileO, , VersionBuffer()
Close #FileO
'get crypto info
strVersion = StrConv(VersionBuffer(), vbUnicode)
If strVersion = FILE_VERSION Then
        'known crypto version
        CheckUltraFile = 1
        Else
        If UCase(Right(Source, 4)) = ".UCC" Then
            CheckUltraFile = 2 'Unknown version
            Else
            CheckUltraFile = 0 'Unprotected"
            End If
        End If
End Function

' ------------------------------------------------------------
'                  Text encryption functions
' ------------------------------------------------------------

Public Function EncodeString(TextIn As String, KeyString As String, PCMstring As String) As String
Dim TextArray() As Byte
Dim DummyString As String
Dim checkByte1 As Byte
Dim checkByte2 As Byte
Dim i As Integer
Screen.MousePointer = 11
AbortUltraRun = False
EncodeString = TextIn
EncodeString = HuffEncodeString(EncodeString)
'create dummy header
DummyString = RandomDummy
checkByte1 = Asc(Mid(DummyString, Len(DummyString) - 1, 1))
checkByte2 = Asc(Mid(DummyString, Len(DummyString), 1))
'add dummy and check bytes
EncodeString = DummyString & EncodeString & Chr(checkByte1) & Chr(checkByte2)
'encode array
Call SetKey(KeyString, PCMstring)
TextArray() = StrConv(EncodeString, vbFromUnicode)
Call EncodeByteArray(TextArray)
EncodeString = StrConv(TextArray(), vbUnicode)
'conter to radix64
EncodeString = EncodeStr64(EncodeString, TEXT_MAXPERLINE)
'add header and trail
EncodeString = TEXT_BEGIN & vbCrLf & TEXT_VERSION & vbCrLf & EncodeString & vbCrLf & TEXT_END
Screen.MousePointer = 0
Call UpdateStatus(0)
End Function

Public Function DecodeString(TextIn As String, KeyString As String, PCMstring As String) As String
Dim TextArray() As Byte
Dim HL As Integer
Dim TL As Integer
Dim DummyString As String
Dim SizeDummy As Integer
Dim checkByte1 As Byte
Dim checkByte2 As Byte
CurrProgresValue = 0
Screen.MousePointer = 11
AbortUltraRun = False
'strip trail and header
HL = Len(TEXT_BEGIN & vbCrLf & TEXT_VERSION & vbCrLf)
TL = Len(vbCrLf & TEXT_END)
DecodeString = Mid(TextIn, HL + 1, Len(TextIn) - HL - TL)
'decode radix64
DecodeString = DecodeStr64(DecodeString)
'decode array
Call SetKey(KeyString, PCMstring)
TextArray() = StrConv(DecodeString, vbFromUnicode)
Call DecodeByteArray(TextArray)
DecodeString = StrConv(TextArray(), vbUnicode)
Screen.MousePointer = 0
'check checkbytes
SizeDummy = Asc(Left(DecodeString, 1))
If SizeDummy > Len(DecodeString) - 2 Then GoTo errDecode
DummyString = Left(DecodeString, SizeDummy)
checkByte1 = Asc(Mid(DummyString, Len(DummyString) - 1, 1))
checkByte2 = Asc(Mid(DummyString, Len(DummyString), 1))
'check decryption
If Asc(Mid(DecodeString, Len(DecodeString) - 1, 1)) = checkByte1 And _
 Asc(Mid(DecodeString, Len(DecodeString), 1)) = checkByte2 Then
    DecodeString = Mid(DecodeString, SizeDummy + 1, (Len(DecodeString) - 2) - SizeDummy)
    DecodeString = HuffDecodeString(DecodeString)
    Else
    GoTo errDecode
    End If
Call UpdateStatus(0)
Screen.MousePointer = 0
Exit Function
errDecode:
DecodeString = ""
UltraReturnValue = 33
Call UpdateStatus(0)
Screen.MousePointer = 0
End Function

Public Function CheckUltraText(ByVal TextIn As String) As Integer
' 0 = not encrypted
' 1 = ultra 1.0.3
' 2 = unknown version
' 3 = incomplete crypto header
Dim HL As Integer
Dim TL As Integer
Dim VL As Integer
TextIn = TrimText(TextIn)
'trim text and cut crlf's
HL = Len(TEXT_BEGIN & vbCrLf)
VL = Len(TEXT_VERSION & vbCrLf)
TL = Len(vbCrLf & TEXT_END)
If Left(TextIn, HL) = TEXT_BEGIN & vbCrLf And Right(TextIn, TL) <> vbCrLf & TEXT_END Then CheckUltraText = 3: Exit Function
If Left(TextIn, HL) <> TEXT_BEGIN & vbCrLf And Right(TextIn, TL) = vbCrLf & TEXT_END Then CheckUltraText = 3: Exit Function
If Len(TextIn) < HL + TL + VL + 1 Then Exit Function
If Left(TextIn, HL) <> TEXT_BEGIN & vbCrLf Then Exit Function
If Right(TextIn, TL) <> vbCrLf & TEXT_END Then Exit Function
If Mid(TextIn, HL + 1, VL) <> TEXT_VERSION & vbCrLf Then CheckUltraText = 2: Exit Function
CheckUltraText = 1
End Function

' ------------------------------------------------------------
'                     Random Dummy generating
' ------------------------------------------------------------

Private Function RandomDummy() As String
'setup dummy string, between 16 and 255 bytes, first byte contains dummylenght
Dim rndKey As String
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim q As Byte
Dim SizeDummy As Integer
RandomDummy = ""
Randomize
SizeDummy = Int(224 * Rnd) + 32
If Len(SeedString) > 0 Then
    For k = 1 To Len(SeedString)
        SizeDummy = SizeDummy Xor Asc(Mid(SeedString, k, 1))
    Next
    End If
Do While SizeDummy > 255
    SizeDummy = SizeDummy - 224
Loop
If SizeDummy < 32 Then SizeDummy = SizeDummy + 224
For k = 1 To SizeDummy - 1
    RandomDummy = RandomDummy & Chr(Int((256 * Rnd)))
Next
j = 1
For k = 1 To 16
    rndKey = ""
    For i = 1 To 16
        q = Int((256 * Rnd))
        If Len(SeedString) > 0 Then q = q Xor Asc(Mid(SeedString, j, 1))
        j = j + 1: If j > Len(SeedString) Then j = 1
        rndKey = rndKey & Chr(q)
    Next i
    Call SetKey(rndKey, "")
    For i = 1 To Len(RandomDummy)
        q = Asc(Mid(RandomDummy, i, 1))
        If k Mod 3 = 0 Then
            q = DecodeByte(q)
            Else
            q = EncodeByte(q)
            End If
        Mid(RandomDummy, i, 1) = Chr(q)
    Next i
Next k
RandomDummy = Chr(SizeDummy) & RandomDummy
End Function

Public Sub RandomFeed(ByVal x As Single, ByVal y As Single)
'this sub enables the user to feed random data to seedstring
Static XP As Single
Static YP As Single
If x = XP And y = YP Then Exit Sub
XP = x: YP = y
SeedString = SeedString & Chr((x Xor y) And 255)
If Len(SeedString) > 251 Then SeedString = Mid(SeedString, 2)
End Sub

' ------------------------------------------------------------
'                   Compression functions
' ------------------------------------------------------------

Private Function HuffDecodeString(Text As String) As String
Dim byteArray() As Byte
byteArray() = StrConv(Text, vbFromUnicode)
Call HuffDecodeByte(byteArray, Len(Text))
HuffDecodeString = StrConv(byteArray(), vbUnicode)
End Function

Private Function HuffEncodeString(Text As String) As String
Dim byteArray() As Byte
byteArray() = StrConv(Text, vbFromUnicode)
Call HuffEncodeByte(byteArray, Len(Text))
HuffEncodeString = StrConv(byteArray(), vbUnicode)
End Function

Private Sub HuffEncodeByte(byteArray() As Byte, ByteLen As Long)
Dim i As Long, j As Long, Char As Byte, BitPos As Byte, lNode1 As Long
Dim lNode2 As Long, lNodes As Long, lLength As Long, Count As Integer
Dim lWeight1 As Long, lWeight2 As Long, Result() As Byte, ByteValue As Byte
Dim ResultLen As Long, bytes As byteArray, NodesCount As Integer, NewProgress As Integer
Dim BitValue(0 To 7) As Byte, CharCount(0 To 255) As Long
Dim Nodes(0 To 511) As HUFFMANTREE, CharValue(0 To 255) As byteArray
'set identification
If (ByteLen = 0) Then
    ReDim Preserve byteArray(0 To ByteLen + 3)
    If (ByteLen > 0) Then Call CopyMem(byteArray(4), byteArray(0), ByteLen)
    byteArray(0) = 72
    byteArray(1) = 69
    byteArray(2) = 48
    byteArray(3) = 13
    Exit Sub
End If
ReDim Result(0 To 522)
Result(0) = 72
Result(1) = 69
Result(2) = 51
Result(3) = 13
ResultLen = 4
'get frequency off all bytes
For i = 0 To (ByteLen - 1)
    CharCount(byteArray(i)) = CharCount(byteArray(i)) + 1
    If (i Mod 1000 = 0) Then
        DoEvents
        If AbortUltraRun = True Then Exit Sub
        '------------------------------------------------------
        'remove the following 5 lines if no progressbar is used
        NewProgress = i / ByteLen * PROGRESS_CALCFREQ
        If (NewProgress <> CurrProgresValue) Then
            CurrProgresValue = NewProgress
            Call UpdateStatus(CurrProgresValue)
        End If
        '------------------------------------------------------
    End If
Next
'put freq in nodes
For i = 0 To 255
    If (CharCount(i) > 0) Then
        With Nodes(NodesCount)
            .Weight = CharCount(i)
            .Value = i
            .LeftNode = -1
            .RightNode = -1
            .ParentNode = -1
        End With
        NodesCount = NodesCount + 1
    End If
Next

For lNodes = NodesCount To 2 Step -1
    lNode1 = -1: lNode2 = -1
    For i = 0 To (NodesCount - 1)
        If (Nodes(i).ParentNode = -1) Then
            If (lNode1 = -1) Then
                lWeight1 = Nodes(i).Weight
                lNode1 = i
            ElseIf (lNode2 = -1) Then
                lWeight2 = Nodes(i).Weight
                lNode2 = i
            ElseIf (Nodes(i).Weight < lWeight1) Then
                If (Nodes(i).Weight < lWeight2) Then
                    If (lWeight1 < lWeight2) Then
                        lWeight2 = Nodes(i).Weight
                        lNode2 = i
                    Else
                        lWeight1 = Nodes(i).Weight
                        lNode1 = i
                    End If
                Else
                    lWeight1 = Nodes(i).Weight
                    lNode1 = i
                End If
            ElseIf (Nodes(i).Weight < lWeight2) Then
                lWeight2 = Nodes(i).Weight
                lNode2 = i
            End If
        End If
    Next
    
    With Nodes(NodesCount)
        .Weight = lWeight1 + lWeight2
        .LeftNode = lNode1
        .RightNode = lNode2
        .ParentNode = -1
        .Value = -1
    End With
    
    Nodes(lNode1).ParentNode = NodesCount
    Nodes(lNode2).ParentNode = NodesCount
    NodesCount = NodesCount + 1
Next
ReDim bytes.Data(0 To 255)
Call CreateBitSequences(Nodes(), NodesCount - 1, bytes, CharValue)
For i = 0 To 255
    If (CharCount(i) > 0) Then lLength = lLength + CharValue(i).Count * CharCount(i)
Next
lLength = IIf(lLength Mod 8 = 0, lLength \ 8, lLength \ 8 + 1)
If ((lLength = 0) Or (lLength > ByteLen)) Then
    ReDim Preserve byteArray(0 To ByteLen + 3)
    Call CopyMem(byteArray(4), byteArray(0), ByteLen)
    byteArray(0) = 72
    byteArray(1) = 69
    byteArray(2) = 48
    byteArray(3) = 13
    Exit Sub
End If
'calculate CRC
Char = 0
For i = 0 To (ByteLen - 1)
    Char = Char Xor byteArray(i)
    If (i Mod 10000 = 0) Then
        DoEvents
        If AbortUltraRun = True Then Exit Sub
        '------------------------------------------------------
        'remove the following 5 lines if no progressbar is used
        NewProgress = i / ByteLen * PROGRESS_CALCCRC + PROGRESS_CALCFREQ
        If (NewProgress <> CurrProgresValue) Then
            CurrProgresValue = NewProgress
            Call UpdateStatus(CurrProgresValue)
        End If
        '------------------------------------------------------
    End If
Next
Result(ResultLen) = Char
ResultLen = ResultLen + 1
Call CopyMem(Result(ResultLen), ByteLen, 4)
ResultLen = ResultLen + 4
BitValue(0) = 2 ^ 0
BitValue(1) = 2 ^ 1
BitValue(2) = 2 ^ 2
BitValue(3) = 2 ^ 3
BitValue(4) = 2 ^ 4
BitValue(5) = 2 ^ 5
BitValue(6) = 2 ^ 6
BitValue(7) = 2 ^ 7
Count = 0
For i = 0 To 255
    If (CharValue(i).Count > 0) Then Count = Count + 1
Next
Call CopyMem(Result(ResultLen), Count, 2)
ResultLen = ResultLen + 2
Count = 0
For i = 0 To 255
    If (CharValue(i).Count > 0) Then
        Result(ResultLen) = i
        ResultLen = ResultLen + 1
        Result(ResultLen) = CharValue(i).Count
        ResultLen = ResultLen + 1
        Count = Count + 16 + CharValue(i).Count
    End If
Next
ReDim Preserve Result(0 To ResultLen + Count \ 8)
BitPos = 0
ByteValue = 0
For i = 0 To 255
    With CharValue(i)
        If (.Count > 0) Then
            For j = 0 To (.Count - 1)
                If (.Data(j)) Then ByteValue = ByteValue + BitValue(BitPos)
                BitPos = BitPos + 1
                If (BitPos = 8) Then
                    Result(ResultLen) = ByteValue
                    ResultLen = ResultLen + 1
                    ByteValue = 0
                    BitPos = 0
                End If
            Next
        End If
    End With
Next
If (BitPos > 0) Then
    Result(ResultLen) = ByteValue
    ResultLen = ResultLen + 1
End If
ReDim Preserve Result(0 To ResultLen - 1 + lLength)
Char = 0
BitPos = 0
For i = 0 To (ByteLen - 1)
    With CharValue(byteArray(i))
        For j = 0 To (.Count - 1)
            If (.Data(j) = 1) Then Char = Char + BitValue(BitPos)
            BitPos = BitPos + 1
            If (BitPos = 8) Then
                Result(ResultLen) = Char
                ResultLen = ResultLen + 1
                BitPos = 0
                Char = 0
            End If
        Next
    End With
    If (i Mod 10000 = 0) Then
        DoEvents
        '------------------------------------------------------
        'remove the following 5 lines if no progressbar is used
        If AbortUltraRun = True Then Exit Sub
        NewProgress = i / ByteLen * PROGRESS_ENCHUFF + PROGRESS_CALCCRC + PROGRESS_CALCFREQ
        If (NewProgress <> CurrProgresValue) Then
            CurrProgresValue = NewProgress
            Call UpdateStatus(CurrProgresValue)
        End If
        '------------------------------------------------------
    End If
Next
If (BitPos > 0) Then
    Result(ResultLen) = Char
    ResultLen = ResultLen + 1
End If
ReDim byteArray(0 To ResultLen - 1)
Call CopyMem(byteArray(0), Result(0), ResultLen)
End Sub

Private Sub HuffDecodeByte(byteArray() As Byte, ByteLen As Long)
Dim i As Long, j As Long, pos As Long, Char As Byte, CurrPos As Long
Dim Count As Integer, CheckSum As Byte, Result() As Byte, BitPos As Integer
Dim NodeIndex As Long, ByteValue As Byte, ResultLen As Long, NodesCount As Long
Dim lResultLen As Long, NewProgress As Integer, BitValue(0 To 7) As Byte
Dim Nodes(0 To 511) As HUFFMANTREE, CharValue(0 To 255) As byteArray
If (byteArray(0) <> 72) Or (byteArray(1) <> 69) Or (byteArray(3) <> 13) Then
ElseIf (byteArray(2) = 48) Then
    Call CopyMem(byteArray(0), byteArray(4), ByteLen - 4)
    ReDim Preserve byteArray(0 To ByteLen - 5)
    Exit Sub
ElseIf (byteArray(2) <> 51) Then
    Err.Raise vbObjectError, "HuffmanDecode()", "The data either was not compressed with HE3 or is corrupt (identification string not found)"
    Exit Sub
End If
CurrPos = 5
CheckSum = byteArray(CurrPos - 1)
CurrPos = CurrPos + 1
Call CopyMem(ResultLen, byteArray(CurrPos - 1), 4)
CurrPos = CurrPos + 4
lResultLen = ResultLen
If (ResultLen = 0) Then Exit Sub
ReDim Result(0 To ResultLen - 1)
Call CopyMem(Count, byteArray(CurrPos - 1), 2)
CurrPos = CurrPos + 2
For i = 1 To Count
    With CharValue(byteArray(CurrPos - 1))
        CurrPos = CurrPos + 1
        .Count = byteArray(CurrPos - 1)
        CurrPos = CurrPos + 1
        ReDim .Data(0 To .Count - 1)
    End With
Next
BitValue(0) = 2 ^ 0
BitValue(1) = 2 ^ 1
BitValue(2) = 2 ^ 2
BitValue(3) = 2 ^ 3
BitValue(4) = 2 ^ 4
BitValue(5) = 2 ^ 5
BitValue(6) = 2 ^ 6
BitValue(7) = 2 ^ 7
ByteValue = byteArray(CurrPos - 1)
CurrPos = CurrPos + 1
BitPos = 0
For i = 0 To 255
    With CharValue(i)
        If (.Count > 0) Then
            For j = 0 To (.Count - 1)
                If (ByteValue And BitValue(BitPos)) Then .Data(j) = 1
                BitPos = BitPos + 1
                If (BitPos = 8) Then
                    ByteValue = byteArray(CurrPos - 1)
                    CurrPos = CurrPos + 1
                    BitPos = 0
                End If
            Next
        End If
    End With
Next
If (BitPos = 0) Then CurrPos = CurrPos - 1
NodesCount = 1
Nodes(0).LeftNode = -1
Nodes(0).RightNode = -1
Nodes(0).ParentNode = -1
Nodes(0).Value = -1
For i = 0 To 255
    Call CreateTree(Nodes(), NodesCount, i, CharValue(i))
Next
ResultLen = 0
    For CurrPos = CurrPos To ByteLen
        ByteValue = byteArray(CurrPos - 1)
        For BitPos = 0 To 7
            If (ByteValue And BitValue(BitPos)) Then NodeIndex = Nodes(NodeIndex).RightNode Else NodeIndex = Nodes(NodeIndex).LeftNode
            If (Nodes(NodeIndex).Value > -1) Then
                Result(ResultLen) = Nodes(NodeIndex).Value
                ResultLen = ResultLen + 1
                If (ResultLen = lResultLen) Then GoTo DecodeFinished
                NodeIndex = 0
            End If
        Next
        If (CurrPos Mod 10000 = 0) Then
            DoEvents
            If AbortUltraRun = True Then Exit Sub
            '------------------------------------------------------
            'remove the following 5 lines if no progressbar is used
            NewProgress = CurrPos / ByteLen * PROGRESS_DECRYPT + PROGRESS_DECHUFF
            If (NewProgress <> CurrProgresValue) Then
                CurrProgresValue = NewProgress
                Call UpdateStatus(CurrProgresValue)
            End If
            '------------------------------------------------------
        End If
    Next
DecodeFinished:
    'check CRC
    Char = 0
    For i = 0 To (ResultLen - 1)
        Char = Char Xor Result(i)
        If (i Mod 10000 = 0) Then
            DoEvents
            If AbortUltraRun = True Then Exit Sub
            '------------------------------------------------------
            'remove the following 5 lines if no progressbar is used
            NewProgress = i / ResultLen * PROGRESS_DECRYPT + PROGRESS_CHECKCRC + PROGRESS_DECHUFF
            If (NewProgress <> CurrProgresValue) Then
                CurrProgresValue = NewProgress
                Call UpdateStatus(CurrProgresValue)
            End If
            '------------------------------------------------------
        End If
    Next
    If (Char <> CheckSum) Then UltraReturnValue = 5
    ReDim byteArray(0 To ResultLen - 1)
    Call CopyMem(byteArray(0), Result(0), ResultLen)
End Sub

Private Sub CreateBitSequences(Nodes() As HUFFMANTREE, ByVal NodeIndex As Integer, bytes As byteArray, CharValue() As byteArray)
    Dim NewBytes As byteArray
    If (Nodes(NodeIndex).Value > -1) Then
        CharValue(Nodes(NodeIndex).Value) = bytes
        Exit Sub
    End If
    If (Nodes(NodeIndex).LeftNode > -1) Then
        NewBytes = bytes
        NewBytes.Data(NewBytes.Count) = 0
        NewBytes.Count = NewBytes.Count + 1
        Call CreateBitSequences(Nodes(), Nodes(NodeIndex).LeftNode, NewBytes, CharValue)
    End If
    If (Nodes(NodeIndex).RightNode > -1) Then
        NewBytes = bytes
        NewBytes.Data(NewBytes.Count) = 1
        NewBytes.Count = NewBytes.Count + 1
        Call CreateBitSequences(Nodes(), Nodes(NodeIndex).RightNode, NewBytes, CharValue)
    End If
End Sub

Private Sub CreateTree(Nodes() As HUFFMANTREE, NodesCount As Long, Char As Long, bytes As byteArray)
Dim a As Integer
Dim NodeIndex As Long
NodeIndex = 0
For a = 0 To (bytes.Count - 1)
    If (bytes.Data(a) = 0) Then
        If (Nodes(NodeIndex).LeftNode = -1) Then
            Nodes(NodeIndex).LeftNode = NodesCount
            Nodes(NodesCount).ParentNode = NodeIndex
            Nodes(NodesCount).LeftNode = -1
            Nodes(NodesCount).RightNode = -1
            Nodes(NodesCount).Value = -1
            NodesCount = NodesCount + 1
        End If
        NodeIndex = Nodes(NodeIndex).LeftNode
    ElseIf (bytes.Data(a) = 1) Then
        If (Nodes(NodeIndex).RightNode = -1) Then
            Nodes(NodeIndex).RightNode = NodesCount
            Nodes(NodesCount).ParentNode = NodeIndex
            Nodes(NodesCount).LeftNode = -1
            Nodes(NodesCount).RightNode = -1
            Nodes(NodesCount).Value = -1
            NodesCount = NodesCount + 1
        End If
        NodeIndex = Nodes(NodeIndex).RightNode
    Else
        Stop
    End If
Next
Nodes(NodeIndex).Value = Char
End Sub

' ------------------------------------------------------------
'                   Base 64 Radix functions
' ------------------------------------------------------------

Private Function PadString(strData As String) As String
Dim nLen As Long
Dim sPad As String
Dim nPad As Integer
nLen = Len(strData)
nPad = ((nLen \ 8) + 1) * 8 - nLen
sPad = String(nPad, Chr(nPad))
PadString = strData & sPad
End Function

Private Function UnpadString(strData As String) As String
Dim nLen As Long
Dim nPad As Long
nLen = Len(strData)
If nLen = 0 Then Exit Function
nPad = Asc(Right(strData, 1))
If nPad > 8 Then nPad = 0
UnpadString = Left(strData, nLen - nPad)
End Function

Private Function EncodeStr64(encString As String, ByVal MaxPerLine As Integer) As String
' Return radix64 of string
Dim abOutput()  As Byte
Dim sLast       As String
Dim b(3)        As Byte
Dim j           As Integer
Dim CharCount   As Integer
Dim iIndex      As Long
Dim Umax        As Long
Dim i As Long, nLen As Long, nQuants As Long
EncodeStr64 = ""
nLen = Len(encString)
nQuants = nLen \ 3
iIndex = 0
If MaxPerLine < 10 Then MaxPerLine = 10
Umax = nQuants + 1
Call MakeEncTab
If (nQuants > 0) Then
    ReDim abOutput(nQuants * 4 - 1)
    For i = 0 To nQuants - 1
        For j = 0 To 2
            b(j) = Asc(Mid(encString, (i * 3) + j + 1, 1))
        Next
        Call EncodeQuantumB(b)
        abOutput(iIndex) = b(0)
        abOutput(iIndex + 1) = b(1)
        abOutput(iIndex + 2) = b(2)
        abOutput(iIndex + 3) = b(3)
        CharCount = CharCount + 4
        ' insert CRLF if max char per line is reached
        If CharCount >= MaxPerLine Then
            ReDim Preserve abOutput(UBound(abOutput) + 2)
            CharCount = 0
            abOutput(iIndex + 4) = 13
            abOutput(iIndex + 5) = 10
            iIndex = iIndex + 6
            Else
            iIndex = iIndex + 4
            End If
    Next
    EncodeStr64 = StrConv(abOutput, vbUnicode)
End If
Select Case nLen Mod 3
Case 0
    sLast = ""
Case 1
    b(0) = Asc(Mid(encString, nLen, 1))
    b(1) = 0
    b(2) = 0
    Call EncodeQuantumB(b)
    sLast = StrConv(b(), vbUnicode)
    sLast = Left(sLast, 2) & "=="
Case 2
    b(0) = Asc(Mid(encString, nLen - 1, 1))
    b(1) = Asc(Mid(encString, nLen, 1))
    b(2) = 0
    Call EncodeQuantumB(b)
    sLast = StrConv(b(), vbUnicode)
    sLast = Left(sLast, 3) & "="
End Select
EncodeStr64 = EncodeStr64 & sLast
End Function

Private Function DecodeStr64(decString As String) As String
' Return string of decoded values from radix64
Dim abDecoded() As Byte
Dim d(3)    As Byte
Dim c       As Integer
Dim di      As Integer
Dim i       As Long
Dim nLen    As Long
Dim iIndex  As Long
Dim Umax    As Long
nLen = Len(decString)
If nLen < 4 Then
    Exit Function
End If
ReDim abDecoded(((nLen \ 4) * 3) - 1)
Umax = nLen
iIndex = 0
di = 0
Call MakeDecTab
For i = 1 To Len(decString)
    c = CByte(Asc(Mid(decString, i, 1)))
    c = aDecTab(c)
    If c >= 0 Then
        d(di) = CByte(c)
        di = di + 1
        If di = 4 Then
            abDecoded(iIndex) = SHL2(d(0)) Or (SHR4(d(1)) And &H3)
            iIndex = iIndex + 1
            abDecoded(iIndex) = SHL4(d(1) And &HF) Or (SHR2(d(2)) And &HF)
            iIndex = iIndex + 1
            abDecoded(iIndex) = SHL6(d(2) And &H3) Or d(3)
            iIndex = iIndex + 1
            If d(3) = 64 Then
                iIndex = iIndex - 1
                abDecoded(iIndex) = 0
            End If
            If d(2) = 64 Then
                iIndex = iIndex - 1
                abDecoded(iIndex) = 0
            End If
            di = 0
        End If
    End If
Next i
DecodeStr64 = StrConv(abDecoded(), vbUnicode)
DecodeStr64 = Left(DecodeStr64, iIndex)
End Function

Private Sub EncodeQuantumB(b() As Byte)
Dim b0 As Byte, b1 As Byte, b2 As Byte, b3 As Byte
b0 = SHR2(b(0)) And &H3F
b1 = SHL4(b(0) And &H3) Or (SHR4(b(1)) And &HF)
b2 = SHL2(b(1) And &HF) Or (SHR6(b(2)) And &H3)
b3 = b(2) And &H3F
b(0) = aEncTab(b0)
b(1) = aEncTab(b1)
b(2) = aEncTab(b2)
b(3) = aEncTab(b3)
End Sub

Private Function MakeDecTab()
Dim t As Integer
Dim c As Integer
For c = 0 To 255
    aDecTab(c) = -1
Next
t = 0
For c = Asc("A") To Asc("Z")
    aDecTab(c) = t
    t = t + 1
Next
For c = Asc("a") To Asc("z")
    aDecTab(c) = t
    t = t + 1
Next
For c = Asc("0") To Asc("9")
    aDecTab(c) = t
    t = t + 1
Next
c = Asc("+")
aDecTab(c) = t
t = t + 1
c = Asc("/")
aDecTab(c) = t
t = t + 1
c = Asc("=")
aDecTab(c) = t
End Function

Private Function MakeEncTab()
Dim i As Integer
Dim c As Integer
i = 0
For c = Asc("A") To Asc("Z")
    aEncTab(i) = c
    i = i + 1
Next
For c = Asc("a") To Asc("z")
    aEncTab(i) = c
    i = i + 1
Next
For c = Asc("0") To Asc("9")
    aEncTab(i) = c
    i = i + 1
Next
c = Asc("+")
aEncTab(i) = c
i = i + 1
c = Asc("/")
aEncTab(i) = c
i = i + 1
End Function

Private Function SHL2(ByVal bytValue As Byte) As Byte
SHL2 = (bytValue * &H4) And &HFF
End Function

Private Function SHL4(ByVal bytValue As Byte) As Byte
SHL4 = (bytValue * &H10) And &HFF
End Function

Private Function SHL6(ByVal bytValue As Byte) As Byte
SHL6 = (bytValue * &H40) And &HFF
End Function

Private Function SHR2(ByVal bytValue As Byte) As Byte
SHR2 = bytValue \ &H4
End Function

Private Function SHR4(ByVal bytValue As Byte) As Byte
SHR4 = bytValue \ &H10
End Function

Private Function SHR6(ByVal bytValue As Byte) As Byte
SHR6 = bytValue \ &H40
End Function

Private Sub SetReturnString()
'get the ultra error descriptions
Select Case UltraReturnValue
Case 0
    UltraReturnString = ""
Case 1
    UltraReturnString = "Cannot continue without text (Error 1)"
Case 2
    UltraReturnString = "Cannot continue without key (Error 2)"
Case 3
    UltraReturnString = "Key too small/is repeating (Error 3)"
Case 4
    UltraReturnString = "Source file not found (Error 4)"
Case 5
    UltraReturnString = "Compression checksum error (Error 5)"
Case 6
    UltraReturnString = "Cannot process empty file (Error 6)"
Case 10
    UltraReturnString = "Crypto version unknown/contains errors (Error 10)"
Case 11
    UltraReturnString = "Encoding has been aborted by user"
Case 12
    UltraReturnString = "File error: " & FileErrDescription & " (Error 12)"
Case 20
    UltraReturnString = "Crypto file version unknown/contains errors (Error 20)"
Case 21
    UltraReturnString = "Decoding has been aborted by user"
Case 22
    UltraReturnString = "File error: " & FileErrDescription & " (Error 22)"
Case 23
    UltraReturnString = "Failed decoding the file (Error 23)"
Case 30
    UltraReturnString = "Crypto header or footer format incomplete/contains errors (Error 30)"
Case 33
    UltraReturnString = "Failed decoding the text (Error 33)"
End Select
End Sub

' ------------------------------------------------------------
'              Miscellanious public functions
' ------------------------------------------------------------

Public Function KeyQuality(ByVal aKey As String) As Integer
' returns an integer value (0 to 100) rating the key quality
Dim QC As Integer
Dim LN As Integer
Dim k As Integer
Dim Uc As Boolean
Dim Lc As Boolean
LN = Len(aKey)
QC = LN * 3
If IsValidKey(aKey) = False Then KeyQuality = 0: Exit Function
'check ucases and lcases
For k = 1 To Len(aKey)
    If Asc(Mid(aKey, k, 1)) > 64 And Asc(Mid(aKey, k, 1)) < 91 Then Uc = True
    If Asc(Mid(aKey, k, 1)) > 96 And Asc(Mid(aKey, k, 1)) < 123 Then Lc = True
Next
If Uc = True And Lc = True Then QC = QC * 1.2
'check numbers
For k = 1 To Len(aKey)
    If Asc(Mid(aKey, k, 1)) > 47 And Asc(Mid(aKey, k, 1)) < 58 Then
        If Uc = True Or Lc = True Then QC = QC * 1.4
        Exit For
        End If
Next
'check signs
For k = 1 To Len(aKey)
    If Asc(Mid(aKey, k, 1)) < 48 Or Asc(Mid(aKey, k, 1)) > 122 Or (Asc(Mid(aKey, k, 1)) > 57 And Asc(Mid(aKey, k, 1)) < 65) Then QC = QC * 1.5: Exit For
Next
If QC > 100 Then QC = 100
KeyQuality = Int(QC)
End Function



Public Function GetFileExt(strFile As String) As String
'returns extension of filename
Dim k   As Integer
Dim pos As Integer
For k = 1 To Len(strFile)
    If Mid(strFile, k, 1) = "." Then pos = k
Next
If pos = Len(strFile) Then pos = 0
If pos = 0 Then
    GetFileExt = ""
    Else
    GetFileExt = LCase(Mid(strFile, pos + 1))
    End If
    
End Function

Public Function GetFilePath(strFile As String) As String
'returns only the path without filename
Dim k As Integer
Dim pos As Integer
For k = 1 To Len(strFile)
    If Mid(strFile, k, 1) = "\" Then pos = k
Next
If pos < 2 Then
    GetFilePath = ""
    Else
    GetFilePath = Left(strFile, pos)
    End If
End Function

Public Function CutFileExt(strFile As String) As String
'returns full path and filename without extension
Dim k As Integer
Dim pos As Integer
For k = 1 To Len(strFile)
    If Mid(strFile, k, 1) = "." Then pos = k
Next
If pos = 0 Then
    CutFileExt = strFile
    Else
    CutFileExt = Left(strFile, pos - 1)
    End If
End Function

Public Function CutFilePath(strFile As String) As String
'returns only the filename without full path
Dim k As Integer
Dim pos As Integer
For k = 1 To Len(strFile)
    If Mid(strFile, k, 1) = "\" Then pos = k
Next
If pos = 0 Then
    CutFilePath = strFile
    Else
    CutFilePath = Mid(strFile, pos + 1)
    End If
End Function

Public Function IsValidKey(ByVal aKey As String) As Boolean
'check if key is at least 5 char long, and doesn't repeat
Dim tmp As String
Dim Wid As Integer
Dim i As Integer
Dim Repro As Boolean
If Len(aKey) < 5 Then Exit Function
For Wid = 1 To Int(Len(aKey) / 2)
    IsValidKey = False
    For i = Wid + 1 To Len(aKey) Step Wid
        If Mid(aKey, 1, Wid) <> Mid(aKey, i, Wid) Then IsValidKey = True: Exit For
    Next
If IsValidKey = False Then Exit For
Next
End Function

Public Function TrimText(ByVal aText As String) As String
'cut off all heading and trailing spaces,tabs,CR's and LF's
Dim tmp As String
BeginCutL:
tmp = Left(aText, 1)
If tmp = Chr(32) Or tmp = Chr(9) Or tmp = Chr(13) Or tmp = Chr(10) Then
    aText = Mid(aText, 2)
    GoTo BeginCutL
    End If
BeginCutR:
tmp = Right(aText, 1)
If tmp = Chr(32) Or tmp = Chr(9) Or tmp = Chr(13) Or tmp = Chr(10) Then
    aText = Left(aText, Len(aText) - 1)
    GoTo BeginCutR
    End If
TrimText = aText
End Function

