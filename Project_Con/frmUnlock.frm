VERSION 5.00
Begin VB.Form frmUnlock 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Password - All File Locker Final Version"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUnlock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6360
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar Bar1 
      Height          =   255
      Left            =   240
      Max             =   100
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Frame fm 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.Timer tmrDeLe 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   5640
         Top             =   360
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Choose key to open"
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   840
         Width           =   3615
      End
      Begin VB.CommandButton cmdSaveFile 
         Caption         =   "Save"
         Height          =   375
         Left            =   5040
         TabIndex        =   6
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton cmdUnlock 
         Caption         =   "Run"
         Height          =   375
         Left            =   5040
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1200
         TabIndex        =   4
         Text            =   "or type password here"
         Top             =   1320
         Width           =   3585
      End
      Begin UnLockFile.Label Label1 
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1085
         BackColor       =   16777215
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UnLockFile.Label lblHints 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1800
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   450
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
      End
      Begin VB.Image picFileIcon 
         Height          =   735
         Left            =   120
         Top             =   1080
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmUnlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Const BIF_RETURNONLYFSDIRS = 1
Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Dim PropBag As New PropertyBag
Dim sRunFile As Boolean

Dim jjjDete As String
Dim strDele As String
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Const GWL_STYLE = (-16)

Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4

Private Const WS_BORDER = &H800000
Private Const WS_CAPTION = &HC00000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_THICKFRAME = &H40000
Dim xyzDetect As String
Dim xMH As Boolean

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Sub Shell32Bit(ByVal fullpath As String)
         Dim hProcess As Long
         Dim RetVal As Long
         hProcess = OpenProcess(&H400, False, Shell(fullpath, 1))
         Do
             GetExitCodeProcess hProcess, RetVal
             DoEvents: Sleep 100
         Loop While RetVal = &H103
End Sub
Private Sub NoBorder()
    'Sizable = No
    Dim lPrevStyle As Long
    lPrevStyle = GetWindowLong(frmUnlock.hWnd, GWL_STYLE)
    Call SetWindowLong(frmUnlock.hWnd, GWL_STYLE, (lPrevStyle And (Not WS_THICKFRAME) And (Not WS_BORDER) And (Not WS_CAPTION) And (Not WS_MINIMIZEBOX) And (Not WS_MAXIMIZEBOX)))
    Call SetWindowPos(frmUnlock.hWnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER)
End Sub

Private Sub YesBorder()
    'Sizable = Yes
    Dim lPrevStyle As Long
    lPrevStyle = GetWindowLong(frmUnlock.hWnd, GWL_STYLE)
    Call SetWindowLong(frmUnlock.hWnd, GWL_STYLE, (lPrevStyle Or WS_THICKFRAME Or WS_BORDER Or WS_CAPTION Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX))
    Call SetWindowPos(frmUnlock.hWnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER)
End Sub




Public Function GetFileType(ByVal sPath As String) As String
GetFileType = Mid(sPath, InStrRev(sPath, ".") + 1)
End Function



Private Sub Command1_Click()
txtPassword.Text = GetMD5(Module1.ShowOpen(, , , "Select Key File to Open this file"))

End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState <> 1 Then
fm.Left = Me.Width / 2 - fm.Width / 2 - 50
fm.Top = Me.Height / 2 - fm.Height / 2 - 300
End If
End Sub

Private Sub cmdSaveFile_Click()
sRunFile = False
UnLockFile

End Sub

Private Sub cmdUnlock_Click()
sRunFile = True
UnLockFile

End Sub

Private Sub Form_Load()

Label1.Caption = "This file has been locked by All File Locker" & vbCrLf & "Please enter your password to run this file."

xMH = False
On Error GoTo Err

    
    Dim BeginPos As Long
    Dim varTemp As Variant
    
    Dim byteArr() As Byte
    Dim AppPath
    AppPath = App.Path
    If Right(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
    Open AppPath & App.EXEName & ".exe" For Binary As #1
        Get #1, LOF(1) - 3, BeginPos

        Seek #1, BeginPos
        Get #1, , varTemp
        
        byteArr = varTemp
        PropBag.Contents = byteArr
    
        PropBag.WriteProperty "LOF", LOF(1)
        PropBag.WriteProperty "BeginPos", BeginPos
    Close #1
    
    Dim Cong1
        Cong1 = FreeFile
    Dim sFileD() As Byte
    
    Dim xDetectCode As String
    
    Dim kDetect As String
    
    Dim xPassword As String
    Dim xHints As String
    Dim xFileName As String
    Dim xFileSize As String
    
    xPassword = GiaiMa(PropBag.ReadProperty(MaHoa("sPassword", 55)), 55)
    xHints = GiaiMa(PropBag.ReadProperty("sHint"), 55)
    xFileName = GiaiMa(PropBag.ReadProperty(MaHoa("sFileName", 55), ""), 55)
    xFileSize = GiaiMa(PropBag.ReadProperty(MaHoa("sFileSize", 55)), 55)
    
    jjjDete = MaHoa(MD5(xPassword & xHints & xFileName), 55)
    If PropBag.ReadProperty("CMHFHK", 0) = 0 Then xMH = False Else xMH = True
    kDetect = PropBag.ReadProperty(MaHoa(xFileSize, 44))
    xyzDetect = kDetect
    xDetectCode = MaHoa(MD5(xPassword & xHints & xFileName & xFileSize), 55)
    If xDetectCode <> kDetect Then
        MsgBox "File's data has been changed!" & vbCrLf & "Can't open file", vbCritical, "Nice try! ;)"
        End
    End If
    Set picFileIcon.Picture = PropBag.ReadProperty(MaHoa("sFileIcon", 55))
    lblHints.Caption = "G" & ChrW(7907) & "i " & ChrW(221) & ": " & xHints
    lblHints.Refresh
Err:

End Sub

Private Sub tmrDeLe_Timer()
On Error Resume Next
SetAttr strDele, vbNormal
DeleteFile (strDele)
If FileExists(strDele) = False Then End
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    sRunFile = True
    UnLockFile
End If
End Sub
Private Sub UnLockFile()
On Error GoTo ErrH
'/////////
DeleOK = True
Dim PropBag As New PropertyBag
    Dim BeginPos As Long
    Dim varTemp As Variant
    
    Dim byteArr() As Byte
    Dim AppPath
    AppPath = App.Path
    If Right(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
    Open AppPath & App.EXEName & ".exe" For Binary As #1
        Get #1, LOF(1) - 3, BeginPos

        Seek #1, BeginPos
        Get #1, , varTemp
        
        byteArr = varTemp
        PropBag.Contents = byteArr
    
        PropBag.WriteProperty "LOF", LOF(1)
        PropBag.WriteProperty "BeginPos", BeginPos
    Close #1
'//////////


Dim sPass As String
sPass = GiaiMa(PropBag.ReadProperty(MaHoa("sPassword", 55), "xxxxx"), 55)
If MD5(txtPassword.Text) = sPass Then

MsgBox "Password OK!"
Bar1.Visible = True
    Dim Cong1
        Cong1 = FreeFile
    Dim sFile() As Byte
    Dim sFileName As String
    Dim sPath As String
    
    sFileName = GiaiMa(PropBag.ReadProperty(MaHoa("sFileName", 55)), 55)
        sFile = PropBag.ReadProperty(xyzDetect)
        sPath = GetTempFolder & "\" & sFileName
        
    If FileExists(sPath) = True Then
    SetAttr sPath, vbNormal
    DeleteFile sPath
    End If
        If sRunFile = True Then
                Open sPath For Binary As Cong1
                    Put Cong1, , sFile
                Close Cong1
            If xMH = True Then
                DecodeFile sPath, sPath & "afl", jjjDete
                DeleteFile sPath
                Name sPath & "afl" As sPath
            End If
                App.TaskVisible = False
                Me.Hide
                ShellExecute Me.hWnd, vbNullString, sPath, vbNullString, "C:\", 1
                tmrDeLe.Enabled = True
                strDele = sPath
            
        Else
            Dim sFileToSave
            Dim sFileType
            sFileType = GetFileType(sFileName)
            sFileToSave = FolderBrowser
            
            If sFileToSave <> "" Then
                If Right(sFileToSave, 1) <> "\" Then sFileToSave = sFileToSave & "\"
                Open sFileToSave & sFileName For Binary As Cong1
                        Put Cong1, , sFile
                Close Cong1
                If xMH = True Then
                    DecodeFile sFileToSave & sFileName, sFileToSave & sFileName & "afl", jjjDete
                    DeleteFile sFileToSave & sFileName
                    Name sFileToSave & sFileName & "afl" As sFileToSave & sFileName
                End If
            End If
        End If
    

Else
    MsgBox "Password Failed!", vbOKOnly, "!"
    txtPassword.Text = ""
End If
ErrH:
End Sub
Function FileExists(sFile As String) As Boolean
On Error Resume Next
FileExists = ((GetAttr(sFile) And vbDirectory) = 0)
End Function

Function FolderBrowser() As String
    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo
    With udtBI
        .hwndOwner = Me.hWnd
        .lpszTitle = lstrcat("C:\", "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        SHGetPathFromIDList lpIDList, sPath
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If

    FolderBrowser = sPath
End Function

