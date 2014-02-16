VERSION 5.00
Object = "{E8FDD05C-3067-4198-8AEC-1A013A46ABDD}#1.0#0"; "FVUnicodeControl.ocx"
Begin VB.Form frmAbout 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About All File Locker"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel8 
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   2520
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      Caption         =   "D. Rijmenants - With your EncodeFile module"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   32768
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel3 
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   2280
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      Caption         =   "http://caulacbovb.com - With Vista Control"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   32768
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel2 
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Caption         =   "Great Thank To:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   32768
   End
   Begin FVUnicodeControl.FVistaUniButton cmdGotoHomePages 
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   -2147483626
      ButtonStyle     =   3
      Caption         =   "Trang Chu3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel5 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      Alignment       =   2
      Caption         =   "Ngo6n ngu74 su73 du5ng:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483635
   End
   Begin FVUnicodeControl.FVistaUniButton cmdExit 
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   2880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BackColor       =   -2147483633
      ButtonStyle     =   3
      Caption         =   "D9o1ng"
      Effects         =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicOpacity      =   0
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel6 
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      Caption         =   "D9inh Quang Trung"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel4 
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      Alignment       =   2
      Caption         =   "Ta1c Gia3:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483635
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel7 
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      Caption         =   "Visual Basic 6.0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel13 
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      Alignment       =   2
      Caption         =   "Website:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483635
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel14 
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   1320
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      Caption         =   "http://phanmemtiengviet.co.cc"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel1 
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Final Version"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
   End
   Begin FVUnicodeControl.FVistaUniLabel lblMain2 
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "Kho1a Ta61t Ca3 Ca1c Loa5i File - Ba3o Ma65t To61t Ho7n"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16576
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1320
      Picture         =   "frmAbout.frx":15162
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2085
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4080
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   3
      Top             =   1080
      Width           =   255
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" _
      Alias "ShellExecuteA" ( _
      ByVal hwnd As Long, _
      ByVal lpOperation As String, _
      ByVal lpFile As String, _
      ByVal lpParameters As String, _
      ByVal lpDirectory As String, _
      ByVal nShowCmd As Long) As Long

Private Sub cmdExit_Click()
Unload Me
End Sub


Private Sub cmdGotoHomePages_Click()
Shell "explorer http://phanmemtiengviet.co.cc"
End Sub

Private Sub Form_Load()

End Sub

Private Sub Label1_Click()
UniMsgBox ChrW$(&H54) & ChrW$(&HE1) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H67) & ChrW$(&H69) & ChrW$(&H1EA3) & ChrW$(&H3A) & ChrW$(&H20) & ChrW$(&H110) & ChrW$(&H69) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H51) & ChrW$(&H75) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H54) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H2D) & ChrW$(&H20) & ChrW$(&H31) & ChrW$(&H32) & ChrW$(&H2F) & ChrW$(&H31) & ChrW$(&H32) & ChrW$(&H2F) & ChrW$(&H31) & ChrW$(&H39) & ChrW$(&H39) & ChrW$(&H33) _
& vbCrLf & vbCrLf _
& ChrW$(&H4C) & ChrW$(&H1EDB) & ChrW$(&H70) & ChrW$(&H20) & ChrW$(&H31) & ChrW$(&H30) & ChrW$(&H54) & ChrW$(&H32) & ChrW$(&H20) & ChrW$(&H54) & ChrW$(&H72) & ChrW$(&H1B0) & ChrW$(&H1EDD) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H54) & ChrW$(&H48) & ChrW$(&H50) & ChrW$(&H54) & ChrW$(&H20) & ChrW$(&H110) & ChrW$(&H1ED3) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H50) & ChrW$(&H68) & ChrW$(&HFA) & ChrW$(&H2E) _
& vbCrLf & vbCrLf & ChrW$(&H4D) & ChrW$(&H1ECD) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H68) & ChrW$(&H1EAF) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&H1EAF) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H78) & ChrW$(&H69) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H6C) & ChrW$(&H69) & ChrW$(&HEA) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H68) & ChrW$(&H1EC7) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H72) & ChrW$(&H1EF1) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H69) & ChrW$(&H1EBF) & ChrW$(&H70) & ChrW$(&H3A) & ChrW$(&H20) & ChrW$(&H44) & ChrW$(&H69) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H51) & ChrW$(&H75) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H54) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H39) & ChrW$(&H30) & ChrW$(&H40) & ChrW$(&H59) & ChrW$(&H61) & ChrW$(&H68) & ChrW$(&H6F) & ChrW$(&H6F) & ChrW$(&H2E) & ChrW$(&H43) & ChrW$(&H6F) & ChrW$(&H6D) _
 , vbOKOnly, "About Me", Me.hwnd

End Sub
