VERSION 5.00
Object = "{E8FDD05C-3067-4198-8AEC-1A013A46ABDD}#1.0#0"; "FVUnicodeControl.ocx"
Begin VB.Form frmHelp 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New version?"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6720
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel1 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   873
      Alignment       =   1
      Caption         =   "Co1 gi2 mo71i trong phie6n ba3n 1.2 ?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   33023
   End
   Begin FVUnicodeControl.FVistaUniButton cmdAbout 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BackColor       =   -2147483633
      ButtonStyle     =   3
      Caption         =   "? Ta1c Gia3 ?"
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
   Begin FVUnicodeControl.FVistaUniButton cmdExit 
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   2280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BackColor       =   -2147483633
      ButtonStyle     =   3
      Caption         =   "D9o1ng"
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
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAbout_Click()
frmAbout.Show
Unload Me
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

