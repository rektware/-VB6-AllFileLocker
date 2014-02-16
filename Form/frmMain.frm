VERSION 5.00
Object = "{E8FDD05C-3067-4198-8AEC-1A013A46ABDD}#1.0#0"; "FVUnicodeControl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Perfect All File Locker - Pro Version *"
   ClientHeight    =   6330
   ClientLeft      =   6990
   ClientTop       =   4605
   ClientWidth     =   5925
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   5925
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel1 
      Height          =   255
      Left            =   4440
      TabIndex        =   47
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Pro Version"
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
   Begin FVUnicodeControl.FVistaUniTabStrip FVistaUniTabStrip1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   8281
      TabCaption(0)   =   "Lock File"
      TabCaption(1)   =   "Option"
      TabCaption(2)   =   "About"
      AutoUni         =   -1  'True
      ActiveTabBackEndColor=   16777215
      ActiveTabBackStartColor=   16777215
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ActiveTabForeColor=   0
      BackColor       =   16777215
      DisabledTabBackColor=   13355721
      DisabledTabForeColor=   10526880
      InActiveTabBackEndColor=   13619151
      InActiveTabBackStartColor=   15461355
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      InActiveTabForeColor=   0
      OuterBorderColor=   9800841
      TabStripBackColor=   -2147483639
      Begin FVUnicodeControl.FVistaUniButton FVistaUniButton1 
         Height          =   375
         Left            =   -19760
         TabIndex        =   48
         Top             =   3720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BackColor       =   -2147483634
         ButtonStyle     =   3
         Caption         =   "Va2o trang chu3"
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
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel2 
         Height          =   495
         Left            =   -18560
         TabIndex        =   46
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   873
         Caption         =   "All File Locker"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel12 
         Height          =   255
         Left            =   -19880
         TabIndex        =   43
         Top             =   2400
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         Alignment       =   1
         Caption         =   "D9inh Quang Trung - dinhquangtrung90@yahoo.com"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin FVUnicodeControl.FVistaUniOption optPass 
         Height          =   195
         Left            =   2040
         TabIndex        =   41
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   344
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Lock by password"
         ForeColor       =   16711680
      End
      Begin FVUnicodeControl.FVistaUniOption optKeyfile 
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   1680
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   344
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
         Caption         =   "Lock by key file"
         ForeColor       =   16711680
      End
      Begin FVUnicodeControl.FVistaUniTextbox txtFileKey 
         Height          =   270
         Left            =   1680
         TabIndex        =   39
         Top             =   2040
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   476
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
         Text            =   ""
         Locked          =   -1  'True
         BorderStyle     =   2
         BorderLine      =   11709605
      End
      Begin FVUnicodeControl.FVistaUniButton cmdFileKey 
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   2040
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         BackColor       =   16744576
         ButtonShape     =   3
         ButtonStyle     =   1
         Caption         =   "Choose key file"
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
      Begin VB.TextBox txtFile 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         OLEDropMode     =   1  'Manual
         TabIndex        =   28
         Top             =   600
         Width           =   3855
      End
      Begin FVUnicodeControl.FVistaUniFrame fm2 
         Height          =   1095
         Left            =   -9760
         TabIndex        =   24
         Top             =   1560
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   1931
         Alignment       =   0
         BackColor       =   16777215
         ForeColor       =   32768
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "With Locked File"
         AutoUnicode     =   -1  'True
         Begin FVUnicodeControl.FVistaUniOption opt3 
            Height          =   195
            Left            =   240
            TabIndex        =   27
            Top             =   720
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   344
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
            Caption         =   "Ask user"
            ForeColor       =   16711680
         End
         Begin FVUnicodeControl.FVistaUniOption opt2 
            Height          =   195
            Left            =   240
            TabIndex        =   26
            Top             =   480
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   344
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Keep original file."
            ForeColor       =   16711680
         End
         Begin FVUnicodeControl.FVistaUniOption opt1 
            Height          =   195
            Left            =   240
            TabIndex        =   25
            Top             =   240
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   344
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Delete original file after lock"
            ForeColor       =   16711680
         End
      End
      Begin FVUnicodeControl.FVistaUniFrame fm1 
         Height          =   855
         Left            =   -9760
         TabIndex        =   21
         Top             =   600
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   1508
         Alignment       =   0
         BackColor       =   16777215
         ForeColor       =   32768
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Encrypt option"
         AutoUnicode     =   -1  'True
         Begin FVUnicodeControl.FVistaUniOption optNo 
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   240
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   344
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Don't encrypt file (Fast)"
            ForeColor       =   16711680
         End
         Begin FVUnicodeControl.FVistaUniOption optYes 
            Height          =   195
            Left            =   240
            TabIndex        =   23
            Top             =   480
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   344
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
            Caption         =   "Encrypt file (Security)"
            ForeColor       =   16711680
         End
      End
      Begin VB.PictureBox picIcon 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF00FF&
         Height          =   975
         Left            =   4320
         ScaleHeight     =   915
         ScaleWidth      =   1155
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin FVUnicodeControl.FVistaUniProgressbar Bar1 
         Height          =   225
         Left            =   120
         Top             =   3720
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   397
         Max             =   100
         Value           =   0
         TStyle          =   2
         Min             =   0
         Style           =   1
         Text            =   "Encrypt"
         Align           =   1
      End
      Begin FVUnicodeControl.FVistaUniLabel txtSafe 
         Height          =   255
         Left            =   3840
         TabIndex        =   1
         Top             =   2760
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         Caption         =   "0 %"
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
      Begin FVUnicodeControl.FVistaUniLabel label99 
         Height          =   495
         Left            =   3840
         TabIndex        =   2
         Top             =   2520
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         BackStyle       =   0
         Caption         =   "Password Quality"
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
      Begin FVUnicodeControl.FVistaUniLabel label77 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   3240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         Alignment       =   2
         Caption         =   "Hints"
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
      Begin FVUnicodeControl.FVistaUniLabel lblFileType 
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   1440
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         Caption         =   ""
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
      Begin FVUnicodeControl.FVistaUniLabel lblFileSize 
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   1200
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         Caption         =   ""
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
      Begin FVUnicodeControl.FVistaUniLabel lblFileName 
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   960
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         Caption         =   ""
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
      Begin FVUnicodeControl.FVistaUniCheckbox lblViewPass 
         Height          =   195
         Left            =   3840
         TabIndex        =   8
         Top             =   3120
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   344
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "View password"
         ForeColor       =   16711680
      End
      Begin FVUnicodeControl.FVistaUniLabel label66 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         Alignment       =   2
         Caption         =   "Retype password "
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
      Begin FVUnicodeControl.FVistaUniLabel label55 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         Alignment       =   2
         Caption         =   "Password"
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
      Begin FVUnicodeControl.FVistaUniLabel Label33 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         Caption         =   "File type"
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
      Begin FVUnicodeControl.FVistaUniLabel Label22 
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         Caption         =   "Size:"
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
      Begin FVUnicodeControl.FVistaUniLabel Label1 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "File name:"
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
      Begin FVUnicodeControl.FVistaUniButton cmdSelectFile 
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         BackColor       =   16744576
         ButtonShape     =   3
         ButtonStyle     =   1
         Caption         =   "Choose file to lock"
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
      Begin FVUnicodeControl.FVistaUniButton cmdLock 
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   3960
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
         BackColor       =   16744576
         ButtonShape     =   3
         ButtonStyle     =   1
         Caption         =   "Lock"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin FVUnicodeControl.FVistaUniButton cmdExit 
         Height          =   495
         Left            =   2880
         TabIndex        =   16
         Top             =   3960
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
         BackColor       =   16744576
         ButtonShape     =   3
         ButtonStyle     =   1
         Caption         =   "Exit"
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
      Begin FVUnicodeControl.FVistaUniTextbox lblEnterPass 
         Height          =   270
         Left            =   1680
         TabIndex        =   17
         Top             =   2520
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   476
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
         Text            =   ""
         PasswordChar    =   "*"
         Enabled         =   0   'False
         BorderStyle     =   2
         BorderLine      =   11709605
      End
      Begin FVUnicodeControl.FVistaUniTextbox lblEnterPassAgain 
         Height          =   270
         Left            =   1680
         TabIndex        =   18
         Top             =   2880
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   476
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
         Text            =   ""
         PasswordChar    =   "*"
         Enabled         =   0   'False
         BorderStyle     =   2
         BorderLine      =   11709605
      End
      Begin FVUnicodeControl.FVistaUniTextbox txtHint 
         Height          =   270
         Left            =   1680
         TabIndex        =   19
         Top             =   3240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   476
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
         Text            =   "No hint!"
         BorderStyle     =   2
         BorderLine      =   11709605
      End
      Begin FVUnicodeControl.FVistaUniButton cmdStop 
         Height          =   255
         Left            =   4800
         TabIndex        =   20
         Top             =   3360
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         BackColor       =   16744576
         ButtonShape     =   3
         ButtonStyle     =   1
         Caption         =   "Stop"
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         State           =   3
      End
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel9 
         Height          =   255
         Left            =   -19280
         TabIndex        =   29
         Top             =   3360
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
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel10 
         Height          =   255
         Left            =   -19280
         TabIndex        =   30
         Top             =   3120
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
      Begin FVUnicodeControl.FVistaUniLabel ll3 
         Height          =   255
         Left            =   -19520
         TabIndex        =   31
         Top             =   2880
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
         Left            =   -15920
         TabIndex        =   32
         Top             =   3720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BackColor       =   -2147483626
         ButtonStyle     =   3
         Caption         =   "Home page"
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
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel13 
         Height          =   255
         Left            =   -17840
         TabIndex        =   33
         Top             =   1680
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "D9inh Quang Trung"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin FVUnicodeControl.FVistaUniLabel ll1 
         Height          =   255
         Left            =   -19160
         TabIndex        =   34
         Top             =   1680
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         Alignment       =   2
         Caption         =   "Author"
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
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel16 
         Height          =   255
         Left            =   -19160
         TabIndex        =   35
         Top             =   1920
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
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel17 
         Height          =   255
         Left            =   -17840
         TabIndex        =   36
         Top             =   1920
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         Caption         =   "http://phanmemvn.net"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin FVUnicodeControl.FVistaUniLabel ll2 
         Height          =   255
         Left            =   -19880
         TabIndex        =   42
         Top             =   2160
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         Alignment       =   1
         Caption         =   "Based on idea of:"
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
      Begin FVUnicodeControl.FVistaUniLabel ll4 
         Height          =   255
         Left            =   -19880
         TabIndex        =   45
         Top             =   1200
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         Alignment       =   1
         Caption         =   "Any file can be locked"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   49152
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
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
         Left            =   -15680
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   37
         Top             =   1680
         Width           =   255
      End
   End
   Begin FVUnicodeControl.FVistaUniCheckbox chkVN 
      Height          =   195
      Left            =   120
      TabIndex        =   44
      Top             =   1200
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Tie61ng Vie65t"
      ForeColor       =   0
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   960
      Picture         =   "frmMain.frx":164A
      Top             =   0
      Width           =   4500
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Dim xStopEncodeFile As Boolean
Public Function GetFileName(ByVal sPath As String) As String
GetFileName = Mid(sPath, InStrRev(sPath, "\") + 1)
End Function
Public Function GetFileType(ByVal sPath As String) As String
GetFileType = Mid(sPath, InStrRev(sPath, ".") + 1)
End Function

Private Sub chkVN_Click()
If chkVN.Value = True Then
cmdSelectFile.Caption = "Cho5n file d9e63 kho1a"
Label1.Caption = "Te6n file:"
Label22.Caption = "Ki1ch thu7o71c:"
Label33.Caption = "Kie63u file:"
optKeyfile.Caption = "Kho1a ba82ng Key File"
optPass.Caption = "Kho1a ba82ng ma65t kha63u"
cmdFileKey.Caption = "Cho5n Key File"
label55.Caption = "Ma65t kha63u"
label66.Caption = "Nha65p la5i ma65t kha63u"
label77.Caption = "Go75i y1"
label99.Caption = "D9o65 an toa2n"
lblViewPass.Caption = "Xem ma65t kha63u"
Bar1.CustomText = "Ma4 ho1a"
cmdLock.Caption = "Kho1a"
cmdExit.Caption = "Thoa1t"
cmdStop.Caption = "Du72ng"
fm1.Caption = "Tu2y cho5n ma4 ho1a"
optNo.Caption = "Kho6ng ma4 ho1a file (Nhanh)"
optYes.Caption = "Ma4 ho1a file (An toa2n)"
fm2.Caption = "D9o61i vo71i file bi5 kho1a"
opt1.Caption = "Xo1a file go61c sau khi kho1a"
opt2.Caption = "Giu73 la5i file go61c sau khi kho1a"
opt3.Caption = "Ho3i y1 kie61n."
ll1.Caption = "Ta1c gia3:"
ll2.Caption = "Chu7o7ng tri2nh d9u7o75c vie61t du75a tre6n y1 tu7o73ng cu3a:"
ll3.Caption = "Gu73i lo72i ca3m o7n d9e61n:"
cmdGotoHomePages.Caption = "Trang chu3"
Else
cmdSelectFile.Caption = "Choose file to lock"
Label1.Caption = "File name:"
Label22.Caption = "Size:"
Label33.Caption = "File type:"
optKeyfile.Caption = "Lock by key file"
optPass.Caption = "Lock by password"
cmdFileKey.Caption = "Choose key file"
label55.Caption = "Password"
label66.Caption = "Retype password"
label77.Caption = "Hints"
label99.Caption = "Password Quality"
lblViewPass.Caption = "View password"
Bar1.CustomText = "Encrypt"
cmdLock.Caption = "Lock"
cmdExit.Caption = "Exit"
cmdStop.Caption = "Stop"
fm1.Caption = "Encrypt option"
optNo.Caption = "Don't encrypt file (Fast)"
optYes.Caption = "Encrypt file (Security)"
fm2.Caption = "With Locked File"
opt1.Caption = "Delete original file after lock"
opt2.Caption = "Keep original file."
opt3.Caption = "Ask user"
ll1.Caption = "Author:"
ll2.Caption = "Based on idea of:"
ll3.Caption = "Great Thank To:"
cmdGotoHomePages.Caption = "Home Page"
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub



Private Sub cmdFileKey_Click()
If txtFile.Text = "" Then
    If chkVN.Value = False Then
        MsgBox "Choose file to be locked, then choose key file"
    Else
        UniMsgBox zToUnicode("Cho5n file d9e63 kho1a tru7o71c, sau d9o1 mo71i cho5n key file!"), , "!"
    End If
    Exit Sub
End If
txtFileKey.Text = DiaLog1.ShowOpen("", , "", "Select Key File")
If txtFileKey.Text = "" Then Exit Sub
'Warning ! Do not use the file going to be locked as key or it will never be opened
If GetMD5(txtFileKey.Text) = GetMD5(txtFile.Text) Then
    
    If chkVN.Value = False Then
    If MsgBox("Do not use the file going to be locked as key or it will never be opened" & vbCrLf & "Do you want to continue?", vbYesNo, "Warning ! ") = vbNo Then
        lblEnterPass.Text = ""
        lblEnterPassAgain.Text = ""
        txtFileKey.Text = ""
        Exit Sub
    End If
    Else
    If UniMsgBox(zToUnicode("Kho6ng d9u7o75c su73 du5ng ca1c file se4 bi5 kho1a d9e63 la2m key file! Ne61u kho6ng, no1 se4 kho6ng bao gio72 co1 the63 mo73a kho1a d9u7o75c." & vbCrLf & "Ba5n co1 muo61n tie61p tu5c kho6ng?"), vbYesNo, "Warning ! ") = vbNo Then
        lblEnterPass.Text = ""
        lblEnterPassAgain.Text = ""
        txtFileKey.Text = ""
        Exit Sub
    End If
    End If
    
    lblEnterPass.Text = GetMD5(txtFileKey.Text)
    lblEnterPassAgain.Text = GetMD5(txtFileKey.Text)

Else
    lblEnterPass.Text = GetMD5(txtFileKey.Text)
    lblEnterPassAgain.Text = GetMD5(txtFileKey.Text)

End If
txtSafe.Caption = KeyQuality(Me.lblEnterPass.Text) & " %"
End Sub

Private Sub cmdGotoHomePages_Click()
Shell "explorer http://phanmemvn.net"
End Sub

Private Sub cmdLock_Click()

If txtHint.Text = "" Then txtHint.Text = "No hint!"
cmdExit.Enabled = False
If txtFile.Text = "" Then
    If chkVN.Value = True Then
    UniMsgBox zToUnicode("Ha4y cho5n file!"), , "!"
    Else
    MsgBox "Please select file!", vbOKOnly, "Error!"
    End If
    cmdExit.Enabled = True
    Exit Sub
End If
If lblEnterPass.Text = "" Or lblEnterPass.Text <> lblEnterPassAgain.Text Then
    If chkVN.Value = True Then
    UniMsgBox zToUnicode("Lo64i ma65t kha63u!"), , "!"
    Else
    MsgBox "Password error!", vbOKOnly, "Error!"
    End If
    cmdExit.Enabled = True
    Exit Sub
End If
If Len(lblEnterPass.Text) < 4 Then
    If chkVN.Value = False Then
    If MsgBox("Password is short! Do you want continue to lock file?", vbYesNo + vbCritical, "!") = vbNo Then Exit Sub
    Else
    If UniMsgBox(zToUnicode("Ma65t kha63u qu1a nga81n. Ba5n co1 muo61n tie61p tu5c kho1a kho6ng?"), vbYesNo + vbCritical, "!") = vbNo Then Exit Sub
    End If
End If

If FileExists(txtFile.Text) = False Then
    If chkVN.Value = False Then
    MsgBox "File not found!", vbOKOnly, "!"
    Else
    UniMsgBox zToUnicode("Ta65p tin kho6ng to62n ta5i!"), , "!"
    End If
    cmdExit.Enabled = True
    Exit Sub
End If


If lblFileType.Caption = "exe" Then
    On Error GoTo EssSS
    Dim PropBag1 As New PropertyBag
    Dim BeginPos As Long
    Dim varTemp As Variant
    Dim byteArr() As Byte
    Open txtFile.Text For Binary As #1
        Get #1, LOF(1) - 3, BeginPos
        Seek #1, BeginPos
        Get #1, , varTemp
        byteArr = varTemp
        PropBag1.Contents = byteArr
        PropBag1.WriteProperty "LOF", LOF(1)
        PropBag1.WriteProperty "BeginPos", BeginPos
    Close #1
    If PropBag1.ReadProperty("sPassword") <> "" Then
        MsgBox "File has been locked before! Can't lock it again!"
        cmdExit.Enabled = True
        Exit Sub
    End If
EssSS:
Close #1
End If

DisableKey
xStopEncodeFile = False

LockFile

EnableKey

End Sub

Private Sub cmdSelectFile_Click()
txtFile.Text = DiaLog1.ShowOpen("", , "", "Select File To Lock")
If txtFile.Text <> "" Then
    GetIconFromFile txtFile.Text, picIcon
    picIcon.AutoSize = True
    lblFileName.Caption = GetFileName(txtFile.Text)
    lblFileSize.Caption = FileLen(txtFile.Text) & " Bytes"
    lblFileType.Caption = GetFileType(txtFile.Text)
End If
End Sub



Private Sub cmdStop_Click()
xStopEncodeFile = True
AbortUltraRun = True
End Sub






Private Sub Form_Load()
If App.PrevInstance = True Then
    End
End If
End Sub

Private Sub Label2_Click()
MsgBox "Author: DinhQuangTrung90@yahoo.com" & vbCrLf & "Website: http://phanmemvn.net", vbOKOnly, "About Me"

End Sub

Private Sub lblEnterPass_GotFocus()
lblEnterPass.SelectTextAll = True
End Sub


Private Sub lblEnterPass_KeyUp(KeyCode As Integer, Shift As Integer)
txtSafe.Caption = KeyQuality(Me.lblEnterPass.Text) & " %"
End Sub

Private Sub lblEnterPassAgain_GotFocus()
lblEnterPassAgain.SelectTextAll = True
End Sub


Private Sub lblViewPass_Click()
If lblViewPass.Value = True Then
lblEnterPass.PasswordChar = ""
lblEnterPassAgain.PasswordChar = ""
Else
lblEnterPass.PasswordChar = "*"
lblEnterPassAgain.PasswordChar = "*"
End If
End Sub


Private Sub LockFile()
On Error Resume Next
SetAttr txtFile.Text, vbNormal
Dim xhDetectCode As String
xhDetectCode = MaHoa(MD5(MD5(lblEnterPass.Text) & txtHint.Text & lblFileName.Caption), 55)

If optYes.Value = True Then
    Me.cmdStop.Enabled = True
    EncodeFile txtFile.Text, txtFile.Text & ".afl", xhDetectCode
    If xStopEncodeFile = True Then Exit Sub
    Me.cmdStop.Enabled = False
    
    If opt1.Value = True Then
        DeleteFile txtFile.Text
    End If
    If opt2.Value = True Then
        Name txtFile.Text As txtFile.Text & ".bak"
    End If
    If opt3.Value = True Then
        If chkVN.Value = False Then
        
            If MsgBox("Do you want to delete original file?", vbYesNo, "OK?") = vbYes Then
                DeleteFile txtFile.Text
            Else
                Name txtFile.Text As txtFile.Text & ".bak"
            End If
        Else
            If UniMsgBox(zToUnicode("Ba5n co1 muo61n xo1a file go61c hay kho6ng?"), vbYesNo, "OK?") = vbYes Then
                DeleteFile txtFile.Text
            Else
                Name txtFile.Text As txtFile.Text & ".bak"
            End If
        End If
    End If

    Name txtFile.Text & ".afl" As txtFile.Text
End If

    Dim BeginPos As Long
    Dim PropBag As New PropertyBag
    Dim varTemp As Variant
    Dim FileName As String
    Dim SendFile()  As Byte
    DoEvents
    Open txtFile.Text For Binary Access Read As #1
        ReDim SendFile(LOF(1) - 1)
        Get #1, , SendFile
    Close #1
    DoEvents

    With PropBag
    
        Dim hDetectCode As String
        hDetectCode = MaHoa(MD5(MD5(lblEnterPass.Text) & txtHint.Text & lblFileName.Caption & FileLen(txtFile.Text)), 55)
        
        Dim Temp As New Collection, i As Integer, NN As Integer
        Dim xKQ As Integer
            For i = 1 To (7) + 1 'Tao gia tri ban dau
                Temp.add (i)
            Next
        For i = 0 To (7) - 1 'Tao so ngau nhien
            Randomize
            NN = Fix(Rnd(1) * (Temp.Count - 1)) + 1
            xKQ = Temp(NN)
            'List1.AddItem xKQ
            If optYes.Value = True Then .WriteProperty "CMHFHK", 1, 0 Else .WriteProperty "CMHFHK", 0, 0
            Select Case xKQ
                Case 1
                    
                    .WriteProperty MaHoa(FileLen(txtFile.Text), 44), hDetectCode
                Case 2
                    
                    .WriteProperty MaHoa("sPassword", 55), MaHoa(MD5(lblEnterPass.Text), 55), ""
                Case 3
                    
                    .WriteProperty "sHint", MaHoa(txtHint.Text, 55), ""
                Case 4
                    
                    .WriteProperty MaHoa("sFileName", 55), MaHoa(lblFileName.Caption, 55), ""
                Case 5
                    
                    .WriteProperty MaHoa("sFileSize", 55), MaHoa(FileLen(txtFile.Text), 55)
                Case 6
                    
                    .WriteProperty hDetectCode, SendFile, ""
                Case 7
                    
                    .WriteProperty MaHoa("sFileIcon", 55), picIcon.Image, ""
            End Select
            
            Temp.Remove (NN)
        Next
    'MsgBox hDetectCode & vbCrLf & xhDetectCode
        'MsgBox MD5(lblEnterPass.Text) & vbCrLf _
         & txtHint.Text & vbCrLf _
         & lblFileName.Caption & vbCrLf _
            & FileLen(txtFile.Text) & vbCrLf
        
    End With
    
    
    SetAttr txtFile.Text, vbNormal
    DeleteFile txtFile.Text
    
    
    'Extract File
    Dim ocxDir$
    ocxDir = txtFile.Text
    If (FileExists(ocxDir) = False) Then
    Dim bytResourceData() As Byte
    bytResourceData = LoadResData(101, "FILE_LOCKED")
    Open ocxDir For Binary Shared As #1
    Put #1, 1, bytResourceData
    Close #1
    End If
    
    
    
    Open txtFile.Text For Binary As #1
        BeginPos = LOF(1)
        varTemp = PropBag.Contents
        Seek #1, LOF(1)
        On Error GoTo ErFileRuN
        Put #1, , varTemp
        Put #1, , BeginPos
    Close #1
    Name txtFile.Text As txtFile.Text & ".exe"
    If chkVN.Value = True Then
    UniMsgBox zToUnicode("D9a4 kho1a tha2nh co6ng!"), , "!"
    Else
    MsgBox "File locked successfully!", vbOKOnly, "OK!"
    End If
    
    Exit Sub
ErFileRuN:
Close #1
UniMsgBox "Error: " & Err & " - " & Err.Description
End Sub



Private Sub txtFile_Changed()

End Sub

Private Sub optKeyfile_Click()
cmdFileKey.Enabled = True
lblEnterPass.Enabled = False
lblEnterPassAgain.Enabled = False

End Sub

Private Sub optPass_Click()
cmdFileKey.Enabled = False
lblEnterPass.Enabled = True
lblEnterPassAgain.Enabled = True

End Sub

Private Sub txtFile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    txtFile.Text = Data.Files(1)
    
    If txtFile.Text <> "" Then
    GetIconFromFile txtFile.Text, picIcon
    picIcon.AutoSize = True
    lblFileName.Caption = GetFileName(txtFile.Text)
    lblFileSize.Caption = FileLen(txtFile.Text) & " Bytes"
    lblFileType.Caption = GetFileType(txtFile.Text)
End If
End Sub

Private Sub txtHint_GotFocus()
txtHint.SelectTextAll = True
End Sub


Sub DisableKey()
cmdSelectFile.Enabled = False
txtFile.Enabled = False
lblEnterPass.Enabled = False
lblEnterPassAgain.Enabled = False
txtHint.Enabled = False
Me.optNo.Enabled = False
Me.optNo.Enabled = False
Me.cmdLock.Enabled = False
Me.optYes.Enabled = False
Me.lblViewPass.Enabled = False
Me.cmdExit.Enabled = False
End Sub
Sub EnableKey()
cmdSelectFile.Enabled = True
txtFile.Enabled = True
lblEnterPass.Enabled = True
lblEnterPassAgain.Enabled = True
txtHint.Enabled = True
Me.optNo.Enabled = True
Me.optNo.Enabled = True
Me.optYes.Enabled = True
Me.cmdLock.Enabled = True
Me.lblViewPass.Enabled = True
Me.cmdExit.Enabled = True
End Sub

