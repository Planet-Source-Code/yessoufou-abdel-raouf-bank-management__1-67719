VERSION 5.00
Begin VB.Form frmLock 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "APPLICATION LOCKED"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4455
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   4215
         Begin VB.TextBox txtPassword 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1200
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   3
            ToolTipText     =   "Enter Password"
            Top             =   720
            Width           =   2895
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "LOCKED"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1200
            TabIndex        =   4
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   4215
         Begin OnlineBanking.lvButtons_H cmdExit 
            Height          =   495
            Left            =   2760
            TabIndex        =   6
            Top             =   150
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            Caption         =   "&Exit"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFHover         =   16777215
            cBhover         =   16711680
            LockHover       =   3
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frmLock.frx":0000
         End
         Begin OnlineBanking.lvButtons_H cmdLogin 
            Default         =   -1  'True
            Height          =   495
            Left            =   1200
            TabIndex        =   7
            Top             =   150
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            Caption         =   "&Ok"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFHover         =   16777215
            cBhover         =   16711680
            LockHover       =   3
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frmLock.frx":031A
         End
      End
   End
End
Attribute VB_Name = "frmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim prompt As String
Dim counter As Integer
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Sub CmdExit_Click()
    If MsgBox("Are you sure you want to exit the application ?", 4 + vbQuestion, title) = vbNo Then
        txtPassword.SetFocus
        GoTo EXITPROCEDURE
        Else
            sndPlaySound App.Path & "\Reminder\reminder.wav", &H1
            End
    End If
EXITPROCEDURE:
    Exit Sub
End Sub


Private Sub Form_Load()
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 3
    Call ConnectMe
End Sub
Private Sub cmdLogin_Click()
    If Trim(txtPassword.Text) = UserPassword Then
        sndPlaySound App.Path & "\Reminder\reminder.wav", &H1
        Unload Me
        Else
            MsgBox "Invalid password, kindly retry ", vbExclamation, title
            txtPassword.SelStart = 0
            txtPassword.SelLength = Len(txtPassword.Text)
            txtPassword.SetFocus
            
    End If
End Sub

