VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4680
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4650
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4455
         Begin VB.TextBox txtPassword 
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
            Left            =   1320
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   6
            Top             =   840
            Width           =   2895
         End
         Begin VB.TextBox txtUserName 
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
            Left            =   1320
            MaxLength       =   20
            TabIndex        =   5
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "User Name"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   1800
         Width           =   4455
         Begin OnlineBanking.lvButtons_H cmdExit 
            Height          =   495
            Left            =   2880
            TabIndex        =   8
            Top             =   160
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
            mIcon           =   "frmLogin.frx":0000
         End
         Begin OnlineBanking.lvButtons_H cmdLogin 
            Default         =   -1  'True
            Height          =   495
            Left            =   1320
            TabIndex        =   7
            Top             =   160
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            Caption         =   "&Login"
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
            mIcon           =   "frmLogin.frx":031A
         End
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Sub CmdExit_Click()

    If MsgBox("Are you sure you want to exit ?", 4 + 32, title) = vbNo Then
        GoTo EXITPROCEDURE
        Else
            End
    End If
    
EXITPROCEDURE:
    Exit Sub
End Sub

Private Sub cmdLogin_Click()
On Error GoTo errhandler
    If txtPassword.Text = "" Or txtUserName.Text = "" Then
        MsgBox "Please enter User Name and password ", vbExclamation, title
        txtPassword.SetFocus
        GoTo EXITPROCEDURE
    End If
    
    txtPassword.Text = LCase(txtPassword.Text)
    recUsers.MoveFirst
    Do While Not recUsers.EOF
        If Trim(recUsers!LoginID) = Trim(txtUserName.Text) And _
        Trim(recUsers!Password) = Trim(txtPassword.Text) Then
            
            sndPlaySound App.Path & "\Reminder\reminder.wav", &H1
            frmMain.status.Panels("Role").Text = Trim(recUsers!Role)
            frmMain.status.Panels("Name").Text = Trim(txtUserName.Text)
            
            UserID = recUsers!EmployeeID
            UserName = recUsers!LoginID
            UserPassword = recUsers!Password
            UserRole = recUsers!Role
                
            Call Enable_Menu
        
            If UserRole = "Employee" Or UserRole = "Teller" Then
                frmMain.mnuReports.Enabled = False
                frmMain.mnuUser.Enabled = False
                frmMain.mnuDate.Enabled = False
                frmMain.mnuCheckBook.Enabled = False
            End If
            
            Unload Me
            GoTo EXITPROCEDURE
        End If
        recUsers.MoveNext
    Loop
    
    If recUsers.EOF Then
        MsgBox "Invalid password, kindly retry", vbExclamation, title
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword.Text)
        txtPassword.SetFocus
        GoTo EXITPROCEDURE
    End If
    
EXITPROCEDURE:
    Exit Sub
    
errhandler:
    MsgBox Err.Description, vbCritical, title
    GoTo EXITPROCEDURE
End Sub

Private Sub Form_Load()
On Error GoTo errhandler

    UserID = ""
    UserName = ""
    UserPassword = ""
    UserRole = ""
    
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 3
    Call ConnectMe
    cmdLogin.Enabled = False
    Call Disable_Menu
    frmMain.status.Panels("Role").Text = ""
    frmMain.status.Panels("Name").Text = ""
    
EXITPROCEDURE:
    Exit Sub
    
errhandler:
    MsgBox Err.Description, vbCritical, title
    GoTo EXITPROCEDURE
End Sub


Private Sub txtPassword_GotFocus()
On Error GoTo errhandler

    recUsers.MoveFirst
    Do While Not recUsers.EOF
        If Trim(recUsers!LoginID) = Trim(txtUserName.Text) Then
            cmdLogin.Enabled = True
            GoTo EXITPROCEDURE
        End If
        recUsers.MoveNext
    Loop

    If recUsers.EOF Then
        MsgBox "Invalid login name, kindly retry", vbExclamation, title
        txtUserName.SelStart = 0
        txtUserName.SelLength = Len(txtUserName.Text)
        txtUserName.SetFocus
        GoTo EXITPROCEDURE
    End If

EXITPROCEDURE:
    Exit Sub
errhandler:
    MsgBox Err.Description, vbCritical, title
    GoTo EXITPROCEDURE
End Sub


Private Sub Disable_Menu()

    frmMain.mnuFile.Enabled = False
    frmMain.mnuAdmin.Enabled = False
    frmMain.mnuTrans.Enabled = False
    frmMain.mnuReports.Enabled = False
    frmMain.mnuHelp.Enabled = False
End Sub

Private Sub Enable_Menu()

    frmMain.mnuFile.Enabled = True
    frmMain.mnuAdmin.Enabled = True
    frmMain.mnuTrans.Enabled = True
    frmMain.mnuReports.Enabled = True
    frmMain.mnuHelp.Enabled = True
    frmMain.mnuUser.Enabled = True
    frmMain.mnuCheckBook.Enabled = True
    frmMain.mnuDate.Enabled = True
    
End Sub
