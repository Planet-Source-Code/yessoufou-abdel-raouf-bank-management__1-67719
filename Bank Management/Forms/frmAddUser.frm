VERSION 5.00
Begin VB.Form frmAddUser 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADD USER"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5070
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5055
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   4815
         Begin VB.TextBox txtLoginName 
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
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   0
            ToolTipText     =   "Login Name"
            Top             =   240
            Width           =   2895
         End
         Begin VB.TextBox txtPass 
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
            Left            =   1800
            PasswordChar    =   "*"
            TabIndex        =   1
            ToolTipText     =   "Password"
            Top             =   720
            Width           =   2895
         End
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
            Left            =   1800
            PasswordChar    =   "*"
            TabIndex        =   2
            ToolTipText     =   "Retype The User Password"
            Top             =   1200
            Width           =   2895
         End
         Begin VB.ComboBox comboRole 
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
            ItemData        =   "frmAddUser.frx":0000
            Left            =   1800
            List            =   "frmAddUser.frx":000A
            TabIndex        =   3
            ToolTipText     =   "Select A Role"
            Top             =   1680
            Width           =   2895
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Login Name"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Comfirm Password"
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Role"
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   1680
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   4815
         Begin OnlineBanking.lvButtons_H cmdExit 
            Height          =   495
            Left            =   3360
            TabIndex        =   11
            Top             =   150
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            Caption         =   "&Close"
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
            mIcon           =   "frmAddUser.frx":0027
         End
         Begin OnlineBanking.lvButtons_H cmdOk 
            Default         =   -1  'True
            Height          =   495
            Left            =   1800
            TabIndex        =   12
            Top             =   150
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            Caption         =   "&Add"
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
            mIcon           =   "frmAddUser.frx":0341
         End
      End
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim control As Object

Private Sub CmdExit_Click()
    If MsgBox("Are you sure you want to close this window ?", vbQuestion + 4, title) = vbNo Then
        GoTo EXITPROCEDURE
        Else
            Unload Me
            frmUsers.Show
    End If
    
EXITPROCEDURE:
    Exit Sub
End Sub

Private Sub cmdOk_Click()
On Error GoTo abdel
    
    If txtLoginName.Text = "" Then
        MsgBox "Please enter user name ", vbExclamation, title
        txtLoginName.SetFocus
        GoTo EXITPROCEDURE
    End If
    
    If txtPass.Text = "" Then
        MsgBox "Please enter password ", vbExclamation, title
        txtPass.SetFocus
        GoTo EXITPROCEDURE
    End If
    
    If txtPassword.Text = "" Then
        MsgBox "Please confirm your password ", vbExclamation, title
        txtPassword.SetFocus
        GoTo EXITPROCEDURE
    End If
    
    If comboRole.Text = "" Then
        MsgBox "Kindly choose your role", vbExclamation, title
        comboRole.SetFocus
        GoTo EXITPROCEDURE
    End If
    
    If Trim(txtPass.Text) <> Trim(txtPassword.Text) Then
        MsgBox "The tow passwords are not the same !", vbExclamation, title
        txtPass.SelStart = 0
        txtPass.SelLength = Len(txtPass.Text)
        txtPass.SetFocus
        GoTo EXITPROCEDURE
    End If
    
    If blUpdateUser = True Then
        recUsers.MoveFirst
        Do While Not recUsers.EOF
            If recUsers!EmployeeID = frmUsers.lstUserID Then
                recUsers!LoginID = Trim(txtLoginName.Text)
                recUsers!Password = Trim(txtPass.Text)
                recUsers!Role = comboRole.Text
                recUsers.UpdateBatch adAffectCurrent
            End If
            recUsers.MoveNext
        Loop
        MsgBox "User's details updated successfully!" & vbCrLf & "You need to login again.", vbExclamation, title
        Unload Me
        frmLogin.Show
    End If
    
    If blAddUser = True Then
        
        If txtLoginName = UserName Then
            MsgBox "User Name already exist, Kindly change the name!", vbExclamation, title
            txtLoginName.SelStart = 0
            txtLoginName.SelLength = Len(txtLoginName.Text)
            txtLoginName.SetFocus
            GoTo EXITPROCEDURE
        End If
    
        recUsers.AddNew
        recUsers!LoginID = Trim(txtLoginName.Text)
        recUsers!EmployeeID = newUserID
        recUsers!Password = Trim(txtPass.Text)
        recUsers!Role = comboRole.Text
        recUsers.Update
        MsgBox "New user added successfully!" & vbCrLf & "You need to login again.", vbExclamation, title
        Unload Me
        frmLogin.Show
    End If
    
EXITPROCEDURE:
Exit Sub
abdel:
    MsgBox "Sorry, transactions unsuccessful", vbExclamation, title
    GoTo EXITPROCEDURE
End Sub



Private Sub comboRole_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Form_Load()
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 3
    Call ConnectMe
End Sub

Private Sub Form_Unload(Cancel As Integer)
    blAddUser = False
    blUpdateUser = False
End Sub

