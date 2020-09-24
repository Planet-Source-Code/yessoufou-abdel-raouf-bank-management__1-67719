VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmUsers 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "USERS"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   5790
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   5535
         Begin MSComctlLib.ListView list 
            Height          =   2655
            Left            =   120
            TabIndex        =   3
            ToolTipText     =   "Double click to modify a uer record"
            Top             =   480
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   4683
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "LoginID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Login Name"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Role"
               Object.Width           =   3969
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Password"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "USERS"
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
            Left            =   1320
            TabIndex        =   4
            Top             =   120
            Width           =   3015
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         TabIndex        =   2
         Top             =   3360
         Width           =   5535
         Begin OnlineBanking.lvButtons_H cmdUpdate 
            Height          =   495
            Left            =   1440
            TabIndex        =   5
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
            Caption         =   "&Update"
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
            mIcon           =   "frmUsers.frx":0000
         End
         Begin OnlineBanking.lvButtons_H cmdNew 
            Height          =   495
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
            Caption         =   "&New"
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
            mIcon           =   "frmUsers.frx":031A
         End
         Begin OnlineBanking.lvButtons_H cmdClose 
            Height          =   495
            Left            =   4200
            TabIndex        =   7
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
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
            mIcon           =   "frmUsers.frx":0634
         End
         Begin OnlineBanking.lvButtons_H cmdDelete 
            Height          =   495
            Left            =   2880
            TabIndex        =   8
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
            Caption         =   "&Delete"
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
            mIcon           =   "frmUsers.frx":094E
         End
      End
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lstUserID As String
Public lstUserName As String
Public lstUserPassword As String
Public lstUserRole As String

Private Sub cmddelete_Click()
On Error GoTo errHandler

    If MsgBox("Are you sure you want to delete user " & list.SelectedItem.ListSubItems(1) & " ?", vbQuestion + 4, title) = vbNo Then
        GoTo EXITPROCEDURE
        Else
            con.Execute "delete from users where employeeid = '" & list.SelectedItem.Text & "'"
            list.ListItems.Remove list.SelectedItem.Index
            MsgBox "User deleted successfully!", vbExclamation, title
    End If

EXITPROCEDURE:
    Exit Sub
errHandler:
    MsgBox "Sorry, user's data could not deleted", vbExclamation, title
    GoTo EXITPROCEDURE
End Sub

Private Sub cmdNew_Click()

    blUpdateUser = False
    blAddUser = True
    newUserID = autogen
    frmAddUser.Show
    Unload Me
    
End Sub

Private Sub cmdClose_Click()
    If MsgBox("Are you sure you want to close this window ?", 4 + vbQuestion, title) = vbNo Then
        GoTo EXITPROCEDURE
        Else
            Unload Me
    End If
    
EXITPROCEDURE:
    Exit Sub
End Sub

Private Sub cmdUpdate_Click()
    
    blUpdateUser = True
    blAddUser = False
    frmAddUser.txtLoginName.Text = lstUserName
    frmAddUser.txtPass.Text = lstUserPassword
    frmAddUser.txtPassword.Text = lstUserPassword
    frmAddUser.comboRole.Text = lstUserRole
    frmAddUser.Show
    Unload Me
    
End Sub

Private Sub Form_Load()

    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 3
    Call ConnectMe
    
    recUsers.MoveFirst
    Do While Not recUsers.EOF
        Set lst = list.ListItems.Add(, , recUsers!EmployeeID)
        lst.ListSubItems.Add , , recUsers!LoginID
        lst.ListSubItems.Add , , recUsers!Role
        lst.ListSubItems.Add , , recUsers!Password
        recUsers.MoveNext
    Loop
    recUsers.MoveFirst
    
End Sub

Public Function autogen() As String
    Dim rec As New Recordset
    
    rec.Open "select max(employeeid) from users", con, adOpenDynamic, adLockOptimistic
    
    If rec.EOF Then
        autogen = "E0001"
        Else
        autogen = "E" & Format(Right(Trim(rec(0)), 4) + 1, "0000")
    End If
    
End Function

Private Sub list_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lstUserID = list.SelectedItem.Text
    lstUserName = list.SelectedItem.ListSubItems(1)
    lstUserRole = list.SelectedItem.ListSubItems(2)
    lstUserPassword = list.SelectedItem.ListSubItems(3)
End Sub
