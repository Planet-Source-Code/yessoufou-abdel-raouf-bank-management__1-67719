VERSION 5.00
Begin VB.Form frmStopPayment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STOP PAYMENT"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   9150
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Customer Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   2280
         TabIndex        =   14
         Top             =   120
         Width           =   4335
         Begin VB.Frame Frame7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   2535
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   3975
            Begin VB.TextBox txtAccountType 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               Left            =   1200
               MaxLength       =   10
               TabIndex        =   20
               Top             =   1560
               Width           =   2775
            End
            Begin VB.TextBox txtLastName 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   19
               Top             =   600
               Width           =   2775
            End
            Begin VB.TextBox txtBalance 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   18
               Top             =   2040
               Width           =   2775
            End
            Begin VB.TextBox txtFirstName 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   17
               Top             =   120
               Width           =   2775
            End
            Begin VB.TextBox txtGender 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               Left            =   1200
               MaxLength       =   10
               TabIndex        =   16
               Top             =   1080
               Width           =   2775
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Account Type"
               Height          =   255
               Left            =   0
               TabIndex        =   25
               Top             =   1560
               Width           =   1215
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Balance"
               Height          =   255
               Left            =   0
               TabIndex        =   24
               Top             =   2040
               Width           =   1095
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Last Name"
               Height          =   255
               Left            =   0
               TabIndex        =   23
               Top             =   600
               Width           =   1815
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "First Name"
               Height          =   255
               Left            =   0
               TabIndex        =   22
               Top             =   120
               Width           =   1455
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Gender"
               Height          =   255
               Left            =   0
               TabIndex        =   21
               Top             =   1080
               Width           =   1815
            End
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Customer Picture"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   5415
         Left            =   6720
         TabIndex        =   12
         Top             =   120
         Width           =   2295
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   2680
            Left            =   120
            ScaleHeight     =   2655
            ScaleWidth      =   2070
            TabIndex        =   13
            Top             =   360
            Width           =   2100
            Begin VB.Image Pic 
               Height          =   2655
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   2070
            End
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Account Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   5415
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   2055
         Begin VB.ComboBox cboAccountNumber 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4860
            Left            =   120
            Style           =   1  'Simple Combo
            TabIndex        =   8
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Instruction Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   2280
         TabIndex        =   1
         Top             =   3120
         Width           =   4335
         Begin VB.TextBox txtInstructions 
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
            Height          =   1395
            IMEMode         =   3  'DISABLE
            Left            =   1320
            MultiLine       =   -1  'True
            PasswordChar    =   "*"
            TabIndex        =   4
            ToolTipText     =   "Retype The User Password"
            Top             =   840
            Width           =   2895
         End
         Begin VB.TextBox txtCheckNo 
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
            Left            =   1320
            TabIndex        =   3
            ToolTipText     =   "Password"
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Instructions"
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Check Number"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         TabIndex        =   2
         Top             =   5520
         Width           =   8895
         Begin OnlineBanking.lvButtons_H cmdClose 
            Height          =   495
            Left            =   7320
            TabIndex        =   9
            Top             =   240
            Width           =   1335
            _extentx        =   2355
            _extenty        =   873
            caption         =   "&Cancel"
            capalign        =   2
            backstyle       =   2
            cgradient       =   0
            font            =   "frmStopPayment.frx":0000
            mode            =   0
            value           =   0   'False
            cfhover         =   16777215
            cback           =   -2147483633
            cbhover         =   16711680
            lockhover       =   3
            micon           =   "frmStopPayment.frx":002C
            mpointer        =   99
         End
         Begin OnlineBanking.lvButtons_H cmdOk 
            Height          =   495
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1335
            _extentx        =   2355
            _extenty        =   873
            caption         =   "&Stop Payment"
            capalign        =   2
            backstyle       =   2
            cgradient       =   0
            font            =   "frmStopPayment.frx":0346
            mode            =   0
            value           =   0   'False
            cfhover         =   16777215
            cback           =   -2147483633
            cbhover         =   16711680
            lockhover       =   3
            micon           =   "frmStopPayment.frx":0372
            mpointer        =   99
         End
         Begin OnlineBanking.lvButtons_H cmdActivate 
            Height          =   495
            Left            =   3600
            TabIndex        =   11
            Top             =   240
            Width           =   1335
            _extentx        =   2355
            _extenty        =   873
            caption         =   "&Activate"
            capalign        =   2
            backstyle       =   2
            cgradient       =   0
            font            =   "frmStopPayment.frx":068C
            mode            =   0
            value           =   0   'False
            cfhover         =   16777215
            cback           =   -2147483633
            cbhover         =   16711680
            lockhover       =   3
            micon           =   "frmStopPayment.frx":06B8
            mpointer        =   99
         End
      End
   End
End
Attribute VB_Name = "frmStopPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdActivate_Click()
On Error GoTo errhandler
    
     If txtCheckNo.Text = "" Then
        MsgBox "Please enter Check Number!", vbExclamation, title
        txtCheckNo.SetFocus
        GoTo EXITPROCEDURE
    End If
    
    
    If Check_StopCheckNo(txtCheckNo.Text, cboAccountNumber.Text) = False Then
        MsgBox "Check Number has already been activated", vbExclamation, title
        txtCheckNo.SelStart = 0
        txtCheckNo.SelLength = Len(txtCheckNo.Text)
        txtCheckNo.SetFocus
        GoTo EXITPROCEDURE
    End If
    
    If MsgBox("Are you sure you want to activate this check  ?", 4 + 32, title) = vbNo Then
        GoTo EXITPROCEDURE
        Else
        con.Execute "Delete from StopPayment WHERE AccountNumber = '" & cboAccountNumber.Text & "' AND CheckNumber = '" & txtCheckNo.Text & "'"

        MsgBox "The check has been Successfully Activated!", vbExclamation, title
        Unload Me
    End If
    
EXITPROCEDURE:
    Exit Sub

errhandler:
    MsgBox "An Error occurred while sending Customer Query ", vbCritical, title
    GoTo EXITPROCEDURE
End Sub

Private Sub cmdClose_Click()
    If MsgBox("Are you sure you want to cancel ?", 4 + 32, title) = vbNo Then
        GoTo EXITPROCEDURE
        Else
            Unload Me
    End If
    
EXITPROCEDURE:
    Exit Sub
End Sub
Private Sub cboAccountNumber_Click()
On Error GoTo errhandler
    Call DisplayCustomerDetails(cboAccountNumber.Text)
    txtBalance.Text = Balance
    txtFirstName.Text = FirstName
    txtLastName.Text = LastName
    txtGender.Text = Gender
    txtAccountType.Text = AccountType
    
    cmdOk.Enabled = True
    cmdActivate.Enabled = True
    txtCheckNo.Enabled = True
    txtInstructions.Enabled = True
    txtInstructions.SetFocus
    Pic.Picture = LoadPicture(App.Path & "\pictures\" & CustomerPicture)
    
EXITPROCEDURE:
    Exit Sub

errhandler:
    Pic.Picture = LoadPicture(App.Path & "\Pictures\" & "NA.GIF")
    GoTo EXITPROCEDURE
End Sub

Private Sub cboAccountNumber_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub cmdOk_Click()
On Error GoTo errhandler
    
     If txtCheckNo.Text = "" Then
        MsgBox "Please enter Check Number!", vbExclamation, title
        txtCheckNo.SetFocus
        GoTo EXITPROCEDURE
    End If

    
    If Check_StopCheckNo(txtCheckNo.Text, cboAccountNumber.Text) = True Then
        MsgBox "Check Number has already been Stopped", vbExclamation, title
        txtCheckNo.SelStart = 0
        txtCheckNo.SelLength = Len(txtCheckNo.Text)
        txtCheckNo.SetFocus
        GoTo EXITPROCEDURE
    End If
    
    If MsgBox("Are you sure you want to stop this check  ?", 4 + 32, title) = vbNo Then
        GoTo EXITPROCEDURE
        Else
        recStopPayment.AddNew
        
            recStopPayment!checkNumber = txtCheckNo.Text
            recStopPayment!AccountNumber = cboAccountNumber.Text
            recStopPayment!Instructions = Trim(txtInstructions.Text)
            
        recStopPayment.Update

        MsgBox "The Check has been registered Successfully!", vbExclamation, title
        Unload Me
    End If
    
EXITPROCEDURE:
    Exit Sub

errhandler:
    MsgBox "An Error occurred while sending Customer Query ", vbCritical, title
    GoTo EXITPROCEDURE
End Sub

Private Sub Form_Load()
On Error GoTo errhandler

    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 3
    Call ConnectMe
    cmdOk.Enabled = False
    cmdActivate.Enabled = False
    txtCheckNo.Enabled = False
    txtInstructions.Enabled = False
    
    Call fill_Combo(cboAccountNumber)

EXITPROCEDURE:
    Exit Sub
    
errhandler:
    MsgBox "An Error occurred while loading the form", vbCritical, title
    GoTo EXITPROCEDURE
End Sub
