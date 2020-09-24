VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTransfer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSFER"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   8880
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   6255
         Left            =   6480
         TabIndex        =   24
         Top             =   240
         Width           =   2295
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   2680
            Left            =   120
            ScaleHeight     =   2655
            ScaleWidth      =   2070
            TabIndex        =   25
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
         TabIndex        =   12
         Top             =   240
         Width           =   4095
         Begin VB.Frame Frame7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   2535
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   3855
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
               Left            =   1320
               MaxLength       =   10
               TabIndex        =   18
               Top             =   1080
               Width           =   2535
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
               Left            =   1320
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   17
               Top             =   120
               Width           =   2535
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
               Left            =   1320
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   16
               Top             =   2040
               Width           =   2535
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
               Left            =   1320
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   15
               Top             =   600
               Width           =   2535
            End
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
               Left            =   1320
               MaxLength       =   10
               TabIndex        =   14
               Top             =   1560
               Width           =   2535
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Gender"
               Height          =   255
               Left            =   0
               TabIndex        =   23
               Top             =   1080
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
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Last Name"
               Height          =   255
               Left            =   0
               TabIndex        =   21
               Top             =   600
               Width           =   1815
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Balance"
               Height          =   255
               Left            =   0
               TabIndex        =   20
               Top             =   2040
               Width           =   1095
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Account Type"
               Height          =   255
               Left            =   0
               TabIndex        =   19
               Top             =   1560
               Width           =   1215
            End
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Transaction Details"
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
         Height          =   3255
         Left            =   2280
         TabIndex        =   3
         Top             =   3240
         Width           =   4095
         Begin VB.ComboBox cboToAccountNo 
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
            ItemData        =   "frmTransfer.frx":0000
            Left            =   1440
            List            =   "frmTransfer.frx":000D
            TabIndex        =   29
            Top             =   2280
            Width           =   2535
         End
         Begin VB.TextBox txtAmount 
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
            Left            =   1440
            MaxLength       =   14
            TabIndex        =   6
            Top             =   1800
            Width           =   2535
         End
         Begin VB.TextBox txtTransactionMode 
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
            Left            =   1440
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   5
            Text            =   "Transfer"
            Top             =   360
            Width           =   2535
         End
         Begin VB.TextBox txtCheckNo 
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
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   4
            Top             =   1320
            Width           =   2535
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   375
            Left            =   1440
            TabIndex        =   7
            Top             =   840
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   91422721
            CurrentDate     =   38954
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "To Account No"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Transaction Mode"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Transaction Date"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Check Number"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   1320
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "From Account No"
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
         Height          =   6255
         Left            =   120
         TabIndex        =   1
         Top             =   240
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
            Height          =   5640
            Left            =   120
            Style           =   1  'Simple Combo
            TabIndex        =   2
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         TabIndex        =   26
         Top             =   6480
         Width           =   8655
         Begin OnlineBanking.lvButtons_H cmdClose 
            Height          =   495
            Left            =   6720
            TabIndex        =   27
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            Caption         =   "&Cancel"
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
            mIcon           =   "frmTransfer.frx":001D
         End
         Begin OnlineBanking.lvButtons_H cmdOk 
            Default         =   -1  'True
            Height          =   495
            Left            =   600
            TabIndex        =   28
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            Caption         =   "&Transfer"
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
            mIcon           =   "frmTransfer.frx":0337
         End
      End
   End
End
Attribute VB_Name = "frmTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboAccountNumber_Click()
On Error GoTo errhandler
    Call DisplayCustomerDetails(cboAccountNumber.Text)
    txtBalance.Text = Balance
    txtFirstName.Text = FirstName
    txtLastName.Text = LastName
    txtGender.Text = Gender
    txtAccountType.Text = AccountType
    
    cmdOk.Enabled = True
    txtAmount.Enabled = True
    txtCheckNo.Enabled = True
    txtCheckNo.SetFocus
    
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

Private Sub cboToAccountNo_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
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

Private Sub cmdOk_Click()
On Error GoTo errhandler

    If txtCheckNo.Text = "" Then
        MsgBox "Please enter check number!", vbExclamation, title
        txtCheckNo.SetFocus
        GoTo EXITPROCEDURE
    End If

    If txtAmount.Text = "" Then
        MsgBox "Please enter Amount!", vbExclamation, title
        txtAmount.SetFocus
        GoTo EXITPROCEDURE
    End If
    
    If cboToAccountNo.Text = "" Then
        MsgBox "Please select the account you are transfering to!", vbExclamation, title
        cboToAccountNo.SetFocus
        GoTo EXITPROCEDURE
    End If
    
    If Check_CheckNo(txtCheckNo.Text) = True Then
        MsgBox "Check Number has already been used", vbExclamation, title
        txtCheckNo.SelStart = 0
        txtCheckNo.SelLength = Len(txtCheckNo.Text)
        txtCheckNo.SetFocus
        GoTo EXITPROCEDURE
    End If

    If Check_StopCheckNo(txtCheckNo.Text, cboAccountNumber.Text) = True Then
        MsgBox "Check Number has been stopped. You can not transfer money with it", vbExclamation, title
        txtCheckNo.SelStart = 0
        txtCheckNo.SelLength = Len(txtCheckNo.Text)
        txtCheckNo.SetFocus
        GoTo EXITPROCEDURE
    End If

    If Trim(txtAmount.Text) > Trim(txtBalance.Text) Then
        MsgBox "The Amount you are withdrawing is more than your current balance!", vbExclamation, title
        txtAmount.SelStart = 0
        txtAmount.SelLength = Len(txtAmount.Text)
        txtAmount.SetFocus
        GoTo EXITPROCEDURE
    End If
    
    If cboAccountNumber.Text = cboToAccountNo.Text Then
        MsgBox "You can transfer from and to the same Account.", vbExclamation, title
        cboAccountNumber.SetFocus
        GoTo EXITPROCEDURE
    End If
    
    If MsgBox("Are you sure you want to transfer " & txtAmount.Text & " from Account Number " & cboAccountNumber.Text & " to Account Number " & cboToAccountNo.Text & " ?", 4 + 32, title) = vbNo Then
        GoTo EXITPROCEDURE
        Else
            recTrans.AddNew
            
                recTrans!TransactionID = autogen
                recTrans!TransactionDate = dtpDate.Value
                recTrans!TransactionType = "Transfer"
                recTrans!TransactionMode = "Check"
                recTrans!FromAccountNumber = cboAccountNumber.Text
                recTrans!ToAccountNumber = cboToAccountNo.Text
                recTrans!checkNumber = txtCheckNo.Text
                recTrans!Amount = txtAmount.Text
                recTrans!status = ""
                
            recTrans.Update
            
            con.Execute "Update Customer set Balance = " & CCur(txtBalance.Text) - CCur(txtAmount.Text) & " Where AccountNumber = '" & Trim(cboAccountNumber.Text) & "'"
            con.Execute "Update Customer set Balance = " & CCur(txtBalance.Text) + CCur(txtAmount.Text) & " Where AccountNumber = '" & Trim(cboToAccountNo.Text) & "'"
            
            MsgBox "Transaction done successfully.", vbExclamation, title
            
            Unload Me
    End If
    
EXITPROCEDURE:
    Exit Sub
    
errhandler:
    MsgBox "An Error occurred. Transaction was unsuccessfull, Try again", vbCritical, title
    GoTo EXITPROCEDURE
End Sub

Private Sub Form_Load()
On Error GoTo errhandler

    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 3
    Call ConnectMe
    cmdOk.Enabled = False
    txtAmount.Enabled = False
    txtCheckNo.Enabled = False
    dtpDate.Value = Date
    
    Call fill_Combo(cboAccountNumber)
    Call fill_Combo(cboToAccountNo)

EXITPROCEDURE:
    Exit Sub
    
errhandler:
    MsgBox "An Error occurred while loading the form", vbCritical, title
    GoTo EXITPROCEDURE
End Sub



Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    Call onlyNumbers(KeyAscii)
End Sub


Public Function autogen() As String
On Error GoTo errhandler

    Dim rec As New Recordset
    
    rec.Open "select max(TransactionID) from Trans", con, adOpenDynamic, adLockOptimistic
    
    If rec.EOF Then
        autogen = 1
        Else
        autogen = Val(rec(0) + 1)
    End If
    
EXITPROCEDURE:
    Exit Function
    
errhandler:
    'MsgBox "An Error occurred. Transaction was unsuccessfull, Try again", vbCritical, title
    GoTo EXITPROCEDURE
End Function


