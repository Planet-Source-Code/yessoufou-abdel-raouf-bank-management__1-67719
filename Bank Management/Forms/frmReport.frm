VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REPORT"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   8895
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
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
         Top             =   840
         Width           =   4095
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
            Width           =   3855
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
               TabIndex        =   20
               Top             =   1560
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
               TabIndex        =   19
               Top             =   600
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
               TabIndex        =   18
               Top             =   2040
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
               TabIndex        =   16
               Top             =   1080
               Width           =   2535
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
            Begin VB.Label Label3 
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
         Height          =   4575
         Left            =   6480
         TabIndex        =   12
         Top             =   840
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
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4575
         Left            =   120
         TabIndex        =   7
         Top             =   840
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
            Height          =   4080
            Left            =   120
            Style           =   1  'Simple Combo
            TabIndex        =   8
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Report Details"
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
         Height          =   1575
         Left            =   2280
         TabIndex        =   4
         Top             =   3840
         Width           =   4095
         Begin MSComCtl2.DTPicker dtpFromDate 
            Height          =   375
            Left            =   1440
            TabIndex        =   5
            Top             =   360
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
            Format          =   19595265
            CurrentDate     =   38954
         End
         Begin MSComCtl2.DTPicker dtpToDate 
            Height          =   375
            Left            =   1440
            TabIndex        =   9
            Top             =   960
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
            Format          =   19595265
            CurrentDate     =   38954
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "From Date"
            Height          =   255
            Left            =   120
            TabIndex        =   6
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
         TabIndex        =   1
         Top             =   5520
         Width           =   8655
         Begin OnlineBanking.lvButtons_H cmdClose 
            Height          =   495
            Left            =   6480
            TabIndex        =   2
            Top             =   240
            Width           =   1575
            _extentx        =   2778
            _extenty        =   873
            caption         =   "&Cancel"
            capalign        =   2
            backstyle       =   2
            cgradient       =   0
            font            =   "frmReport.frx":0000
            mode            =   0
            value           =   0   'False
            cfhover         =   16777215
            cback           =   -2147483633
            cbhover         =   16711680
            lockhover       =   3
            micon           =   "frmReport.frx":002C
            mpointer        =   99
         End
         Begin OnlineBanking.lvButtons_H cmdOk 
            Default         =   -1  'True
            Height          =   495
            Left            =   600
            TabIndex        =   3
            Top             =   240
            Width           =   1575
            _extentx        =   2778
            _extenty        =   873
            caption         =   "&Preview Report"
            capalign        =   2
            backstyle       =   2
            cgradient       =   0
            font            =   "frmReport.frx":0346
            mode            =   0
            value           =   0   'False
            cfhover         =   16777215
            cback           =   -2147483633
            cbhover         =   16711680
            lockhover       =   3
            micon           =   "frmReport.frx":0372
            mpointer        =   99
         End
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   390
         Left            =   3435
         TabIndex        =   11
         Top             =   240
         Width           =   105
      End
   End
End
Attribute VB_Name = "frmReport"
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
    
    Pic.Picture = LoadPicture(App.Path & "\pictures\" & CustomerPicture)
    
EXITPROCEDURE:
    Exit Sub

errhandler:
    Pic.Picture = LoadPicture(App.Path & "\Pictures\" & "NA.GIF")
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

Private Sub cmdOk_Click()
    If dtpFromDate.Value > dtpToDate.Value Then
        MsgBox "The Second Date Should Not Be More Than The First One", vbExclamation, title
        dtpFromDate.SetFocus
        GoTo EXITPROCEDURE
    End If
    
    If blQuery = True And blDeposit = False And blTransfer = False And blWithdrawal = False And blStop = False Then
        Set QueryReport.DataSource = con.Execute("Select * from CustomerQuery where accountnumber = '" & cboAccountNumber.Text & "' and DateReceived between #" & dtpFromDate.Value & "# and #" & dtpToDate.Value & "#")
        QueryReport.Show
        Set QueryReport = Nothing
    End If
    
    If blQuery = False And blDeposit = True And blTransfer = False And blWithdrawal = False And blStop = False Then
        Set DepositReport.DataSource = con.Execute("Select * from Trans where Toaccountnumber = '" & cboAccountNumber.Text & "' and transactionType = 'Deposit' and transactionDate between #" & dtpFromDate.Value & "# and #" & dtpToDate.Value & "#")
        DepositReport.Show
        Set DepositReport = Nothing
    End If
    
    If blQuery = False And blDeposit = False And blTransfer = True And blWithdrawal = False And blStop = False Then
        Set TransferReport.DataSource = con.Execute("Select * from Trans where FromAccountnumber = '" & cboAccountNumber.Text & "' and transactionType = 'Transfer' and transactionDate between #" & dtpFromDate.Value & "# and #" & dtpToDate.Value & "#")
        TransferReport.Show
        Set TransferReport = Nothing
    End If
    
    If blQuery = False And blDeposit = False And blTransfer = False And blWithdrawal = True And blStop = False Then
        Set WithdrawalReport.DataSource = con.Execute("Select * from Trans where Fromaccountnumber = '" & cboAccountNumber.Text & "' and transactionType = 'Withdrawal' and transactionDate between #" & dtpFromDate.Value & "# and #" & dtpToDate.Value & "#")
        WithdrawalReport.Show
        Set WithdrawalReport = Nothing
    End If

    If blQuery = False And blDeposit = False And blTransfer = False And blWithdrawal = False And blStop = True Then
        Set StopPaymentReport.DataSource = con.Execute("Select * from StopPayment where accountnumber = '" & cboAccountNumber.Text & "'")
        StopPaymentReport.Show
        Set StopPaymentReport = Nothing
    End If

EXITPROCEDURE:
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo errhandler

    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 3
    Call ConnectMe
    
    dtpFromDate.Value = Date
    dtpToDate.Value = Date
    cmdOk.Enabled = False
    
    
    Call fill_Combo(cboAccountNumber)

EXITPROCEDURE:
    Exit Sub
    
errhandler:
    MsgBox "An Error occurred while loading the form", vbCritical, title
    GoTo EXITPROCEDURE
End Sub

Private Sub Form_Unload(Cancel As Integer)
    blQuery = False
    blDeposit = False
    blTransfer = False
    blWithdrawal = False
    blStop = False
End Sub
