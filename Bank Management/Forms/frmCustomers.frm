VERSION 5.00
Begin VB.Form frmCustomers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CUSTOMERS"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   13830
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13815
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4815
         Left            =   11400
         TabIndex        =   47
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
            TabIndex        =   48
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
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   3240
         TabIndex        =   5
         Top             =   5160
         Width           =   10455
         Begin OnlineBanking.lvButtons_H cmdClose 
            Height          =   495
            Left            =   8880
            TabIndex        =   45
            Top             =   240
            Width           =   1335
            _extentx        =   2355
            _extenty        =   873
            caption         =   "&Close"
            capalign        =   2
            backstyle       =   2
            cgradient       =   0
            font            =   "frmCustomers.frx":0000
            mode            =   0
            value           =   0   'False
            cfhover         =   16777215
            cback           =   -2147483633
            cbhover         =   16711680
            lockhover       =   3
            micon           =   "frmCustomers.frx":002C
            mpointer        =   99
         End
         Begin OnlineBanking.lvButtons_H cmdDelete 
            Height          =   495
            Left            =   4560
            TabIndex        =   44
            Top             =   240
            Width           =   1335
            _extentx        =   2355
            _extenty        =   873
            caption         =   "&Delete"
            capalign        =   2
            backstyle       =   2
            cgradient       =   0
            font            =   "frmCustomers.frx":0346
            mode            =   0
            value           =   0   'False
            cfhover         =   16777215
            cback           =   -2147483633
            cbhover         =   16711680
            lockhover       =   3
            micon           =   "frmCustomers.frx":0372
            mpointer        =   99
         End
         Begin OnlineBanking.lvButtons_H cmdUpdate 
            Height          =   495
            Left            =   2520
            TabIndex        =   43
            Top             =   240
            Width           =   1335
            _extentx        =   2355
            _extenty        =   873
            caption         =   "&Update"
            capalign        =   2
            backstyle       =   2
            cgradient       =   0
            font            =   "frmCustomers.frx":068C
            mode            =   0
            value           =   0   'False
            cfhover         =   16777215
            cback           =   -2147483633
            cbhover         =   16711680
            lockhover       =   3
            micon           =   "frmCustomers.frx":06B8
            mpointer        =   99
         End
         Begin OnlineBanking.lvButtons_H cmdNew 
            Height          =   495
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   1335
            _extentx        =   2355
            _extenty        =   873
            caption         =   "&New"
            capalign        =   2
            backstyle       =   2
            cgradient       =   0
            font            =   "frmCustomers.frx":09D2
            mode            =   0
            value           =   0   'False
            cfhover         =   16777215
            cback           =   -2147483633
            cbhover         =   16711680
            lockhover       =   3
            micon           =   "frmCustomers.frx":09FE
            mpointer        =   99
         End
         Begin OnlineBanking.lvButtons_H cmdSearch 
            Height          =   495
            Left            =   6720
            TabIndex        =   46
            Top             =   240
            Width           =   1335
            _extentx        =   2355
            _extenty        =   873
            caption         =   "&Search"
            capalign        =   2
            backstyle       =   2
            cgradient       =   0
            font            =   "frmCustomers.frx":0D18
            mode            =   0
            value           =   0   'False
            cfhover         =   16777215
            cback           =   -2147483633
            cbhover         =   16711680
            lockhover       =   3
            micon           =   "frmCustomers.frx":0D44
            mpointer        =   99
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Account Details"
         Enabled         =   0   'False
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
         Height          =   4815
         Left            =   7320
         TabIndex        =   4
         Top             =   240
         Width           =   3975
         Begin VB.TextBox txtDateOpened 
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
            Left            =   1200
            TabIndex        =   39
            Top             =   3240
            Width           =   2655
         End
         Begin VB.TextBox txtAccessCode 
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
            Left            =   1200
            TabIndex        =   38
            Top             =   2760
            Width           =   2655
         End
         Begin VB.TextBox txtBalance 
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
            Left            =   1200
            TabIndex        =   37
            Top             =   2280
            Width           =   2655
         End
         Begin VB.TextBox txtExpiryDate 
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
            Left            =   1200
            TabIndex        =   36
            Top             =   1800
            Width           =   2655
         End
         Begin VB.TextBox txtATM 
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
            Left            =   1200
            TabIndex        =   35
            Top             =   1320
            Width           =   2655
         End
         Begin VB.TextBox txtAccountType 
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
            Left            =   1200
            TabIndex        =   34
            Top             =   840
            Width           =   2655
         End
         Begin VB.TextBox txtAccountNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   33
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Opened"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   3240
            Width           =   1815
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Access Code"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   2760
            Width           =   1815
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Balance"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Expiry Date"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "ATM Card No"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Account Type"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Account No"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Personal Details"
         Enabled         =   0   'False
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
         Height          =   4815
         Left            =   3240
         TabIndex        =   2
         Top             =   240
         Width           =   3975
         Begin VB.TextBox txtGender 
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
            Left            =   1200
            TabIndex        =   41
            Top             =   2760
            Width           =   2655
         End
         Begin VB.TextBox txtDateOfBirth 
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
            Left            =   1200
            TabIndex        =   40
            Top             =   2280
            Width           =   2655
         End
         Begin VB.TextBox txtHomePhone 
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
            Left            =   1200
            TabIndex        =   25
            Top             =   4200
            Width           =   2655
         End
         Begin VB.TextBox txtOfficePhone 
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
            Left            =   1200
            TabIndex        =   24
            Top             =   3720
            Width           =   2655
         End
         Begin VB.TextBox txtEmail 
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
            Left            =   1200
            TabIndex        =   23
            Top             =   3240
            Width           =   2655
         End
         Begin VB.TextBox txtHomeAddress 
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
            Left            =   1200
            TabIndex        =   22
            Top             =   1800
            Width           =   2655
         End
         Begin VB.TextBox txtOfficeAddress 
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
            Left            =   1200
            TabIndex        =   21
            Top             =   1320
            Width           =   2655
         End
         Begin VB.TextBox txtLastName 
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
            Left            =   1200
            TabIndex        =   20
            Top             =   840
            Width           =   2655
         End
         Begin VB.TextBox txtFirstName 
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
            Left            =   1200
            TabIndex        =   19
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Home Phone"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   4200
            Width           =   1815
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Office Phone"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   3720
            Width           =   1815
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "E-mail Address"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   3240
            Width           =   1815
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Gender"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   2760
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Of Birth"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Home Address"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Office Address"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Last Name"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "First Name"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5775
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3015
         Begin OnlineBanking.lvButtons_H cmdLast 
            Height          =   495
            Left            =   2280
            TabIndex        =   9
            Top             =   5160
            Width           =   615
            _extentx        =   1085
            _extenty        =   873
            caption         =   ">>"
            capalign        =   2
            backstyle       =   2
            cgradient       =   0
            font            =   "frmCustomers.frx":105E
            mode            =   0
            value           =   0   'False
            cfhover         =   16777215
            cback           =   -2147483633
            cbhover         =   16711680
            lockhover       =   3
         End
         Begin OnlineBanking.lvButtons_H cmdNext 
            Height          =   495
            Left            =   1560
            TabIndex        =   8
            Top             =   5160
            Width           =   615
            _extentx        =   1085
            _extenty        =   873
            caption         =   ">"
            capalign        =   2
            backstyle       =   2
            cgradient       =   0
            font            =   "frmCustomers.frx":108A
            mode            =   0
            value           =   0   'False
            cfhover         =   16777215
            cback           =   -2147483633
            cbhover         =   16711680
            lockhover       =   3
         End
         Begin OnlineBanking.lvButtons_H cmdPrevious 
            Height          =   495
            Left            =   840
            TabIndex        =   7
            Top             =   5160
            Width           =   615
            _extentx        =   1085
            _extenty        =   873
            caption         =   "<"
            capalign        =   2
            backstyle       =   2
            cgradient       =   0
            font            =   "frmCustomers.frx":10B6
            mode            =   0
            value           =   0   'False
            cfhover         =   16777215
            cback           =   -2147483633
            cbhover         =   16711680
            lockhover       =   3
         End
         Begin OnlineBanking.lvButtons_H cmdFirst 
            Height          =   495
            Left            =   120
            TabIndex        =   6
            Top             =   5160
            Width           =   615
            _extentx        =   1085
            _extenty        =   873
            caption         =   "<<"
            capalign        =   2
            backstyle       =   2
            cgradient       =   0
            font            =   "frmCustomers.frx":10E2
            mode            =   0
            value           =   0   'False
            cfhover         =   16777215
            cback           =   -2147483633
            cbhover         =   16711680
            lockhover       =   3
         End
         Begin VB.ListBox lst 
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
            Height          =   4515
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   2775
         End
      End
   End
End
Attribute VB_Name = "frmCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim position As Integer
Public AccountNumber As String
Dim strPictureName As String
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmddelete_Click()
    If MsgBox("Are you sure you want to delete " & Trim(txtFirstName.Text) & " " & Trim(txtLastName) & " ?", 4 + 32, title) = vbNo Then
        GoTo EXITPROCEDURE
        Else
            con.Execute "Delete from customer where FirstName = '" & Trim(txtFirstName.Text) & "' and lastname = '" & Trim(txtLastName.Text) & "'"
            Call EmptyTextFields
            Call fillList
    End If
    
EXITPROCEDURE:
    Exit Sub
End Sub

Private Sub cmdFirst_Click()
    recCustomers.MoveFirst
    Call display
End Sub

Private Sub cmdLast_Click()
    recCustomers.MoveLast
    Call display
End Sub

Private Sub cmdNew_Click()
    blAddCustomer = True
    blUpdateCustomer = False
    frmAddCustomer.Show
    AccountNumber = autogen
    frmAddCustomer.txtAccountNo.Text = AccountNumber
    Unload frmCustomers
End Sub

Private Sub cmdNext_Click()
    recCustomers.MoveNext
    If recCustomers.EOF Then
        recCustomers.MoveLast
    End If
    Call display
End Sub

Private Sub cmdPrevious_Click()
    recCustomers.MovePrevious
    If recCustomers.BOF Then
        recCustomers.MoveFirst
    End If
    Call display
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo errhandler

    blAddCustomer = False
    blUpdateCustomer = True
    
        frmAddCustomer.txtFirstName.Text = txtFirstName.Text
        frmAddCustomer.txtLastName.Text = txtLastName.Text
        frmAddCustomer.txtOfficeAddress.Text = txtOfficeAddress.Text
        frmAddCustomer.txtHomeAddress.Text = txtHomeAddress.Text
        frmAddCustomer.dtpDateOfBirth.Value = txtDateOfBirth.Text
        frmAddCustomer.cboGender.Text = txtGender.Text
        frmAddCustomer.txtEmail.Text = txtEmail.Text
        frmAddCustomer.txtOfficePhone.Text = txtOfficePhone.Text
        frmAddCustomer.txtHomePhone.Text = txtHomePhone.Text
        frmAddCustomer.txtAccountNo.Text = txtAccountNo.Text
        frmAddCustomer.cboAccountType.Text = txtAccountType.Text
        frmAddCustomer.txtATM.Text = txtATM.Text
        
        If txtExpiryDate.Text = "" Then
            frmAddCustomer.dtpExpiryDate.Value = Date
            Else
                frmAddCustomer.dtpExpiryDate.Value = txtExpiryDate.Text
        End If
        
        frmAddCustomer.txtBalance.Text = txtBalance.Text
        frmAddCustomer.txtAccessCode.Text = txtAccessCode.Text
        If txtDateOpened.Text = "" Then
            frmAddCustomer.dtpOpenDate.Value = Date
            Else
                frmAddCustomer.dtpOpenDate.Value = txtDateOpened.Text
        End If
        frmAddCustomer.Pic.Picture = LoadPicture(strPictureName)
        frmAddCustomer.strPath = strPictureName

        Unload Me
    
EXITPROCEDURE:
    Unload Me
    Exit Sub

errhandler:
    Pic.Picture = LoadPicture(App.Path & "\Pictures\" & "NA.GIF")
    GoTo EXITPROCEDURE
End Sub

Private Sub Form_Load()
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 3
    Call ConnectMe

    
    Call fillList
    
End Sub

Private Sub fillList()
    lst.Clear
    recCustomers.Requery
    If Not recCustomers.EOF Then
    
        recCustomers.MoveFirst
        Do While Not recCustomers.EOF
            lst.AddItem recCustomers!FirstName & "  " & recCustomers!LastName
            recCustomers.MoveNext
        Loop
        
        recCustomers.MoveFirst
        Call display
        
        Else
            cmdFirst.Enabled = False
            cmdNext.Enabled = False
            cmdPrevious.Enabled = False
            cmdLast.Enabled = False
            cmdUpdate.Enabled = False
            cmdPreview.Enabled = False
            cmdDelete.Enabled = False
    End If

End Sub

Public Sub display()
On Error GoTo errhandler

    txtFirstName.Text = recCustomers!FirstName & ""
    txtLastName.Text = recCustomers!LastName & ""
    txtOfficeAddress.Text = recCustomers!OfficeAddress & ""
    txtHomeAddress.Text = recCustomers!HomeAddress & ""
    txtDateOfBirth.Text = recCustomers!DateOfBirth & ""
    txtGender.Text = recCustomers!Gender & ""
    txtEmail.Text = recCustomers!Email & ""
    txtOfficePhone.Text = recCustomers!OfficePhone & ""
    txtHomePhone.Text = recCustomers!HomePhone & ""
    txtAccountNo.Text = recCustomers!AccountNumber & ""
    txtAccountType.Text = recCustomers!AccountType & ""
    txtATM.Text = recCustomers!ATMCardNumber & ""
    txtExpiryDate.Text = recCustomers!ExpiryDate & ""
    txtBalance.Text = Format(recCustomers!Balance, "currency") & ""
    txtAccessCode.Text = recCustomers!AccessCode & ""
    txtDateOpened.Text = recCustomers!DateOpened & ""
    
    strPictureName = recCustomers!CustomerPicture & ""
    
    If strPictureName = "" Then
        strPictureName = App.Path & "\Pictures\NA.GIF"
        Else
        
    End If
    
    Pic.Picture = LoadPicture(App.Path & "\Pictures\" & strPictureName)
    
EXITPROCEDURE:
    Exit Sub

errhandler:
    Pic.Picture = LoadPicture(App.Path & "\Pictures\" & "NA.GIF")
    GoTo EXITPROCEDURE
End Sub

Private Sub lst_Click()
    position = lst.ListIndex
    recCustomers.MoveFirst
    recCustomers.Move position
    Call display
End Sub

Private Sub EmptyTextFields()

    Dim ctl As Object
    
    For Each ctl In Me
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        End If
    Next
End Sub

Public Function autogen() As String
    Dim rec As New Recordset
    
    rec.Open "select max(accountNumber) from Customer", con, adOpenDynamic, adLockOptimistic
    
    If rec.EOF Then
        autogen = "C0001"
        Else
        autogen = "C" & Format(Right(Trim(rec(0)), 4) + 1, "0000")
    End If
    
End Function

Private Sub txtBalance_KeyPress(KeyAscii As Integer)
    Call onlyNumbers(KeyAscii)
End Sub
