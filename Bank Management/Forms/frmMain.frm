VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   3480
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1800
      Top             =   1320
   End
   Begin MSComDlg.CommonDialog ComDlg 
      Left            =   1440
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2715
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   12
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   884
            MinWidth        =   884
            Picture         =   "frmMain.frx":0000
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2822
            MinWidth        =   2822
            Key             =   "Role"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "frmMain.frx":629A
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1060
            MinWidth        =   1060
            Text            =   "Name"
            TextSave        =   "Name"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Key             =   "Name"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "frmMain.frx":BA8C
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1940
            MinWidth        =   1940
            Text            =   "Login Date"
            TextSave        =   "Login Date"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "26/08/2006"
            Key             =   "D"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "Time"
            TextSave        =   "Time"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "17:32"
            Key             =   "T"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLock 
         Caption         =   "Lock Apllication"
      End
      Begin VB.Menu mnuLogoff 
         Caption         =   "LogOff"
      End
      Begin VB.Menu mnuLogin 
         Caption         =   "Login"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuSpet 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "&Administrator"
      Begin VB.Menu mnuUser 
         Caption         =   "User"
         Begin VB.Menu mnuAllUsers 
            Caption         =   "All Users"
         End
      End
      Begin VB.Menu mnuCustomer 
         Caption         =   "Customer"
         Begin VB.Menu mnuAllCustomers 
            Caption         =   "All Customers"
         End
         Begin VB.Menu mnuCustomerQuery 
            Caption         =   "Customer Query"
         End
      End
   End
   Begin VB.Menu mnuTrans 
      Caption         =   "Transaction"
      Begin VB.Menu mnuDeposit 
         Caption         =   "Deposit"
      End
      Begin VB.Menu mnuWithdraw 
         Caption         =   "Withdrawal"
      End
      Begin VB.Menu mnuTransfer 
         Caption         =   "Transfer"
      End
      Begin VB.Menu mnuStopPayment 
         Caption         =   "Stop Payment"
      End
      Begin VB.Menu mnuCheckBook 
         Caption         =   "Issue Check Book"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Begin VB.Menu mnuCustomerQueryReport 
         Caption         =   "Customer Query Report"
      End
      Begin VB.Menu mnuDepositReport 
         Caption         =   "Deposit Report"
      End
      Begin VB.Menu mnuTransferReport 
         Caption         =   "Transfer Report"
      End
      Begin VB.Menu mnuWithdrawalReport 
         Caption         =   "Withdrawal Report"
      End
      Begin VB.Menu mnuStopPaymentReport 
         Caption         =   "Stop Payment Report"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Tools"
      Begin VB.Menu mnuCalculator 
         Caption         =   "Calculator"
      End
      Begin VB.Menu mnuCalendar 
         Caption         =   "Calendar"
      End
      Begin VB.Menu mnuNotepad 
         Caption         =   "Notepad"
      End
      Begin VB.Menu mnuDate 
         Caption         =   "Date and Time"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim currentLength As Byte
Const msg  As String = "Online Banking"
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If MsgBox("Are you sure you want to exit ?", 4 + 32, title) = vbNo Then
        GoTo EXITPROCEDURE
        Else
            End
    End If
    
EXITPROCEDURE:
    Exit Sub
    
End Sub



Private Sub mnuAllCustomers_Click()
    frmCustomers.Show
End Sub

Private Sub mnuAllUsers_Click()
    Load frmUsers
    frmUsers.Show
End Sub

Private Sub mnuCalculator_Click()
On Error GoTo errhandler

    Shell "calc"
    
    Exit Sub
    
errhandler:
    MsgBox "Calculator is not available for now...", vbExclamation, title
End Sub

Private Sub mnuCalendar_Click()
On Error GoTo errhandler

    Load frmCalendar
    frmCalendar.Show
    
    Exit Sub
    
errhandler:
    MsgBox "The calendar is not available for now...", vbExclamation, title
End Sub

Private Sub mnuCheckBook_Click()
    frmCheckBook.Show
End Sub

Private Sub mnuCustomerQuery_Click()
    frmCustomerQuery.Show
End Sub

Private Sub mnuCustomerQueryReport_Click()

    blQuery = True
    blDeposit = False
    blTransfer = False
    blWithdrawal = False
    blStop = False
    
    frmReport.lbl.Caption = "Customer Query"
    frmReport.Show
End Sub

Private Sub mnuDate_Click()
On Error GoTo errhandler
    Dim dblReturn As Double

    dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl", 5)

    Exit Sub
errhandler:
    MsgBox "The date and time is not available for now", vbExclamation, title

End Sub

Private Sub mnuDeposit_Click()
    frmDeposit.Show
End Sub

Private Sub mnuDepositReport_Click()

    blQuery = False
    blDeposit = True
    blTransfer = False
    blWithdrawal = False
    blStop = False
    frmReport.lbl.Caption = "Deposit Report"
    frmReport.Show
    
End Sub

Private Sub mnuExit_Click()

    If MsgBox("Are you sure you want to exit ?", 4 + 32, title) = vbNo Then
        GoTo EXITPROCEDURE
        Else
            End
    End If
    
EXITPROCEDURE:
    Exit Sub
    
End Sub

Private Sub mnuLock_Click()
    Load frmLock
    frmLock.Show
End Sub

Private Sub mnuLogin_Click()
    Load frmLogin
    frmLogin.Show
End Sub

Private Sub mnuLogoff_Click()
    If MsgBox("Are you sure you want to log off ?", 4 + 32, title) = vbNo Then
        GoTo EXITPROCEDURE
        Else
            frmLogin.Show
    End If
    
EXITPROCEDURE:
    Exit Sub
End Sub

Private Sub mnuNotepad_Click()
On Error GoTo errhandler
    Shell "notepad.exe", vbNormalFocus
    Exit Sub
errhandler:
    MsgBox "Notepad not available for now", vbExclamation, title
End Sub

Private Sub mnuStopPayment_Click()
    frmStopPayment.Show
End Sub

Private Sub mnuStopPaymentReport_Click()

    blQuery = False
    blDeposit = False
    blTransfer = False
    blWithdrawal = False
    blStop = True
    frmReport.Frame3.Visible = False

    frmReport.lbl.Caption = "Stop Payment Report"
    frmReport.Show
End Sub

Private Sub mnuTransfer_Click()
    frmTransfer.Show
End Sub

Private Sub mnuTransferReport_Click()

    blQuery = False
    blDeposit = False
    blTransfer = True
    blWithdrawal = False
    blStop = False
    
    frmReport.lbl.Caption = "TRansfer Report"
    frmReport.Show
End Sub

Private Sub mnuWithDraw_Click()
    frmWithdrawal.Show
End Sub

Private Sub mnuWithdrawalReport_Click()

    blQuery = False
    blDeposit = False
    blTransfer = False
    blWithdrawal = True
    blStop = False
    
    frmReport.lbl.Caption = "Withdrawal Report"
    frmReport.Show
End Sub

Private Sub Timer2_Timer()
    Caption = Left(msg, currentLength)
    currentLength = (currentLength + 1) Mod (Len(msg) + 1)
End Sub
