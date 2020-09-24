VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   9555
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   11640
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmStart.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmStart.frx":000C
   ScaleHeight     =   9555
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6330
      Left            =   1200
      TabIndex        =   0
      Top             =   1440
      Width           =   9465
      Begin VB.Timer timer 
         Interval        =   5000
         Left            =   7680
         Top             =   3960
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Abdel Soft"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   8
         Top             =   2040
         Width           =   3255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "License To BTI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   7
         Top             =   2760
         Width           =   3375
      End
      Begin VB.Image Image1 
         Height          =   1440
         Left            =   360
         Picture         =   "frmStart.frx":55054
         Stretch         =   -1  'True
         Top             =   480
         Width           =   2040
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         Height          =   1215
         Left            =   3240
         Shape           =   2  'Oval
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright 2006"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   2
         Top             =   5280
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "This product is protected BY ISO and ANSI.  All rights reserved"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   720
         TabIndex        =   1
         Top             =   3960
         Width           =   3285
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5520
         TabIndex        =   3
         Top             =   3480
         Width           =   1635
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "windows family"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5280
         TabIndex        =   4
         Top             =   3120
         Width           =   2265
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "nlineBanking"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   4440
         TabIndex        =   6
         Top             =   1800
         Width           =   4710
      End
      Begin VB.Label B 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banking System Software"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4680
         TabIndex        =   5
         Top             =   360
         Width           =   3570
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         Height          =   1935
         Left            =   3120
         Shape           =   2  'Oval
         Top             =   1080
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub timerSplash_Timer()

End Sub

Private Sub Form_Load()
    
    Frame1.Move (Screen.Width - Frame1.Width) / 2, (Screen.Height - Frame1.Height) / 2

End Sub

Private Sub timer_Timer()
    Unload Me
    frmLogin.Show
End Sub
