VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HELP MANUAL"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar VScroll1 
      Height          =   3735
      Left            =   6240
      TabIndex        =   25
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox pctFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3705
      Left            =   0
      ScaleHeight     =   3675
      ScaleWidth      =   6225
      TabIndex        =   1
      Top             =   480
      Width           =   6255
      Begin VB.PictureBox pctHelp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   10335
         Left            =   0
         ScaleHeight     =   10305
         ScaleWidth      =   6225
         TabIndex        =   2
         Top             =   0
         Width           =   6255
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "MSECURE DIRECTORY/FOLDER SECURITY"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   120
            Width           =   5895
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Designed by : MONU"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   5895
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Designed In : Microsoft Visual Basic 6.0"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   600
            Width           =   5895
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "E-mail : manaruchimohapatra@yahoo.co.in"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   21
            Top             =   960
            Width           =   5895
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Program File Name : MSECUREDS.exe"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   20
            Top             =   1200
            Width           =   5895
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "INTRODUCTION"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   1560
            Width           =   5895
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmHelp.frx":0CCA
            Height          =   975
            Index           =   4
            Left            =   120
            TabIndex        =   18
            Top             =   1800
            Width           =   5895
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "HOW TO USE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   2640
            Width           =   5895
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "THE LOGIN SCREEN"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   2880
            Width           =   5895
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmHelp.frx":0E0A
            Height          =   975
            Index           =   5
            Left            =   120
            TabIndex        =   15
            Top             =   3120
            Width           =   5895
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Note  : If you lost or forget your password, you can't recover it. So it's good to make a note of it."
            ForeColor       =   &H000000FF&
            Height          =   495
            Index           =   6
            Left            =   120
            TabIndex        =   14
            Top             =   4200
            Width           =   5895
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "               After entering the password. click on &GO . The the program will load and you will see the MAIN SCREEN."
            Height          =   495
            Index           =   7
            Left            =   120
            TabIndex        =   13
            Top             =   4680
            Width           =   5895
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "THE MAIN SCREEN"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Top             =   5160
            Width           =   5895
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmHelp.frx":0F86
            Height          =   975
            Index           =   8
            Left            =   120
            TabIndex        =   11
            Top             =   5400
            Width           =   5895
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "PROTECTING A FOLDER"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   10
            Top             =   5880
            Width           =   5895
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmHelp.frx":101B
            Height          =   1215
            Index           =   9
            Left            =   120
            TabIndex        =   9
            Top             =   6120
            Width           =   5895
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "UN-PROTECTING A FOLDER"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   8
            Top             =   7320
            Width           =   5895
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmHelp.frx":11F9
            Height          =   615
            Index           =   10
            Left            =   120
            TabIndex        =   7
            Top             =   7560
            Width           =   5895
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmHelp.frx":12E7
            ForeColor       =   &H000000FF&
            Height          =   495
            Index           =   11
            Left            =   120
            TabIndex        =   6
            Top             =   8160
            Width           =   5895
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "CHANGING A PASSWORD"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   5
            Top             =   8640
            Width           =   5895
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmHelp.frx":136F
            Height          =   855
            Index           =   12
            Left            =   120
            TabIndex        =   4
            Top             =   8880
            Width           =   6015
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   6120
            Y1              =   9720
            Y2              =   9720
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "© MSOFT sft."
            Height          =   255
            Left            =   4440
            TabIndex        =   3
            Top             =   9840
            Width           =   1695
         End
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MSECURE DIRECTORY SECURITY > HELP MANUAL"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6375
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   0
      Left            =   0
      Picture         =   "frmHelp.frx":14B1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6840
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If pctHelp.Height > pctFrame.Height Then
    VScroll1.Enabled = True
End If
End Sub

Private Sub VScroll1_Change()
    VScroll1.Max = pctHelp.Height - pctFrame.Height
    pctHelp.Top = -VScroll1.Value
End Sub
