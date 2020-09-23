VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmSplash.frx":0CCA
   ScaleHeight     =   1980
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2400
      Top             =   1200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MSECURE DIRECTORY SECURITY"
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
      Top             =   120
      Width           =   6375
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   120
      Picture         =   "frmSplash.frx":ECB9C
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   120
      Picture         =   "frmSplash.frx":ED753
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   5055
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Image2.Width = 0
End Sub

Private Sub Timer1_Timer()
Image2.Width = Image2.Width + 20
If Image2.Width = 5055 Then


    Unload Me
    frmMain.Show
    
End If
End Sub
