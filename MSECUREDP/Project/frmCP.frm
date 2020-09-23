VERSION 5.00
Begin VB.Form frmCP 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MSECURE CHANGE PASWORD"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5775
   ControlBox      =   0   'False
   Icon            =   "frmCP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin FolderSecurity.Button Button1 
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   7
      TX              =   "&Change"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCP.frx":0CCA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2400
      Width           =   5295
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1680
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   960
      Width           =   5295
   End
   Begin FolderSecurity.Button Button2 
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   7
      TX              =   "C&ancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCP.frx":0CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Verify New Password :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "New Password :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Current Password :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1665
   End
   Begin VB.Image Image2 
      Height          =   2895
      Left            =   120
      Picture         =   "frmCP.frx":0D02
      Stretch         =   -1  'True
      Top             =   600
      Width           =   5535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CHANGE PASSWORD"
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
      Index           =   6
      Left            =   0
      Picture         =   "frmCP.frx":ECBD4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6840
   End
End
Attribute VB_Name = "frmCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Button1_Click()
If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" Then
    MsgBox "This Following Error has Occured:" + vbCrLf + "* Current Password Field couldn't be left blank" + vbCrLf + "* New Password Field couldn't be left blank" + vbCrLf + "* Verify New Password Field couldn't be left blank", vbExclamation, "Error"

ElseIf Text1.Text = "" And Text2.Text = "" Then
    MsgBox "This Following Error has Occured:" + vbCrLf + "* Current Password Field couldn't be left blank" + vbCrLf + "* New Password Field couldn't be left blank", vbExclamation, "Error"
ElseIf Text1.Text = "" And Text3.Text = "" Then
    MsgBox "This Following Error has Occured:" + vbCrLf + "* Current Password Field couldn't be left blank" + vbCrLf + "* Verify New Password Field couldn't be left blank", vbExclamation, "Error"
ElseIf Text1.Text <> Label2.Caption Then
    MsgBox "This Following Error has Occured:" + vbCrLf + "* Wrong Password Entered", vbExclamation, "Error"
Else

    If Text2.Text <> Text3.Text Then
        MsgBox "This Following Error has Occured:" + vbCrLf + "* New Password doesn't match with Verify Password", vbExclamation, "Error"
    Else
         SaveSetting App.Title, "Settings", "Password", Text3.Text
         Unload Me
    End If
End If
End Sub

Private Sub Button2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Label2.Caption = GetSetting(App.Title, "Settings", "Password")
End Sub

