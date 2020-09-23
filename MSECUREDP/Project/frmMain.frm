VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MSECURE - FOLDER SECURITY"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin FolderSecurity.Button Button4 
      Height          =   375
      Left            =   1440
      TabIndex        =   20
      Top             =   5760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "Help"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   16777215
      FCOLO           =   33023
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":0CCA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin FolderSecurity.Button Button3 
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   5760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "About"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   16777215
      FCOLO           =   33023
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":0CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   120
      ScaleHeight     =   4545
      ScaleWidth      =   6825
      TabIndex        =   14
      Top             =   1080
      Width           =   6855
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1590
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   6495
      End
      Begin FolderSecurity.Button Command1 
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   21
         Top             =   3600
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Remove Protection"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":0D02
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Protected Directory List:"
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
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   2085
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":0D1E
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   6495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Directory"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         Left            =   240
         TabIndex        =   15
         Top             =   140
         Width           =   1920
      End
      Begin VB.Image Image7 
         Height          =   375
         Index           =   1
         Left            =   120
         Picture         =   "frmMain.frx":0DA9
         Stretch         =   -1  'True
         Top             =   120
         Width           =   6495
      End
      Begin VB.Image Image6 
         Height          =   4575
         Index           =   1
         Left            =   0
         Picture         =   "frmMain.frx":2F03
         Stretch         =   -1  'True
         Top             =   0
         Width           =   6855
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   120
      ScaleHeight     =   4545
      ScaleWidth      =   6825
      TabIndex        =   3
      Top             =   1080
      Width           =   6855
      Begin FolderSecurity.Button Command1 
         Height          =   375
         Index           =   0
         Left            =   3960
         TabIndex        =   13
         Top             =   3840
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Add Directory For Protection"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":53DD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   3840
         TabIndex        =   12
         Top             =   1600
         Width           =   2775
      End
      Begin VB.TextBox txtPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3840
         TabIndex        =   9
         Top             =   960
         Width           =   2895
      End
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2565
         Left            =   120
         TabIndex        =   8
         Top             =   1640
         Width           =   3255
      End
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Height          =   345
         Left            =   120
         TabIndex        =   6
         Top             =   920
         Width           =   3255
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   0
         Left            =   3840
         Picture         =   "frmMain.frx":53F9
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   465
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   1
         Left            =   4440
         Picture         =   "frmMain.frx":583B
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   2
         Left            =   5040
         Picture         =   "frmMain.frx":5C7D
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   3
         Left            =   5640
         Picture         =   "frmMain.frx":60BF
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   4
         Left            =   6240
         Picture         =   "frmMain.frx":6EAD
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Directory Name [ With Path ]:"
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
         Index           =   3
         Left            =   3840
         TabIndex        =   11
         Top             =   680
         Width           =   2565
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Show when Directory Opened:"
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
         Left            =   3840
         TabIndex        =   10
         Top             =   1320
         Width           =   2625
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   3600
         X2              =   3600
         Y1              =   600
         Y2              =   4200
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         X1              =   240
         X2              =   3120
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Directory:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Drive:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Directory"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   135
         Width           =   1800
      End
      Begin VB.Image Image7 
         Height          =   375
         Index           =   0
         Left            =   120
         Picture         =   "frmMain.frx":72EF
         Stretch         =   -1  'True
         Top             =   120
         Width           =   6495
      End
      Begin VB.Image Image6 
         Height          =   4575
         Index           =   0
         Left            =   0
         Picture         =   "frmMain.frx":9449
         Stretch         =   -1  'True
         Top             =   0
         Width           =   6855
      End
   End
   Begin FolderSecurity.Button Button1 
      Height          =   375
      Left            =   5400
      TabIndex        =   22
      Top             =   5760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "&Change Password"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   16777215
      FCOLO           =   33023
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":B923
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
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PROTECT FOLDER"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   780
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UNPROTECT FOLDER"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   780
      Width           =   2295
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   2280
      Picture         =   "frmMain.frx":B93F
      Stretch         =   -1  'True
      Top             =   720
      Width           =   2295
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   120
      Picture         =   "frmMain.frx":DA99
      Stretch         =   -1  'True
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MSECURE DIRECTORY SECURITY WINDOW"
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
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   6375
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   6
      Left            =   120
      Picture         =   "frmMain.frx":FBF3
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6840
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   120
      Picture         =   "frmMain.frx":107AA
      Stretch         =   -1  'True
      Top             =   720
      Width           =   2175
   End
   Begin VB.Image Image5 
      Height          =   375
      Left            =   2280
      Picture         =   "frmMain.frx":12410
      Stretch         =   -1  'True
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Control_P = ".{21EC2020-3AEA-1069-A2DD-08002B30309D}"
Const My_COMP = ".{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
Const Desk_TOP = ".{9E56BE61-C50F-11CF-9A2C-00A0C90A90CE}"
Const NetWork = ".{208D2C60-3AEA-1069-A2D7-08002B30309D}"
Const IE = ".{FBF23B42-E3F0-101B-8488-00AA003E56F8}"
Const RBin = ".{645FF040-5081-101B-9F08-00AA002F954E}"
Const Printer = ".{2227A280-3AEA-1069-A2DE-08002B30309D}"
Const HTMLDoc = ".{25336920-03F9-11CF-8FD0-00AA00686F13}"
Const TaskS = ".{255b3f60-829e-11cf-8d8b-00aa0060f5bf}"
Const WaveFile = ".{0003000D-0000-0000-C000-000000000046}"
Const MovClip = ".{00022602-0000-0000-C000-000000000046}"
Const WinIcon = ".{00021401-0000-0000-C000-000000000046}"

Private Sub Button1_Click()
frmCP.Show vbModal

End Sub

Private Sub Button3_Click()
frmAbout.Show vbModal
End Sub

Private Sub Button4_Click()
frmHelp.Show

End Sub

Private Sub Command1_Click(Index As Integer)
Dim Str, Attrib, Icon As String
Dim I, count As Integer

On Error GoTo ErrHnd:
Close #1

'Initialize Icon variable
Select Case Combo1.ListIndex
Case 0: Icon = NetWork
Case 1: Icon = Control_P
Case 2: Icon = RBin
Case 3: Icon = HTMLDoc
Case 4: Icon = Printer
End Select

Select Case Index

'If "Add Directory" button clicked
Case 0:

'Save Folder path, It's look and Extension
Open "c:\flist.dat" For Append As #1
Write #1, txtPath.Text, Combo1.List(Combo1.ListIndex), Icon
Close #1

'Rename Folder
Name txtPath.Text As txtPath.Text & Icon

List1.Clear

'Call "Form_Load" Event
Form_Load

'If "Remove Protection" button clicked
Case 1:

'Get the selected item's Index form the List Box
count = List1.ListIndex

'Open "flist.dat" for information
Open "c:\flist.dat" For Input As #1

I = 0
While Not EOF(1)
Input #1, Str, Attrib, Icon

'"flist.dat" and ListBox maintain same serial of the item
If I = count Then
Name Str & Icon As Str
Close #1
GoTo next_s
End If
I = I + 1
Wend

next_s:
Close #1

List1.Clear
Form_Load

'Save new List
Open "c:\flist.dat" For Input As #1
Open "c:\flist.tmp" For Append As #2
I = 0
While Not EOF(1)

If I = count Then
Input #1, Str, Attrib, Icon
GoTo 2:
End If
Input #1, Str, Attrib, Icon
Write #2, Str, Attrib, Icon
I = I + 1
Wend

2:
Close #1
Close #2

Kill "C:\flist.dat"
Name "C:\flist.tmp" As "C:\flist.dat"
List1.Clear
Form_Load
End Select

ErrHnd:
If Err.Number <> 0 Then
MsgBox Err.Description, vbCritical, Err.Number
End If
End Sub

Private Sub Dir1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtPath.Text = Dir1.List(Dir1.ListIndex)
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
Picture2.Visible = False
Picture1.Visible = True

Image2.Visible = False
Image3.Visible = True
On Error GoTo errnum
Dim Str, Attrib, Icon As String
Combo1.Clear
List1.Clear

Combo1.AddItem "My Network Places"
Combo1.AddItem "Control Panel"
Combo1.AddItem "Recyclebin"
Combo1.AddItem "HTML document"
Combo1.AddItem "Printers"
Combo1.ListIndex = 0

txtPath.Text = Dir1.List(Dir1.ListIndex)

'Following file has changed Folders information
Open "c:\flist.dat" For Input As #1

While Not EOF(1)
Input #1, Str, Attrib, Icon
List1.AddItem Str & " [ Icon: " & Attrib & " ]"
Wend

Close #1
If List1.ListIndex <> -1 Then
List1.ListIndex = 0
End If

errnum:

If Err.Number = 53 Then
Exit Sub
End If
If Err.Number <> 0 Then
MsgBox Err.Description, vbCritical, Err.Number
End If

End Sub




Private Sub Image2_Click()
Label2_Click
End Sub

Private Sub Image3_Click()
Label3_Click
End Sub

Private Sub Label2_Click()

Image2.Visible = False
Image3.Visible = True
Picture2.Visible = False
Picture1.Visible = True

End Sub



Private Sub Label3_Click()

Image3.Visible = False
Image2.Visible = True
Picture1.Visible = False
Picture2.Visible = True
End Sub
Private Sub Combo1_Click()
Dim I As Integer
For I = 0 To 4
Image1(I).BorderStyle = 0
Next I
Image1(Combo1.ListIndex).BorderStyle = 1
End Sub

