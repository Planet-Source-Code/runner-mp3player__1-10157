VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Mp3Player-Liste"
   ClientHeight    =   4140
   ClientLeft      =   4230
   ClientTop       =   2910
   ClientWidth     =   4590
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   4590
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1200
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Verzeichnis "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   3720
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   2520
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Liste laden"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Liste Speichern"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      MaskColor       =   &H00000000&
      TabIndex        =   5
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Liste löschen"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   960
      Left            =   360
      MultiSelect     =   2  'Erweitert
      TabIndex        =   3
      Top             =   2400
      Width           =   3855
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FFFF&
      Height          =   2040
      Left            =   2400
      Pattern         =   "*.mp3;*.wav"
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H80000006&
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FFFF&
      Height          =   1665
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAdd_Click()
Form3.Show
End Sub
Private Sub Command1_Click()
List1.Clear
End Sub
Private Sub Command2_Click()
Dim naampje As String
naampje = InputBox("Geben Sie den Namen der Liste ein ?", "Listen Name")
naampje = naampje & ".mp3l"
Open (App.Path & "\" & naampje) For Output As #1
       Dim i%
       For i = 0 To List1.ListCount - 1
       Print #1, List1.List(i)
       Next
       Close #1
End Sub
Private Sub Command3_Click()
Dim File As String
CommonDialog2.DialogTitle = "Load your list."
   CommonDialog2.MaxFileSize = 16384
   CommonDialog2.FileName = ""
   CommonDialog2.Filter = "Mp3-Listen|*.mp3l"
   CommonDialog2.ShowOpen
If CommonDialog2.FileName = "" Then Exit Sub
File = CommonDialog2.FileName
Dim A As String
Dim X As String
On Error GoTo Error
Open File For Input As #1
Do Until EOF(1)
Input #1, A$
List1.AddItem A$
Loop
Close 1
Exit Sub
Error:
X = MsgBox("Liste nicht gefunden!", vbOKOnly, "Fehler!")
End Sub
Private Sub Dir_Click()
Form3.Show vbModal
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub
Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub
Private Sub File1_Click()
Dim backslash
 backslash = Right(File1.Path, 1)
 If backslash = "\" Then
 List1.AddItem File1.Path & File1.FileName
 Else
List1.AddItem File1.Path & "\" & File1.FileName
End If
End Sub
Private Sub List1_DblClick()
Form1.MediaPlayer1.FileName = List1.Text
If Form1.MediaPlayer1.PlayState = mpPlaying Then
Form1.MediaPlayer1.CurrentPosition = 0
Form1.MediaPlayer1.Play
End If
Form1.Caption = List1.Text
End Sub
