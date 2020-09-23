VERSION 5.00
Begin VB.Form Dialog 
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Mp3Player"
   ClientHeight    =   3195
   ClientLeft      =   9225
   ClientTop       =   1125
   ClientWidth     =   6030
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H80000007&
      Caption         =   "&Schließen"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "E-Mail:   RunnerSp@gmx.de"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   2520
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Version 1.1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "Mp3Player"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "CopyRight (c) Stefan Schloßmacher"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   1920
      Width           =   3855
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub
