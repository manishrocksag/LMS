VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19095
   LinkTopic       =   "Form9"
   ScaleHeight     =   7125
   ScaleWidth      =   19095
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Height          =   7095
      Left            =   14160
      Picture         =   "Form9.frx":0000
      ScaleHeight     =   7035
      ScaleWidth      =   4875
      TabIndex        =   3
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton Command6 
         Caption         =   "LOGOUT"
         Height          =   615
         Left            =   2400
         TabIndex        =   9
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Height          =   615
         Left            =   3840
         Picture         =   "Form9.frx":288A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Height          =   735
         Left            =   3600
         Picture         =   "Form9.frx":2D8E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5400
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Height          =   735
         Left            =   3600
         Picture         =   "Form9.frx":368A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3960
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000014&
         Caption         =   "Student Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   720
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Form9.frx":3F86
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5400
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000014&
         Caption         =   "Books Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   720
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Form9.frx":4309
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3960
         Width           =   2895
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   7095
      Left            =   0
      Picture         =   "Form9.frx":468C
      ScaleHeight     =   7035
      ScaleWidth      =   14115
      TabIndex        =   0
      Top             =   0
      Width           =   14175
      Begin VB.PictureBox Picture3 
         Height          =   5535
         Left            =   14160
         ScaleHeight     =   5535
         ScaleWidth      =   15
         TabIndex        =   2
         Top             =   0
         Width           =   15
      End
      Begin VB.PictureBox Picture2 
         Height          =   5895
         Left            =   14160
         ScaleHeight     =   5895
         ScaleWidth      =   15
         TabIndex        =   1
         Top             =   0
         Width           =   15
      End
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Form10.Show
Unload Me

End Sub

Private Sub Command2_Click()
Form11.Show
Unload Me

End Sub


Private Sub Command3_Click()
MsgBox "The student can address all their queries in relation to search for books in the library", vbInformation
End Sub

Private Sub Command4_Click()
MsgBox "Students can get all the information of all the books issued in their name and their due date", vbInformation
End Sub

Private Sub Command6_Click()
Form1.Show
Unload Me

End Sub
