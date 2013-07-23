VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   8790
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17535
   DrawMode        =   15  'Merge Pen Not
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8749.863
   ScaleMode       =   0  'User
   ScaleWidth      =   17535
   Begin VB.CommandButton cmdBooks 
      Height          =   735
      Index           =   3
      Left            =   15960
      Picture         =   "Form2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6700
      Width           =   615
   End
   Begin VB.CommandButton cmdADD 
      Height          =   735
      Index           =   2
      Left            =   15960
      Picture         =   "Form2.frx":08FC
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5490
      Width           =   615
   End
   Begin VB.CommandButton cmdHELPR 
      Height          =   735
      Index           =   1
      Left            =   15960
      Picture         =   "Form2.frx":11F8
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4290
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Height          =   8775
      Left            =   0
      Picture         =   "Form2.frx":1AF4
      ScaleHeight     =   8715
      ScaleWidth      =   17595
      TabIndex        =   0
      Top             =   0
      Width           =   17655
      Begin VB.PictureBox Picture5 
         Height          =   375
         Left            =   16080
         Picture         =   "Form2.frx":2FC89
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   7
         Top             =   1920
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00404000&
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H8000000B&
         Height          =   5295
         Left            =   12600
         ScaleHeight     =   5235
         ScaleWidth      =   4515
         TabIndex        =   2
         Top             =   2400
         Width           =   4575
         Begin VB.CommandButton cmdHELPI 
            Height          =   735
            Index           =   0
            Left            =   3280
            Picture         =   "Form2.frx":3018D
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   600
            Width           =   615
         End
         Begin VB.CommandButton cmdAddBooks 
            BackColor       =   &H80000012&
            Caption         =   "ADD BOOKS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   3
            Left            =   480
            TabIndex        =   6
            Top             =   4210
            Width           =   2775
         End
         Begin VB.CommandButton cmdAddMembers 
            BackColor       =   &H8000000C&
            Caption         =   "ADD MEMBERS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   2
            Left            =   480
            TabIndex        =   5
            Top             =   3010
            Width           =   2775
         End
         Begin VB.CommandButton cmdReturn 
            BackColor       =   &H80000010&
            Caption         =   "RETURN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   1
            Left            =   480
            TabIndex        =   4
            Top             =   1810
            Width           =   2775
         End
         Begin VB.CommandButton cmdIssue 
            BackColor       =   &H80000012&
            Caption         =   "ISSUE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   0
            Left            =   480
            MaskColor       =   &H00FFFF00&
            TabIndex        =   3
            Top             =   610
            Width           =   2775
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "LOGOUT"
         Height          =   375
         Left            =   14760
         TabIndex        =   1
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   4215
         Left            =   9480
         Picture         =   "Form2.frx":30A89
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdADD_Click(Index As Integer)
MsgBox "This option allows you to add members to the library system.The default member id starts from 1001.You also have the option to delete members from the library", vbInformation, "ADD MEMBERS"

End Sub

Private Sub cmdAddBooks_Click(Index As Integer)
Form6.Show
Form2.Hide

End Sub

Private Sub cmdAddMembers_Click(Index As Integer)
Form5.Show
Form2.Hide

End Sub

Private Sub cmdBooks_Click(Index As Integer)
MsgBox "This option makes you to add books to the database.The BookId starts from 1000.This option also allows you to delete books from the library!!!", vbInformation, "ADD BOOKS"

End Sub

Private Sub cmdHELPI_Click(Index As Integer)
MsgBox "This option allows you to issue books to the library members.Please update after issuing the book", vbInformation, "ISSUe BOOKS"

End Sub

Private Sub cmdHELPR_Click(Index As Integer)
MsgBox "This option allows you return books to the library.Please update the system after returning the book.", vbInformation, "RETURN BOOKS"

End Sub

Private Sub cmdIssue_Click(Index As Integer)
Form3.Show
Form2.Hide


End Sub

Private Sub cmdReturn_Click(Index As Integer)
Form4.Show
Form2.Hide

End Sub

Private Sub Command1_Click()
Form1.Show
Form2.Hide


End Sub

Private Sub Command3_Click(Index As Integer)

End Sub

Private Sub Command2_Click(Index As Integer)

End Sub
