VERSION 5.00
Begin VB.Form Form11 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form11"
   ClientHeight    =   7965
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   19530
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   19530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   7935
      Left            =   0
      Picture         =   "Form11.frx":0000
      ScaleHeight     =   7875
      ScaleWidth      =   19515
      TabIndex        =   0
      Top             =   0
      Width           =   19575
      Begin VB.CommandButton Command3 
         Height          =   1095
         Left            =   16320
         Picture         =   "Form11.frx":195AC
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4920
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Height          =   1095
         Left            =   16320
         Picture         =   "Form11.frx":1E437
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Left            =   15360
         Picture         =   "Form11.frx":2324E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   12840
         TabIndex        =   2
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000B&
         Height          =   495
         Left            =   12960
         TabIndex        =   11
         Top             =   5880
         Width           =   2535
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000B&
         Height          =   495
         Left            =   12960
         TabIndex        =   10
         Top             =   4920
         Width           =   2535
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000B&
         Height          =   495
         Left            =   12960
         TabIndex        =   9
         Top             =   3960
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000B&
         Height          =   495
         Left            =   12960
         TabIndex        =   8
         Top             =   3000
         Width           =   2535
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "Date Of Return:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   7
         Top             =   5760
         Width           =   3015
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "Date Of Issue:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   6
         Top             =   4920
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Title:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   5
         Top             =   4080
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Book ID:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   4
         Top             =   3120
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Enter Members' Library Identity :"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9240
         TabIndex        =   1
         Top             =   1920
         Width           =   4095
      End
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
 Dim conn As New ADODB.Connection
  Dim cmd As New ADODB.Command
  Dim rst As New ADODB.Recordset
  
 
  conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
    App.Path & "\" & "LibraryManagement.mdb;Mode=Read|Write"
 
  conn.CursorLocation = adUseClient
  conn.Open
  
  
  With cmd
    .ActiveConnection = conn
    .CommandText = "SELECT * From issueBooks where MemberID LIKE '" & Text1.Text & "'"
    .CommandType = adCmdText
  End With
 
  With rst
    .CursorType = adOpenStatic
    .CursorLocation = adUseClient
    .LockType = adLockOptimistic
    .Open cmd
  End With
 
  If rst.EOF = False Then
  
    rst.MoveFirst
   
    Do
      'Displays found record in Message Box
      Label6.Caption = rst!BookID
      Label7.Caption = rst!Title
      Label8.Caption = rst!dateOfIssue
      
      Label9.Caption = rst!dateOfReturn
      
      rst.MoveNext
    Loop Until rst.EOF = True
    
    rst.Close
  Else
    MsgBox "No records were found"
  End If
  
  conn.Close
  
  Set conn = Nothing
  Set cmd = Nothing
  Set rst = Nothing

End Sub

Private Sub Command2_Click()
Form10.Show
Unload Me

End Sub

Private Sub Command3_Click()
Form9.Show
Unload Me

End Sub

