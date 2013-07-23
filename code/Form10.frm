VERSION 5.00
Begin VB.Form Form10 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   8100
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19170
   DrawStyle       =   6  'Inside Solid
   Icon            =   "Form10.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   19170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   8175
      Left            =   0
      Picture         =   "Form10.frx":03CE
      ScaleHeight     =   8115
      ScaleWidth      =   19155
      TabIndex        =   0
      Top             =   0
      Width           =   19215
      Begin VB.CommandButton Command5 
         Height          =   375
         Left            =   17640
         Picture         =   "Form10.frx":1A199
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Height          =   1095
         Left            =   16200
         Picture         =   "Form10.frx":1A69D
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Height          =   1095
         Left            =   16200
         Picture         =   "Form10.frx":1F528
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Height          =   375
         Left            =   6480
         Picture         =   "Form10.frx":2433F
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   7320
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   6480
         Picture         =   "Form10.frx":24B33
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6720
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   3840
         TabIndex        =   5
         Top             =   7320
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   3840
         TabIndex        =   4
         Top             =   6720
         Width           =   2415
      End
      Begin VB.Label Label15 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11880
         TabIndex        =   20
         Top             =   6480
         Width           =   3495
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11880
         TabIndex        =   19
         Top             =   5280
         Width           =   3495
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11880
         TabIndex        =   18
         Top             =   4320
         Width           =   3495
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11880
         TabIndex        =   17
         Top             =   3480
         Width           =   3495
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11880
         TabIndex        =   16
         Top             =   2640
         Width           =   3495
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11880
         TabIndex        =   15
         Top             =   1800
         Width           =   3495
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000E&
         Caption         =   "Availability:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7650
         TabIndex        =   14
         Top             =   6480
         Width           =   2295
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000E&
         Caption         =   "Shelf No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7632
         TabIndex        =   13
         Top             =   5400
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000E&
         Caption         =   "Publisher:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7614
         TabIndex        =   12
         Top             =   4440
         Width           =   2295
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         Caption         =   "ISBN Number:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7596
         TabIndex        =   11
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "Author:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7560
         TabIndex        =   10
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "Title:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   7578
         TabIndex        =   8
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Author:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   7320
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Title:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   6720
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Search By Category:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   1
         Top             =   6240
         Width           =   2655
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "Form10"
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
    .CommandText = "SELECT * From addBooks where Title LIKE '" & Text1.Text & "'"
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
      Label10.Caption = rst!Title
      Label11.Caption = rst!Author
      Label12.Caption = rst!ISBNNumber
      Label13.Caption = rst!Publisher
      Label14.Caption = rst!ShelfNumber
      Label15.Caption = rst!Availability
      
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
 Dim conn As New ADODB.Connection
  Dim cmd As New ADODB.Command
  Dim rst As New ADODB.Recordset
  
 
  conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
    App.Path & "\" & "LibraryManagement.mdb;Mode=Read|Write"
 
  conn.CursorLocation = adUseClient
  conn.Open
  
  
  With cmd
    .ActiveConnection = conn
    .CommandText = "SELECT * From addBooks where Author LIKE '" & Text2.Text & "'"
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
      Label10.Caption = rst!Title
      Label11.Caption = rst!Author
      Label12.Caption = rst!ISBNNumber
      Label13.Caption = rst!Publisher
      Label14.Caption = rst!ShelfNumber
      Label15.Caption = rst!Availability
      
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

Private Sub Command3_Click()
Form11.Show
Unload Me

End Sub

Private Sub Command4_Click()
Form11.Show
Unload Me

End Sub

Private Sub Command5_Click()
Form1.Show
Unload Me

End Sub
