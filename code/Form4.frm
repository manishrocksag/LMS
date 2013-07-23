VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form4 
   Caption         =   "RETURN BOOKS"
   ClientHeight    =   9120
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20040
   LinkTopic       =   "Form4"
   ScaleHeight     =   9120
   ScaleWidth      =   20040
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   6
      Left            =   13440
      TabIndex        =   11
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   5
      Left            =   13440
      TabIndex        =   10
      Top             =   3600
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   1
      Left            =   13440
      TabIndex        =   9
      Top             =   4320
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      Height          =   9135
      Left            =   0
      Picture         =   "Form4.frx":0000
      ScaleHeight     =   9075
      ScaleWidth      =   20235
      TabIndex        =   0
      Top             =   0
      Width           =   20295
      Begin VB.CommandButton Command4 
         Caption         =   "SET"
         Height          =   375
         Left            =   18960
         TabIndex        =   24
         Top             =   7920
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "GET"
         Height          =   375
         Left            =   18960
         TabIndex        =   23
         Top             =   7320
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   18120
         TabIndex        =   22
         Top             =   7920
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Height          =   405
         Left            =   16920
         TabIndex        =   21
         Top             =   7920
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   16920
         TabIndex        =   20
         Top             =   7320
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   13440
         TabIndex        =   18
         Top             =   5760
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Height          =   615
         Left            =   15240
         Picture         =   "Form4.frx":21852
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   5680
         Width           =   615
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   13440
         TabIndex        =   16
         Top             =   5040
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   873
         _Version        =   393216
         Format          =   99549185
         CurrentDate     =   41020
      End
      Begin VB.CommandButton cmdRight 
         Height          =   1095
         Left            =   15240
         Picture         =   "Form4.frx":21EE8
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   7200
         Width           =   1215
      End
      Begin VB.CommandButton cmdLeft 
         Height          =   1095
         Left            =   13440
         Picture         =   "Form4.frx":26D73
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   7200
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Height          =   1095
         Left            =   11640
         Picture         =   "Form4.frx":2BB8A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   7200
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Left            =   17040
         Picture         =   "Form4.frx":2CBEE
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2050
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Index           =   0
         Left            =   13440
         TabIndex        =   8
         Top             =   2160
         Width           =   3375
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000014&
         Caption         =   "Update Status:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   16920
         TabIndex        =   25
         Top             =   6720
         Width           =   2535
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Image Image1 
         Height          =   660
         Left            =   14520
         Picture         =   "Form4.frx":2D3E2
         Top             =   5640
         Width           =   570
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         Caption         =   "Fine:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11640
         TabIndex        =   7
         Top             =   5880
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "Return Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11640
         TabIndex        =   6
         Top             =   5160
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "Due Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11640
         TabIndex        =   5
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Member Id:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11640
         TabIndex        =   4
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Title :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11640
         TabIndex        =   3
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Book Id :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   11640
         TabIndex        =   1
         Top             =   2280
         Width           =   1815
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Book Id :"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cn As New ADODB.Connection


Private Sub cmdLeft_Click()
Form2.Show
Form4.Hide

End Sub

Private Sub cmdRight_Click()
Form5.Show
Form4.Hide

End Sub

Private Sub cmdSave_Click()
Dim msg
Dim rst As New ADODB.Recordset
Dim Response As Integer


rst.Open "returnBooks", cn, adOpenStatic, adLockOptimistic, adCmdTable
With rst
.AddNew
.Fields("BookId") = Me.Text1(0).Text
.Fields("Title") = Me.Text1(6).Text
.Fields("MemberId") = Me.Text1(5).Text
.Fields("DueDate") = Me.Text1(1).Text
.Fields("ReturnDate") = Me.DTPicker1.Value
.Fields("Fine") = Me.Text3.Text
 On Error GoTo er
    Response = MsgBox("Do you want to save this record in the database", vbYesNoCancel + vbQuestion, "Enter your Choice")
        If (Response = vbYes) Then
         
        .Update
       Response = MsgBox("Your record has been saved in the database", vbOK + vbInformation, "Library Management System")
       
       If (Response = vbOK) Then
       Form2.Show
        Form4.Hide

        
       
        
        
        ElseIf (Response = vbNo) Then
        MsgBox "Your record has not been saved in the database.Try Again", vbInformation, "Library Management System"
        Else
        Form2.Show
        Unload Me
        End If
        End If
        
        
er:
If (Err.Number = 0) Then


msg = MsgBox("Successful!", vbInformation, "Library Management System")

Else
msg = MsgBox(Err.Description, vbCritical, "OOPS!Try Again ")


End If



End With
End Sub

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
    .CommandText = "SELECT * From issueBooks where BookID LIKE '" & Text1(0).Text & "'"
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
      Text1(6).Text = rst!Title
      Text1(5).Text = rst!MemberID
      Text1(1).Text = rst!dateOfReturn
     
      
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
Dim due As Date
Dim ret As Date
Dim dif As String
Dim fine As Integer

due = Text1(1).Text
ret = DTPicker1.Value

dif = DateDiff("d", due, ret)
Label7.Caption = dif
fine = Label7.Caption * 2
Text3.Text = fine
MsgBox "Fine is calculated @ Rs 2 per due day!!", vbInformation, "FINE"



End Sub

Private Sub Command3_Click()
 Dim conn As New ADODB.Connection
  Dim cmd As New ADODB.Command
  Dim rst As New ADODB.Recordset
  
 
  conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
    App.Path & "\" & "LibraryManagement.mdb;Mode=Read|Write"
 
  conn.CursorLocation = adUseClient
  conn.Open
  
  
  With cmd
    .ActiveConnection = conn
    .CommandText = "SELECT * From addBooks where BookId LIKE '" & Text2.Text & "'"
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
      Text4.Text = rst!Availability
      Text5.Text = rst!NoOfCopies
      
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

Private Sub Command4_Click()
On Error GoTo er

 Dim conn As New ADODB.Connection
  Dim cmd As New ADODB.Command
  Dim rst As New ADODB.Recordset
  Dim updstr
  Dim msg
  
  
 
  conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
    App.Path & "\" & "LibraryManagement.mdb;Mode=Read|Write"
 
  conn.CursorLocation = adUseClient
  conn.Open
  On Error GoTo er

updstr = "update addBooks set Availability='" & Text4.Text & "',NoOfCopies=" & Text5.Text & " where BookId=" & Text2.Text





        
cmd.ActiveConnection = conn
cmd.CommandType = adCmdText
cmd.CommandText = updstr
cmd.Execute
    
    msg = MsgBox("BookRecord Successfully Modified", vbInformation, "ISSUE BOOKS")
    
    
Exit Sub

er:

msg = MsgBox(Err.Description, vbCritical, "ISSUE BOOKS")
  
  
End Sub

Private Sub Form_Load()
cn.ConnectionString = "Provider=Microsoft.Jet.OleDb.4.0; Data Source=" & App.Path & "\LibraryManagement.mdb"
    cn.Open
End Sub

Private Sub Form_Unload(Cancel As Integer)
cn.Close
    Set cn = Nothing

End Sub
