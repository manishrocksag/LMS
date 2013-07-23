VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form3 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "ISSUE BOOKS"
   ClientHeight    =   8805
   ClientLeft      =   -8880
   ClientTop       =   5115
   ClientWidth     =   20025
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   20025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRight 
      Height          =   1095
      Index           =   3
      Left            =   12960
      Picture         =   "Form3.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdLEft 
      Height          =   1095
      Index           =   2
      Left            =   10680
      Picture         =   "Form3.frx":4E8B
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Height          =   1095
      Index           =   1
      Left            =   8520
      Picture         =   "Form3.frx":9CA2
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7440
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000E&
      Height          =   8895
      Left            =   0
      Picture         =   "Form3.frx":AD06
      ScaleHeight     =   8835
      ScaleWidth      =   20235
      TabIndex        =   0
      Top             =   0
      Width           =   20295
      Begin VB.TextBox Text8 
         Height          =   525
         Left            =   15960
         TabIndex        =   3
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   16320
         TabIndex        =   33
         Top             =   6000
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   15000
         TabIndex        =   2
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Height          =   405
         Left            =   17160
         TabIndex        =   32
         Top             =   6700
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   195
         Left            =   15600
         TabIndex        =   28
         Top             =   8160
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   15600
         TabIndex        =   27
         Top             =   7560
         Width           =   255
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "SET"
         Height          =   255
         Left            =   18480
         TabIndex        =   26
         Top             =   6720
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Height          =   1095
         Index           =   0
         Left            =   6120
         Picture         =   "Form3.frx":1E7D8
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   7440
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Height          =   405
         Left            =   7800
         TabIndex        =   21
         Top             =   4200
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   7800
         TabIndex        =   20
         Top             =   3480
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   7800
         TabIndex        =   1
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   16320
         TabIndex        =   19
         Top             =   6720
         Width           =   615
      End
      Begin VB.CommandButton cmdGet 
         Caption         =   "GET"
         Height          =   255
         Left            =   18480
         TabIndex        =   18
         Top             =   6120
         Width           =   735
      End
      Begin VB.CommandButton cmdSearch2 
         Height          =   615
         Index           =   1
         Left            =   18120
         Picture         =   "Form3.frx":1F55A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2450
         Width           =   615
      End
      Begin VB.CommandButton cmdSearch1 
         Height          =   615
         Index           =   0
         Left            =   10680
         Picture         =   "Form3.frx":1FD4E
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2520
         Width           =   615
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   7800
         TabIndex        =   14
         Top             =   6120
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         Format          =   99418113
         CurrentDate     =   41019
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   7800
         TabIndex        =   12
         Top             =   5040
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         Format          =   99418113
         CurrentDate     =   41019
      End
      Begin VB.Line Line4 
         X1              =   15000
         X2              =   17640
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Line Line3 
         X1              =   17640
         X2              =   17640
         Y1              =   3240
         Y2              =   5640
      End
      Begin VB.Line Line2 
         X1              =   15000
         X2              =   17640
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line Line1 
         X1              =   15000
         X2              =   15000
         Y1              =   3240
         Y2              =   5760
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000E&
         Caption         =   "Show All Member Records"
         Height          =   375
         Left            =   16080
         TabIndex        =   31
         Top             =   8160
         Width           =   3375
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000E&
         Caption         =   "Show All Books Record"
         Height          =   375
         Index           =   0
         Left            =   16080
         TabIndex        =   29
         Top             =   7560
         Width           =   3375
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
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
         Left            =   13680
         TabIndex        =   17
         Top             =   6120
         Width           =   3015
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "Date Of Return:"
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
         Left            =   5520
         TabIndex        =   13
         Top             =   6120
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "Date Of Issue:"
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
         Left            =   5520
         TabIndex        =   11
         Top             =   5160
         Width           =   2055
      End
      Begin VB.Image Image1 
         Height          =   2415
         Left            =   15000
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Member Id:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13440
         TabIndex        =   10
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label2 
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
         Height          =   375
         Left            =   5520
         TabIndex        =   9
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label1 
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
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   5
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Book ID:"
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
         Index           =   0
         Left            =   5520
         TabIndex        =   4
         Top             =   2640
         Width           =   1695
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Index           =   4
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Book ID:"
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
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cn As New ADODB.Connection

Private Sub Command1_Click(Index As Integer)
End Sub

Private Sub Check1_Click()
Form8.Show
Unload Me




End Sub

Private Sub Check2_Click()
Form7.Show
Unload Me



End Sub

Private Sub cmdAdd_Click(Index As Integer)
Text1.SetFocus

End Sub

Private Sub cmdGet_Click()
 Dim conn As New ADODB.Connection
  Dim cmd As New ADODB.Command
  Dim rst As New ADODB.Recordset
  
 
  conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
    App.Path & "\" & "LibraryManagement.mdb;Mode=Read|Write"
 
  conn.CursorLocation = adUseClient
  conn.Open
  
  
  With cmd
    .ActiveConnection = conn
    .CommandText = "SELECT * From addBooks where BookId LIKE '" & Text3.Text & "'"
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
      Text5.Text = rst!Availability
      Text7.Text = rst!NoOfCopies
      
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

Private Sub cmdLeft_Click(Index As Integer)
Form2.Show
Unload Me

End Sub

Private Sub cmdRight_Click(Index As Integer)
Form4.Show
Unload Me

End Sub

Private Sub cmdSave_Click(Index As Integer)
Dim msg


Dim arrImage() As Byte
Dim Response As Integer
Dim fNum As Integer
Dim rst As New ADODB.Recordset
Dim photoPath As String
Dim isImage As Boolean

If Me.Image1.Picture <> LoadPicture("") Then
SavePicture Me.Image1.Picture, App.Path & "\picture.jpg"
photoPath = App.Path & "\picture.jpg"
ReDim arrImage(FileLen(photoPath))
fNum = FreeFile()
Open photoPath For Binary As #fNum
Get #fNum, , arrImage
Close fNum
isImage = True
End If
rst.Open "issueBooks", cn, adOpenStatic, adLockOptimistic, adCmdTable
With rst
.AddNew
.Fields("BookID") = Me.Text1.Text
.Fields("Title") = Me.Text2.Text
.Fields("MemberID") = Me.Text4.Text
.Fields("dateOfIssue") = Me.DTPicker1.Value
.Fields("dateOfReturn") = Me.DTPicker1.Value
If isImage = True Then
.Fields("Picture").AppendChunk arrImage
 On Error GoTo er
    Response = MsgBox("Do you want to save this record in the database", vbYesNoCancel + vbQuestion, "Enter your Choice")
        If (Response = vbYes) Then
         
        .Update
       Response = MsgBox("Your record has been saved in the database", vbOK + vbInformation, "Library Management System")
       
       If (Response = vbOK) Then
        Me.Text1.Text = vbNullString
        Me.Text2.Text = vbNullString
        Me.Text4.Text = vbNullString
        Me.Text8.Text = vbNullString
        Me.Text6.Text = vbNullString
        
       
       Form3.Show
       
        

        
       
        
        
        ElseIf (Response = vbNo) Then
        MsgBox "Your record has not been saved in the database.Try Again", vbInformation, "Library Management System"
        Else
        Form2.Show
        Unload Me
        End If
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

Private Sub cmdSearch1_Click(Index As Integer)
   Dim conn As New ADODB.Connection
  Dim cmd As New ADODB.Command
  Dim rst As New ADODB.Recordset
  
 
  conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
    App.Path & "\" & "LibraryManagement.mdb;Mode=Read|Write"
 
  conn.CursorLocation = adUseClient
  conn.Open
  
  
  With cmd
    .ActiveConnection = conn
    .CommandText = "SELECT * From addBooks where BookId LIKE '" & Text1.Text & "'"
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
      Text2.Text = rst!Title
      Text6.Text = rst!Availability
      
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

Private Sub cmdSearch2_Click(Index As Integer)
On Error GoTo er

 Dim rst As New ADODB.Recordset
    Dim arrImageByte() As Byte
    Dim fNum As Integer
    Dim strPhotoPath As String
    strPhotoPath = App.Path & "\Picture.jpg"
    
    rst.Open "SELECT * FROM addMember WHERE [Name]='" & Me.Text8.Text & "'", _
            cn, adOpenStatic, adLockOptimistic
    If Not (rst.EOF And rst.BOF) Then
        If rst.Fields("Picture").ActualSize <> 0 Then
            arrImageByte = rst.Fields("Picture").GetChunk(rst.Fields("Picture").ActualSize)
            
            fNum = FreeFile()
            Open strPhotoPath For Binary As #fNum
            Put #fNum, , arrImageByte
            Close fNum
            Set Me.Image1.Picture = LoadPicture(strPhotoPath)
            
    
        Else
            MsgBox "No photo to view!", vbInformation
        End If
    End If
er:
    MsgBox "Successful", vbInformation, "Add Books"
    
    
    
    
End Sub

Private Sub cmdSet_Click()
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

updstr = "update addBooks set Availability='" & Text5.Text & "',NoOfCopies=" & Text7.Text & " where BookId=" & Text3.Text





        
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
 cn.CursorLocation = adUseClient
    cn.Open
End Sub

Private Sub Form_Unload(Cancel As Integer)
cn.Close
    Set cn = Nothing
End Sub

