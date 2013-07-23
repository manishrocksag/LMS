VERSION 5.00
Begin VB.Form Form6 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16185
   DrawMode        =   15  'Merge Pen Not
   DrawStyle       =   4  'Dash-Dot-Dot
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   16185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRight 
      Height          =   1095
      Index           =   1
      Left            =   14760
      Picture         =   "Form6.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdLeft 
      Height          =   1095
      Index           =   0
      Left            =   14760
      Picture         =   "Form6.frx":4E8B
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      DownPicture     =   "Form6.frx":9CA2
      Height          =   1095
      Index           =   5
      Left            =   14760
      Picture         =   "Form6.frx":A4BD
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   1095
      Index           =   4
      Left            =   14760
      Picture         =   "Form6.frx":B521
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      DownPicture     =   "Form6.frx":C3D2
      Height          =   1095
      Index           =   3
      Left            =   14760
      Picture         =   "Form6.frx":CA8D
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1560
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   8175
      Left            =   0
      Picture         =   "Form6.frx":D80F
      ScaleHeight     =   8115
      ScaleWidth      =   16155
      TabIndex        =   0
      Top             =   0
      Width           =   16215
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Left            =   5160
         TabIndex        =   32
         Top             =   7440
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   8160
         TabIndex        =   8
         Top             =   6840
         Width           =   4935
      End
      Begin VB.PictureBox Picture7 
         Height          =   975
         Left            =   14520
         ScaleHeight     =   975
         ScaleWidth      =   15
         TabIndex        =   28
         Top             =   0
         Width           =   15
      End
      Begin VB.PictureBox Picture6 
         Height          =   2415
         Left            =   14520
         ScaleHeight     =   2415
         ScaleWidth      =   15
         TabIndex        =   27
         Top             =   0
         Width           =   15
      End
      Begin VB.PictureBox Picture5 
         Height          =   6855
         Left            =   14520
         ScaleHeight     =   6855
         ScaleWidth      =   15
         TabIndex        =   26
         Top             =   0
         Width           =   15
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   1095
         Left            =   14520
         TabIndex        =   22
         Top             =   0
         Width           =   75
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   1095
         Left            =   14520
         TabIndex        =   21
         Top             =   0
         Width           =   75
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   7455
         Left            =   14520
         TabIndex        =   20
         Top             =   0
         Width           =   15
      End
      Begin VB.PictureBox Picture4 
         Height          =   7455
         Left            =   14520
         ScaleHeight     =   7455
         ScaleWidth      =   15
         TabIndex        =   19
         Top             =   0
         Width           =   15
      End
      Begin VB.PictureBox Picture3 
         Height          =   7335
         Left            =   14520
         ScaleHeight     =   7335
         ScaleWidth      =   15
         TabIndex        =   18
         Top             =   0
         Width           =   15
      End
      Begin VB.TextBox Text7 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   405
         Index           =   6
         Left            =   8160
         TabIndex        =   7
         Top             =   6120
         Width           =   4935
      End
      Begin VB.TextBox Text6 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   8160
         TabIndex        =   6
         Top             =   5400
         Width           =   4935
      End
      Begin VB.TextBox Text5 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   8160
         TabIndex        =   5
         Top             =   4680
         Width           =   4935
      End
      Begin VB.TextBox Text4 
         Height          =   405
         Index           =   3
         Left            =   8160
         TabIndex        =   4
         Top             =   3960
         Width           =   4935
      End
      Begin VB.TextBox Text3 
         Height          =   405
         Index           =   2
         Left            =   8160
         TabIndex        =   3
         Top             =   3240
         Width           =   4935
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Index           =   1
         Left            =   8160
         TabIndex        =   2
         Top             =   2520
         Width           =   4935
      End
      Begin VB.TextBox Txt1 
         Height          =   405
         Index           =   0
         Left            =   8160
         TabIndex        =   1
         Top             =   1680
         Width           =   4935
      End
      Begin VB.PictureBox Picture2 
         Height          =   5775
         Left            =   360
         Picture         =   "Form6.frx":1A2B0
         ScaleHeight     =   5715
         ScaleWidth      =   2955
         TabIndex        =   9
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label label3 
         BackColor       =   &H8000000B&
         Caption         =   "Show All Books Record"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   33
         Top             =   7440
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "No Of Copies:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   31
         Top             =   6840
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Purchase Price:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   4560
         TabIndex        =   17
         Top             =   6120
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Shelf No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   4560
         TabIndex        =   16
         Top             =   5520
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "ISBN No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   4560
         TabIndex        =   15
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Publisher Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   4560
         TabIndex        =   14
         Top             =   3960
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Author:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   4560
         TabIndex        =   13
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Title:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   4560
         TabIndex        =   12
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "BOOK ID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   4560
         TabIndex        =   10
         Top             =   1680
         Width           =   2295
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cn As New ADODB.Connection


Private Sub Check1_Click()
Form8.Show
Unload Me


End Sub

Private Sub cmdAdd_Click(Index As Integer)
Txt1(0).SetFocus


End Sub

Private Sub cmdDelete_Click(Index As Integer)
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
updstr = "delete from addBooks where BookId=" & Txt1(0).Text
        
cmd.ActiveConnection = conn
cmd.CommandType = adCmdText
cmd.CommandText = updstr
cmd.Execute
    
msg = MsgBox("The Book with the given id is successfully deleted from database", vbInformation, "ADD BOOKS")
er:
MsgBox "Successful", vbInformation

    
End Sub

Private Sub cmdLeft_Click(Index As Integer)
Form2.Show
Form6.Hide

End Sub

Private Sub cmdRight_Click(Index As Integer)
Form2.Show
Form6.Hide

End Sub

Private Sub cmdSave_Click(Index As Integer)
 On Error GoTo er
Dim msg
Dim rst As New ADODB.Recordset
Dim Response As Integer
rst.Open "addBooks", cn, adOpenStatic, adLockOptimistic, adCmdTable
With rst
.AddNew
.Fields("BookId") = Me.Txt1(0).Text
.Fields("Title") = Me.Text2(1).Text
.Fields("Author") = Me.Text3(2).Text
.Fields("Publisher") = Me.Text4(3).Text
.Fields("ISBNNumber") = Me.Text5(4).Text
.Fields("ShelfNumber") = Me.Text6(5).Text
.Fields("PurchasePrice") = Me.Text7(6).Text
.Fields("NoOfCopies") = Me.Text1.Text




  
    Response = MsgBox("Do you want to add this book  in the database", vbYesNoCancel + vbQuestion, "Enter your Choice")
        If (Response = vbYes) Then
                .Update
       Response = MsgBox("Your record has been saved in the database", vbOK + vbInformation, "Library Management System")
       
       If (Response = vbOK) Then
    Me.Txt1(0).Text = vbNullString
    Me.Text2(1).Text = vbNullString
    Me.Text3(2).Text = vbNullString
    Me.Text4(3).Text = vbNullString
    Me.Text5(4).Text = vbNullString
    Me.Text6(5).Text = vbNullString
    Me.Text7(6).Text = vbNullString
    Me.Text1.Text = vbNullString
       
       Form6.Show
        ElseIf (Response = vbNo) Then
        MsgBox "Your record has not been saved in the database.Try Again", vbInformation, "Library Management System"
        Else
        Form2.Show
        Unload Me
        End If
        End If
er:
If (Err.Number = 0) Then


msg = MsgBox("Successful!", vbInformation, "ADD BOOKS")

ElseIf (Err.Number = -2147352571) Then
msg = MsgBox("You cannot have an empty field.Please fill all of them", vbExclamation, "ADD BOOKS")
Else
msg = MsgBox("A book with this id already exists.Duplicate ids' for books cannot be stored in the database", vbCritical, "ADD BOOKS")





End If



End With


End Sub

Private Sub Form_Load()
cn.ConnectionString = "Provider=Microsoft.Jet.OleDb.4.0; Data Source=" & App.Path & "\LibraryManagement.mdb"
    cn.Open
End Sub

Private Sub Form_Unload(Cancel As Integer)
cn.Close
    Set cn = Nothing
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'Only allow Numerical Values and a hyphen.If we input other values there will be a beep
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = 45) Or KeyAscii = 8 Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
'Only allow string values.If we input other values there will be a beep
If (KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 32) Or KeyAscii = 8 Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub Text4_KeyPress(Index As Integer, KeyAscii As Integer)
'Only allow string values.If we input other values there will be a beep
If (KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 32) Or KeyAscii = 8 Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub Text5_KeyPress(Index As Integer, KeyAscii As Integer)
'Only allow Numerical Values and a hyphen.If we input other values there will be a beep
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = 45) Or KeyAscii = 8 Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub Text6_KeyPress(Index As Integer, KeyAscii As Integer)
'Only allow Numeric Values,a hyphen(45) or constants from A to Z.If we input other values there will be a beep
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii >= 65 And KeyAscii <= 91) Or (KeyAscii = 45) Or KeyAscii = 8 Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub Text7_KeyPress(Index As Integer, KeyAscii As Integer)
'Only allow Numeric Values.If we input other values there will be a beep
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub Txt1_KeyPress(Index As Integer, KeyAscii As Integer)
'Only allow string values.If we input other values there will be a beep
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub
