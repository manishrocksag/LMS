VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form5 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form5"
   ClientHeight    =   9585
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18240
   DrawStyle       =   6  'Inside Solid
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   18240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExBkwd 
      Height          =   975
      Index           =   3
      Left            =   11520
      Picture         =   "Form5.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8160
      Width           =   1095
   End
   Begin VB.CommandButton cmdExFwd 
      Height          =   1095
      Index           =   2
      Left            =   15720
      Picture         =   "Form5.frx":4CC0
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8160
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      DownPicture     =   "Form5.frx":5912
      Height          =   1095
      Index           =   1
      Left            =   15840
      Picture         =   "Form5.frx":612D
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Height          =   1095
      Index           =   6
      Left            =   6240
      TabIndex        =   7
      Top             =   7800
      Width           =   4575
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Index           =   5
      Left            =   6240
      TabIndex        =   6
      Top             =   7080
      Width           =   4575
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Index           =   4
      Left            =   6240
      TabIndex        =   5
      Top             =   6000
      Width           =   4575
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Index           =   3
      Left            =   6240
      TabIndex        =   4
      Top             =   5280
      Width           =   4575
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Index           =   2
      Left            =   6240
      TabIndex        =   3
      Top             =   4440
      Width           =   4575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   1
      Left            =   6240
      TabIndex        =   2
      Top             =   3720
      Width           =   4575
   End
   Begin VB.PictureBox Picture1 
      Height          =   9615
      Left            =   0
      Picture         =   "Form5.frx":7191
      ScaleHeight     =   9555
      ScaleWidth      =   18195
      TabIndex        =   0
      Top             =   360
      Width           =   18255
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   6240
         TabIndex        =   1
         Top             =   2640
         Width           =   4575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   11760
         TabIndex        =   23
         Top             =   7080
         Width           =   135
      End
      Begin VB.CommandButton Command1 
         Height          =   1095
         Left            =   15840
         Picture         =   "Form5.frx":1E34C
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   6120
         Width           =   1335
      End
      Begin MSComDlg.CommonDialog cd 
         Left            =   2040
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdFwd 
         Height          =   975
         Index           =   1
         Left            =   14280
         Picture         =   "Form5.frx":1F1FD
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   7800
         Width           =   1095
      End
      Begin VB.CommandButton cmdBkwd 
         Height          =   975
         Index           =   0
         Left            =   12840
         Picture         =   "Form5.frx":23CBD
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   7800
         Width           =   1095
      End
      Begin VB.CommandButton cmdBrowse 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12600
         Picture         =   "Form5.frx":24691
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   5760
         Width           =   1935
      End
      Begin VB.CommandButton cmdAdd 
         DownPicture     =   "Form5.frx":25625
         Height          =   1095
         Index           =   0
         Left            =   15840
         Picture         =   "Form5.frx":25CE0
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Show All Member Records"
         Height          =   495
         Left            =   12240
         TabIndex        =   24
         Top             =   7080
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Address (Maximum 250 characters):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   6
         Left            =   4560
         TabIndex        =   14
         Top             =   7560
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Phone No:"
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
         Index           =   5
         Left            =   4560
         TabIndex        =   13
         Top             =   6780
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Validity Of Membership:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   4560
         TabIndex        =   12
         Top             =   5760
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Department:"
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
         TabIndex        =   11
         Top             =   4980
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Semester:"
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
         TabIndex        =   10
         Top             =   4200
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Name:"
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
         Index           =   1
         Left            =   4560
         TabIndex        =   9
         Top             =   3420
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Member ID:"
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
         Index           =   0
         Left            =   4560
         TabIndex        =   8
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   3015
         Left            =   12120
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   2895
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cn As New ADODB.Connection


Private Sub Command2_Click()


End Sub

Private Sub Check1_Click()
Form7.Show
Unload Me


End Sub

Private Sub cmdAdd_Click(Index As Integer)
Text1.SetFocus
End Sub

Private Sub cmdBkwd_Click(Index As Integer)
Form2.Show
Form5.Hide

End Sub

Private Sub cmdBrowse_Click()
Me.cd.ShowOpen
If Me.cd.FileName & "" <> "" Then
Set Me.Image1.Picture = LoadPicture(Me.cd.FileName)
End If
End Sub

Private Sub cmdExBkwd_Click(Index As Integer)
Form3.Show
Form5.Hide

End Sub

Private Sub cmdExFwd_Click(Index As Integer)
Form2.Show
Form5.Hide

End Sub

Private Sub cmdFwd_Click(Index As Integer)
Form6.Show
Unload Me
End Sub

Private Sub cmdSave_Click(Index As Integer)
 On Error GoTo er
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
rst.Open "addMember", cn, adOpenStatic, adLockOptimistic, adCmdTable
With rst
.AddNew
.Fields("MemberId") = Me.Text1.Text
.Fields("Name") = Me.Text2(1).Text
.Fields("Semester") = Me.Text3(2).Text
.Fields("Department") = Me.Text4(3).Text
.Fields("Validity Of Membership") = Me.Text5(4).Text
.Fields("Phone No") = Me.Text6(5).Text
.Fields("Address") = Me.Text7(6).Text
If isImage = True Then
.Fields("Picture").AppendChunk arrImage
  If (Me.Text1.Text = "" Or Me.Text2(1).Text = "" Or Me.Text3(2).Text = "" Or Me.Text4(3).Text = "" Or Me.Text5(4).Text = "" Or Me.Text6(5).Text = "" Or Me.Text7(6).Text = "") Then
  msg = MsgBox("You cannot have an empty field.Please fill all of them", vbExclamation, "ADD MEMBERS")
  Else
    Response = MsgBox("Do you want to save this record in the database", vbYesNoCancel + vbQuestion, "Enter your Choice")
        If (Response = vbYes) Then
         
        .Update
       Response = MsgBox("Your record has been saved in the database", vbOK + vbInformation, "Library Management System")
       
       If (Response = vbOK) Then
    Me.Text1.Text = vbNullString
    Me.Text2(1).Text = vbNullString
    Me.Text3(2).Text = vbNullString
    Me.Text4(3).Text = vbNullString
    Me.Text5(4).Text = vbNullString
    Me.Text6(5).Text = vbNullString
    Me.Text7(6).Text = vbNullString
       Form5.Show
    

        
       
        
        
        ElseIf (Response = vbNo) Then
        MsgBox "Your record has not been saved in the database.Try Again", vbInformation, "Library Management System"
        Else
        Form2.Show
        Unload Me
        End If
        End If
        End If
        End If
        
        

er:
If (Err.Number = 0) Then


msg = MsgBox("Successful!", vbInformation, "Library Management System")

Else
msg = MsgBox("A Member with this id already exists.Duplicate ids' for members cannot be stored in the database", vbCritical, "ADD BOOKS")
End If
End With


End Sub

Private Sub Command3_Click(Index As Integer)
Form6.Show
Form5.Hide

End Sub

Private Sub Command1_Click()
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
updstr = "delete from addMember where MemberId=" & (Text1.Text)

cmd.ActiveConnection = conn
cmd.CommandType = adCmdText
cmd.CommandText = updstr
cmd.Execute
    
msg = MsgBox("The Member with the given id and name is successfully deleted from database", vbInformation, "ADD BOOKS")
er:
MsgBox "Successful", vbInformation

End Sub

Private Sub Form_Load()

cn.ConnectionString = "Provider= Microsoft.Jet.OleDb.4.0;Data Source=" & App.Path & "\LibraryManagement.mdb"
cn.Open


End Sub

Private Sub Form_Unload(Cancel As Integer)
cn.Close
    Set cn = Nothing
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
'Only allow string values,a space amd a backspace.If we input other values there will be a beep
If (KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 32) Or KeyAscii = 8 Then

Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
'Only allow string values,a space amd a backspace.If we input other values there will be a beep
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 32) Or KeyAscii = 8 Then

Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub Text4_KeyPress(Index As Integer, KeyAscii As Integer)
'Only allow string values,a space amd a backspace.If we input other values there will be a beep
If (KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 32) Or KeyAscii = 8 Then

Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub Text5_KeyPress(Index As Integer, KeyAscii As Integer)
'Only allow string values,a space amd a backspace.If we input other values there will be a beep
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 32) Or (KeyAscii = 45) Or (KeyAscii = 47) Or KeyAscii = 8 Then

Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub Text6_KeyPress(Index As Integer, KeyAscii As Integer)
'Only allow string values,a space amd a backspace.If we input other values there will be a beep
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = 32) Or (KeyAscii = 45) Or (KeyAscii = 47) Or KeyAscii = 8 Then

Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub
