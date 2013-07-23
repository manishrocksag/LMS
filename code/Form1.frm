VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000E&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   7080
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   17160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7080
   ScaleWidth      =   17160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   13680
      Top             =   4560
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   12960
      Top             =   4440
   End
   Begin VB.Frame Frame2 
      Height          =   4800
      Left            =   8310
      TabIndex        =   8
      Top             =   2250
      Width           =   3950
      Begin VB.CommandButton Command4 
         Height          =   135
         Left            =   0
         TabIndex        =   26
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Validate"
         Height          =   375
         Left            =   1920
         TabIndex        =   15
         Top             =   3960
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   0
         PasswordChar    =   "*"
         TabIndex        =   13
         Tag             =   "password"
         Top             =   2640
         Width           =   3395
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   0
         TabIndex        =   11
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label Label8 
         Caption         =   "Trouble Signing In"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   14
         Top             =   3360
         Width           =   3615
      End
      Begin VB.Label Label7 
         Caption         =   "PASSWORD"
         Height          =   495
         Left            =   0
         TabIndex        =   12
         Top             =   2160
         Width           =   3615
      End
      Begin VB.Label Label6 
         Caption         =   "ENTER YOUR ROLL NO"
         Height          =   495
         Left            =   0
         TabIndex        =   10
         Top             =   900
         Width           =   3375
      End
      Begin VB.Label Label5 
         Caption         =   "Sign Is As Student"
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
         Left            =   0
         TabIndex        =   9
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4800
      Left            =   4600
      TabIndex        =   0
      Top             =   2250
      Width           =   3700
      Begin VB.CommandButton Command3 
         Height          =   135
         Left            =   0
         TabIndex        =   25
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Validate"
         Height          =   375
         Left            =   1920
         Picture         =   "Form1.frx":AAE5
         TabIndex        =   7
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   0
         PasswordChar    =   "*"
         TabIndex        =   5
         Tag             =   "password"
         Top             =   2640
         Width           =   3395
      End
      Begin VB.TextBox txtUsername 
         Height          =   495
         Left            =   0
         TabIndex        =   3
         Tag             =   "admin"
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label Label4 
         Caption         =   "Trouble Signing In"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   3360
         Width           =   3615
      End
      Begin VB.Label Label3 
         Caption         =   "PASSWORD"
         Height          =   495
         Left            =   0
         TabIndex        =   4
         Top             =   2040
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "USERNAME"
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Sign In As Adminstrator"
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
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Image Image3 
      Height          =   7095
      Index           =   4
      Left            =   0
      Picture         =   "Form1.frx":10D6F
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Image Image3 
      Height          =   7095
      Index           =   3
      Left            =   0
      Picture         =   "Form1.frx":1603E
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Image Image3 
      Height          =   7095
      Index           =   2
      Left            =   0
      Picture         =   "Form1.frx":209FD
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Image Image3 
      Height          =   7095
      Index           =   1
      Left            =   0
      Picture         =   "Form1.frx":33B13
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Image Image3 
      Height          =   7095
      Index           =   0
      Left            =   0
      Picture         =   "Form1.frx":36327
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Image Image2 
      Height          =   7095
      Left            =   0
      Picture         =   "Form1.frx":4211A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000E&
      Caption         =   "Ritushree Rathi(104562)"
      Height          =   375
      Index           =   3
      Left            =   12480
      TabIndex        =   24
      Top             =   6650
      Width           =   4695
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000E&
      Caption         =   "Raushan Kumar(105459)"
      Height          =   375
      Index           =   2
      Left            =   12480
      TabIndex        =   23
      Top             =   6200
      Width           =   4695
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000E&
      Caption         =   "Manish Agarwal(105456)"
      Height          =   375
      Index           =   1
      Left            =   12480
      TabIndex        =   22
      Top             =   5730
      Width           =   4695
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000D&
      Caption         =   "CREDITS:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   12480
      TabIndex        =   21
      Top             =   5280
      Width           =   4695
   End
   Begin VB.Label lblYear 
      BackColor       =   &H8000000E&
      Caption         =   "2012"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   15360
      TabIndex        =   20
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label lblMonth 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      Caption         =   "April"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15120
      TabIndex        =   19
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label lblDay 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "TUESDAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   12480
      TabIndex        =   18
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label lblNumber 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15360
      TabIndex        =   17
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblTime 
      BackColor       =   &H8000000E&
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12480
      TabIndex        =   16
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   5235
      Left            =   -375
      Picture         =   "Form1.frx":473E9
      Top             =   0
      Width           =   17550
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim startTime As Variant


Private Sub Command1_Click()
'This procedure checks the input password
Dim Response As Integer
If txtPassword.Text = "" And txtUsername.Text = "" Then
MsgBox "Please Enter Username and Password!", vbOKOnly + vbExclamation, "Try Again"
ElseIf txtPassword.Text = "" Then
MsgBox "Please Enter your Password! It cannot be blank", vbOKOnly + vbExclamation, "Try Again"
ElseIf txtUsername.Text = "" Then
MsgBox "Please Enter your Username! It cannot be blank", vbOKOnly + vbExclamation, "Try Again"
ElseIf txtPassword.Text = txtPassword.Tag And txtUsername.Text <> txtUsername.Tag Then
'If correct, display message box
MsgBox "Incorrect Username", vbOKOnly + vbExclamation, "Try Again"
ElseIf txtPassword.Text <> txtPassword.Tag And txtUsername.Text = txtUsername.Tag Then
MsgBox "Incorrect Password", vbOKOnly + vbExclamation, "Try Again"
ElseIf txtPassword.Text <> txtPassword.Tag And txtUsername.Text <> txtUsername.Tag Then
MsgBox "Incorrect Username and Password", vbOKOnly + vbExclamation, "Try Again"
Me.txtPassword.Text = vbNullString
Me.txtUsername.Text = vbNullString



Else
'If incorrect, give option to try again
Response = MsgBox("Access Granted", vbOKOnly, "Access Granted")

If Response = vbOK Then
Me.txtPassword.Text = vbNullString
Me.txtUsername.Text = vbNullString
Form2.Show
Form1.Hide
Else
End

End If
End If


End Sub

Private Sub Command2_Click()
'This procedure checks the input password
Dim Response As Integer
If Text4.Text = "" And Text3.Text = "" Then
MsgBox "Please Enter Username and Password!", vbOKOnly + vbExclamation, "Try Again"
ElseIf Text4.Text = "" Then
MsgBox "Please Enter your Password! It cannot be blank", vbOKOnly + vbExclamation, "Try Again"
ElseIf Text3.Text = "" Then
MsgBox "Please Enter your Username! It cannot be blank", vbOKOnly + vbExclamation, "Try Again"
ElseIf Text4.Text <> txtPassword.Tag Then
MsgBox "Incorrect Password", vbOKOnly + vbExclamation, "Try Again"
Me.Text4.Text = vbNullString

Else
'If incorrect, give option to try again
Response = MsgBox("Access Granted", vbOKOnly, "Access Granted")

If Response = vbOK Then
Me.Text3.Text = vbNullString
Me.Text4.Text = vbNullString
Form9.Show
Form1.Hide
Else
End

End If
End If

End Sub

Private Sub Command3_Click()
MsgBox "The default username is admin and default password is password", vbOKOnly + vbExclamation, ""
End Sub

Private Sub Command4_Click()
MsgBox "The default password is password", vbInformation, "Student Login"

End Sub

Private Sub Form_Activate()
txtUsername.SetFocus
End Sub

Private Sub Image2_Click()
Timer2.Enabled = Not (Timer2.Enabled)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
'Only allow numerical values.If we input other values there will be a beep
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub Timer1_Timer()
Dim today As Variant
today = Now

lblTime.Caption = Format(today, "h:mm:ss ampm")
lblNumber.Caption = Format(today, "d")
lblDay.Caption = Format(today, "dddd")
lblMonth.Caption = Format(today, "mmmm")
lblYear.Caption = Format(today, "yyyy")


End Sub

Private Sub Timer2_Timer()
Static PicNum As Integer
PicNum = PicNum + 1
If PicNum > 4 Then PicNum = 0
Image2.Picture = Image3(PicNum).Picture
End Sub
