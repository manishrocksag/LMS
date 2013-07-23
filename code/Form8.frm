VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form8 
   BackColor       =   &H8000000E&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "BOOKS RECORD"
   ClientHeight    =   7860
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   18240
   DrawMode        =   3  'Not Merge Pen
   HasDC           =   0   'False
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   18240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   8175
      Left            =   -120
      Picture         =   "Form8.frx":0000
      ScaleHeight     =   8115
      ScaleWidth      =   18315
      TabIndex        =   0
      Top             =   0
      Width           =   18375
      Begin VB.CommandButton Command2 
         Height          =   1095
         Left            =   14520
         Picture         =   "Form8.frx":4C61
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   6720
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Height          =   1095
         Left            =   12720
         Picture         =   "Form8.frx":9AEC
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   6720
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   2040
         Top             =   7080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\MANISH\Downloads\Project\LibraryManagement.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\MANISH\Downloads\Project\LibraryManagement.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select BookId,Title,Author,ISBNNumber,Publisher,ShelfNumber,Availability,NoOfCopies from addBooks"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form8.frx":E903
         Height          =   5295
         Left            =   1680
         TabIndex        =   1
         Top             =   1320
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   9340
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483634
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "BOOKS RECORD"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Form6.Show

Unload Me

End Sub

Private Sub Command2_Click()
Form3.Show
Unload Me


End Sub
