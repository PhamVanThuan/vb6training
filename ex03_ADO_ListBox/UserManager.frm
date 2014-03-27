VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmUsers 
   Caption         =   "Form1"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLogin 
      Appearance      =   0  'Flat
      Caption         =   "Login"
      Height          =   495
      Left            =   5640
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox txtUsername 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1920
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Check user existance"
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7335
      Begin VB.Label Label2 
         Caption         =   "Password"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Username"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.ListBox lstUsers 
      Height          =   1425
      Left            =   360
      TabIndex        =   4
      Top             =   3000
      Width           =   6855
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load all users"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "User management"
      Height          =   6735
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   7335
      Begin VB.CommandButton cmdLoadToGrid 
         Caption         =   "Load to grid"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   2880
         Width           =   1815
      End
      Begin MSDataGridLib.DataGrid dgdUsers 
         Height          =   2775
         Left            =   240
         TabIndex        =   9
         Top             =   3600
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   4895
         _Version        =   393216
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
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection

Private Sub cmdLoad_Click()

Dim rcs As New ADODB.Recordset

rcs.Open "SELECT * FROM ex03_users", cnn, adOpenDynamic, adLockOptimistic

Do While Not rcs.EOF
    lstUsers.AddItem rcs.Fields(0)
    rcs.MoveNext
Loop


rcs.Close

End Sub

Private Sub cmdLoadToGrid_Click()
Dim rcs As New ADODB.Recordset
rcs.Open "SELECT * FROM ex03_users", cnn, adOpenStatic, adLockReadOnly

Set dgdUsers.DataSource = rcs

End Sub

Private Sub cmdLogin_Click()

Dim rcs As New ADODB.Recordset
rcs.Open "SELECT * FROM ex03_users WHERE username='" + _
    txtUsername.Text + "' and password='" + txtPassword.Text + "'", cnn, adOpenStatic, adLockReadOnly

If rcs.RecordCount > 0 Then
    MsgBox "Record Found"
Else
    MsgBox "No Record Found"
End If

rcs.Close

End Sub



Private Sub Form_Load()
cnn.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=vb_example;uid=sa;pwd=123456;"

End Sub

Private Sub Form_Unload(Cancel As Integer)
cnn.Close
End Sub
