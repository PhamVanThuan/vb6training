VERSION 5.00
Begin VB.Form frmTest 
   AutoRedraw      =   -1  'True
   Caption         =   "Test"
   ClientHeight    =   5370
   ClientLeft      =   615
   ClientTop       =   1650
   ClientWidth     =   8205
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   8205
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   4335
      TabIndex        =   2
      Top             =   4500
      Width           =   1875
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   2355
      TabIndex        =   1
      Top             =   4500
      Width           =   1875
   End
   Begin VB.CommandButton cmdTryIt 
      Caption         =   "Try It"
      Height          =   495
      Left            =   375
      TabIndex        =   0
      Top             =   4500
      Width           =   1875
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdClear_Click()
    Cls
End Sub


Private Sub cmdExit_Click()
    End
End Sub


Private Sub cmdTryIt_Click()

    Dim strEmpFileName  As String
    Dim strBackSlash  As String
    Dim intEmpFileNbr As Integer
    
    Dim strEmpName    As String
    Dim intDeptNbr    As Integer
    Dim strJobTitle   As String
    Dim dtmHireDate   As Date
    Dim sngHrlyRate   As Single
    
    strBackSlash = IIf(Right$(App.Path, 1) = "\", "", "\")
    strEmpFileName = App.Path & strBackSlash & "EMPLOYEE.DAT"
    intEmpFileNbr = FreeFile
    
    Open strEmpFileName For Input As #intEmpFileNbr
    
    Do Until EOF(intEmpFileNbr)
        Input #intEmpFileNbr, strEmpName, intDeptNbr, strJobTitle, dtmHireDate, sngHrlyRate
        Print strEmpName; _
              Tab(25); Format$(intDeptNbr, "@@@@"); _
              Tab(35); strJobTitle; _
              Tab(55); Format$(dtmHireDate, "mm/dd/yyyy"); _
              Tab(70); Format$(Format$(sngHrlyRate, "Standard"), "@@@@@@@")
    Loop
    
    Close #intEmpFileNbr

End Sub

