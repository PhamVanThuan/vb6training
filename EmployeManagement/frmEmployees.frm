VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.Form frmEmployees 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employees"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   11940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Remove"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New "
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5520
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvEmployees 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9340
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Firstname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Lastname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Middlename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Department"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmEmployees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim selectedId As String

Private Sub cmdEdit_Click()

If lvEmployees.SelectedItem Is Nothing Then
    MsgBox "No record selected to edit!", vbExclamation, "Edit Status"
Else
    'Show UpdateDialog
    frmEditEmployee.tbId.Text = lvEmployees.SelectedItem.Text
    frmEditEmployee.Show vbModal
End If

End Sub

Private Sub cmdNew_Click()

frmAddEmployee.tbId.Enabled = True
frmAddEmployee.Show vbModal

End Sub

Private Sub Form_Load()
'STEP1
lvEmployees.LabelEdit = lvwManual 'make data not editable
lvEmployees.FullRowSelect = True
lvEmployees.GridLines = True
lvEmployees.View = lvwReport      'set the view mode to REPORT mode. (show columns headers in order)

'To view/edit the columns, go right-click lvList, -> properties, -> in "column headers" Tab...

Call refreshEmployeeList

End Sub
