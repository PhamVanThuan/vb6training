VERSION 5.00
Begin VB.Form frmAddEmployee 
   Caption         =   "New Employee"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   3360
      TabIndex        =   13
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   2640
      Width           =   1335
   End
   Begin VB.ComboBox cmbStatus 
      Height          =   315
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2040
      Width           =   1815
   End
   Begin VB.ComboBox cmbDeparment 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox tbLastname 
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox tbMiddleName 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox tbFirstname 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox tbId 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Status:"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Department:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Lastname:"
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Middle name:"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Firstname:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "ID"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmAddEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdReset_Click()
Clear

End Sub

Private Sub Clear()

Dim ctrl As Control

For Each ctrl In Me
    If TypeOf ctrl Is TextBox Then
        ctrl.Text = vbNullString
    ElseIf TypeOf ctrl Is ComboBox Then
        ctrl.ListIndex = -1
    End If
Next ctrl

End Sub

Private Sub cmdSave_Click()

On Error GoTo onSaveError

'Validate inputs
If tbId.Text = vbNullString Or tbLastname.Text = vbNullString Or tbFirstname.Text = vbNullString _
    Or cmbDeparment.ListIndex = -1 Or cmbStatus.ListIndex = -1 Then
    
    MsgBox "Please fill all of the required fields in order to proceed.", vbExclamation, "Saving Status"
    Exit Sub:

End If

Dim query As String
query = "INSERT INTO fnl_Employees(Id, Firstname, Middlename, Lastname, Department, Status)" _
    & " VALUES ('" & tbId.Text & "', '" & tbFirstname.Text & "', '" & tbMiddleName.Text _
    & "', '" & tbLastname.Text & "', '" & cmbDeparment & "', " & cmbStatus.ListIndex & ");"

executeQuery query

Call refreshEmployeeList
Call Clear

If MsgBox("New records has been added now in your database, Do you want to enter another record?", vbYesNo + vbQuestion, "Confirmation Dialog") = vbNo Then
    Unload Me 'close frmAdd
End If

Exit Sub:
onSaveError:

MsgBox Err.Description, vbCritical, "Error"
Set connection = Nothing
Set recordset = Nothing
End 'End application

End Sub

Private Sub Form_Load()

cmbDeparment.Clear
cmbDeparment.AddItem "Technical Team"
cmbDeparment.AddItem "IT Dept."
cmbDeparment.AddItem "HR Dept."
cmbDeparment.AddItem "MY Dept."

cmbStatus.Clear
cmbStatus.AddItem "False"
cmbStatus.AddItem "True"

End Sub

