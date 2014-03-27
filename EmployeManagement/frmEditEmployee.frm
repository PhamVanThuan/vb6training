VERSION 5.00
Begin VB.Form frmEditEmployee 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Employee"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tbId 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox tbFirstname 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox tbMiddleName 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox tbLastname 
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.ComboBox cmbDeparment 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2040
      Width           =   2415
   End
   Begin VB.ComboBox cmbStatus 
      Height          =   315
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "ID"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Firstname:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Middle name:"
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Lastname:"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   960
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
   Begin VB.Label Label6 
      Caption         =   "Status:"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   1800
      Width           =   735
   End
End
Attribute VB_Name = "frmEditEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()

On Error GoTo saveError

If tbId.Text = vbNullString Or tbFirstname.Text = vbNullString Or tbLastname.Text = vbNullString _
    Or cmbDeparment.ListIndex = -1 Or cmbStatus.ListIndex = -1 Then
    
    MsgBox "Please fill all of the required fields in order to proceed.", vbExclamation, "Saving Status"
    Exit Sub:
End If

'execute the sql command for Update
Dim query As String
query = "UPDATE fnl_Employees SET Lastname = '" & tbLastname & "', Firstname = '" & tbFirstname _
    & "', Middlename = '" & tbMiddleName & "', Department = '" & cmbDeparment _
    & "', Status =  " & cmbStatus.ListIndex & " WHERE Id = '" & tbId & "';"
executeQuery query

MsgBox "Record successfully Updated.", vbInformation, "Update Success"
Call refreshEmployeeList
Unload Me

Exit Sub:
saveError:
MsgBox Err.Description, vbExclamation, "Error"
Set connection = Nothing
'Set recordset = Nothing
End

End Sub

Private Sub Form_Activate()

cmbDeparment.Clear
cmbDeparment.AddItem "Technical Team"
cmbDeparment.AddItem "IT Dept."
cmbDeparment.AddItem "HR Dept."
cmbDeparment.AddItem "MY Dept."

cmbStatus.Clear
cmbStatus.AddItem "False"
cmbStatus.AddItem "True"

'get the data of the selected employee to edit and update(NOTE: ID Number is no editable)
initRecordset "SELECT * FROM fnl_Employees WHERE Id = '" & frmEmployees.lvEmployees.SelectedItem.Text & "';"

If Not recordset.BOF = True Or recordset.EOF = True Then
    tbLastname.Text = recordset.Fields("Lastname")
    tbFirstname.Text = recordset.Fields("Firstname")
    tbMiddleName.Text = recordset.Fields("Middlename")
    
    Dim strDepartment As String
    Dim blStatus As Boolean
    blStatus = recordset.Fields("Status")
    strDepartment = recordset.Fields("Department")
    If blStatus Then
        cmbStatus.ListIndex = 1
    Else
        cmbStatus.ListIndex = 0
    End If
    
    For i = 0 To cmbDeparment.ListCount
        If strDepartment = cmbDeparment.List(i) Then
            cmbDeparment.ListIndex = i
            Exit For
        End If
    Next
    
Else
End If

End Sub

