Attribute VB_Name = "mdlDataConnection"
'Step 2: Add new variables
Public connection As ADODB.connection
Public recordset As ADODB.recordset

'Step 3: Sub for openconnection
Public Sub openConnection()

Set connection = New ADODB.connection
connection.ConnectionString = "Driver=SQL Server;Server=.\SQLEXPRESS;Database=vb_example;uid=sa;pwd=123456;"
connection.Open

End Sub

'Step :
Public Sub executeQuery(ByVal sql As String)

On Error GoTo connectionErr

Set connection = New ADODB.connection

If connection.State = adStateOpen Then
    connection.Close
End If

openConnection

connection.Execute sql, adOpenDynamic, adCmdText

Exit Sub:

'catch error
connectionErr:
MsgBox Err.Description, vbCritical, "Error"
End

End Sub

'STEP 5
Public Sub refreshEmployeeList()
Dim listItem As listItem

initRecordset "SELECT * FROM fnl_Employees ORDER BY Id"

frmEmployees.lvEmployees.ListItems.Clear

While Not recordset.EOF = True 'While myRs did not reach the End Of the File(last record in the database)
Set listItem = frmEmployees.lvEmployees.ListItems.Add(, , recordset.Fields("ID")) 'put the first data from tblInfo(ID Number) into the first column of lvList(ID Number)
listItem.SubItems(1) = recordset.Fields("Firstname") 'put second data from tblInof(Surname) into the second column of lvList(Surname)
listItem.SubItems(2) = recordset.Fields("Lastname") 'put third data from tblInofFirstname) into the third column of lvList(Surname)
listItem.SubItems(3) = recordset.Fields("Middlename") 'put fourth data from tblInof(Middlename) into the fourth column of lvList(Surname)
listItem.SubItems(4) = recordset.Fields("Department") 'put fifth data from tblInof(Department) into the fifth column of lvList(Surname)
listItem.SubItems(5) = recordset.Fields("Status") 'put sixth data from tblInof(Status) into the second sixth of lvList(Surname)
recordset.MoveNext 'move to the next record
Wend 'end While
End Sub

'STEP 4
Public Sub initRecordset(ByVal query As String)

On Error GoTo recordsetErr

Set recordset = New ADODB.recordset
If recordset.State = adStateOpen Then
    recordset.Close
End If

openConnection

recordset.Open query, connection, adOpenDynamic, adLockOptimistic

Exit Sub:

'catch error
recordsetErr:

MsgBox Err.Description, vbCritical, "Error"
Set recordset = Nothing
End 'terminate the program when an error occured

End Sub




