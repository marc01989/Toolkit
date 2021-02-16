VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRCNOViewNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim arrayNotes As Variant
Dim recordMax As Long, recordCounter As Long



Private Sub Form_Open(Cancel As Integer)
      
      Dim searchVal As String
      
    If Not IsNull(Forms!frmRCNOViewNotes.OpenArgs) Then
        searchVal = Forms!frmRCNOViewNotes.OpenArgs
        If Len(searchVal) > 0 Then
            txtSearch.Value = searchVal
            btnSearch_Click
        End If
    Else
    End If

End Sub

Private Sub btnSearch_Click()

Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim criteria As String

clearBoxes

Select Case optSearchType
Case 1 'string
    criteria = "FFMExchangeAssignedID = '" & Trim(txtSearch.Value) & "'"
Case 2 'long
    criteria = "InventoryNumber = " & Trim(txtSearch.Value)
End Select

Set db = CurrentDb
Set rs = db.OpenRecordset("SELECT * FROM tblRCNONotes WHERE " & criteria, dbOpenSnapshot, dbReadOnly)

With rs
    If .recordCount = 0 Then
        MsgBox "No records found"
        recordMax = 0
        recordCounter = 0
    Else
        .MoveFirst
        .MoveLast
        .MoveFirst
        txtRecordCount.Value = .recordCount
        recordMax = .recordCount - 1
        recordCounter = 0
        arrayNotes = .GetRows(.recordCount)
        Call populateBoxes(recordCounter)
    End If
    
End With

rs.Close: Set rs = Nothing
db.Close: Set db = Nothing


End Sub

Private Sub btnBack_Click()

If recordMax = 0 Then Exit Sub

If recordCounter = 0 Then
    Exit Sub
Else
    recordCounter = recordCounter - 1
    Call populateBoxes(recordCounter)
End If

End Sub

Private Sub btnForward_Click()

If recordMax = 0 Then Exit Sub

If recordCounter = recordMax Then
    Exit Sub
Else
    recordCounter = recordCounter + 1
    Call populateBoxes(recordCounter)
End If

End Sub

Public Sub populateBoxes(i As Long)

    txtERRCaseID.Value = arrayNotes(1, i)
    txtInventoryNumber = arrayNotes(2, i)
    txtBatchNumber = arrayNotes(3, i)
    txtNote = arrayNotes(4, i)
    txtFFMExchangeAssignedID = arrayNotes(5, i)
    txtFFMPolicyID = arrayNotes(6, i)
    txtEmployee = arrayNotes(7, i)
    txtSubmitDate = arrayNotes(8, i)
    txtRecordCounter = recordCounter + 1

End Sub

Public Sub clearBoxes()

    txtERRCaseID.Value = Null
    txtInventoryNumber = Null
    txtBatchNumber = Null
    txtNote = Null
    txtFFMExchangeAssignedID = Null
    txtFFMPolicyID = Null
    txtEmployee = Null
    txtSubmitDate = Null

End Sub

