VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm1095ASearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdSearch_Click()
Dim strtab As String: strtab = "     "
Dim strResults As String: strResults = "MMBR Number" & strtab & " EX SUB NO  " & strtab & "    EX MMB NO" & strtab & "    last name" & strtab & "first name" & strtab & "date of birth     " & strtab & "SSN"
Dim db As Database
Dim rs As Recordset
Set db = CurrentDb
If txtSearch <> "" Or Not IsNull(txtSearch.Value) Then
    Select Case grpSearch
    Case 1: Set rs = db.OpenRecordset("Select * From tbl1095Disputes Where exchange_subscriber_number = " & "'" & txtSearch.Value & "';", dbOpenSnapshot)
    Case 2: Set rs = db.OpenRecordset("Select * From tbl1095Disputes Where exchange_member_number = " & "'" & txtSearch.Value & "';", dbOpenSnapshot)
    Case 3: Set rs = db.OpenRecordset("Select * From tbl1095Disputes Where social_security_number = " & "'" & txtSearch.Value & "';", dbOpenSnapshot)
    Case 4: Set rs = db.OpenRecordset("Select * From tbl1095Disputes Where member_number = " & "'" & txtSearch.Value & "';", dbOpenSnapshot)
    End Select
    If Not rs.EOF Then
        Do Until rs.EOF = True
            strResults = strResults & vbCrLf & rs![member_number] & strtab & strtab & rs![exchange_subscriber_number] & "         " & rs![exchange_member_number] & _
                        "         " & rs![member_last_name] & "       " & rs![member_first_name] & "            " & rs![date_of_birth] & "         " & rs![social_security_number]
             rs.MoveNext
        
        Loop
    Else: rs.Close: Set rs = Nothing: db.Close: Set db = Nothing: MsgBox " no Results"
    End If
    
    txtResults.Value = strResults
Else: MsgBox "You must provide a search value": db.Close: Set db = Nothing
    
End If

End Sub
