VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmPAF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public varPAFArray As Variant
Public currentIndex As Integer
Public numRecords As Integer

'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'form for searching and viewing records on the CMS preaudit file (PAF)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------


Private Sub cmdNext_Click()
'next record
Dim strSearch As String
Dim strRecordCount As String
strSearch = txtSearch.Value
strRecordCount = txtRecordCount.Value
    If currentIndex + 1 >= numRecords Then
        MsgBox "No more Records"
        Exit Sub
    Else: currentIndex = currentIndex + 1
        ClearForm
        txtSearch.Value = strSearch
        txtRecordCount.Value = strRecordCount
        For i = 1 To 104
        Controls("txt" & i).Value = varPAFArray(i - 1, currentIndex)
        Next i
          txt10.Value = txt10.Value & "-" & txt94.Value
    DateFormat txt63: DateFormat txt64
    DateFormat txt77: DateFormat txt78
    DateFormat txt80: DateFormat txt81
    DateFormat txt83: DateFormat txt84
    DateFormat txt8
    TimeFormat txt9
    End If
   
End Sub

Private Sub cmdPrev_Click()
'previous record
Dim strSearch As String
Dim strRecordCount As String
strSearch = txtSearch.Value
strRecordCount = txtRecordCount.Value

    If currentIndex <= 0 Then
        MsgBox "No more records"
        Exit Sub
    Else: currentIndex = currentIndex - 1
    ClearForm
    txtSearch.Value = strSearch
    txtRecordCount.Value = strRecordCount
     For i = 1 To 104
        Controls("txt" & i).Value = varPAFArray(i - 1, currentIndex)
        Next i
    txt10.Value = txt10.Value & "-" & txt94.Value
    DateFormat txt63: DateFormat txt64
    DateFormat txt77: DateFormat txt78
    DateFormat txt80: DateFormat txt81
    DateFormat txt83: DateFormat txt84
    DateFormat txt8
    TimeFormat txt9
    End If
   
End Sub

Private Sub Command1107_Click()
'search button
Dim db As Database
Dim rs As Recordset
Dim strFile As String

On Error GoTo err1:

cleanText txtSearch
currentIndex = 0
strSearch = txtSearch.Value
ClearForm
txtSearch.Value = strSearch
Set db = CurrentDb

'table names
    Select Case grpYear.Value
        Case 1: strFile = "PAF17"
        Case 2: strFile = "PAF18"
        Case 3: strFile = "BARCN"
        Case 4: strFile = "ISIPA18"
        Case 5: strFile = "PAF16"
        Case 6: strFile = "ISIPA16"
        Case 7: strFile = "PAF15"
        Case 8: strFile = "MISC8DOB_PREM"
        Case 9: strFile = "PAF19"
        Case 10: strFile = "ISIPA17"
        Case 11: strFile = "PAF20"
        
        Case Else: MsgBox "Please Select a Year": Exit Sub
        
    End Select
   
'sets criteria selected radio in the search optgroup
    Select Case grpSearchBy.Value
        Case 1: Set rs = db.OpenRecordset("Select * From " & strFile & " Where [Exchange-Assigned Subscriber ID] = " & "'" & txtSearch.Value & "'", dbOpenSnapshot)
        Case 2: Set rs = db.OpenRecordset("Select * From " & strFile & " Where [Exchange-Assigned Member ID] = " & "'" & txtSearch.Value & "'", dbOpenSnapshot)
        Case 3: Set rs = db.OpenRecordset("Select * From " & strFile & " Where [Issuer-Assigned Subscriber ID] = " & "'" & txtSearch.Value & "'", dbOpenSnapshot)
        Case 4: Set rs = db.OpenRecordset("Select * From " & strFile & " Where [Issuer-Assigned Member ID] = " & "'" & txtSearch.Value & "'", dbOpenSnapshot)
        Case 5: Set rs = db.OpenRecordset("Select * From " & strFile & " Where [Application ID]  = " & "'" & txtSearch.Value & "'", dbOpenSnapshot)
        Case 6: Set rs = db.OpenRecordset("Select * From " & strFile & " Where [Social Security Number (SSN)] = " & "'" & txtSearch.Value & "'", dbOpenSnapshot)
        Case 7: Set rs = db.OpenRecordset("Select * From " & strFile & " Where [Issuer-Assigned Policy Number] = " & "'" & txtSearch.Value & "'", dbOpenSnapshot)
        Case 8: Set rs = db.OpenRecordset("Select * From " & strFile & " Where [Exchange-Assigned Policy Number] = " & "'" & txtSearch.Value & "'", dbOpenSnapshot)
        Case 9: Set rs = db.OpenRecordset("Select * From " & strFile & " Where [QI Last Name] Like " & "'*" & txtSearch.Value & "*'", dbOpenSnapshot)
        
    End Select
    
'if records > 0, save records into varPAFArray
    If rs.recordCount > 0 Then
        
        rs.MoveFirst
        rs.MoveLast
        rs.MoveFirst
        varPAFArray = rs.GetRows(rs.recordCount)
        numRecords = rs.recordCount
        txtRecordCount.Value = rs.recordCount
    Else: MsgBox "No Records Found": rs.Close: Set rs = Nothing: Exit Sub
    End If
    
'get column count of table
    Dim tbl As DAO.TableDef
    Dim colCount As Integer
    Set tbl = db.TableDefs(strFile)
    colCount = tbl.Fields.Count
    Set tbl = Nothing
    
'populates txtBoxes to value within varPAFArray, depending on array index #
    For i = 1 To colCount
        Controls("txt" & i).Value = varPAFArray(i - 1, currentIndex)
        Next i
    rs.Close: Set rs = Nothing
  
'reformat dates/times
    txt10.Value = txt10.Value & "-" & txt94.Value
    DateFormat txt6
    DateFormat txt36
    DateFormat txt63: DateFormat txt64
    DateFormat txt71: DateFormat txt72
    DateFormat txt74: DateFormat txt75
    DateFormat txt77: DateFormat txt78
    DateFormat txt80: DateFormat txt81
    DateFormat txt83: DateFormat txt84
    DateFormat txt8
    
   
    TimeFormat txt7
    TimeFormat txt9
    If txt38.Value = 2 Then
        txt38.Value = "N"
    ElseIf txt38.Value = 1 Then
        txt38.Value = "T"
    End If
    
'relationship codes
    Set rs = db.OpenRecordset("tblRelCodes", dbOpenSnapshot)
    rs.MoveFirst
    rs.FindFirst ("[ASC X12 Value] = " & "'" & Right(txt17, 2)) & "'"
    txt17.ControlTipText = rs![Description]
    rs.Close: Set rs = Nothing
    db.Close
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmPAF; Command1107_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub


Sub DateFormat(ctrl As Control)
    ctrl.Value = Mid(ctrl.Value, 5, 2) & "/" & Right(ctrl.Value, 2) & "/" & Left(ctrl.Value, 4)
    
End Sub


Sub ClearForm()
Dim ctrl As Control
    For Each ctrl In Controls
        If TypeOf ctrl Is TextBox Then
            ctrl.Value = ""
        End If
        Next
        
End Sub
Sub TimeFormat(ctrl As Control)
    ctrl.Value = Left(ctrl.Value, 2) & ":" & Mid(ctrl.Value, 3, 2) & ":" & Mid(ctrl.Value, 6, 2) & ":" & Right(ctrl.Value, 3)
End Sub

Sub cleanText(ctrl As Control)
    If ctrl.Value <> "" And Not IsNull(ctrl.Value) Then
    
        ctrl.Value = Replace(ctrl.Value, "-", "")
        ctrl.Value = Trim(ctrl.Value)
    End If
    
    
End Sub

'if other form links to PAF form, passes value into openArgs
'9 is 2019
Private Sub Form_Open(Cancel As Integer)
Dim strRefSubscriberID As String
    If Not IsNull(Forms!frmPAF.OpenArgs) Then
    
    strRefSubscriberID = Forms!frmPAF.OpenArgs
    If Len(strRefSubscriberID) > 0 Then
        strRefSubscriberID = Left(strRefSubscriberID, 9)
        txtSearch.Value = strRefSubscriberID
        '9 is 2019
        '2 is 2018
        grpYear.Value = 9
        Call grpYear_Click
        grpSearchBy.Value = 3
        Command1107_Click
    End If
    End If
End Sub
'sets back color by year
Private Sub grpYear_Click()
    Select Case grpYear.Value
    Case 1: Detail.BackColor = RGB(136, 176, 75): Label2041.BackColor = RGB(136, 176, 75):  Label2169.BackColor = RGB(136, 176, 75)
    Case 2: Detail.BackColor = RGB(95, 75, 139): Label2041.BackColor = RGB(95, 75, 139): Label2169.BackColor = RGB(95, 75, 139)
    Case 3: Detail.BackColor = RGB(175, 175, 175): Label2041.BackColor = RGB(175, 175, 175): Label2169.BackColor = RGB(175, 175, 175)
    Case 5: Detail.BackColor = RGB(257, 202, 201): Label2041.BackColor = RGB(257, 202, 201): Label2169.BackColor = RGB(257, 202, 201)
    Case 7: Detail.BackColor = RGB(150, 79, 76): Label2041.BackColor = RGB(150, 79, 76): Label2169.BackColor = RGB(150, 79, 76)
    Case 9: Detail.BackColor = RGB(60, 80, 79): Label2041.BackColor = RGB(60, 80, 79): Label2169.BackColor = RGB(60, 80, 79)
    Case 11: Detail.BackColor = RGB(188, 169, 225): Label2041.BackColor = RGB(188, 169, 225): Label2169.BackColor = RGB(188, 169, 225)
    End Select
    
End Sub

