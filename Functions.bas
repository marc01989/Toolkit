Attribute VB_Name = "Functions"
Option Compare Database

Public Sub LogError(strError, modName As String)
'list of errors
'https://msdn.microsoft.com/en-us/library/bb221208(v=office.12).aspx

    Dim strPath As String, comp As String
    Dim fs As Object
    Dim a As Object

    comp = Environ$("username")
    strPath = "X:\Member Enrollment\Member Enrollment(Custom)\Marketplace\Database\db utilities\errors"

    Set fs = CreateObject("Scripting.FileSystemObject")
        If fs.FileExists(strPath & "\errorLogQA.txt") = True Then
            Set a = fs.Opentextfile(strPath & "\errorLogQA.txt", 8)
        Else
            Set a = fs.createtextfile(strPath & "\errorLogQA.txt")
        End If
    
        a.writeline Date + Time & "|ERROR: " & strError & "|USER: " & comp & "|MODULE: " & modName
        a.Close
    Set fs = Nothing
End Sub
Public Function getCurrentWeekId()
    Dim today As String
    Dim wkd As Integer
    today = Date
    wkd = Weekday(Date)
    wkd = wkd - 1
    today = Date - wkd
    today = Format(today, "mmmm d, yyyy")
    getCurrentWeekId = DLookup("[week_id]", "[weeks]", "[week_start] = '" & today & "'")
End Function
Public Function ScrubMemberId(memberId As String) As String
    memberId = Replace(memberId, "-", "")
    memberId = Replace(memberId, "_", "")
    memberId = Replace(memberId, " ", "")
    memberId = Trim(memberId)
    ScrubMemberId = memberId
End Function
Public Sub LogUserOff(formModule As String)
    Dim obj As AccessObject, dbs As Object
    
    MsgBox ("User not found - please login again")
    Set dbs = Application.CurrentProject
    For Each obj In dbs.AllForms
        If obj.IsLoaded = True Then
          DoCmd.Close acForm, obj.Name, acSaveNo
        End If
    Next obj
    DoCmd.OpenForm "Login", acNormal, , , , acWindowNormal
    Call LogError(0 & " " & "User Id not found or not passed to home screen", formModule)
End Sub

Public Sub GetScore(employeeId As Integer, weekId As Integer, processType As Integer)

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sumPointVal As Double, distinctItems As Double
    Dim processes(0 To 8) As Variant
    Dim criteria As String
    processes(0) = "hics"
    processes(1) = "on_term_job"
    processes(2) = "off_term_job"
    processes(3) = "cancel_job"
    processes(4) = "other"
    processes(5) = "recon"
    processes(6) = "cutlog_demo_changes"
    processes(7) = "chat"
    processes(8) = "external_qa"
    
    
On Error GoTo err1:
    
    '--get sum of error point values
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT SUM(point_value) " & _
    " FROM errors INNER JOIN review_items_mkt ON errors.error_id = review_items_mkt.error_id " & _
    " WHERE (week_id = " & weekId & " AND employee_id = " & employeeId & " AND  process_id = " & processType & ");")
        With rs
            If .recordCount > 0 Then
                If Not IsNull(.Fields(0)) Then
                    sumPointVal = .Fields(0)
                End If
            End If
        End With
        rs.Close
    
    '--get distinct items per process
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM " & _
    " (SELECT DISTINCT member_id FROM review_items_mkt " & _
    " WHERE (week_id = " & weekId & " AND employee_id = " & employeeId & " AND  process_id = " & processType & "));")
        With rs
            If .recordCount > 0 Then
                If Not IsNull(.Fields(0)) Then
                    distinctItems = .Fields(0)
                End If
            End If
        End With
        rs.Close
        
        Set rs = db.OpenRecordset("scores_mkt")
            With rs
                .FindFirst ("[week_id] = " & weekId & " AND [employee_id] = " & employeeId)
                    If .NoMatch Then
                        .AddNew
                            rs.Fields(processes(processType)).Value = calculateScore(sumPointVal, distinctItems, processType)
                            rs![employee_id] = employeeId
                            rs![week_id] = weekId
                            rs![submit_date] = Now()
                        .Update
                    Else
                        .Edit
                            rs.Fields(processes(processType)).Value = calculateScore(sumPointVal, distinctItems, processType)
                            rs![submit_date] = Now()
                        .Update
                    End If
            End With
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
        

err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Functions; GetScore()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select

End Sub
Public Function Get6MonthScore(employeeId As Integer, weekId As Integer) As Variant

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sumPointVal As Double, distinctItems As Double
    Dim i As Integer
    Dim processes(0 To 8) As Variant
    Dim criteria As String
'    processes(0) = "hics"
'    processes(1) = "on_term_job"
'    processes(2) = "off_term_job"
'    processes(3) = "cancel_job"
'    processes(4) = "other"
'    processes(5) = "recon"
'    processes(6) = "cutlog_demo_changes"
'    processes(7) = "chat"
'    processes(8) = "external_qa"
On Error GoTo err1:

    
    For i = 0 To 8
    
        '--get sum of error point values
        Set db = CurrentDb
        
        Set rs = db.OpenRecordset("SELECT SUM(point_value) " & _
        " FROM errors INNER JOIN review_items_mkt ON errors.error_id = review_items_mkt.error_id " & _
        " WHERE (week_id >= " & weekId - 26 & " AND week_id <= " & weekId & " AND employee_id = " & employeeId & " AND  process_id = " & i & ");")
            With rs
                If .recordCount > 0 Then
                    If Not IsNull(.Fields(0)) Then
                        sumPointVal = .Fields(0)
                    End If
                End If
            End With
            rs.Close
        
        '--get distinct items per process
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM " & _
        " (SELECT DISTINCT member_id FROM review_items_mkt " & _
        " WHERE (week_id >= " & weekId - 26 & " AND week_id <= " & weekId & " AND employee_id = " & employeeId & " AND  process_id = " & i & "));")
            With rs
                If .recordCount > 0 Then
                    If Not IsNull(.Fields(0)) Then
                        distinctItems = .Fields(0)
                    End If
                End If
            End With
            rs.Close
            
            processes(i) = calculateScore(sumPointVal, distinctItems, i)

    Next i
    
    Set rs = Nothing
    db.Close: Set db = Nothing
        
    Get6MonthScore = processes
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Functions; Get6MonthScore()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Function
            Exit Function
    End Select

End Function
Function calculateScore(sumPointVal As Double, distinctItems As Double, processType As Integer) As Variant

    Dim score As Double
        Select Case processType
            Case 0
                If distinctItems = 0 Then
                    calculateScore = Null
                Else
                    score = ((52 * distinctItems) - sumPointVal)
                    totes = (52 * distinctItems)
                    calculateScore = Round(score / totes, 4) * 100
                End If
            Case 1 To 8
                If distinctItems = 0 Then
                    calculateScore = Null
                Else
                    score = ((38 * distinctItems) - sumPointVal)
                    totes = (38 * distinctItems)
                    calculateScore = Round(score / totes, 4) * 100
                End If
        End Select

End Function
