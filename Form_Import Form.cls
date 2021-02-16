VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Import Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim onDefaultDateTime As String
Dim offDefaultDateTime As String
Dim ONFILEPATH As String
Dim OFFFILEPATH As String

Private Sub Form_Load()

On Error GoTo resetGlobal:

'--- set constants to filepaths of respective files for use throughout module
ONFILEPATH = "\\acct.upmchs.net\hp\EDI\EDI_Shared\UPMPRD_Batch834\HWEGIP\ACK997\HIX834IND-HWEGIP.txt"
OFFFILEPATH = "\\acct.upmchs.net\hp\EDI\EDI_Shared\UPMPRD_Batch834\HWEGIP\ACK997\HPOFFEX834-HWEGIP.txt"

'--- set global variables to default date/time of each file
onDefaultDateTime = FileDateTime(ONFILEPATH)
offDefaultDateTime = FileDateTime(OFFFILEPATH)

resetGlobal:
    Select Case Err.Number
    Case 0
    Case Else
        Call LogError(Err.Number & " " & Err.Description, "Form_Load()")
        onDefaultDateTime = ""
        offDefaultDateTime = ""
        ONFILEPATH = ""
        OFFFILEPATH = ""
        Call Form_Load
    End Select

End Sub

Private Sub Form_Timer()

'------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'timed subroutine will test the below if-statement every 30,000 miliseconds (30 seconds)
'UPDATE 9/27/2018 - TIMER INTERVAL UPPED TO 60,000 MILISECONDS (1 MINS) TO ACCOUNT FOR LARGE BAR FILE IMPORT
'UPDATE 11/8/2018 - TIMER INTERVAL LOWERED TO 15,000 MILISECONDS (15 SECS) TO TRY AND CAPTURE SEPARATE 2018 AND 2018 OFF GIPS, WHICH DROP AT SAME TIME
'file dateTime of GIP txt file is saved as global variable.
'timed sub checks if file defaultDateTime has changed in the past 1.6 mins, and imports the new file if it has changed
'UPDATE 8/12/209 - TIMER INTERVAL CHANGED BACK TO 60,000 (1 MINS).
'CONT'D: ADDED ERROR HANDLER TO Form_load() SUBROUTINE TO RESET ALL GLOBAL VALUES AND RERUN THE PROCESS IN THE EVENT OF VBA ERROR BEING THROWN
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Dim newDateTime As String, timeStamp As String, filePath As String
Dim i As Integer

On Error GoTo resetGlobal

For i = 1 To 2

    Select Case i

    Case 1
        newDateTime = FileDateTime(ONFILEPATH)

        If onDefaultDateTime <> newDateTime Then
            importFile (i)
            onDefaultDateTime = FileDateTime(ONFILEPATH)
        End If
        
    Case 2
        newDateTime = FileDateTime(OFFFILEPATH)
        
        If offDefaultDateTime <> newDateTime Then
            importFile (i)
            offDefaultDateTime = FileDateTime(OFFFILEPATH)
        End If
        
    End Select

Next i

Exit Sub


resetGlobal:
    Select Case Err.Number
    Case 0
    Case Else
        Call LogError(Err.Number & " " & Err.Description, "Form_Timer()")
        Call Form_Load
    End Select
    
End Sub

Private Sub importFile(fileType As Integer)

Dim filePath As String, tableName As String, importSpec As String, exportPath As String

'check which file type is being processed and assign appropriate variables
If fileType = 1 Then        '----- ON -------
    filePath = ONFILEPATH
    tableName = "HIX834IND-HWEGIP"
    importSpec = "gipSpec2"
    exportPath = "X:\Member Enrollment\Member Enrollment(Custom)\Marketplace\834 Checklists and File Reports\834 Database\Archived 834 Files\On\"
ElseIf fileType = 2 Then    '----- OFF ------
    filePath = OFFFILEPATH
    tableName = "HPOFFEX834-HWEGIP"
    importSpec = "gipSpec3"
    exportPath = "X:\Member Enrollment\Member Enrollment(Custom)\Marketplace\834 Checklists and File Reports\834 Database\Archived 834 Files\Off\"
End If

DoCmd.SetWarnings False

On Error GoTo resetGlobal

'delete temp table, populate, export excel workbook
DoCmd.runSQL ("DELETE * FROM [" & tableName & "_temp];")
DoCmd.TransferText acImportFixed, importSpec, tableName & "_temp", filePath
DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, tableName & "_temp", exportPath & tableName & "_" & Format(Now(), "yyyymmddss")

'add to archive table and insert date
DoCmd.TransferText acImportFixed, importSpec, tableName, filePath
DoCmd.runSQL ("UPDATE [" & tableName & "] SET FileDate = '" & Format(Now(), "yyyymmddhhmmss") & "' WHERE FileDate IS NULL;")

'run queries on the latest file imported to the temp table
runQueries (fileType)

Exit Sub


resetGlobal:
    Call LogError(Err.Number & " " & Err.Description, "importFile()")
    Call Form_Load

End Sub

Private Sub runQueries(fileType As Integer)

Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim qd As DAO.QueryDef
Dim onQueries(0 To 5) As Variant
'onQueries(0) = "Add and Cancel"
'onQueries(1) = "Appeals"
'onQueries(2) = "Broker Update"
'onQueries(3) = "Duplicate EMP Count"
'onQueries(4) = "Name Change"
'onQueries(5) = "SSN Change"
'onQueries(6) = "Broker Change"
'onQueries(7) = "Broker Change Full"
onQueries(0) = "Broker Change_19"
onQueries(1) = "Broker Change_18"
onQueries(2) = "Broker Change_Full"
onQueries(3) = "qryAppID_Mismath19"
onQueries(4) = "qryPolicyID_Mismatch19"
onQueries(5) = "qryStartDate_MulipleSpans"
Dim offQueries(0 To 1) As Variant
offQueries(0) = "OFF Child Only"
offQueries(1) = "OFF Test Members"

'---cycle through all saved queries and export them to the same excel workbook, tabbed out
timeStamp = Format(Now(), "yyyymmddss")
Set db = CurrentDb

Select Case fileType
Case 1 '------------------------------ON exchange queries--------------------
    For i = 0 To 5
        Set rs = db.OpenRecordset(onQueries(i))
        
        If rs.recordCount > 0 Then
        
            Set qd = db.QueryDefs(onQueries(i))
            With qd
                .ReturnsRecords = True
            End With
            
            DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, onQueries(i), "X:\Member Enrollment\Member Enrollment(Custom)\Marketplace\834 Checklists and File Reports\834 Database\Reports\ON 834 Reports_" & timeStamp, 1
        End If
    Next i
Case 2 '------------------------------OFF exchange queries--------------------
'    For i = 0 To 1
'        Set rs = db.OpenRecordset(offQueries(i))
'
'        If rs.RecordCount > 0 Then
'
'            Set qd = db.QueryDefs(offQueries(i))
'            With qd
'                .ReturnsRecords = True
'            End With
'
'            DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, offQueries(i), "X:\Member Enrollment\Member Enrollment(Custom)\MarketplaceDatabaseProjects\834\Test\834 Review\OFF 834 Review_" & timeStamp, 1
'        End If
'    Next i
End Select

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Manual import subroutines, on-click controls on from
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Private Sub btnImport_Click()

Dim filePath As String, tableName As String, importSpec As String, exportPath As String

If optLOB.Value = 1 Then        '----- ON --------
    filePath = "\\acct.upmchs.net\hp\EDI\EDI_Shared\UPMPRD_Batch834\HWEGIP\ACK997\HIX834IND-HWEGIP.txt"
    tableName = "HIX834IND-HWEGIP"
    importSpec = "gipSpec2"
    exportPath = "X:\Member Enrollment\Member Enrollment(Custom)\Marketplace\834 Checklists and File Reports\834 Database\Archived 834 Files\On\"
ElseIf optLOB.Value = 2 Then    '----- OFF -------
    filePath = "\\acct.upmchs.net\hp\EDI\EDI_Shared\UPMPRD_Batch834\HWEGIP\ACK997\HPOFFEX834-HWEGIP.txt"
    tableName = "HPOFFEX834-HWEGIP"
    importSpec = "gipSpec3"
    exportPath = "X:\Member Enrollment\Member Enrollment(Custom)\Marketplace\834 Checklists and File Reports\834 Database\Archived 834 Files\Off\"
End If

DoCmd.SetWarnings False
If MsgBox("Would you like to import the 834 file?", vbExclamation + vbYesNo) = vbYes Then

    'delete temp table, populate, export excel workbook
    DoCmd.runSQL ("DELETE * FROM [" & tableName & "_temp];")
    DoCmd.TransferText acImportFixed, importSpec, tableName & "_temp", filePath
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, tableName & "_temp", exportPath & tableName & "_" & Format(Now(), "yyyymmddss")
    
    'add to archive table and insert date
    DoCmd.TransferText acImportFixed, importSpec, tableName, filePath
    DoCmd.runSQL ("UPDATE [" & tableName & "] SET FileDate = '" & Format(Now(), "yyyymmddhhmmss") & "' WHERE FileDate IS NULL;")
        
runQueries (optLOB.Value)
        
Else
    End
End If

End Sub

Private Sub btnGenerateReport_Click()

If optLOB.Value = 1 Then
    Call runQueries(1)
ElseIf optLOB.Value = 2 Then
    Call runQueries(2)
End If

If MsgBox("Report successfully exported. Go to folder?", vbYesNo, "Success") = vbYes Then
    Shell "explorer.exe " & "X:\Member Enrollment\Member Enrollment(Custom)\MarketplaceDatabaseProjects\834\Test\834 Review\", vbNormalFocus
    End
Else
    End
End If

End Sub

