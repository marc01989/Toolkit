VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmHICSUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'where extracts of casework notes and cutlog notes are generated/formatted for upload into hics.cms.gov
'-----------------------------------------------------------------------------------------------------------------------------------------------------------


Private Sub cmdCompleteScans_Click()
    DoCmd.OpenQuery ("qupdScansCompleted")
End Sub

Private Sub cmdGenerateScans_Click()
    'EXPORT PATH
    Dim oXLSXSTRS As String
    oXLSXSTRS = "X:\Member Enrollment\Member Enrollment(Custom)\MarketplaceDatabaseProjects\HICS DBs\UPLOADS\" & "SCANS_UPLOAD_" & Format(Now(), "mm_dd_yyyy_hhnnss")
    
    'grab scans from 'results' excel sheet, append to tblScans
    DoCmd.OpenQuery ("qappScans")
    'use join to append Member ID to new records on scans table where match on Case ID
    DoCmd.OpenQuery ("qappScans_IDs")
    'export the query that has scans not uploaded
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "qryScansUpload", oXLSXSTRS
        
    If MsgBox("Go to HICS Upload folder?", vbYesNo) = vbYes Then
        Shell "explorer.exe " & "X:\Member Enrollment\Member Enrollment(Custom)\MarketplaceDatabaseProjects\HICS DBs\UPLOADS\", vbNormalFocus
    Else: Exit Sub
    End If
        
End Sub

'------------------------------------------------------------
' Generate HICS and Cutlog UploadS
'------------------------------------------------------------
Private Sub Command0_Click()
On Error GoTo Command0_Click_Err
Dim oXLSXSTR As String

'FILE PATHS AND FILE NAMES
oXLSXSTRY = "X:\Member Enrollment\Member Enrollment(Custom)\MarketplaceDatabaseProjects\HICS DBs\UPLOADS\" & "HICS_UPLOAD_RESOLVED_" & Format(Now(), "mm_dd_yyyy_hhnnss")
oXLSXSTRN = "X:\Member Enrollment\Member Enrollment(Custom)\MarketplaceDatabaseProjects\HICS DBs\UPLOADS\" & "HICS_UPLOAD_UNRESOLVED_" & Format(Now(), "mm_dd_yyyy_hhnnss")
oXLSXSTRC = "X:\Member Enrollment\Member Enrollment(Custom)\MarketplaceDatabaseProjects\HICS DBs\UPLOADS\" & "CUTLOG_UPLOAD_" & Format(Now(), "mm_dd_yyyy_hhnnss")
oXLSXSTRE = "X:\Member Enrollment\Member Enrollment(Custom)\MarketplaceDatabaseProjects\HICS DBs\UPLOADS\" & "ERR_UPLOAD_" & Format(Now(), "mm_dd_yyyy_hhnnss")

'EXPORT THE QUERIES AS XLSX FILES (EXCEL)
DoCmd.TransferSpreadsheet acExport, 9, "qryHICSUploadResolved", oXLSXSTRY
DoCmd.TransferSpreadsheet acExport, 9, "qryHICSUploadUnresolved", oXLSXSTRN
DoCmd.TransferSpreadsheet acExport, 9, "qryCutlogUpload", oXLSXSTRC
DoCmd.TransferSpreadsheet acExport, 9, "qryERRUpload", oXLSXSTRE

'MARK ALL RECORDS AS [extracted] = 'Y'
DoCmd.SetWarnings False
DoCmd.runSQL ("UPDATE tblCaseWorkNotes SET extracted = 'Y' WHERE extracted = 'N' AND is_res = 'Y';")
DoCmd.runSQL ("UPDATE tblCaseWorkNotes SET extracted = 'Y' WHERE extracted = 'N' AND is_res = 'N';")
DoCmd.runSQL ("UPDATE tblCutlogNotes SET extracted = 'Y' WHERE extracted = 'N';")
DoCmd.runSQL ("UPDATE tblERRNotes SET extracted = 'Y' WHERE extracted = 'N';")

If MsgBox("Go to HICS Upload folder?", vbYesNo) = vbYes Then
    Shell "explorer.exe " & "X:\Member Enrollment\Member Enrollment(Custom)\MarketplaceDatabaseProjects\HICS DBs\UPLOADS\", vbNormalFocus
End If


Command0_Click_Exit:
    Exit Sub

Command0_Click_Err:
    MsgBox Error$
    Resume Command0_Click_Exit

End Sub

'------------------------------------------------------------
'complete uploads button
'user must click this after uploading each set of notes into respective system
'------------------------------------------------------------
Private Sub Command5_Click()
On Error GoTo Command5_Click_Err

    DoCmd.SetWarnings True
    Select Case optCompleteUpload
        Case 1 'unresolved
            DoCmd.OpenQuery "qupdCaseworkNotes", acViewNormal, acEdit
            MsgBox ("Complete")
        Case 2 'resolved
            DoCmd.OpenQuery "qupdResolutionNotes", acViewNormal, acEdit
            MsgBox ("Complete")
        Case 3 'ERR
            DoCmd.OpenQuery "qupdERRUpload", acViewNormal, acEdit
            MsgBox ("Complete")
        Case 4 'cutlogs
            DoCmd.OpenQuery "qupdCutlogNotes", acViewNormal, acEdit
            MsgBox ("Complete")
    End Select

Command5_Click_Exit:
    Exit Sub

Command5_Click_Err:
    MsgBox Error$
    Resume Command5_Click_Exit

End Sub

Private Sub Form_Unload(Cancel As Integer)

'------------------------------------------------------------------------------------------------------
    'CHECK THAT EACH EXTRACTED RECORD IS MARKED [uploaded] = 'Y'. IF NOT, PROMPT USER
'------------------------------------------------------------------------------------------------------

    Dim msgStr As String
    If DCount("case_id", "tblCaseWorkNotes", "extracted = 'Y' AND uploaded = 'N' AND is_res = 'N'") > 0 Then
        msgStr = "   Unresolved Notes" & vbCrLf
    End If
    If DCount("case_id", "tblCaseWorkNotes", "extracted = 'Y' AND uploaded = 'N' AND is_res = 'Y'") > 0 Then
        msgStr = msgStr & "   Resolved Notes" & vbCrLf
    End If
    If DCount("CaseID", "tblERRNotes", "extracted = 'Y' AND Uploaded = 0") > 0 Then
        msgStr = msgStr & "   ER&R Notes" & vbCrLf
    End If
    If DCount("case_id", "tblCutlogNotes", "extracted = 'Y' AND uploaded = 'N'") > 0 Then
        msgStr = msgStr & "   Cutlog Notes" & vbCrLf
    End If
    
    If msgStr <> "" Then
        If MsgBox("Attention: the following uploads have not been marked complete:" & _
        vbCrLf & vbCrLf & msgStr & vbCrLf & "Please mark complete before exiting.", vbExclamation + vbOKCancel, "Warning") = vbOK Then
            Cancel = True
        End If
    Else
        'do nothing - allow form to close
    End If
End Sub
