VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmERR_Reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'this section shows users what outstanding disputes need exported for CMS contractor upload
'all disputes are exported in pipe-delimited format, as required by CMS
'queries grab disputes that have [Extracted] = 'N' and append to a transition table that is formatted in with correct naming convention and order of headers. then
'an import spec is used to export the results in pipe-delimited format
'-----------------------------------------------------------------------------------------------------------------------------------------------------------


' users selects which report they want  o generate for ERR and this clears the relevant extract table then appends all records that need to be exported and updates the table that holds all disputes to show that they have been extracted
Private Sub cmdGenerateReport_Click()

Dim cuFilePath As String
cuFilePath = "X:\Member Enrollment\Member Enrollment(Custom)\MarketplaceDatabaseProjects\RCNO DB\Current Year Disputes\95216.COG.ERRS.D" & Format(Now(), "yymmdd") & ".T" & Format(Now(), "hhnn")

DoCmd.SetWarnings False
    Select Case grpERRDisputes.Value
    Case 1  'discrepancy
        If DCount("*", "qryDiscDisputeExtract") <> 0 Then
        
            '-------------------------------------------------------
            'REPEAT THE BELOW STEPS FOR EACH OF THE 7 TYPES OF DISPUTES
            '-------------------------------------------------------
            
            'delete old records from the transition table
            DoCmd.runSQL "Delete * From tblDiscrepancyDisputeExtract"
            
            'run insert query to append disputes to transition table. only dipsutes with [Extracted] = 'N' are appended
            DoCmd.OpenQuery "qappDiscrepancyDisputeExtract", acViewNormal
            
            'use import/export spec to export the table in pipe-delimited format
            DoCmd.TransferText acExportDelim, "DiscDisputeExtract", "tblDiscrepancyDisputeExtract", cuFilePath & "11111.P.txt", 1
            
            'mark the disputes as [Extracted] = 'Y' along with timestamp
            DoCmd.runSQL "UPDATE tblDiscrepancyDispute SET tblDiscrepancyDispute.Extracted = 'Y', tblDiscrepancyDispute.Extract_Date = Now() WHERE (tblDiscrepancyDispute.Extracted = 'N');"
        End If
           
    Case 2  'reinstatement
        If DCount("*", "qryReinstatementEndDateExtract") <> 0 Then
            DoCmd.runSQL "Delete * From tblReinstatementEndDateExtract"
            DoCmd.OpenQuery "qappReinstatementEndDate", acViewNormal
        
            DoCmd.TransferText acExportDelim, "ReinstatementEndDateExtract", "tblReinstatementEndDateExtract", cuFilePath & "22222.P.txt", 1
            DoCmd.runSQL "UPDATE tblReinstatementEndDate SET tblReinstatementEndDate.Extracted = 'Y', tblReinstatementEndDate.Extract_Date = Now() WHERE (tblReinstatementEndDate.Extracted = 'N');"
        End If
        
        
    Case 3  'Newborn
        If DCount("*", "qryNewbornExtract") <> 0 Then
            DoCmd.runSQL "Delete * From tblNewbornPremiumUpdatesExtract"
            DoCmd.OpenQuery "qappNewbornExtract", acViewNormal
        
            DoCmd.TransferText acExportDelim, "NewbornExtract", "tblNewbornPremiumUpdatesExtract", cuFilePath & "33333.P.txt", 1
            DoCmd.runSQL "UPDATE tblNewbornPremiumUpdates SET tblNewbornPremiumUpdates.Extracted = 'Y', tblNewbornPremiumUpdates.extract_date = Now() WHERE (tblNewbornPremiumUpdates.Extracted = 'N' Or tblNewbornPremiumUpdates.Extracted = '');"
        End If
        
    
    Case 4  'Enrollment blocker
        If DCount("*", "qryEnrollmentBlockerExtract") <> 0 Then
            DoCmd.runSQL "Delete * From tblEnrollmentBlockerExtract"
            DoCmd.OpenQuery "qapEnrollmentBlockerExtract", acViewNormal
            
            DoCmd.TransferText acExportDelim, "EnrollmentExtract", "tblEnrollmentBlockerExtract", cuFilePath & "44444.P.txt", 1
            DoCmd.runSQL "UPDATE tblEnrollmentBlocker SET tblEnrollmentBlocker.Extracted = 'Y'  WHERE (tblEnrollmentBlocker.Extracted = 'N' Or tblEnrollmentBlocker.Extracted = '');"
        End If
        

    Case 5  'Rejected enrollments
        If DCount("*", "qryRejectedEnrollmentsExtract") <> 0 Then
            DoCmd.runSQL "Delete * From tblRejectedEnrollmentsExtract"
            DoCmd.OpenQuery "qappRejectedEnrollmentsExtract", acViewNormal
        
            DoCmd.TransferText acExportDelim, "RejectedEnrollmentsExtract", "tblRejectedEnrollmentsExtract", cuFilePath & "55555.P.txt", 1
            DoCmd.runSQL "UPDATE tblRejectedEnrollments SET tblRejectedEnrollments.Extracted = 'Y', tblRejectedEnrollments.extract_date = Now() WHERE (tblRejectedEnrollments.Extracted = 'N' Or tblRejectedEnrollments.Extracted = '');"
        End If
        

    Case 6  'mailing address
        If DCount("*", "qryMailingAddressExtract") <> 0 Then
            DoCmd.runSQL "Delete * From tblMailingAddressDisputeExtract"
            DoCmd.OpenQuery "qappMailingAddressExtract", acViewNormal
        
            DoCmd.TransferText acExportDelim, "MailingAddressExtract", "tblMailingAddressDisputeExtract", cuFilePath & "66666.P.txt", 1
            DoCmd.runSQL "UPDATE tblMailingAddressDispute SET tblMailingAddressDispute.Extracted = 'Y', tblMailingAddressDispute.Extract_Date = Now() WHERE (tblMailingAddressDispute.Extracted = 'N' Or tblMailingAddressDispute.Extracted = '');"
        End If
        

    Case 7   'agent/broker
        If DCount("*", "qryAgentBrokerExtract") <> 0 Then
            DoCmd.runSQL "Delete * From tblAgentBrokerExtract"
            DoCmd.OpenQuery "qappAgentBrokerExtract", acViewNormal
        
            DoCmd.TransferText acExportDelim, "AgentBrokerExtract", "tblAgentBrokerDisputesExtract", "X:\Member Enrollment\Member Enrollment(Custom)\MarketplaceDatabaseProjects\RCNO DB\95216.COG.ERRS.D" & Format(Now(), "yymmdd") & ".T" & Format(Now(), "hhnn") & "77777.P.txt", 1
            DoCmd.runSQL "UPDATE tblAgentBrokerDisputes SET tblAgentBrokerDisputes.Extracted = 'Y', tblAgentBrokerDisputes.Extract_Date = Now() WHERE (((tblAgentBrokerDisputes.Extracted)='N' Or (tblAgentBrokerDisputes.Extracted)=''));"
        
        Else: MsgBox "No Records to Export"
        End If
        
    End Select
    
    lstTotals.Requery
End Sub


