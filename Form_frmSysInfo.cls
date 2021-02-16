VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSysInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


'Created by:             Michael Mahoney: TPA Operations - mahoneym2@upmc.edu
'                        Marco Caruso: Marketplace Enrollment - carusomj@upmc.edu
'
'Modified Date:          6/28/2019
'Version Go Live:        7/19/2019
'Purpose:                To log, quantify and display data for HICS and RECON processes
'
'
'
'                        Current Version = 2.5
'
'                        Change Log
'                        1.0: Stable version:
'                            Implemented mainly by Michael Mahoney in Dec 2018
'
'                        2.0: HICS:
'                            New interface for frmHICSUpdate; removed frmHICSResolution; removed/updated several queries; updated frmAdmin & frmQA
'                            See documentation for detailed changes: X:\Member Enrollment\Member Enrollment(Custom)\Marketplace\Toolkits\Employee DB\Marco Caruso\HICs\System Design
'
'                        2.1: frmRCNOView
'                            Added validation for all dispute submissions (CheckSubmission())
'                            Enabled textboxs for editing for purpose of populating fields necessary for certain dispute types
'                            Updated queries on LandingPage - originally the listbox.rowsource was set to a SQL statement - this was opening the 'HICS 2.0 BE' and preventing me from updating tables while
'                                specialists had the toolkit open. replaced them w/ recordset snapshot queries to the records are released after execution
'                            Updated frmHICSUpload to prompt user to mark uploads as complete for each upload type before closing the form
'
'                        2.2: frmHICSUpload
'                            Added [extracted] Yes/No field to all notes tables (tblCaseworkNotes, tblResolutionNotes, tblERRNotes, tblCutlogNotes).
'                                When doing morning note upload, specialist extracts note reports, changing [extracted] from N to Y. After upload each extract into respective system,
'                                specialist must mark each extract as uploaded in frmHICSUpload - this insures upload timestamp in toolkit is close to actual upload time into CMS and MC400
'                            "qryFraudCases" added to Morning reports in frmAdmin. Pulls cases denoted by caseworker name "Ellen Witman"
'                            Added unauthorized enrollments (eu), AKA fraud cases to tblHICSCaseType, tblHICSCaseAction and tblHICSLetterResponses. made minor updates to frmHICSUpdate to accomodate this
'                            Removed frmHICSEditNotes, replacing with frmHICSRecordEdit
'
'                        2.3 Merged tblCAseWorkNotes and tblResolutionNotes
'                            Merged records from tables into one, using [is_res] to distinguish btw regular note and resolution note.
'                            Updated code in frmHICSUpdate, frmHICSQA and frmHICSNotesUpdate, as well as all queries to remove dependency from tblResolutionNotes
'                                See backups with (v2.2) for how backend was set up before this update
'                            Changed qryHICSEmailPrecursor --> qryHICSEmailPrecursor_daily. Updated version pulls all non-response members regardless of # of outreaches
'
'                        2.4 Added [casework_note_id] fields to following tables: tblInternalNotes, tblCutlogNotes and tblERRNotes
'                                Took the ID from each record on tblCaseWorkNotes and inserted into above tables, by joining records on submit_date
'                                This field ties the main casework note to each optional note stored on other tables. Makes iterating through notes cleaner
'
'                        2.5 Appeal Letter functionality added
'                            Appeal form optBox added to frmHICUpdate, also added new letter template and updated mailemerge query to pull in appeal form based on added field.






Private Sub Label0_Click()
        CreateObject("Shell.application").Open "https://www.youtube.com/watch?v=MZmTxcmLOCg"


End Sub
