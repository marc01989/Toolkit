Attribute VB_Name = "Module1"
Option Compare Database

Public Sub TestClose()
'list of errors
'https://msdn.microsoft.com/en-us/library/bb221208(v=office.12).aspx

    Dim strPath As String, comp As String
    Dim fs As Object
    Dim a As Object

    comp = Environ$("username")
    strPath = "X:\Member Enrollment\Member Enrollment(Custom)\Marketplace\Toolkits\Employee DB\Marco Caruso\Database"

    Set fs = CreateObject("Scripting.FileSystemObject")
        If fs.FileExists(strPath & "\test_close.txt") = True Then
            Set a = fs.Opentextfile(strPath & "\test_close.txt", 8)
        Else
            Set a = fs.createtextfile(strPath & "\test_close.txt")
        End If
    
        a.writeline Date + Time & comp
        a.Close
    Set fs = Nothing
End Sub


' Auto Mail Merge With VBA and Access (Late Binding)


Sub startMerge(iOpt)
    Dim oWord As Object
    Dim oWdoc As Object
    Dim wdInputName As String
    Dim wdOutputName As String
    Dim outFileName As String
    
On Error GoTo ErrorHandler
    
    ' Set Template Path
    '------------------------------------------------
    wdInputName = "X:\Member Enrollment\Member Enrollment(Custom)\Marketplace\HICS DB Project\Letters\Letter Template\Letter Template.docx"
    
    ' Create unique save filename with minutes and seconds to prevent overwrite
    '------------------------------------------------
    outFileName = "HICS Resolution Letters " & " " & Format(Now(), "yyyymmddhhnnss")
    
    ' Output File Path w/outFileName
    '------------------------------------------------
    wdOutputName = "X:\Member Enrollment\Member Enrollment(Custom)\Marketplace\HICS\Resolution Letters and Report\db\"
    
    
    Set oWord = CreateObject("Word.Application")
    Set oWdoc = oWord.Documents.Open(wdInputName)
    
    ' Start mail merge
    '------------------------------------------------
    With oWdoc.MailMerge
        .MainDocumentType = 0 'wdFormLetters
        .OpenDataSource _
            Name:=CurrentProject.FullName, _
            AddToRecentFiles:=False, _
            LinkToSource:=True, _
            SQLStatement:="SELECT * FROM [tblMaileMerge] WHERE [appeal_form] = 0;"
        .Destination = 0 'wdSendToNewDocument
        .Execute Pause:=False
    End With
    
    ' Hide Word During Merge
    '------------------------------------------------
    oWord.Visible = False
    
Select Case iOpt

    Case 1:
    ' Save as PDF if startMerge(1)
    oWord.ActiveDocument.SaveAs2 wdOutputName & outFileName & ".pdf", 17
    
     Case 2:
    ' Save as WordDoc if startMerger(2)
    oWord.ActiveDocument.SaveAs2 wdOutputName & outFileName & ".docx", 16
    
End Select
    
              
    ' Quit Word to Save Memory
    '------------------------------------------------
    oWord.Quit savechanges:=False
       
    ' Clean up memory
    '------------------------------------------------
    Set oWord = Nothing
    Set oWdoc = Nothing
    
Exit Sub

ErrorHandler:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Module1; startMerge()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    oWord.Quit savechanges:=False
    Set oWord = Nothing
    Set oWdoc = Nothing
    
    
End Sub
' Auto Mail Merge With VBA and Access (Late Binding)


Sub startMergeAppealForm(iOpt)
    Dim oWord As Object
    Dim oWdoc As Object
    Dim wdInputName As String
    Dim wdOutputName As String
    Dim outFileName As String
    
On Error GoTo ErrorHandler
    
    ' Set Template Path
    '------------------------------------------------
    wdInputName = "X:\Member Enrollment\Member Enrollment(Custom)\Marketplace\HICS DB Project\Letters\Letter Template\Letter Template with Appeal Form.docx"
    
    ' Create unique save filename with minutes and seconds to prevent overwrite
    '------------------------------------------------
    outFileName = "HICS Appeal Letters " & " " & Format(Now(), "yyyymmddhhnnss")
    
    ' Output File Path w/outFileName
    '------------------------------------------------
    wdOutputName = "X:\Member Enrollment\Member Enrollment(Custom)\Marketplace\HICS\Resolution Letters and Report\db\"
    
    
    Set oWord = CreateObject("Word.Application")
    Set oWdoc = oWord.Documents.Open(wdInputName)
    
    ' Start mail merge
    '------------------------------------------------
    With oWdoc.MailMerge
        .MainDocumentType = 0 'wdFormLetters
        .OpenDataSource _
            Name:=CurrentProject.FullName, _
            AddToRecentFiles:=False, _
            LinkToSource:=True, _
            SQLStatement:="SELECT * FROM [tblMaileMerge] WHERE appeal_form = -1;"
        .Destination = 0 'wdSendToNewDocument
        .Execute Pause:=False
    End With
    
    ' Hide Word During Merge
    '------------------------------------------------
    oWord.Visible = False
    
Select Case iOpt

    Case 1:
    ' Save as PDF if startMerge(1)
    oWord.ActiveDocument.SaveAs2 wdOutputName & outFileName & ".pdf", 17
    
     Case 2:
    ' Save as WordDoc if startMerger(2)
    oWord.ActiveDocument.SaveAs2 wdOutputName & outFileName & ".docx", 16
    
End Select
    
              
    ' Quit Word to Save Memory
    '------------------------------------------------
    oWord.Quit savechanges:=False
       
    ' Clean up memory
    '------------------------------------------------
    Set oWord = Nothing
    Set oWdoc = Nothing
    
Exit Sub

ErrorHandler:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Module1; startMergeAppealForm()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    oWord.Quit savechanges:=False
    Set oWord = Nothing
    Set oWdoc = Nothing
    
    
End Sub
Function CleanPhoneNumber(txtClean As String)
        txtClean = Replace(txtClean, vbLf, "")
        txtClean = Replace(txtClean, vbTab, "")
        txtClean = Replace(txtClean, vbCr, "")
        txtClean = Replace(txtClean, vbCrLf, "")
        txtClean = Replace(txtClean, vbNewLine, "")
        txtClean = Replace(txtClean, ";", ":")
        txtClean = Replace(txtClean, "<", "(")
        txtClean = Replace(txtClean, ">", ")")
        txtClean = Replace(txtClean, Chr(160), "")
        txtClean = Replace(txtClean, "'", " ")
        txtClean = Replace(txtClean, Chr(146), "")
        txtClean = Replace(txtClean, Chr(39), "")
        txtClean = Replace(txtClean, "(", "")
        txtClean = Replace(txtClean, ")", "")
        txtClean = Replace(txtClean, "-", "")
        txtClean = Replace(txtClean, "|", "  ")
         txtClean = Replace(txtClean, " ", "")
        Trim (txtClean)
        CleanPhoneNumber = txtClean
End Function

Public Sub LogError(strError, modName As String)
'list of errors
'https://msdn.microsoft.com/en-us/library/bb221208(v=office.12).aspx

    Dim strPath As String, comp As String
    Dim fs As Object
    Dim a As Object

    comp = Environ$("username")
    strPath = "X:\Member Enrollment\Member Enrollment(Custom)\Marketplace\Database\db utilities\errors"

    Set fs = CreateObject("Scripting.FileSystemObject")
        If fs.FileExists(strPath & "\errorLogToolkit.txt") = True Then
            Set a = fs.Opentextfile(strPath & "\errorLogToolkit.txt", 8)
        Else
            Set a = fs.createtextfile(strPath & "\errorLogToolkit.txt")
        End If
    
        a.writeline Date + Time & "|ERROR: " & strError & "|USER: " & comp & "|MODULE: " & modName
        a.Close
    Set fs = Nothing
End Sub



Public Function cleanText(inputTxt As Variant) As Variant

If IsNull(inputTxt) Then
    inputTxt = ""
End If

inputTxt = Replace(inputTxt, "'", "''")
inputTxt = Replace(inputTxt, "<b>", "")
inputTxt = Replace(inputTxt, "<ul>", "")
inputTxt = Replace(inputTxt, "</ul>", vbCrLf & vbCrLf)
inputTxt = Replace(inputTxt, "</b>", "")

cleanText = inputTxt
End Function

Public Function CheckSubmission() As Boolean

Dim ctl As Control
Dim isBlank As Boolean

isBlank = False

For Each ctl In Forms.[QA Tracker]!Controls
    If Left(ctl.Name, 3) = "ext" Then
    MsgBox ("works")
End If
Next ctl

CheckSubmission = isBlank

End Function

Function testTable(tableName As String, week_id As Integer, employee_id As Integer) As Boolean
Dim recordCount As Integer

recordCount = DCount(employee_id, tableName, "([week_id] = " & week_id & " AND [employee_id] = " & employee_id & ")")

    If recordCount = 0 Then
        testTable = True
    Else
        testTable = False
    End If

End Function

Public Sub Forward(maxRecord As Integer, currentRecord As Integer, arRecords As Variant, txtBox As Control, txtSubmitDate As Control, txtSpecialist As Control)
      If currentRecord + 1 > maxRecord - 1 Then
        'MsgBox "No more Records"
        Exit Sub
        
    ElseIf currentRecord <= (maxRecord - 1) Then
        currentRecord = currentRecord + 1
        txtBox.Value = arRecords(0, currentRecord)
        txtSubmitDate.Value = arRecords(1, currentRecord)
        txtSpecialist.Value = arRecords(2, currentRecord)
    End If
End Sub

Public Sub Previous(maxRecord As Integer, currentRecord As Integer, arRecords As Variant, txtBox As Control, txtSubmitDate As Control, txtSpecialist As Control)
    If currentRecord - 1 < 0 Then
        'MsgBox "No more Record"
        Exit Sub
   ElseIf currentRecord - 1 >= 0 Then
        currentRecord = currentRecord - 1
        txtBox.Value = arRecords(0, currentRecord)
        txtSubmitDate.Value = arRecords(1, currentRecord)
        txtSpecialist.Value = arRecords(2, currentRecord)
    End If
End Sub
Public Function ScrubMemberId(memberId As String) As String
    memberId = Replace(memberId, "-", "")
    memberId = Replace(memberId, "_", "")
    memberId = Replace(memberId, " ", "")
    memberId = Trim(memberId)
    ScrubMemberId = memberId
End Function

Public Function checkQHPID(userInput As String) As Boolean

userInput = Trim(userInput)

test = Len(userInput)

If Len(userInput) <> 16 Then
    checkQHPID = False
ElseIf (Left(userInput, 7) <> "16322PA" And Left(userInput, 7) <> "62560PA") Then
    checkQHPID = False
ElseIf IsNull(userInput) Then
    checkQHPID = False
Else
    checkQHPID = True
End If

End Function

Public Function getQHPID(yearType As Integer, userInput As String) As Boolean

Dim lookupResult As Integer

Select Case yearType
    Case 1
        lookupResult = DCount("[2017_qhpid]", "[2017_qhpid_tbl]", "[2017_qhpid] = '" & userInput & "'")
    Case 2
        lookupResult = DCount("[2018_qhpid]", "[2018_qhpid_tbl]", "[2018_qhpid] = '" & userInput & "' AND [is_crosswalked] = 1")
    Case 3
        lookupResult = DCount("[2019_qhpid]", "[2019_qhpid_tbl]", "[2019_qhpid] = '" & userInput & "' AND [is_crosswalked] = 1")
End Select

If lookupResult > 0 Then
    getQHPID = True
Else
    getQHPID = False
End If

End Function

Public Function getQHPIDSingle(yearType As Integer, userInput As String) As Boolean

Dim lookupResult As Integer

Select Case yearType
    Case 1
        lookupResult = DCount("[2017_qhpid]", "[2017_qhpid_tbl]", "[2017_qhpid] = '" & userInput & "'")
    Case 2
        lookupResult = DCount("[2018_qhpid]", "[2018_qhpid_tbl]", "[2018_qhpid] = '" & userInput & "'")
    Case 3
        lookupResult = DCount("[2019_qhpid]", "[2019_qhpid_tbl]", "[2019_qhpid] = '" & userInput & "'")
End Select

If lookupResult > 0 Then
    getQHPIDSingle = True
Else
    getQHPIDSingle = False
End If

End Function

Public Function getGroupSubgroup(yearType As Integer, userGroup As String, userSubgroup As String) As Boolean

Dim lookupResult As Integer

Select Case yearType
    Case 1
        lookupResult = DCount("[2017_qhpid]", "[2017_qhpid_tbl]", "[2017_group] = '" & userGroup & "' AND [2017_subgroup] = '" & userSubgroup & "'")
    Case 2
        lookupResult = DCount("[2018_qhpid]", "[2018_qhpid_tbl]", "[2018_group] = '" & userGroup & "' AND [2018_subgroup] = '" & userSubgroup & "' AND [is_crosswalked] = 1")
    Case 3
        lookupResult = DCount("[2019_qhpid]", "[2019_qhpid_tbl]", "[2019_group] = '" & userGroup & "' AND [2019_subgroup] = '" & userSubgroup & "' AND [is_crosswalked] = 1")
End Select

If lookupResult > 0 Then
    getGroupSubgroup = True
Else
    getGroupSubgroup = False
End If

End Function

Public Function getGroupSubgroupSingle(yearType As Integer, userGroup As String, userSubgroup As String) As Boolean

Dim lookupResult As Integer

Select Case yearType
    Case 1
        lookupResult = DCount("[2017_qhpid]", "[2017_qhpid_tbl]", "[2017_group] = '" & userGroup & "' AND [2017_subgroup] = '" & userSubgroup & "'")
    Case 2
        lookupResult = DCount("[2018_qhpid]", "[2018_qhpid_tbl]", "[2018_group] = '" & userGroup & "' AND [2018_subgroup] = '" & userSubgroup & "'")
    Case 3
        lookupResult = DCount("[2019_qhpid]", "[2019_qhpid_tbl]", "[2019_group] = '" & userGroup & "' AND [2019_subgroup] = '" & userSubgroup & "'")
End Select

If lookupResult > 0 Then
    getGroupSubgroupSingle = True
Else
    getGroupSubgroupSingle = False
End If

End Function
Public Function RegexSwitch(strSubcategory As String) As String
'depending on type of hics case, return the regex string needed to reformat the narrative
Dim strRegex As String

    Select Case strSubcategory
        Case "Alternative Format Request Issue"
            strRegex = "(?:ALTERNATIVE FORMAT ISSUE REQUEST|The type of materials the consumer needs\:|Exchange Assigned Policy)"
        Case "Auto Re-Enrollment or Renewal"
            strRegex = "(?:Consumer attests that misrepresentation or misinformation has occurred, which resulted in the consumer not enrolling\.|Please verify and determine SEP eligibility\.|Misleading Information consumer received\:|The individual who gave wrong information\:|Exchange Assigned Policy ID)"
        Case "Cancellation/Termination Request"
            strRegex = "(?:Consumer attests that misrepresentation or misinformation has occurred, which resulted in the consumer not enrolling\.|Consumer is requesting a retroactive termination date due to overlapping coverage with a different insurance company\.|Please verify and determine SEP eligibility\.|Type of overlapping insurance\:|Actions taken\:|The consumer's application ID is|Information of consumer\(s\) to be made nonapplicants\:|Reason consumer has overlapping coverage\:|" & _
        "Details regarding what directions the consumer provided\:|The individual who failed to follow the request\:|Is the consumer eligible for APTC\/CSR\:|Exchange Assigned Policy ID|Date to set the termination\:|Consumer requested to terminate coverage\.|Consumer was removed from Marketplace coverage [0-9]{10} and other members remain in coverage on the application\.|Last App Updated By\:|Original App Source\:|" & _
        "Consumer requested to terminate coverage\. The Marketplace is unable to process this request due to system error 500\.280\. Please terminate the consumer's coverage\.|Consumer\(s\) can be reached at|The consumer's intended termination date is|but functionality limitations provided a termination date of|Issuer Action\: Consumer requested cancelation or termination more than 30 days ago but the plan has no record of the request\. Please investigate and provide the consumer with confirmation of termination, if applicable\.)"
        Case "Consumer Believes APTC not Awarded Properly"
            strRegex = "(?:Consumer attests that misrepresentation or misinformation has occurred, which resulted in the consumer not enrolling\. Please verify and determine SEP eligibility\.|Consumer attests that misrepresentation or misinformation has occurred, which resulted in the consumer not receiving financial help\. Please verify and determine SEP eligibility\.|Misleading information provided\:|Details regarding what directions the consumer provided\:|The individual who failed to follow the request\:|Is the consumer eligible for APTC\/CSR\:|" & _
        "Exchange Assigned Policy ID|Last App Updated By\:|Original App Source\:|Issuer Action\: Consumer believes the plan is billing them for the balance of the premium that the Marketplace should pay the plan directly, per the consumer's APTC determination\. Please investigate and contact CMS regarding any errors\.|The consumer's application ID)"
        Case "Cost-Sharing"
            strRegex = "(?:Issuer Action\: Consumer believes the cost-sharing amount for this plan is different from what they saw on the plan compare website\. Please investigate and contact CMS regarding any errors\.|The consumer's application ID is|Consumer's current cost-sharing amount is\:|Consumer saw a cost-sharing amount in Plan Compare of\:|Exchange Assigned Policy ID|Last App Updated By\:|Original App Source\:)"
        Case "Eligibility Appeals Related (OHI Use Only)"
            strRegex = "(?:^.*granted retroactivity for plan enrollment and\/or APTC\/CSR amount[^.]+|The Issuer is to update its internal systems to reflect changes to enrollment and to any advance premium tax credits or CSRs, as applicable\.|Monthly amount of APTC and effective date\:|Retroactive enrollment date\:|Cost Sharing Reductions and effective date\:)"
        Case "Issuer Customer Service"
            strRegex = "(?:Issuer Action\: Marketplace records indicate consumer's coverage was canceled but the plan has reinstated coverage, according to the consumer\. If the issuer agrees the policy is current, issuer needs to send reinstatement \(or coverage end date dispute\) as an enrollment dispute to the ER R contractor so the Marketplace application can be updated\. Enrollment updates for enrollments pending reinstatement may come via HICS case\.|Application ID\:|Wrong information on application\:|Correct enrollment information according to consumer\:|" & _
        "Consumer's termination date\:|Consumer's monthly premium amount\:|Consumer confirmed with plan they are still enrolled\:|Exchange Assigned Policy ID|Issuer Action: The plan's consumer record does not match the consumer's Marketplace application\. Reconcile the consumer's plan record with the Marketplace application\.|Wrong information on record\:|Correction Information as listed on Application\:|The consumer's application ID is|Last App Updated By\:|Original App Source\:)"
        Case "Issuer Enrollment/Disenrollment"
            strRegex = "(?:Consumer attests that misrepresentation or misinformation has occurred, which resulted in the consumer not enrolling\. Please verify and determine SEP eligibility\.|Misleading Information consumer received\:|The individual who gave wrong information\:|Exchange Assigned Policy ID|" & _
        "Is the consumer eligible for APTC\/CSR\:|Last App Updated By\:|Original App Source\:|CSR accidentally terminated consumer's coverage when canceling coverage for another enrollee on the application\. Please consider a special enrollment period to reinstate coverage for|Name of consumer\(s\) who lost coverage\:|" & _
        "Issuer Action\: Consumer unable to submit application due to change in circumstance enrollment confirmation blocker\. Please make the updates in your system and submit changes to the Marketplace through the ER R process\.|Reason for change in circumstance\:|Name and date of birth of person being removed from coverage\:|" & _
        "Effective date the person should be removed from coverage\:|New APTC amount\:|Cost\-sharing reduction variant\:|Additional information\:|Issuer Action\: Consumer was unable to submit application due to error 302100, 500\.300588\.|Reason for change in circumstance\:|Name of new enrollee\:|" & _
        "Date of birth\/adoption\/marriage\:|Coverage effective date\:|Process coverage termination retroactive to the consumer's date of death\.|Deceased Consumer Action\:|SEP \- Consumer is eligible for a retroactive start date\.|Information about consumer's issue\:|The consumer's requested start date is)"
        Case "Premium Payment"
            strRegex = "(?:Consumer believes they are being billed for a premium that differs from what was listed on plan compare\. Please investigate and contact CMS regarding any errors\.|The consumer's application ID|The consumer's current premium amount is|The consumer believes the correct premium amount should be|Exchange Assigned Policy ID|Issuer Action\: Consumer believes the plan double\-billed them for their premium because of a duplicate enrollment\. Please investigate and determine if consumer has duplicate enrollment\. If applicable, remove the duplicate enrollment and contact CMS regarding any errors\.|" & _
        "Issuer Action\: Consumer continues to be billed by their old plan when they are enrolled in a new plan\. Please investigate and reconcile the consumer's account\.|Last App Updated By\:|Original App Source\:)"
        Case "Reinstatement/Re-enrollment Request"
            strRegex = "(?:Consumer attests that misrepresentation or misinformation has occurred, which resulted in the consumer not enrolling\.  Please verify and determine SEP eligibility\.|Misleading Information consumer received\:|The individual who gave wrong information\:|Is the consumer eligible for APTC\/CSR\:|Exchange Assigned Policy ID|Last App Updated By\:|Original App Source\:|CSR accidentally terminated consumer's coverage when canceling coverage for another enrollee on the application\. Please consider a special enrollment period to reinstate coverage for|" & _
        "Name of consumer\(s\) who lost coverage\:|Exchange Assigned Policy ID|Issuer Action\: Review case for reinstatement\. The consumer is requesting reinstatement because the consumer feels they were incorrectly terminated\/canceled by their plan\. Per the QHP issuer call, the issuer needs to determine if the consumer was canceled\/terminated in error and if it is determined the consumer was canceled\/terminated in error, the issuer should reinstate the consumer to the original effective date [^.]+\.|The consumer's original application ID is|Reason why the plan ended coverage\:|Reason why consumer believes this is wrong\:)"
        Case "Special Enrollment Period (Issuer Action Required)"
            strRegex = "(?:Consumer attests that misrepresentation or misinformation has occurred, which resulted in the consumer not enrolling\. Please verify and determine SEP eligibility\.|Details regarding what directions the consumer provided\:|The individual who failed to follow the request\:|Is the consumer eligible for APTC\/CSR\:|Exchange Assigned Policy ID|Last App Updated By\:|Original App Source\:|SEP \- Consumer is a [0-9]{4} unaffiliated issuer enrollment \(issuer orphan\) or [0-9]{4} missing enrollment and is eligible for a retroactive [^.]+\.|The consumer's application ID is)"
        Case Else: strRegex = "(?:Does this match)"
    End Select

RegexSwitch = strRegex
End Function

Public Function RegexNarrative(strNarrative As String, strRegex As String) As String
'REFERENCE REQUIRED - Microsoft VBScript Regular Expression 5.5
'https://regex101.com/

    Dim regex As RegExp
    Dim resultStr As String
    Dim intPos As Integer
    
    'set pattern - what you are/aren't looking for
    'to find string where this expression is not present:  ^((?! PHRASE HERE  ).)*$
    Set regex = New RegExp
    With regex
        .MultiLine = False
        .Global = True
        .IgnoreCase = True
        .Pattern = strRegex
    End With
    
    If regex.test(strNarrative) = True Then
        tempStr = strNarrative
        tempStr = regex.Replace(tempStr, "<br><br><b>$&</b><br>")
        
        If Len(tempStr) > 0 Then
            intPos = InStr(tempStr, "<br>")
            If intPos = 1 Then
                tempStr = Right(tempStr, Len(tempStr) - 8)
            End If
        End If
        RegexNarrative = tempStr
    Else
        RegexNarrative = strNarrative
    End If
    
End Function



